import { readdir, stat } from "node:fs/promises";
import { basename, extname, join, relative, resolve } from "node:path";
import { LOGICAL_RELATIONS, SLIDE_ROLES } from "../../domain/types.js";
import type { LogicalRelation, SlideRole } from "../../domain/types.js";
import { createSafeDeckId, isValidDeckId } from "../../knowledge/deck-id.js";
import { ingestDeckFromPptx } from "../../knowledge/ingest/ingest.js";
import { JsonKnowledgeStore } from "../../knowledge/store/json-store.js";
import type { KnowledgeQuery } from "../../knowledge/retrieval/retrieval.js";
import {
  boolFlag,
  numberFlag,
  parseArgs,
  stringFlag,
  CliArgumentError,
} from "../args.js";

export const KNOWLEDGE_USAGE = `
semantic-deck knowledge <ingest|search|stats> [options]

  ingest <path...>      過去 PPTX（ファイルまたはディレクトリ）を Knowledge へ取り込む
    --store <dir>       保存先 (既定: ./knowledge)
    --deck-id <id>      deck id（既定: ファイル名。複数ファイル時はファイル名ごと）
    --organization <s>  組織名
    --department <s>    部署名
    --purpose <s>       資料の目的（例: 投資判断）
    --audience <s>      想定読者（例: 経営会議）
    --created-at <s>    作成日
    --imported-at <s>   記録する取り込み日時

  search                過去スライドを検索する（AI 不要の決定論的検索）
    --store <dir>       検索対象 (既定: ./knowledge)
    --text <s>          自由文
    --role <role>       ${SLIDE_ROLES.join(" | ")}
    --relation <rel>    ${LOGICAL_RELATIONS.join(" | ")}
    --purpose <s>       目的で絞る
    --audience <s>      読者で絞る
    --pattern <id>      過去に使われた Pattern
    --limit <n>         件数 (既定 5)
    --strict            メタデータ条件を必須にする
    --json              JSON で出力する

  stats                 蓄積状況を表示する
    --store <dir>
`.trim();

async function collectPptxFiles(target: string): Promise<string[]> {
  const absolute = resolve(target);
  const info = await stat(absolute).catch(() => undefined);
  if (!info) throw new CliArgumentError(`パスが見つかりません: ${target}`);
  if (info.isFile()) return [absolute];

  const entries = await readdir(absolute, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && extname(entry.name).toLowerCase() === ".pptx")
    .map((entry) => join(absolute, entry.name))
    .sort();
}

/**
 * 取り込むファイルの deckId を決める。
 *
 * 日本語のファイル名（例: 経営会議_生成AI戦略.pptx）でも必ず有効な ID になるよう、
 * 自動生成は `createSafeDeckId` に一本化する。--deck-id の明示指定はそのまま使い、
 * 不正な場合は黙って書き換えず呼び出し側でエラーにする。
 */
function deckIdFor(file: string, explicit: string | undefined, multiple: boolean): string {
  if (explicit && !multiple) return explicit;
  const stem = basename(file, extname(file));
  return createSafeDeckId(explicit ? `${explicit}-${stem}` : stem);
}

export async function runKnowledge(argv: readonly string[]): Promise<number> {
  const sub = argv[0];
  const rest = argv.slice(1);

  if (!sub || sub === "--help" || sub === "-h") {
    console.log(KNOWLEDGE_USAGE);
    return sub ? 0 : 1;
  }

  if (sub === "ingest") {
    const args = parseArgs(rest, {
      boolean: ["help", "quiet"],
      string: [
        "store",
        "deck-id",
        "organization",
        "department",
        "purpose",
        "audience",
        "created-at",
        "imported-at",
      ],
    });
    if (boolFlag(args, "help") || args.positional.length === 0) {
      console.log(KNOWLEDGE_USAGE);
      return args.positional.length === 0 ? 1 : 0;
    }

    const explicitDeckId = stringFlag(args, "deck-id");
    if (explicitDeckId && !isValidDeckId(explicitDeckId)) {
      throw new CliArgumentError(
        `--deck-id が不正です: ${explicitDeckId}\n` +
          `  英数字と - _ . のみ使用でき、先頭は英数字である必要があります` +
          `（例: ${createSafeDeckId(explicitDeckId)}）`,
      );
    }

    const store = new JsonKnowledgeStore(resolve(stringFlag(args, "store") ?? "knowledge"));
    const files = (await Promise.all(args.positional.map(collectPptxFiles))).flat();
    if (files.length === 0) {
      console.error("取り込める .pptx が見つかりませんでした");
      return 1;
    }

    let total = 0;
    for (const file of files) {
      const result = await ingestDeckFromPptx(file, {
        deckId: deckIdFor(file, explicitDeckId, files.length > 1),
        metadata: {
          ...(stringFlag(args, "organization")
            ? { organization: stringFlag(args, "organization")! }
            : {}),
          ...(stringFlag(args, "department") ? { department: stringFlag(args, "department")! } : {}),
          ...(stringFlag(args, "purpose") ? { purpose: stringFlag(args, "purpose")! } : {}),
          ...(stringFlag(args, "audience") ? { audience: stringFlag(args, "audience")! } : {}),
          ...(stringFlag(args, "created-at") ? { createdAt: stringFlag(args, "created-at")! } : {}),
        },
        ...(stringFlag(args, "imported-at") ? { importedAt: stringFlag(args, "imported-at")! } : {}),
        sourceLabel: relative(process.cwd(), file) || file,
      });
      await store.add(result.entries);
      total += result.entries.length;

      console.log(`取り込み: ${relative(process.cwd(), file)} → ${result.entries.length} スライド`);
      for (const warning of result.report.warnings) console.log(`  警告: ${warning}`);
      for (const item of result.report.unsupportedFeatures) console.log(`  非対応: ${item}`);
    }

    const stats = await store.stats();
    console.log("");
    console.log(`保存先: ${resolve(stringFlag(args, "store") ?? "knowledge")}`);
    console.log(`蓄積状況: ${stats.decks} デッキ / ${stats.slides} スライド（今回 ${total} 件）`);
    return 0;
  }

  if (sub === "search") {
    const args = parseArgs(rest, {
      boolean: ["help", "json", "strict"],
      string: ["store", "text", "role", "relation", "purpose", "audience", "pattern", "limit"],
    });
    if (boolFlag(args, "help")) {
      console.log(KNOWLEDGE_USAGE);
      return 0;
    }

    const role = stringFlag(args, "role");
    if (role && !(SLIDE_ROLES as readonly string[]).includes(role)) {
      throw new CliArgumentError(`--role が不正です: ${role}`);
    }
    const relation = stringFlag(args, "relation");
    if (relation && !(LOGICAL_RELATIONS as readonly string[]).includes(relation)) {
      throw new CliArgumentError(`--relation が不正です: ${relation}`);
    }

    const store = new JsonKnowledgeStore(resolve(stringFlag(args, "store") ?? "knowledge"));
    const query: KnowledgeQuery = {
      ...(stringFlag(args, "text") ? { text: stringFlag(args, "text")! } : {}),
      ...(role ? { slideRole: role as SlideRole } : {}),
      ...(relation ? { logicalRelation: relation as LogicalRelation } : {}),
      ...(stringFlag(args, "purpose") ? { purpose: stringFlag(args, "purpose")! } : {}),
      ...(stringFlag(args, "audience") ? { audience: stringFlag(args, "audience")! } : {}),
      ...(stringFlag(args, "pattern") ? { patternId: stringFlag(args, "pattern")! } : {}),
      ...(numberFlag(args, "limit") !== undefined ? { limit: numberFlag(args, "limit")! } : {}),
      ...(boolFlag(args, "strict") ? { strict: true } : {}),
    };

    const results = await store.search(query);
    if (boolFlag(args, "json")) {
      console.log(JSON.stringify(results, null, 2));
      return 0;
    }
    if (results.length === 0) {
      console.log("一致する過去スライドはありませんでした");
      return 0;
    }
    for (const result of results) {
      const entry = result.entry;
      console.log(`${entry.slideId}  score=${result.score.toFixed(3)}`);
      console.log(`  ${entry.semantic?.keyMessage ?? "(keyMessage なし)"}`);
      console.log(
        `  role=${entry.semantic?.slideRole ?? "-"} relation=${entry.semantic?.logicalRelation ?? "-"} pattern=${entry.visual?.patternId ?? "-"}`,
      );
      console.log(`  一致: ${result.matched.join(", ")}`);
    }
    return 0;
  }

  if (sub === "stats") {
    const args = parseArgs(rest, { boolean: ["help"], string: ["store"] });
    if (boolFlag(args, "help")) {
      console.log(KNOWLEDGE_USAGE);
      return 0;
    }
    const dir = resolve(stringFlag(args, "store") ?? "knowledge");
    const store = new JsonKnowledgeStore(dir);
    const stats = await store.stats();
    console.log(`${dir}`);
    console.log(`  ${stats.decks} デッキ / ${stats.slides} スライド`);
    return 0;
  }

  throw new CliArgumentError(`未知のサブコマンドです: knowledge ${sub}`);
}
