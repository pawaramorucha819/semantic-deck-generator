import { relative, resolve } from "node:path";
import {
  capturePatternFromPptx,
  writePatternPackage,
  type CaptureOverrides,
} from "../../patterns/capture/capture.js";
import { PatternRegistry } from "../../patterns/registry.js";
import { getTheme } from "../../theme/theme.js";
import {
  boolFlag,
  listFlag,
  numberFlag,
  parseArgs,
  requireString,
  stringFlag,
  CliArgumentError,
} from "../args.js";

export const PATTERN_USAGE = `
semantic-deck pattern <import|list> [options]

  import <file.pptx>    PowerPoint の 1 スライドから Pattern 下書きを作る
    --slide <n>         取り込むスライド番号（必須）
    --id <id>           作成する pattern id（必須）
    --out-dir <dir>     出力先ルート (既定: ./patterns/company)
    --scope <scope>     company | team | project | user (既定: company)
    --theme <id>        色・サイズを寄せる先のテーマ (既定: default)
    --organization <s>  組織名
    --description <s>   Pattern の説明
    --headline-index <n>  見出しにする要素インデックスを明示指定
    --subline-index <n>   補足メッセージにする要素インデックス
    --eyebrow-index <n>   役割ラベルにする要素インデックス
    --source-index <n>    出典行にする要素インデックス
    --static <n>          固定デザイン扱いにする要素インデックス (複数可)
    --imported-at <s>     記録する取り込み日時

    出力される pattern.json は必ず status=draft です。
    人間がレビューして metadata.status を "approved" に変えるまで登録できません。

  list                  登録済み Pattern を一覧する
    --pattern-dir <d>   外部 Pattern ディレクトリ (複数可)
    --include-drafts    draft も一覧に含める
    --json              JSON で出力する
`.trim();

export async function runPattern(argv: readonly string[]): Promise<number> {
  const sub = argv[0];
  const rest = argv.slice(1);

  if (!sub || sub === "--help" || sub === "-h") {
    console.log(PATTERN_USAGE);
    return sub ? 0 : 1;
  }

  if (sub === "import") {
    const args = parseArgs(rest, {
      boolean: ["help"],
      string: [
        "slide",
        "id",
        "out-dir",
        "scope",
        "theme",
        "organization",
        "description",
        "headline-index",
        "subline-index",
        "eyebrow-index",
        "source-index",
        "imported-at",
      ],
      repeatable: ["static"],
    });
    if (boolFlag(args, "help") || args.positional.length === 0) {
      console.log(PATTERN_USAGE);
      return args.positional.length === 0 ? 1 : 0;
    }

    const input = resolve(args.positional[0]!);
    const id = requireString(args, "id");
    const slideNumber = numberFlag(args, "slide");
    if (slideNumber === undefined) throw new CliArgumentError("--slide は必須です");

    const scope = stringFlag(args, "scope") ?? "company";
    if (!["company", "team", "project", "user"].includes(scope)) {
      throw new CliArgumentError(`--scope には company | team | project | user を指定してください`);
    }

    const overrides: CaptureOverrides = {
      ...(numberFlag(args, "headline-index") !== undefined
        ? { headlineIndex: numberFlag(args, "headline-index")! }
        : {}),
      ...(numberFlag(args, "subline-index") !== undefined
        ? { sublineIndex: numberFlag(args, "subline-index")! }
        : {}),
      ...(numberFlag(args, "eyebrow-index") !== undefined
        ? { eyebrowIndex: numberFlag(args, "eyebrow-index")! }
        : {}),
      ...(numberFlag(args, "source-index") !== undefined
        ? { sourceIndex: numberFlag(args, "source-index")! }
        : {}),
      ...(listFlag(args, "static").length > 0
        ? {
            staticIndexes: listFlag(args, "static")
              .flatMap((value) => value.split(","))
              .map((value) => Number(value.trim()))
              .filter((value) => Number.isInteger(value) && value >= 0),
          }
        : {}),
    };

    const captured = await capturePatternFromPptx(input, {
      id,
      slideNumber,
      scope: scope as "company" | "team" | "project" | "user",
      theme: getTheme(stringFlag(args, "theme")),
      ...(stringFlag(args, "organization")
        ? { organization: stringFlag(args, "organization")! }
        : {}),
      ...(stringFlag(args, "description") ? { description: stringFlag(args, "description")! } : {}),
      ...(stringFlag(args, "imported-at") ? { importedAt: stringFlag(args, "imported-at")! } : {}),
      ...(Object.keys(overrides).length > 0 ? { overrides } : {}),
      sourceLabel: relative(process.cwd(), input) || input,
    });

    const outDir = resolve(stringFlag(args, "out-dir") ?? "patterns/company");
    const written = await writePatternPackage(outDir, captured);

    console.log(`取り込み: ${relative(process.cwd(), input) || input} (スライド ${slideNumber})`);
    console.log(
      `  要素: text=${captured.report.elements.text} shape=${captured.report.elements.shape} image=${captured.report.elements.image}`,
    );
    console.log(`  推定: ${captured.report.inference.inferencer}`);
    for (const note of captured.report.inference.notes) console.log(`    - ${note}`);
    console.log(`  slot: ${captured.document.slots.map((s) => s.id).join(", ")}`);

    if (captured.report.unsupportedFeatures.length > 0) {
      console.log("");
      console.log("非対応（取り込んでいない要素）:");
      for (const item of captured.report.unsupportedFeatures) console.log(`  - ${item}`);
    }
    if (captured.report.warnings.length > 0) {
      console.log("");
      console.log("警告:");
      for (const item of captured.report.warnings) console.log(`  - ${item}`);
    }

    console.log("");
    console.log(`生成: ${relative(process.cwd(), written.patternPath)} (status=draft)`);
    if (written.reportPath) console.log(`生成: ${relative(process.cwd(), written.reportPath)}`);
    for (const path of written.assetPaths) console.log(`生成: ${relative(process.cwd(), path)}`);

    console.log("");
    console.log("次にやること（人間のレビュー）:");
    for (const item of captured.report.reviewChecklist) console.log(`  - ${item}`);
    return 0;
  }

  if (sub === "list") {
    const args = parseArgs(rest, {
      boolean: ["help", "json", "include-drafts"],
      repeatable: ["pattern-dir"],
    });
    if (boolFlag(args, "help")) {
      console.log(PATTERN_USAGE);
      return 0;
    }

    const registry = new PatternRegistry();
    for (const dir of listFlag(args, "pattern-dir")) {
      await registry.registerDirectory(dir, { allowDraft: boolFlag(args, "include-drafts") });
    }
    const list = registry.list();

    if (boolFlag(args, "json")) {
      console.log(JSON.stringify(list, null, 2));
      return 0;
    }
    for (const item of list) {
      console.log(
        `${item.id.padEnd(24)} ${item.kind.padEnd(11)} ${item.scope.padEnd(8)} v${item.version}  ${item.description}`,
      );
      if (item.source) console.log(`  ${item.source}`);
      if (item.previewImagePath) console.log(`  preview: ${item.previewImagePath}`);
    }
    return 0;
  }

  throw new CliArgumentError(`未知のサブコマンドです: pattern ${sub}`);
}
