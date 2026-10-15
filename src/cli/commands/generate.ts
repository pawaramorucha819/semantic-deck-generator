import { relative } from "node:path";
import {
  buildGenerationContext,
  loadGenerationConfig,
  type BuildContextOptions,
  type GenerationContext,
} from "../../config/context.js";
import { GENERATION_CONFIG_SCHEMA_VERSION } from "../../config/schema.js";
import { createSafeDeckId } from "../../knowledge/deck-id.js";
import { ingestDeckFromGeneration } from "../../knowledge/ingest/ingest.js";
import { generateFromFile, type GenerateFilesResult } from "../../pipeline.js";
import type { LayoutIssue } from "../../render/validate.js";
import { builtinThemes } from "../../theme/theme.js";
import {
  boolFlag,
  listFlag,
  parseArgs,
  stringFlag,
  type ParsedArgs,
} from "../args.js";
import { CliArgumentError } from "../args.js";

export const GENERATE_USAGE = `
semantic-deck generate <presentation.json> [options]

  presentation.json（意味・論理）から PowerPoint を生成します。
  生成方法の設定は generation.config.json 側に分離されています。

Options
  --config <file>       generation.config.json（Template / Pattern / Knowledge / AI をまとめて指定）
  --out-dir <dir>       出力先 (既定: 入力と同階層の generated/)
  --theme <id>          テーマ (${Object.keys(builtinThemes).join(" | ")})
  --template <id>       企業テンプレート
  --template-dir <dir>  外部テンプレートディレクトリ (複数可)
  --pattern-dir <dir>   外部 Pattern ディレクトリ (複数可)
  --knowledge-dir <dir> Knowledge ストア (複数可)
  --ai-mode <mode>      off | manual | assist | auto (既定: config の値、無ければ off)
  --section <slideId>   そのスライドを章扉 (section master) にする (複数可)
  --no-title-slide      表紙を付けない
  --no-candidates       Visual Plan に候補一覧を含めない
  --report              配置と意味構造の Review Report を書き出す
  --advisor-report      Advisor Report（候補・参照・AI 推薦の詳細）を書き出す
  --strict              レイアウト警告が 1 件でもあれば終了コード 1
  --quiet               ログを最小限にする
  -h, --help            このヘルプ

Examples
  semantic-deck generate examples/inventory-system.json
  semantic-deck generate deck.json --template corporate --template-dir ./templates
  semantic-deck generate deck.json --config generation.config.json --ai-mode assist
`.trim();

const GENERATE_FLAGS = {
  boolean: [
    "help",
    "no-title-slide",
    "no-candidates",
    "report",
    "advisor-report",
    "strict",
    "quiet",
  ],
  string: ["config", "out-dir", "theme", "template", "ai-mode"],
  repeatable: ["template-dir", "pattern-dir", "knowledge-dir", "section"],
} as const;

function formatIssue(issue: LayoutIssue): string {
  const mark = issue.severity === "error" ? "ERROR" : "WARN ";
  return `  [${mark}] p${issue.slideNumber} (${issue.patternId}) ${issue.code}: ${issue.message}`;
}

/** --config を使わない場合でも同じ経路を通すための、空の config */
function inlineConfigContext(): Parameters<typeof buildGenerationContext>[0] {
  return {
    config: { schemaVersion: GENERATION_CONFIG_SCHEMA_VERSION },
    baseDir: process.cwd(),
    path: "(inline)",
  };
}

function overridesFrom(args: ParsedArgs): BuildContextOptions {
  const aiMode = stringFlag(args, "ai-mode");
  if (aiMode && !["off", "manual", "assist", "auto"].includes(aiMode)) {
    throw new CliArgumentError(
      `--ai-mode には off | manual | assist | auto を指定してください: ${aiMode}`,
    );
  }
  return {
    ...(stringFlag(args, "template") ? { templateId: stringFlag(args, "template")! } : {}),
    ...(stringFlag(args, "theme") ? { themeId: stringFlag(args, "theme")! } : {}),
    ...(aiMode ? { aiMode: aiMode as BuildContextOptions["aiMode"] } : {}),
    ...(stringFlag(args, "out-dir") ? { outputDir: stringFlag(args, "out-dir")! } : {}),
    ...(listFlag(args, "template-dir").length > 0
      ? { templateDirs: listFlag(args, "template-dir") }
      : {}),
    ...(listFlag(args, "pattern-dir").length > 0
      ? { patternDirs: listFlag(args, "pattern-dir") }
      : {}),
    ...(listFlag(args, "knowledge-dir").length > 0
      ? { knowledgeDirs: listFlag(args, "knowledge-dir") }
      : {}),
    ...(boolFlag(args, "report") ? { writeRenderReport: true } : {}),
    ...(boolFlag(args, "advisor-report") ? { writeAdvisorReport: true } : {}),
    ...(boolFlag(args, "no-title-slide") ? { includeTitleSlide: false } : {}),
    ...(boolFlag(args, "no-candidates") ? { includeCandidates: false } : {}),
    ...(listFlag(args, "section").length > 0
      ? { sectionSlideIds: listFlag(args, "section") }
      : {}),
  };
}

/** 生成結果を Knowledge へ登録する（recordKnowledge 設定時） */
async function recordKnowledge(
  context: GenerationContext,
  result: GenerateFilesResult,
  quiet: boolean,
): Promise<void> {
  const record = context.config.recordKnowledge;
  if (!record) return;

  const store = context.knowledgeStores[record.storeIndex ?? 0];
  if (!store) return;

  // 資料タイトルは日本語であることが普通なので、必ず安全な ID へ変換する
  const deckId = record.deckId ?? createSafeDeckId(result.plan.title);
  const ingest = ingestDeckFromGeneration(result.spec, result.plan, {
    deckId,
    metadata: {
      ...(record.purpose ? { purpose: record.purpose } : {}),
      ...(record.audience ? { audience: record.audience } : {}),
      ...(record.organization ? { organization: record.organization } : {}),
    },
    sourceLabel: result.pptxPath,
  });
  await store.add(ingest.entries);
  if (!quiet) {
    console.log(`Knowledge 登録: ${deckId} (${ingest.entries.length} スライド)`);
  }
}

export async function runGenerate(argv: readonly string[]): Promise<number> {
  const args = parseArgs(argv, GENERATE_FLAGS);
  const quiet = boolFlag(args, "quiet");

  if (boolFlag(args, "help") || args.positional.length === 0) {
    console.log(GENERATE_USAGE);
    return args.positional.length === 0 ? 1 : 0;
  }

  const input = args.positional[0]!;
  const configPath = stringFlag(args, "config");
  const loaded = configPath ? await loadGenerationConfig(configPath) : inlineConfigContext();
  const context = await buildGenerationContext(loaded, overridesFrom(args));

  if (!quiet) {
    for (const note of context.notes) console.log(note);
  }

  const result = await generateFromFile(input, {
    ...context.generateOptions,
    ...(context.outputDir ? { outDir: context.outputDir } : {}),
    writeReport: context.writeRenderReport,
    writeAdvisorReport: context.writeAdvisorReport,
  });

  const cwd = process.cwd();
  const rel = (p: string) => relative(cwd, p) || p;

  if (!quiet) {
    console.log(`入力    : ${rel(result.inputPath)}`);
    console.log(`テーマ  : ${result.plan.themeId}`);
    if (result.plan.templateId) console.log(`テンプレ: ${result.plan.templateId}`);
    if (context.policy) console.log(`選択方式: ${context.policy.mode}`);
    if (context.aiProvider) console.log(`AI      : ${context.aiProvider.id}`);
    if (context.knowledgeStore) console.log(`知識    : ${context.knowledgeStore.id}`);
    console.log(`スライド: ${result.plan.slides.length} 枚`);
    console.log("");
    console.log("Pattern 選択結果");
    for (const slide of result.plan.slides) {
      const fallback = slide.fallbackUsed ? " (fallback)" : "";
      const by = slide.decisionTrace ? ` [${slide.decisionTrace.decision.selectedBy}]` : "";
      console.log(
        `  ${String(slide.slideNumber).padStart(2, " ")}. ${slide.slideId}  →  ${slide.patternId}/${slide.variantId}  score=${slide.score.toFixed(2)}${fallback}${by}`,
      );
      console.log(`      理由: ${slide.reason}`);
      const ai = slide.decisionTrace?.ai;
      if (ai) {
        console.log(
          `      AI  : ${ai.status}${ai.recommendedPatternId ? ` → ${ai.recommendedPatternId}` : ""}${
            ai.rejectionReason ? ` (${ai.rejectionReason})` : ""
          }`,
        );
      }
      for (const note of slide.notes ?? []) console.log(`      注記: ${note}`);
    }
    console.log("");
  }

  const errors = result.issues.filter((i) => i.severity === "error");
  const warnings = result.issues.filter((i) => i.severity === "warning");

  if (result.issues.length > 0) {
    console.log(`レイアウト検査: エラー ${errors.length} 件 / 警告 ${warnings.length} 件`);
    for (const issue of result.issues) console.log(formatIssue(issue));
    console.log("");
  } else if (!quiet) {
    console.log("レイアウト検査: 問題なし");
    console.log("");
  }

  await recordKnowledge(context, result, quiet);

  console.log(`生成: ${rel(result.pptxPath)} (${(result.pptxBuffer.length / 1024).toFixed(1)} KB)`);
  console.log(`生成: ${rel(result.planPath)}`);
  if (result.reportPath) console.log(`生成: ${rel(result.reportPath)}`);
  if (result.structureReviewPath) console.log(`生成: ${rel(result.structureReviewPath)}`);
  if (result.advisorReportPath) console.log(`生成: ${rel(result.advisorReportPath)}`);

  if (errors.length > 0) return 1;
  if (boolFlag(args, "strict") && warnings.length > 0) return 1;
  return 0;
}
