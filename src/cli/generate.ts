#!/usr/bin/env node
import { relative } from "node:path";
import { PresentationValidationError } from "../domain/schema.js";
import { generateFromFile } from "../pipeline.js";
import { builtinThemes } from "../theme/theme.js";
import type { LayoutIssue } from "../render/validate.js";

interface CliArgs {
  input?: string;
  outDir?: string;
  themeId?: string;
  noTitleSlide: boolean;
  noCandidates: boolean;
  quiet: boolean;
  strict: boolean;
  help: boolean;
}

function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    noTitleSlide: false,
    noCandidates: false,
    quiet: false,
    strict: false,
    help: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i]!;
    switch (token) {
      case "-h":
      case "--help":
        args.help = true;
        break;
      case "--out-dir":
        args.outDir = argv[++i];
        break;
      case "--theme":
        args.themeId = argv[++i];
        break;
      case "--no-title-slide":
        args.noTitleSlide = true;
        break;
      case "--no-candidates":
        args.noCandidates = true;
        break;
      case "--quiet":
        args.quiet = true;
        break;
      case "--strict":
        args.strict = true;
        break;
      default:
        if (token.startsWith("-")) {
          throw new Error(`未知のオプションです: ${token}`);
        }
        args.input ??= token;
    }
  }
  return args;
}

const USAGE = `
Semantic Deck Generator — presentation.json から PowerPoint を生成します

  npm run generate -- <presentation.json> [options]

Options
  --out-dir <dir>     出力先ディレクトリ (既定: 入力と同階層の generated/)
  --theme <id>        テーマ (${Object.keys(builtinThemes).join(" | ")})
  --no-title-slide    表紙スライドを付けない
  --no-candidates     Visual Plan に Pattern 候補一覧を含めない
  --strict            レイアウト警告が 1 件でもあれば終了コード 1
  --quiet             ログを最小限にする
  -h, --help          このヘルプ

Example
  npm run generate -- examples/inventory-system.json
`.trim();

function formatIssue(issue: LayoutIssue): string {
  const mark = issue.severity === "error" ? "ERROR" : "WARN ";
  return `  [${mark}] p${issue.slideNumber} (${issue.patternId}) ${issue.code}: ${issue.message}`;
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));

  if (args.help || !args.input) {
    console.log(USAGE);
    return args.input ? 0 : 1;
  }

  const result = await generateFromFile(args.input, {
    outDir: args.outDir,
    themeId: args.themeId,
    includeTitleSlide: !args.noTitleSlide,
    includeCandidates: !args.noCandidates,
  });

  const cwd = process.cwd();
  const rel = (p: string) => relative(cwd, p) || p;

  if (!args.quiet) {
    console.log(`入力    : ${rel(result.inputPath)}`);
    console.log(`テーマ  : ${result.plan.themeId}`);
    console.log(`スライド: ${result.plan.slides.length} 枚`);
    console.log("");
    console.log("Pattern 選択結果");
    for (const slide of result.plan.slides) {
      const fallback = slide.fallbackUsed ? " (fallback)" : "";
      console.log(
        `  ${String(slide.slideNumber).padStart(2, " ")}. ${slide.slideId}  →  ${slide.patternId}/${slide.variantId}  score=${slide.score.toFixed(2)}${fallback}`,
      );
      console.log(`      理由: ${slide.reason}`);
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
  } else if (!args.quiet) {
    console.log("レイアウト検査: 問題なし");
    console.log("");
  }

  console.log(`生成: ${rel(result.pptxPath)} (${(result.pptxBuffer.length / 1024).toFixed(1)} KB)`);
  console.log(`生成: ${rel(result.planPath)}`);

  if (errors.length > 0) return 1;
  if (args.strict && warnings.length > 0) return 1;
  return 0;
}

main()
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error: unknown) => {
    if (error instanceof PresentationValidationError) {
      console.error(`\n${error.message}`);
      console.error(
        "\nヒント: presentation.json には座標・色・フォントサイズなど描画に関する情報を書けません。",
      );
    } else {
      console.error(`\n${error instanceof Error ? error.message : String(error)}`);
    }
    process.exitCode = 1;
  });
