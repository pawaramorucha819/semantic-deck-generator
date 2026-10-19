#!/usr/bin/env node
import { runArgumentReview, ARGUMENT_USAGE } from "./commands/argument.js";
import { PresentationValidationError } from "../domain/schema.js";
import { CliArgumentError } from "./args.js";
import { runGenerate, GENERATE_USAGE } from "./commands/generate.js";
import { runKnowledge, KNOWLEDGE_USAGE } from "./commands/knowledge.js";
import { runPattern, PATTERN_USAGE } from "./commands/pattern.js";
import { runTemplate, TEMPLATE_USAGE } from "./commands/template.js";
import { runStructure, STRUCTURE_USAGE } from "./commands/structure.js";

/**
 * Semantic Deck Generator CLI（V2）。
 *
 *   generate   presentation.json → .pptx
 *   template   既存 PPTX から企業テンプレートを取り込む / 一覧
 *   pattern    既存 PPTX の 1 スライドから Pattern 下書きを作る / 一覧
 *   knowledge  過去資料を蓄積する / 検索する
 */

const USAGE = `
Semantic Deck Generator — 意味構造から決定論的に PowerPoint を生成します

  semantic-deck <command> [options]

Commands
  generate <presentation.json>   PowerPoint を生成する
  argument-review <file.json>    論拠設計の未承認レビュー票を出力する
  structure <command>           原文から分解候補を検査・レビュー・コンパイルする（オフライン）
  template import <file.pptx>    既存 PPTX から企業テンプレートを取り込む
  template list                  テンプレート一覧
  pattern import <file.pptx>     PPTX の 1 スライドから Pattern 下書きを作る
  pattern list                   Pattern 一覧
  knowledge ingest <path...>     過去資料を Knowledge へ蓄積する
  knowledge search               過去スライドを検索する
  knowledge stats                蓄積状況を表示する

  各コマンドの詳細は  semantic-deck <command> --help

Examples
  npm run sdg -- generate examples/inventory-system.json
  npm run sdg -- generate examples/enterprise/presentation.json --config examples/enterprise/generation.config.json
  npm run sdg -- template import company.pptx --id corporate
  npm run sdg -- pattern import sample.pptx --slide 3 --id company-three-column
  npm run sdg -- knowledge ingest ./past-decks --store ./knowledge
  npm run sdg -- knowledge search --relation comparison --role recommendation
`.trim();

const HELP_BY_COMMAND: Record<string, string> = {
  generate: GENERATE_USAGE,
  "argument-review": ARGUMENT_USAGE,
  template: TEMPLATE_USAGE,
  pattern: PATTERN_USAGE,
  knowledge: KNOWLEDGE_USAGE,
  structure: STRUCTURE_USAGE,
};

export async function main(argv: readonly string[]): Promise<number> {
  const command = argv[0];

  if (!command || command === "--help" || command === "-h" || command === "help") {
    const target = argv[1];
    console.log(target && HELP_BY_COMMAND[target] ? HELP_BY_COMMAND[target] : USAGE);
    return command ? 0 : 1;
  }

  switch (command) {
    case "argument-review":
      return runArgumentReview(argv.slice(1));
    case "generate":
      return runGenerate(argv.slice(1));
    case "structure":
      return runStructure(argv.slice(1));
    case "template":
      return runTemplate(argv.slice(1));
    case "pattern":
      return runPattern(argv.slice(1));
    case "knowledge":
      return runKnowledge(argv.slice(1));
    default:
      // 後方互換: `semantic-deck presentation.json` を generate として扱う
      if (!command.startsWith("-") && command.endsWith(".json")) {
        return runGenerate(argv);
      }
      console.error(`未知のコマンドです: ${command}\n`);
      console.error(USAGE);
      return 1;
  }
}

/** エラーを利用者向けに整形する（原因と次の一手が分かるように） */
export function reportError(error: unknown): void {
  if (error instanceof PresentationValidationError) {
    console.error(`\n${error.message}`);
    console.error(
      "\nヒント: presentation.json には座標・色・フォントサイズ・patternId など" +
        "生成方法に関する情報を書けません。それらは generation.config.json 側で指定します。",
    );
    return;
  }
  if (error instanceof CliArgumentError) {
    console.error(`\n${error.message}`);
    console.error("\nヒント: semantic-deck <command> --help で使い方を確認できます。");
    return;
  }
  console.error(`\n${error instanceof Error ? error.message : String(error)}`);
}

// 直接実行されたときだけ動かす（テストからは main() を import して使う）
const invokedDirectly =
  process.argv[1] !== undefined && /[\\/]cli[\\/]main\.(ts|js)$/.test(process.argv[1]);

if (invokedDirectly) {
  main(process.argv.slice(2))
    .then((code) => {
      process.exitCode = code;
    })
    .catch((error: unknown) => {
      reportError(error);
      process.exitCode = 1;
    });
}
