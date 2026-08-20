import { relative, resolve } from "node:path";
import {
  importTemplateFromPptx,
  writeTemplatePackage,
} from "../../templates/importer/import-pptx.js";
import { TemplateRegistry } from "../../templates/registry.js";
import {
  boolFlag,
  listFlag,
  parseArgs,
  requireString,
  stringFlag,
  CliArgumentError,
} from "../args.js";

export const TEMPLATE_USAGE = `
semantic-deck template <import|list> [options]

  import <file.pptx>    既存 PowerPoint から企業テンプレートを取り込む
    --id <id>           作成する template id（必須。英小文字・数字・- _）
    --out-dir <dir>     出力先ルート (既定: ./templates)
    --name <name>       表示名
    --organization <s>  組織名（Master の {{organization}} に入る）
    --department <s>    部署名
    --confidential <s>  機密表示（Master の {{confidentiality}} に入る）
    --imported-at <s>   記録する取り込み日時（決定論性のため任意指定）

  list                  登録済みテンプレートを一覧する
    --template-dir <d>  外部テンプレートディレクトリ (複数可)
    --json              JSON で出力する

取り込める範囲と、取り込めない要素（SmartArt / グラフ / 表 / アニメーション等）は
import-report.json に必ず記録されます。
`.trim();

export async function runTemplate(argv: readonly string[]): Promise<number> {
  const sub = argv[0];
  const rest = argv.slice(1);

  if (!sub || sub === "--help" || sub === "-h") {
    console.log(TEMPLATE_USAGE);
    return sub ? 0 : 1;
  }

  if (sub === "import") {
    const args = parseArgs(rest, {
      boolean: ["help"],
      string: ["id", "out-dir", "name", "organization", "department", "confidential", "imported-at"],
    });
    if (boolFlag(args, "help") || args.positional.length === 0) {
      console.log(TEMPLATE_USAGE);
      return args.positional.length === 0 ? 1 : 0;
    }

    const input = resolve(args.positional[0]!);
    const id = requireString(args, "id");
    const outDir = resolve(stringFlag(args, "out-dir") ?? "templates");

    const imported = await importTemplateFromPptx(input, {
      id,
      ...(stringFlag(args, "name") ? { name: stringFlag(args, "name")! } : {}),
      ...(stringFlag(args, "organization")
        ? { organization: stringFlag(args, "organization")! }
        : {}),
      ...(stringFlag(args, "department") ? { department: stringFlag(args, "department")! } : {}),
      ...(stringFlag(args, "confidential")
        ? { confidentialityLabel: stringFlag(args, "confidential")! }
        : {}),
      sourceLabel: relative(process.cwd(), input) || input,
    });

    const written = await writeTemplatePackage(outDir, imported);

    console.log(`取り込み: ${relative(process.cwd(), input) || input}`);
    console.log(`  スライドサイズ: ${imported.report.slideSize.widthIn} x ${imported.report.slideSize.heightIn} in`);
    console.log(`  マスター      : ${imported.report.extracted.masters}`);
    console.log(`  固定要素      : ${imported.report.extracted.elements}`);
    console.log(`  アセット      : ${imported.report.extracted.assets}`);

    if (imported.report.unsupportedFeatures.length > 0) {
      console.log("");
      console.log("非対応（取り込んでいない要素）:");
      for (const item of imported.report.unsupportedFeatures) console.log(`  - ${item}`);
    }
    if (imported.report.warnings.length > 0) {
      console.log("");
      console.log("警告:");
      for (const item of imported.report.warnings) console.log(`  - ${item}`);
    }

    console.log("");
    console.log(`生成: ${relative(process.cwd(), written.templatePath)}`);
    console.log(`生成: ${relative(process.cwd(), written.reportPath)}`);
    for (const path of written.assetPaths) {
      console.log(`生成: ${relative(process.cwd(), path)}`);
    }
    console.log("");
    console.log(
      `使い方: semantic-deck generate deck.json --template ${id} --template-dir ${relative(process.cwd(), outDir) || "."}`,
    );
    return 0;
  }

  if (sub === "list") {
    const args = parseArgs(rest, { boolean: ["help", "json"], repeatable: ["template-dir"] });
    if (boolFlag(args, "help")) {
      console.log(TEMPLATE_USAGE);
      return 0;
    }

    const registry = new TemplateRegistry();
    for (const dir of listFlag(args, "template-dir")) {
      await registry.registerDirectory(dir);
    }
    const list = registry.list();

    if (boolFlag(args, "json")) {
      console.log(JSON.stringify(list, null, 2));
      return 0;
    }
    for (const item of list) {
      console.log(
        `${item.id.padEnd(20)} ${item.source.padEnd(9)} v${item.version}  masters=${item.masters}  ${item.name}`,
      );
      if (item.dir) console.log(`  ${item.dir}`);
    }
    return 0;
  }

  throw new CliArgumentError(`未知のサブコマンドです: template ${sub}`);
}
