import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";
import { parsePresentation } from "./domain/schema.js";
import type { PresentationSpec } from "./domain/types.js";
import type { PatternRegistry } from "./patterns/registry.js";
import {
  buildVisualPlan,
  serializeVisualPlan,
  type VisualPlan,
} from "./plan/visual-plan.js";
import {
  buildRenderReport,
  renderVisualPlan,
  serializeRenderReport,
  writePptxFile,
  type RenderReport,
  type SlideRenderReport,
} from "./render/renderer.js";
import type { LayoutIssue } from "./render/validate.js";
import {
  defaultTemplateRegistry,
  type TemplateRegistry,
} from "./templates/registry.js";
import type { LoadedTemplateAsset, PresentationTemplate } from "./templates/template.js";
import { getTheme, type PresentationTheme } from "./theme/theme.js";

/**
 * End-to-End パイプライン
 *
 *   presentation.json
 *     ↓ parsePresentation      (Schema 検証)
 *     ↓ buildVisualPlan        (Selector → Visual Plan)
 *     ↓ renderVisualPlan       (決定論的 Renderer)
 *   .pptx
 */

export interface GenerateOptions {
  /** ビルトインテーマ ID（'default' | 'midnight'） */
  themeId?: string;
  /** 直接 Theme を渡す場合（themeId より優先） */
  theme?: PresentationTheme;
  /** 企業テンプレート ID（TemplateRegistry から解決される） */
  templateId?: string;
  /** Template を直接渡す場合（templateId より優先） */
  template?: PresentationTemplate;
  /** options.template と併用する画像アセット */
  templateAssets?: LoadedTemplateAsset[];
  /** templateId の解決に使う Registry（既定: ビルトインのみ） */
  templateRegistry?: TemplateRegistry;
  /** 章扉（section master）として描くスライド id */
  sectionSlideIds?: readonly string[];
  registry?: PatternRegistry;
  includeTitleSlide?: boolean;
  showSlideNumbers?: boolean;
  includeCandidates?: boolean;
  minimumScore?: number;
}

export interface GenerateResult {
  spec: PresentationSpec;
  plan: VisualPlan;
  reports: SlideRenderReport[];
  issues: LayoutIssue[];
  renderReport: RenderReport;
  pptxBuffer: Buffer;
}

export interface GenerateFilesResult extends GenerateResult {
  inputPath: string;
  pptxPath: string;
  planPath: string;
  /** --report 指定時のみ */
  reportPath?: string;
}

interface ResolvedTemplateSelection {
  template?: PresentationTemplate;
  assets: LoadedTemplateAsset[];
}

function templateOf(options: GenerateOptions): ResolvedTemplateSelection {
  if (options.template) {
    return { template: options.template, assets: options.templateAssets ?? [] };
  }
  if (!options.templateId) return { assets: [] };
  const entry = (options.templateRegistry ?? defaultTemplateRegistry).get(options.templateId);
  return { template: entry.template, assets: entry.assets };
}

/**
 * Theme の決定順序:
 *  1. options.theme（直接指定）
 *  2. options.themeId（--theme による明示的な上書き）
 *  3. Template の theme
 *  4. ビルトインの default テーマ
 */
function themeOf(options: GenerateOptions, template?: PresentationTemplate): PresentationTheme {
  if (options.theme) return options.theme;
  if (options.themeId) return getTheme(options.themeId);
  return template?.theme ?? getTheme(undefined);
}

/** 検証済み spec からメモリ上で生成する */
export async function generate(
  spec: PresentationSpec,
  options: GenerateOptions = {},
): Promise<GenerateResult> {
  const { template, assets } = templateOf(options);
  const theme = themeOf(options, template);
  const plan = buildVisualPlan(spec, {
    registry: options.registry,
    themeId: theme.id,
    ...(template ? { templateId: template.id } : {}),
    ...(options.sectionSlideIds ? { sectionSlideIds: options.sectionSlideIds } : {}),
    includeCandidates: options.includeCandidates,
    minimumScore: options.minimumScore,
  });
  const result = renderVisualPlan(plan, {
    theme,
    registry: options.registry,
    ...(template ? { template, templateAssets: assets } : {}),
    includeTitleSlide: options.includeTitleSlide,
    ...(options.showSlideNumbers !== undefined
      ? { showSlideNumbers: options.showSlideNumbers }
      : {}),
  });
  const renderReport = buildRenderReport(plan, result);
  const pptxBuffer = (await result.pptx.write({ outputType: "nodebuffer" })) as Buffer;
  return { spec, plan, reports: result.reports, issues: result.issues, renderReport, pptxBuffer };
}

export interface GenerateFromFileOptions extends GenerateOptions {
  /** 出力先ディレクトリ（既定: 入力ファイルと同階層の generated/） */
  outDir?: string;
  /** Render Report（配置結果 + 検査結果）も書き出すか */
  writeReport?: boolean;
}

/** presentation.json ファイルから .pptx と .visual-plan.json を出力する */
export async function generateFromFile(
  inputPath: string,
  options: GenerateFromFileOptions = {},
): Promise<GenerateFilesResult> {
  const absoluteInput = resolve(inputPath);
  const raw = await readFile(absoluteInput, "utf8");

  let json: unknown;
  try {
    json = JSON.parse(raw);
  } catch (error) {
    throw new Error(
      `JSON として読み込めませんでした: ${absoluteInput}\n${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const spec = parsePresentation(json);
  const result = await generate(spec, options);

  const outDir = options.outDir ?? join(dirname(absoluteInput), "generated");
  const stem = basename(absoluteInput, extname(absoluteInput));
  const pptxPath = join(outDir, `${stem}.pptx`);
  const planPath = join(outDir, `${stem}.visual-plan.json`);

  await mkdir(outDir, { recursive: true });
  await writeFile(pptxPath, result.pptxBuffer);
  await writeFile(planPath, serializeVisualPlan(result.plan), "utf8");

  let reportPath: string | undefined;
  if (options.writeReport) {
    reportPath = join(outDir, `${stem}.render-report.json`);
    await writeFile(reportPath, serializeRenderReport(result.renderReport), "utf8");
  }

  return { ...result, inputPath: absoluteInput, pptxPath, planPath, reportPath };
}

export { writePptxFile };
