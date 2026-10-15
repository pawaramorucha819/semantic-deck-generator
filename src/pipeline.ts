import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";
import { parsePresentation } from "./domain/schema.js";
import type { PresentationSpec } from "./domain/types.js";
import type { PatternRegistry } from "./patterns/registry.js";
import {
  buildAdvisorReport,
  serializeAdvisorReport,
  type AdvisorReport,
} from "./plan/advisor-report.js";
import {
  buildVisualPlan,
  buildVisualPlanWithPolicy,
  serializeVisualPlan,
  type VisualPlan,
} from "./plan/visual-plan.js";
import type { SelectionPolicyOptions, SlideSelectionOutcome } from "./selector/selection-policy.js";
import {
  buildRenderReport,
  renderVisualPlan,
  serializeRenderReport,
  writePptxFile,
  type RenderReport,
  type SlideRenderReport,
} from "./render/renderer.js";
import type { LayoutIssue } from "./render/validate.js";
import { buildStructureReview } from "./review/structure.js";
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
  /**
   * Selection Policy（manual / assist / auto と AI・Knowledge の接続）。
   * 未指定なら V1 と同じ決定論的選択だけが走る。
   */
  policy?: SelectionPolicyOptions;
  /** Advisor Report に記録する Knowledge Store 識別子 */
  knowledgeStoreId?: string;
}

export interface GenerateResult {
  spec: PresentationSpec;
  plan: VisualPlan;
  reports: SlideRenderReport[];
  issues: LayoutIssue[];
  renderReport: RenderReport;
  pptxBuffer: Buffer;
  /** Selection Policy を使った場合のみ */
  advisorReport?: AdvisorReport;
  outcomes?: SlideSelectionOutcome[];
}

export interface GenerateFilesResult extends GenerateResult {
  inputPath: string;
  pptxPath: string;
  planPath: string;
  /** --report 指定時のみ */
  reportPath?: string;
  /** --report also records semantic decomposition and relationship review. */
  structureReviewPath?: string;
  /** --advisor-report 指定時のみ */
  advisorReportPath?: string;
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

  const planOptions = {
    registry: options.registry,
    themeId: theme.id,
    ...(template ? { templateId: template.id } : {}),
    ...(options.sectionSlideIds ? { sectionSlideIds: options.sectionSlideIds } : {}),
    includeCandidates: options.includeCandidates,
    minimumScore: options.minimumScore,
  };

  // Selection Policy（AI / Knowledge / 人間の指定）を使うかどうかで入口が変わる。
  // どちらの経路でも、出来上がった Visual Plan 以降の処理は完全に同じで決定論的。
  let plan: VisualPlan;
  let outcomes: SlideSelectionOutcome[] | undefined;
  let advisorReport: AdvisorReport | undefined;

  if (options.policy) {
    const result = await buildVisualPlanWithPolicy(spec, { ...planOptions, policy: options.policy });
    plan = result.plan;
    outcomes = result.outcomes;
    advisorReport = buildAdvisorReport(result.outcomes, {
      mode: options.policy.mode,
      ...(options.policy.advisor ? { provider: options.policy.advisor.providerId } : {}),
      ...(options.knowledgeStoreId ? { knowledgeStore: options.knowledgeStoreId } : {}),
    });
  } else {
    plan = buildVisualPlan(spec, planOptions);
  }
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
  return {
    spec,
    plan,
    reports: result.reports,
    issues: result.issues,
    renderReport,
    pptxBuffer,
    ...(advisorReport ? { advisorReport } : {}),
    ...(outcomes ? { outcomes } : {}),
  };
}

export interface GenerateFromFileOptions extends GenerateOptions {
  /** 出力先ディレクトリ（既定: 入力ファイルと同階層の generated/） */
  outDir?: string;
  /** Render Report（配置結果 + 検査結果）も書き出すか */
  writeReport?: boolean;
  /** Advisor Report（候補・参照・AI 推薦の詳細）も書き出すか */
  writeAdvisorReport?: boolean;
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
  let structureReviewPath: string | undefined;
  if (options.writeReport) {
    reportPath = join(outDir, `${stem}.render-report.json`);
    await writeFile(reportPath, serializeRenderReport(result.renderReport), "utf8");
    structureReviewPath = join(outDir, `${stem}.structure-review.json`);
    await writeFile(structureReviewPath, `${JSON.stringify(buildStructureReview(spec, result.plan), null, 2)}\n`, "utf8");
  }

  let advisorReportPath: string | undefined;
  if (options.writeAdvisorReport && result.advisorReport) {
    advisorReportPath = join(outDir, `${stem}.advisor-report.json`);
    await writeFile(advisorReportPath, serializeAdvisorReport(result.advisorReport), "utf8");
  }

  return {
    ...result,
    inputPath: absoluteInput,
    pptxPath,
    planPath,
    ...(reportPath ? { reportPath } : {}),
    ...(structureReviewPath ? { structureReviewPath } : {}),
    ...(advisorReportPath ? { advisorReportPath } : {}),
  };
}

export { writePptxFile };
