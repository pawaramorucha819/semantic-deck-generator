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
import { renderVisualPlan, writePptxFile, type SlideRenderReport } from "./render/renderer.js";
import type { LayoutIssue } from "./render/validate.js";
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
  registry?: PatternRegistry;
  includeTitleSlide?: boolean;
  includeCandidates?: boolean;
  minimumScore?: number;
}

export interface GenerateResult {
  spec: PresentationSpec;
  plan: VisualPlan;
  reports: SlideRenderReport[];
  issues: LayoutIssue[];
  pptxBuffer: Buffer;
}

export interface GenerateFilesResult extends GenerateResult {
  inputPath: string;
  pptxPath: string;
  planPath: string;
}

function themeOf(options: GenerateOptions): PresentationTheme {
  return options.theme ?? getTheme(options.themeId);
}

/** 検証済み spec からメモリ上で生成する */
export async function generate(
  spec: PresentationSpec,
  options: GenerateOptions = {},
): Promise<GenerateResult> {
  const theme = themeOf(options);
  const plan = buildVisualPlan(spec, {
    registry: options.registry,
    themeId: theme.id,
    includeCandidates: options.includeCandidates,
    minimumScore: options.minimumScore,
  });
  const { pptx, reports, issues } = renderVisualPlan(plan, {
    theme,
    registry: options.registry,
    includeTitleSlide: options.includeTitleSlide,
  });
  const pptxBuffer = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  return { spec, plan, reports, issues, pptxBuffer };
}

export interface GenerateFromFileOptions extends GenerateOptions {
  /** 出力先ディレクトリ（既定: 入力ファイルと同階層の generated/） */
  outDir?: string;
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

  return { ...result, inputPath: absoluteInput, pptxPath, planPath };
}

export { writePptxFile };
