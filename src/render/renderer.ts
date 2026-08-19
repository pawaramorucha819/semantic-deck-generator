import { mkdir, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import PptxGenJS from "pptxgenjs";
import type { PresentationSpec } from "../domain/types.js";
import type { PatternRegistry } from "../patterns/registry.js";
import { defaultRegistry } from "../patterns/registry.js";
import { buildVisualPlan, type BuildVisualPlanOptions, type VisualPlan } from "../plan/visual-plan.js";
import { defaultTheme, normalizeHex, resolveTheme, type PresentationTheme } from "../theme/theme.js";
import { rect, type Rect } from "./geometry.js";
import { SlideSurface, type Placement } from "./surface.js";
import { validatePlacements, type LayoutIssue } from "./validate.js";

/** LAYOUT_16x9 の実寸 (inch) */
export const SLIDE_WIDTH_IN = 10;
export const SLIDE_HEIGHT_IN = 5.625;

export interface RenderOptions {
  theme?: PresentationTheme;
  registry?: PatternRegistry;
  /** 表紙スライドを付けるか（既定 true） */
  includeTitleSlide?: boolean;
  /** ページ番号を入れるか（既定 true） */
  showSlideNumbers?: boolean;
}

export interface SlideRenderReport {
  slideId: string;
  slideNumber: number;
  patternId: string;
  variantId: string;
  placements: Placement[];
  issues: LayoutIssue[];
}

export interface RenderResult {
  pptx: PptxGenJS;
  reports: SlideRenderReport[];
  issues: LayoutIssue[];
}

/**
 * Visual Plan → PPTX。
 *
 * ここは完全に決定論的。Renderer は bindings しか見ないので、
 * 同じ Visual Plan からは常に同じスライド構造が出力される。
 */
export function renderVisualPlan(plan: VisualPlan, options: RenderOptions = {}): RenderResult {
  const registry = options.registry ?? defaultRegistry;
  const theme = resolveTheme(options.theme ?? defaultTheme);
  const includeTitleSlide = options.includeTitleSlide ?? true;
  const showSlideNumbers = options.showSlideNumbers ?? true;

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = plan.title;
  pptx.subject = plan.subtitle ?? "";
  pptx.author = plan.author ?? "";
  pptx.company = "";

  const canvas: Rect = rect(0, 0, SLIDE_WIDTH_IN, SLIDE_HEIGHT_IN);
  const frame: Rect = rect(
    theme.margins.left,
    theme.margins.top,
    SLIDE_WIDTH_IN - theme.margins.left - theme.margins.right,
    SLIDE_HEIGHT_IN - theme.margins.top - theme.margins.bottom,
  );

  const reports: SlideRenderReport[] = [];
  const total = plan.slides.length;

  if (includeTitleSlide) {
    const slide = pptx.addSlide();
    slide.background = { color: normalizeHex(theme.colors.background) };
    const surface = new SlideSurface(slide, theme, canvas, frame, "title", 0, total);
    renderTitleSlide(surface, plan);
    reports.push({
      slideId: "__title__",
      slideNumber: 0,
      patternId: "__title__",
      variantId: "standard",
      placements: surface.placements,
      issues: validatePlacements(surface.placements, {
        slideId: "__title__",
        slideNumber: 0,
        patternId: "__title__",
        canvas,
        frame,
      }),
    });
  }

  for (const planSlide of plan.slides) {
    const pattern = registry.get(planSlide.patternId);
    const slide = pptx.addSlide();
    slide.background = { color: normalizeHex(theme.colors.background) };

    const surface = new SlideSurface(
      slide,
      theme,
      canvas,
      frame,
      planSlide.variantId,
      planSlide.slideNumber,
      total,
    );

    pattern.render(planSlide.bindings, surface);

    if (showSlideNumbers) {
      drawSlideNumberChrome(slide, theme, planSlide.slideNumber, total);
    }

    reports.push({
      slideId: planSlide.slideId,
      slideNumber: planSlide.slideNumber,
      patternId: planSlide.patternId,
      variantId: planSlide.variantId,
      placements: surface.placements,
      issues: validatePlacements(surface.placements, {
        slideId: planSlide.slideId,
        slideNumber: planSlide.slideNumber,
        patternId: planSlide.patternId,
        canvas,
        frame,
      }),
    });
  }

  return { pptx, reports, issues: reports.flatMap((r) => r.issues) };
}

/** presentation.json から一気に PPTX まで（CLI / テスト用の便宜関数） */
export function renderPresentation(
  spec: PresentationSpec,
  options: RenderOptions & BuildVisualPlanOptions = {},
): RenderResult & { plan: VisualPlan } {
  const plan = buildVisualPlan(spec, {
    registry: options.registry,
    themeId: options.theme?.id ?? options.themeId,
    includeCandidates: options.includeCandidates,
    minimumScore: options.minimumScore,
  });
  const result = renderVisualPlan(plan, options);
  return { ...result, plan };
}

export async function writePptxFile(pptx: PptxGenJS, filePath: string): Promise<number> {
  const buffer = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  await mkdir(dirname(filePath), { recursive: true });
  await writeFile(filePath, buffer);
  return buffer.length;
}

export async function toPptxBuffer(pptx: PptxGenJS): Promise<Buffer> {
  return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
}

function renderTitleSlide(ctx: SlideSurface, plan: VisualPlan): void {
  const { theme, canvas, frame } = ctx;

  ctx.shape("rect", { x: 0, y: 0, w: 0.34, h: canvas.h }, {
    fill: "primary",
    layer: "background",
    name: "title-band",
  });

  const titleSize = Math.round(theme.fontSizes.keyMessage * 1.8);
  const titleW = frame.w - 0.4;
  const titleH = Math.max(
    0.7,
    ctx.measureText(plan.title, { widthIn: titleW, fontSize: titleSize, lineHeight: 1.24 }),
  );

  const blockH =
    titleH +
    (plan.subtitle ? theme.spacing.md + 0.36 : 0) +
    (plan.author ? theme.spacing.lg + 0.26 : 0);
  let cursor = frame.y + Math.max(0, (frame.h - blockH) / 2);

  ctx.text({ x: frame.x + 0.4, y: cursor, w: titleW, h: titleH }, plan.title, {
    role: "keyMessage",
    fontSize: titleSize,
    color: "text",
    bold: true,
    lineHeight: 1.24,
    layer: "content",
    name: "deck-title",
  });
  cursor += titleH;

  if (plan.subtitle) {
    cursor += theme.spacing.md;
    ctx.text({ x: frame.x + 0.4, y: cursor, w: titleW, h: 0.36 }, plan.subtitle, {
      role: "heading",
      color: "mutedText",
      layer: "content",
      name: "deck-subtitle",
    });
    cursor += 0.36;
  }

  if (plan.author) {
    cursor += theme.spacing.lg;
    ctx.text({ x: frame.x + 0.4, y: cursor, w: titleW, h: 0.26 }, plan.author, {
      role: "caption",
      color: "mutedText",
      layer: "content",
      name: "deck-author",
    });
  }
}

/**
 * ページ番号は Pattern の描画対象ではなく Renderer 側の「紙面装飾」なので、
 * RenderContext を経由せず直接描く（レイアウト検査の対象外）。
 */
function drawSlideNumberChrome(
  slide: PptxGenJS.Slide,
  theme: ReturnType<typeof resolveTheme>,
  slideNumber: number,
  total: number,
): void {
  slide.addText(`${slideNumber} / ${total}`, {
    x: SLIDE_WIDTH_IN - theme.margins.right - 1.2,
    y: SLIDE_HEIGHT_IN - theme.margins.bottom + 0.02,
    w: 1.2,
    h: 0.26,
    align: "right",
    valign: "middle",
    margin: 0,
    fontFace: theme.fonts.numeric,
    fontSize: theme.fontSizes.caption,
    color: normalizeHex(theme.colors.mutedText),
    isTextBox: true,
    objectName: `slide-number-${slideNumber}`,
  });
}
