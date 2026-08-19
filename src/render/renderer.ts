import { mkdir, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import PptxGenJS from "pptxgenjs";
import type { PresentationSpec } from "../domain/types.js";
import type { PatternRegistry } from "../patterns/registry.js";
import { defaultRegistry } from "../patterns/registry.js";
import {
  buildVisualPlan,
  parseVisualPlan,
  type BuildVisualPlanOptions,
  type SlideProvenance,
  type VisualPlan,
} from "../plan/visual-plan.js";
import { getTheme, normalizeHex, resolveTheme, type PresentationTheme } from "../theme/theme.js";
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
 * Visual Plan → PPTX（描画バックエンド）。
 *
 * ここは完全に決定論的。Renderer は bindings しか見ないので、
 * 同じ Visual Plan からは常に同じスライド構造が出力される。
 *
 * 再生の契約:
 *  - 入口で必ず `parseVisualPlan()` を通す。schemaVersion / 構造 / patternId /
 *    patternVersion / variantId / bindings のすべてが、渡された Registry を基準に
 *    検証される。検証を通らない Plan は描画されず必ず例外になる。
 *  - Theme は既定で `plan.themeId` から解決される（options.theme で上書き可）。
 *    ビルトイン以外のテーマで生成した Plan は、同じテーマを options.theme で
 *    再度渡さない限り再生できない（未知の themeId は例外になる）。
 *  「黙って違う出力になる」ことはない。
 */
export function renderVisualPlan(plan: VisualPlan, options: RenderOptions = {}): RenderResult {
  const registry = options.registry ?? defaultRegistry;

  // 描画の境界で検証する。parseVisualPlan 済みの Plan を渡しても冪等。
  // これにより「parseVisualPlan は拒否するが renderVisualPlan は受け入れる」
  // という抜け道が塞がれる。
  const validatedPlan = parseVisualPlan(plan, { registry });

  const theme = resolveTheme(options.theme ?? getTheme(validatedPlan.themeId));
  const includeTitleSlide = options.includeTitleSlide ?? true;
  const showSlideNumbers = options.showSlideNumbers ?? true;

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = validatedPlan.title;
  pptx.subject = validatedPlan.subtitle ?? "";
  pptx.author = validatedPlan.author ?? "";
  pptx.company = "";

  const canvas: Rect = rect(0, 0, SLIDE_WIDTH_IN, SLIDE_HEIGHT_IN);
  const frame: Rect = rect(
    theme.margins.left,
    theme.margins.top,
    SLIDE_WIDTH_IN - theme.margins.left - theme.margins.right,
    SLIDE_HEIGHT_IN - theme.margins.top - theme.margins.bottom,
  );

  const reports: SlideRenderReport[] = [];
  const total = validatedPlan.slides.length;

  if (includeTitleSlide) {
    const slide = pptx.addSlide();
    slide.background = { color: normalizeHex(theme.colors.background) };
    const surface = new SlideSurface(slide, theme, canvas, frame, "title", 0, total);
    renderTitleSlide(surface, validatedPlan);
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

  // ここから先の planSlide は parseVisualPlan により
  // patternId / patternVersion / variantId / bindings がすべて検証済み。
  for (const planSlide of validatedPlan.slides) {
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

    // 完全な出典（URL・参照箇所込み）はスピーカーノートへ格納する。
    // フッターの短い出典行と役割を分ける（表示用 / 検証用）。
    const notesText = buildNotesText(planSlide.provenance);
    if (notesText) slide.addNotes(notesText);

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

/**
 * スピーカーノート用の完全な出典テキストを組み立てる。
 * フッターの表示用出典（bindings.sourceText）と違い、URL・参照箇所・取得日を落とさない。
 */
function buildNotesText(provenance: SlideProvenance | undefined): string | undefined {
  if (!provenance) return undefined;
  const lines: string[] = [];

  if (provenance.sources.length > 0) {
    lines.push("出典:");
    for (const s of provenance.sources) {
      const bits = [s.title];
      if (s.publisher) bits.push(s.publisher);
      if (s.date) bits.push(s.date);
      if (s.locator) bits.push(`参照箇所: ${s.locator}`);
      if (s.url) bits.push(s.url);
      if (s.accessedAt) bits.push(`取得日: ${s.accessedAt}`);
      lines.push(`- [${s.id}] ${bits.join(" | ")}`);
    }
  }

  if (provenance.evidence.length > 0) {
    if (lines.length > 0) lines.push("");
    lines.push("根拠:");
    for (const e of provenance.evidence) {
      const ref = e.sourceIds?.length ? ` (出典: ${e.sourceIds.join(", ")})` : "";
      lines.push(`- [${e.id}] ${e.statement}${ref}`);
    }
  }

  return lines.length > 0 ? lines.join("\n") : undefined;
}

/**
 * Render Report — 配置結果（矩形・検査結果）の機械可読な記録。
 *
 * Visual Plan は「Pattern 呼び出し + 検証済み props」であり最終座標を持たない。
 * 最終的なレイアウトの検証・AI Visual Review にはこちらを使う。
 */
export interface RenderReport {
  schemaVersion: "1.0";
  title: string;
  themeId: string;
  slides: SlideRenderReport[];
  issues: LayoutIssue[];
}

export function buildRenderReport(plan: VisualPlan, result: RenderResult): RenderReport {
  return {
    schemaVersion: "1.0",
    title: plan.title,
    themeId: plan.themeId,
    slides: result.reports,
    issues: result.issues,
  };
}

export function serializeRenderReport(report: RenderReport): string {
  return `${JSON.stringify(report, null, 2)}\n`;
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
