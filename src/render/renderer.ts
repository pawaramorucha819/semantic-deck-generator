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
import { defineTemplateMasters } from "../templates/master.js";
import { createAssetResolver } from "../templates/loader.js";
import { defaultTemplateRegistry, type TemplateRegistry } from "../templates/registry.js";
import {
  contentFrameOf,
  masterForRole,
  type LoadedTemplateAsset,
  type MasterRole,
  type PresentationTemplate,
  type SlideMasterDefinition,
  type TemplateTextVariables,
} from "../templates/template.js";
import { getTheme, normalizeHex, resolveTheme, type PresentationTheme, type ResolvedTheme } from "../theme/theme.js";
import { rect, type Rect } from "./geometry.js";
import { SlideSurface, type Placement } from "./surface.js";
import { validatePlacements, type LayoutIssue } from "./validate.js";

/** LAYOUT_16x9 の実寸 (inch) */
export const SLIDE_WIDTH_IN = 10;
export const SLIDE_HEIGHT_IN = 5.625;

/** Template が独自スライドサイズを持つ場合に定義するレイアウト名 */
const CUSTOM_LAYOUT_NAME = "SDG_TEMPLATE";

export interface RenderOptions {
  theme?: PresentationTheme;
  registry?: PatternRegistry;
  /**
   * 企業テンプレートを明示的に渡す（`plan.templateId` より優先）。
   * 未指定かつ `plan.templateId` も無い場合は Template 非使用（V1 と同じ紙面）。
   */
  template?: PresentationTemplate;
  /** Template の画像アセット（ロゴ等）。options.template と併用する。 */
  templateAssets?: LoadedTemplateAsset[];
  /** `plan.templateId` の解決に使う Registry（既定: ビルトインのみ） */
  templateRegistry?: TemplateRegistry;
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
  /** 使用した Template（未使用なら undefined） */
  templateId?: string;
}

interface ActiveTemplate {
  template: PresentationTemplate;
  assets: LoadedTemplateAsset[];
}

/**
 * 描画に使う Template を決める。
 *  1. options.template（明示指定）
 *  2. plan.templateId → Registry
 *  3. なし（V1 と同じ紙面）
 * plan.templateId が未登録の場合は例外。黙って Template 無しにはしない。
 */
function resolveActiveTemplate(
  plan: VisualPlan,
  options: RenderOptions,
): ActiveTemplate | undefined {
  if (options.template) {
    return { template: options.template, assets: options.templateAssets ?? [] };
  }
  if (!plan.templateId) return undefined;
  const entry = (options.templateRegistry ?? defaultTemplateRegistry).get(plan.templateId);
  return { template: entry.template, assets: entry.assets };
}

/**
 * Theme の解決順序（V1 の契約をそのまま維持する）:
 *  1. options.theme（明示指定）
 *  2. Template の theme（plan.themeId と一致する場合のみ = --theme で上書きされていない）
 *  3. plan.themeId からビルトインテーマを解決（未知の id は例外）
 */
function resolveThemeFor(
  plan: VisualPlan,
  active: ActiveTemplate | undefined,
  options: RenderOptions,
): ResolvedTheme {
  if (options.theme) return resolveTheme(options.theme);
  const templateTheme = active?.template.theme;
  if (templateTheme && templateTheme.id === plan.themeId) return resolveTheme(templateTheme);
  return resolveTheme(getTheme(plan.themeId));
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

  const active = resolveActiveTemplate(validatedPlan, options);
  const theme = resolveThemeFor(validatedPlan, active, options);
  const includeTitleSlide = options.includeTitleSlide ?? true;
  const showSlideNumbers = options.showSlideNumbers ?? true;

  const slideSize = active?.template.slideSize ?? {
    widthIn: SLIDE_WIDTH_IN,
    heightIn: SLIDE_HEIGHT_IN,
  };

  const pptx = new PptxGenJS();
  if (slideSize.widthIn === SLIDE_WIDTH_IN && slideSize.heightIn === SLIDE_HEIGHT_IN) {
    pptx.layout = "LAYOUT_16x9";
  } else {
    // Template が独自のスライドサイズを持つ場合のみカスタムレイアウトを定義する
    pptx.defineLayout({
      name: CUSTOM_LAYOUT_NAME,
      width: slideSize.widthIn,
      height: slideSize.heightIn,
    });
    pptx.layout = CUSTOM_LAYOUT_NAME;
  }
  pptx.title = validatedPlan.title;
  pptx.subject = validatedPlan.subtitle ?? "";
  pptx.author = validatedPlan.author ?? "";
  pptx.company = active?.template.metadata?.organization ?? "";

  const canvas: Rect = rect(0, 0, slideSize.widthIn, slideSize.heightIn);
  const defaultFrame: Rect = rect(
    theme.margins.left,
    theme.margins.top,
    slideSize.widthIn - theme.margins.left - theme.margins.right,
    slideSize.heightIn - theme.margins.top - theme.margins.bottom,
  );

  // 画像アセットは Template と Pattern パッケージの両方から来る。
  // Pattern 側は `pattern:<id>:<assetId>` へ名前空間化済みなので衝突しない。
  const allAssets = [...(active?.assets ?? []), ...registry.assets()];
  const assets = allAssets.length > 0 ? createAssetResolver(allAssets) : undefined;
  const masterNames =
    active
      ? defineTemplateMasters({
          pptx,
          template: active.template,
          theme,
          assets: assets ?? createAssetResolver([]),
          includeSlideNumbers: showSlideNumbers,
          variables: templateVariablesOf(validatedPlan, active.template),
        })
      : undefined;

  /** 役割に応じた Master でスライドを作り、本文領域を返す */
  const openSlide = (role: MasterRole): { slide: PptxGenJS.Slide; frame: Rect } => {
    if (!active || !masterNames) {
      const slide = pptx.addSlide();
      slide.background = { color: normalizeHex(theme.colors.background) };
      return { slide, frame: defaultFrame };
    }
    const master = masterForRole(active.template, role);
    const masterName = masterNames.get(master.id);
    const slide = pptx.addSlide(masterName ? { masterName } : undefined);
    return { slide, frame: contentFrameOf(master, theme, slideSize) };
  };

  const reports: SlideRenderReport[] = [];
  const total = validatedPlan.slides.length;

  if (includeTitleSlide) {
    const opened = openSlide("title");
    const surface = new SlideSurface({
      slide: opened.slide,
      theme,
      canvas,
      frame: opened.frame,
      variantId: "title",
      slideNumber: 0,
      totalSlides: total,
      ...(assets ? { assets } : {}),
    });
    // Template 使用時は帯・ロゴ等が Master 側にあるので、表紙は文字だけを描く
    renderTitleSlide(surface, validatedPlan, {
      drawAccentBand: !active,
      onColoredBackground: isColoredTitleBackground(active, theme),
    });
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
        frame: opened.frame,
      }),
    });
  }

  // ここから先の planSlide は parseVisualPlan により
  // patternId / patternVersion / variantId / bindings がすべて検証済み。
  for (const planSlide of validatedPlan.slides) {
    const pattern = registry.get(planSlide.patternId);
    const opened = openSlide(planSlide.masterRole ?? "content");

    const surface = new SlideSurface({
      slide: opened.slide,
      theme,
      canvas,
      frame: opened.frame,
      variantId: planSlide.variantId,
      slideNumber: planSlide.slideNumber,
      totalSlides: total,
      ...(assets ? { assets } : {}),
    });

    pattern.render(planSlide.bindings, surface);

    // Template 使用時のページ番号は Slide Master 側（PowerPoint のページ番号
    // フィールド）が担当するため、ここでは描かない。
    if (showSlideNumbers && !active) {
      drawSlideNumberChrome(opened.slide, theme, planSlide.slideNumber, total, slideSize);
    }

    // 完全な出典（URL・参照箇所込み）はスピーカーノートへ格納する。
    // フッターの短い出典行と役割を分ける（表示用 / 検証用）。
    const notesText = buildNotesText(planSlide.provenance);
    if (notesText) opened.slide.addNotes(notesText);

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
        frame: opened.frame,
      }),
    });
  }

  return {
    pptx,
    reports,
    issues: reports.flatMap((r) => r.issues),
    ...(active ? { templateId: active.template.id } : {}),
  };
}

/** Master のテキストで使える変数を Plan と Template から組み立てる */
function templateVariablesOf(
  plan: VisualPlan,
  template: PresentationTemplate,
): TemplateTextVariables {
  return {
    deckTitle: plan.title,
    ...(plan.subtitle ? { deckSubtitle: plan.subtitle } : {}),
    ...(plan.author ? { author: plan.author } : {}),
    ...(template.metadata?.organization ? { organization: template.metadata.organization } : {}),
    ...(template.metadata?.department ? { department: template.metadata.department } : {}),
    ...(template.metadata?.confidentialityLabel
      ? { confidentiality: template.metadata.confidentialityLabel }
      : {}),
  };
}

/**
 * 表紙 Master の背景が濃色かどうか。
 * 濃色なら表紙の文字色を onPrimary 側へ切り替える（Theme トークンのみで判定する）。
 */
function isColoredTitleBackground(
  active: { template: PresentationTemplate } | undefined,
  theme: ResolvedTheme,
): boolean {
  if (!active) return false;
  const master = masterForRole(active.template, "title");
  const token = master.background?.color;
  if (!token || typeof token !== "string") return false;
  return token === "primary" || token === "secondary" || token === "accent" || token === "text";
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

function renderTitleSlide(
  ctx: SlideSurface,
  plan: VisualPlan,
  options: { drawAccentBand: boolean; onColoredBackground: boolean },
): void {
  const { theme, canvas, frame } = ctx;
  const titleColor = options.onColoredBackground ? "onPrimary" : "text";
  const subColor = options.onColoredBackground ? "onPrimary" : "mutedText";

  if (options.drawAccentBand) {
    ctx.shape("rect", { x: 0, y: 0, w: 0.34, h: canvas.h }, {
      fill: "primary",
      layer: "background",
      name: "title-band",
    });
  }

  const titleSize = Math.round(theme.fontSizes.keyMessage * 1.8);
  const textOffset = options.drawAccentBand ? 0.4 : 0;
  const titleW = frame.w - textOffset;
  const titleH = Math.max(
    0.7,
    ctx.measureText(plan.title, { widthIn: titleW, fontSize: titleSize, lineHeight: 1.24 }),
  );

  const blockH =
    titleH +
    (plan.subtitle ? theme.spacing.md + 0.36 : 0) +
    (plan.author ? theme.spacing.lg + 0.26 : 0);
  let cursor = frame.y + Math.max(0, (frame.h - blockH) / 2);

  ctx.text({ x: frame.x + textOffset, y: cursor, w: titleW, h: titleH }, plan.title, {
    role: "keyMessage",
    fontSize: titleSize,
    color: titleColor,
    bold: true,
    lineHeight: 1.24,
    layer: "content",
    name: "deck-title",
  });
  cursor += titleH;

  if (plan.subtitle) {
    cursor += theme.spacing.md;
    ctx.text({ x: frame.x + textOffset, y: cursor, w: titleW, h: 0.36 }, plan.subtitle, {
      role: "heading",
      color: subColor,
      layer: "content",
      name: "deck-subtitle",
    });
    cursor += 0.36;
  }

  if (plan.author) {
    cursor += theme.spacing.lg;
    ctx.text({ x: frame.x + textOffset, y: cursor, w: titleW, h: 0.26 }, plan.author, {
      role: "caption",
      color: subColor,
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
  slideSize: { widthIn: number; heightIn: number },
): void {
  slide.addText(`${slideNumber} / ${total}`, {
    x: slideSize.widthIn - theme.margins.right - 1.2,
    y: slideSize.heightIn - theme.margins.bottom + 0.02,
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
