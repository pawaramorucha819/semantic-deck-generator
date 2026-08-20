import type PptxGenJS from "pptxgenjs";
import type { ResolvedTheme } from "../theme/theme.js";
import type {
  AssetResolver,
  ChartSpec,
  ColorToken,
  ImageSpec,
  Layer,
  RenderContext,
  ShapeKind,
  ShapeStyle,
  TextContent,
  TextRole,
  TextStyle,
} from "./context.js";
import { estimateTextHeight, roundRectValues, type Rect } from "./geometry.js";
import { defaultFontRoleFor, resolveColor, resolveFontFace, tintToken } from "./tokens.js";

/** 配置された 1 要素の記録（レイアウト検査に使う） */
export interface Placement {
  kind: "text" | "shape" | "chart" | "image";
  layer: Layer;
  box: Rect;
  name: string;
  /** kind === 'text' のときのみ */
  text?: string;
  /** kind === 'text' のときのみ (pt) */
  fontSize?: number;
  lineHeight?: number;
  paddingPt?: number;
  shrink?: boolean;
}

const DEFAULT_LINE_HEIGHT = 1.3;

export interface SlideSurfaceOptions {
  slide: PptxGenJS.Slide;
  theme: ResolvedTheme;
  canvas: Rect;
  frame: Rect;
  variantId: string;
  slideNumber: number;
  totalSlides: number;
  /** ロゴ等の埋め込みアセット。未指定なら image() は必ず例外になる。 */
  assets?: AssetResolver;
}

/** PptxGenJS の Slide に描き込みつつ、配置矩形をすべて記録する RenderContext 実装 */
export class SlideSurface implements RenderContext {
  readonly placements: Placement[] = [];
  private counter = 0;

  private readonly slide: PptxGenJS.Slide;
  readonly theme: ResolvedTheme;
  readonly canvas: Rect;
  readonly frame: Rect;
  readonly variantId: string;
  readonly slideNumber: number;
  readonly totalSlides: number;
  private readonly assets: AssetResolver | undefined;

  constructor(options: SlideSurfaceOptions) {
    this.slide = options.slide;
    this.theme = options.theme;
    this.canvas = options.canvas;
    this.frame = options.frame;
    this.variantId = options.variantId;
    this.slideNumber = options.slideNumber;
    this.totalSlides = options.totalSlides;
    this.assets = options.assets;
  }

  fontSizeOf(role: TextRole): number {
    return this.theme.fontSizes[role];
  }

  /** 背景色方向へ混色する。ダークテーマでも「淡くなる」意味が保たれる。 */
  tint(token: ColorToken, amount: number): ColorToken {
    return tintToken(this.theme, token, amount);
  }

  private fontFace(role: "heading" | "body" | "numeric" = "body"): string {
    return resolveFontFace(this.theme, role);
  }

  private color(token: ColorToken | undefined, fallback: ColorToken = "text"): string {
    return resolveColor(this.theme, token, fallback);
  }

  private nextName(prefix: string, explicit?: string): string {
    this.counter += 1;
    return explicit ?? `${prefix}-${this.slideNumber}-${this.counter}`;
  }

  private static flatten(content: TextContent): string {
    return typeof content === "string"
      ? content
      : content.map((run) => run.text + (run.breakLine ? "\n" : "")).join("");
  }

  text(box: Rect, content: TextContent, style: TextStyle = {}): void {
    const role = style.role ?? "body";
    const fontSize = style.fontSize ?? this.fontSizeOf(role);
    const lineHeight = style.lineHeight ?? DEFAULT_LINE_HEIGHT;
    const paddingPt = style.paddingPt ?? 0;
    const name = this.nextName("text", style.name);
    const b = roundRectValues(box);

    const base: PptxGenJS.TextPropsOptions = {
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      fontFace: this.fontFace(style.font ?? defaultFontRoleFor(role)),
      fontSize,
      color: this.color(style.color, "text"),
      bold: style.bold ?? (role === "keyMessage" || role === "display"),
      italic: style.italic ?? false,
      align: style.align ?? "left",
      valign: style.valign ?? "top",
      margin: paddingPt,
      lineSpacingMultiple: Number(lineHeight.toFixed(2)),
      charSpacing: style.charSpacing,
      wrap: style.wrap ?? true,
      isTextBox: true,
      objectName: name,
      ...(style.shrink ? { fit: "shrink" as const } : {}),
    };

    if (typeof content === "string") {
      this.slide.addText(content, base);
    } else {
      const runs: PptxGenJS.TextProps[] = content.map((run) => ({
        text: run.text,
        options: {
          fontSize: run.style?.fontSize,
          fontFace: run.style?.font ? this.fontFace(run.style.font) : undefined,
          color: run.style?.color ? this.color(run.style.color) : undefined,
          bold: run.style?.bold,
          italic: run.style?.italic,
          charSpacing: run.style?.charSpacing,
          breakLine: run.breakLine ?? false,
        },
      }));
      this.slide.addText(runs, base);
    }

    this.placements.push({
      kind: "text",
      layer: style.layer ?? "content",
      box: b,
      name,
      text: SlideSurface.flatten(content),
      fontSize,
      lineHeight,
      paddingPt,
      shrink: style.shrink ?? false,
    });
  }

  shape(kind: ShapeKind, box: Rect, style: ShapeStyle = {}): void {
    const name = this.nextName("shape", style.name);
    const b = roundRectValues(box);
    const opts: PptxGenJS.ShapeProps = {
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      objectName: name,
    };

    if (style.fill && style.fill !== "none") {
      opts.fill = { color: this.color(style.fill), transparency: style.fillTransparency };
    } else if (style.fill === "none") {
      opts.fill = { type: "none" };
    }

    if (style.line) {
      opts.line = {
        color: this.color(style.line.color),
        width: style.line.width ?? this.theme.shape.borderWidth,
        dashType: style.line.dashType ?? "solid",
      };
    }

    if (kind === "roundRect") {
      opts.rectRadius = style.cornerRadius ?? this.theme.shape.cornerRadius;
    }
    if (style.rotate !== undefined) opts.rotate = style.rotate;
    if (style.flipH) opts.flipH = true;
    if (style.flipV) opts.flipV = true;

    this.slide.addShape(kind as PptxGenJS.SHAPE_NAME, opts);

    this.placements.push({
      kind: "shape",
      layer: style.layer ?? "decoration",
      box: b,
      name,
    });
  }

  chart(box: Rect, spec: ChartSpec): void {
    const name = this.nextName("chart", spec.name);
    const b = roundRectValues(box);
    const colors = (spec.colors ?? ["primary", "secondary", "accent", "mutedText"]).map((c) =>
      this.color(c),
    );

    const data = spec.series.map((series) => ({
      name: series.name,
      labels: spec.categories,
      values: series.values,
    }));

    const chartType = (spec.kind === "bar" ? "bar" : spec.kind) as PptxGenJS.CHART_NAME;

    this.slide.addChart(chartType, data, {
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      objectName: name,
      barDir: spec.barDir ?? "col",
      chartColors: colors,
      showLegend: spec.showLegend ?? false,
      legendPos: "b",
      legendColor: this.color("mutedText"),
      legendFontSize: this.theme.fontSizes.caption,
      showValue: spec.showValue ?? true,
      dataLabelColor: this.color("text"),
      dataLabelFontSize: this.theme.fontSizes.caption,
      dataLabelFontFace: this.fontFace("numeric"),
      dataLabelFormatCode: spec.valueUnit ? `#,##0"${spec.valueUnit}"` : "#,##0",
      catAxisLabelColor: this.color("mutedText"),
      catAxisLabelFontFace: this.fontFace("body"),
      catAxisLabelFontSize: this.theme.fontSizes.caption,
      valAxisLabelColor: this.color("mutedText"),
      valAxisLabelFontFace: this.fontFace("numeric"),
      valAxisLabelFontSize: this.theme.fontSizes.caption,
      valGridLine: { color: this.color("border"), style: "solid", size: 0.5 },
      catGridLine: { style: "none" },
      valAxisLineShow: false,
      catAxisLineShow: true,
      catAxisLineColor: this.color("border"),
      barGapWidthPct: 60,
      chartArea: { fill: { color: this.color("background") } },
      plotArea: { fill: { color: this.color("background") } },
    });

    this.placements.push({
      kind: "chart",
      layer: spec.layer ?? "content",
      box: b,
      name,
    });
  }

  image(box: Rect, spec: ImageSpec): void {
    const name = this.nextName("image", spec.name);
    const b = roundRectValues(box);
    if (!this.assets) {
      throw new Error(
        `画像アセット '${spec.assetId}' を解決できません: この描画面にはアセットが登録されていません`,
      );
    }
    const asset = this.assets.resolve(spec.assetId);

    const sizing = spec.sizing ?? "contain";
    this.slide.addImage({
      data: asset.data,
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      objectName: name,
      transparency: spec.transparency,
      ...(sizing === "stretch"
        ? {}
        : { sizing: { type: sizing, w: b.w, h: b.h } as PptxGenJS.ImageProps["sizing"] }),
    });

    this.placements.push({
      kind: "image",
      layer: spec.layer ?? "decoration",
      box: b,
      name,
    });
  }

  measureText(
    content: TextContent,
    opts: {
      widthIn: number;
      role?: TextRole;
      fontSize?: number;
      lineHeight?: number;
      paddingPt?: number;
    },
  ): number {
    const fontSize = opts.fontSize ?? this.fontSizeOf(opts.role ?? "body");
    return estimateTextHeight(SlideSurface.flatten(content), {
      widthIn: opts.widthIn,
      fontSize,
      lineHeight: opts.lineHeight ?? DEFAULT_LINE_HEIGHT,
      horizontalPaddingPt: (opts.paddingPt ?? 0) * 2,
      verticalPaddingPt: (opts.paddingPt ?? 0) * 2,
    });
  }
}
