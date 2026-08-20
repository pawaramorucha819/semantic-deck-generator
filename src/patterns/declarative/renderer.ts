import type { ColorToken, RenderContext, TextStyle } from "../../render/context.js";
import {
  columns as splitColumns,
  columnsWeighted,
  inset,
  type Rect,
} from "../../render/geometry.js";
import { resolveSpacing, type SpacingTokenName } from "../../render/tokens.js";
import { drawCard, drawFooter, drawHeader } from "../shared/layout.js";
import type {
  DeclarativeColor,
  DeclarativeLayout,
  DeclarativeTextStyle,
  HeightSpec,
  LayoutNode,
} from "./schema.js";
import {
  evaluateCondition,
  resolveValueRef,
  type BindingScope,
} from "./values.js";

/**
 * Declarative Pattern の描画。
 *
 * JSON から PptxGenJS を直接叩くことは絶対にしない。すべて RenderContext 経由なので、
 * はみ出し・重なり・実効フォントサイズの検査、Theme トークン制約、Render Report への
 * 記録がコード Pattern とまったく同じように機能する。
 *
 * レイアウトは「縦積み(stack) / 横並び(columns) / カード(card)」の単純なフロー計算と、
 * 取り込んだスライド向けの absolute 配置の 2 モードだけを持つ。
 */

const DEFAULT_SHAPE_HEIGHT_IN = 0.08;

export interface DeclarativeRenderOptions {
  /** 画像アセット ID の名前空間（Pattern パッケージのアセット衝突を避ける） */
  assetPrefix?: string;
  /** 要素名の接頭辞（Render Report 上の識別子） */
  namePrefix?: string;
}

export class DeclarativeRenderError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "DeclarativeRenderError";
  }
}

function color(ctx: RenderContext, value: DeclarativeColor | undefined): ColorToken | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") return value;
  return ctx.tint(value.tint, value.amount);
}

function spacing(ctx: RenderContext, value: SpacingTokenName | number | undefined, fallback = 0): number {
  return resolveSpacing(ctx.theme, value, fallback);
}

function textStyleOf(
  ctx: RenderContext,
  style: DeclarativeTextStyle | undefined,
  name: string,
): TextStyle & { fontSize: number; lineHeight: number } {
  const role = style?.role ?? "body";
  const base = style?.fontSize ?? ctx.fontSizeOf(role);
  const fontSize = Math.max(1, Math.round(base * (style?.fontScale ?? 1)));
  const resolved = color(ctx, style?.color);
  return {
    role,
    fontSize,
    ...(style?.font ? { font: style.font } : {}),
    ...(resolved ? { color: resolved } : {}),
    ...(style?.bold !== undefined ? { bold: style.bold } : {}),
    ...(style?.italic !== undefined ? { italic: style.italic } : {}),
    ...(style?.align ? { align: style.align } : {}),
    ...(style?.valign ? { valign: style.valign } : {}),
    lineHeight: style?.lineHeight ?? 1.32,
    ...(style?.charSpacing !== undefined ? { charSpacing: style.charSpacing } : {}),
    ...(style?.shrink ? { shrink: true } : {}),
    layer: "content",
    name,
  };
}

/** repeat 対象の要素配列を取り出す */
function repeatItems(node: Extract<LayoutNode, { type: "columns" }>, scope: BindingScope): Array<Record<string, unknown>> {
  if (!node.repeat) return [];
  const value = scope.bindings[node.repeat.slot];
  if (!Array.isArray(value)) {
    throw new DeclarativeRenderError(
      `repeat が参照する slot '${node.repeat.slot}' が配列ではありません`,
    );
  }
  const max = node.repeat.max ?? value.length;
  return value.slice(0, max) as Array<Record<string, unknown>>;
}

function visible(node: LayoutNode, scope: BindingScope): boolean {
  return node.when ? evaluateCondition(node.when, scope) : true;
}

class DeclarativeRenderer {
  private counter = 0;

  constructor(
    private readonly ctx: RenderContext,
    private readonly options: DeclarativeRenderOptions,
  ) {}

  private name(kind: string): string {
    this.counter += 1;
    const prefix = this.options.namePrefix ?? "decl";
    return `${prefix}-${kind}-${this.counter}`;
  }

  /** ノードの必要高さ (inch)。"fill" は「残りを分け合う」を意味する。 */
  measure(node: LayoutNode, widthIn: number, scope: BindingScope): number | "fill" {
    if (!visible(node, scope)) return 0;
    const explicit = node.type === "spacer" ? undefined : (node.height as HeightSpec | undefined);
    if (explicit === "fill") return "fill";
    if (typeof explicit === "number") return explicit;

    switch (node.type) {
      case "spacer":
        return spacing(this.ctx, node.height);
      case "text": {
        const value = resolveValueRef(node.value, scope);
        if (value === undefined) return 0;
        const style = textStyleOf(this.ctx, node.style, "measure");
        return this.ctx.measureText(value, {
          widthIn,
          fontSize: style.fontSize,
          lineHeight: style.lineHeight,
        });
      }
      case "shape":
        return DEFAULT_SHAPE_HEIGHT_IN;
      case "image":
        return node.height;
      case "stack": {
        const pad = spacing(this.ctx, node.padding);
        const gap = spacing(this.ctx, node.gap);
        const inner = Math.max(0, widthIn - pad * 2);
        const visibleChildren = node.children.filter((child) => visible(child, scope));
        let total = 0;
        for (const child of visibleChildren) {
          const measured = this.measure(child, inner, scope);
          if (measured !== "fill") total += measured;
        }
        return total + Math.max(0, visibleChildren.length - 1) * gap + pad * 2;
      }
      case "columns": {
        const gap = spacing(this.ctx, node.gap);
        if (node.repeat && node.item) {
          const items = repeatItems(node, scope);
          if (items.length === 0) return 0;
          const colWidth = (widthIn - gap * (items.length - 1)) / items.length;
          return items.reduce((max, item, index) => {
            const child = this.measure(node.item!, colWidth, { ...scope, item, index });
            return child === "fill" ? max : Math.max(max, child);
          }, 0);
        }
        const children = (node.children ?? []).filter((child) => visible(child, scope));
        if (children.length === 0) return 0;
        const colWidth = (widthIn - gap * (children.length - 1)) / children.length;
        return children.reduce((max, child) => {
          const measured = this.measure(child, colWidth, scope);
          return measured === "fill" ? max : Math.max(max, measured);
        }, 0);
      }
      case "card": {
        const pad = this.ctx.theme.spacing.md;
        const child = this.measure(node.child, Math.max(0, widthIn - pad * 2), scope);
        return (child === "fill" ? 0 : child) + pad * 2;
      }
      case "absolute":
        return "fill";
    }
  }

  render(node: LayoutNode, box: Rect, scope: BindingScope): void {
    if (!visible(node, scope)) return;
    if (box.w <= 0 || box.h <= 0) return;

    switch (node.type) {
      case "spacer":
        return;

      case "text": {
        const value = resolveValueRef(node.value, scope);
        if (value === undefined) return;
        this.ctx.text(box, value, textStyleOf(this.ctx, node.style, this.name("text")));
        return;
      }

      case "shape": {
        const width = node.width ?? box.w;
        const x =
          node.align === "center"
            ? box.x + (box.w - width) / 2
            : node.align === "right"
              ? box.x + box.w - width
              : box.x;
        const style = node.style;
        this.ctx.shape(
          node.shape,
          { x, y: box.y, w: width, h: box.h },
          {
            ...(color(this.ctx, style?.fill) ? { fill: color(this.ctx, style?.fill)! } : {}),
            ...(style?.fillTransparency !== undefined
              ? { fillTransparency: style.fillTransparency }
              : {}),
            ...(style?.line
              ? {
                  line: {
                    color: color(this.ctx, style.line.color)!,
                    ...(style.line.width !== undefined ? { width: style.line.width } : {}),
                    ...(style.line.dashType ? { dashType: style.line.dashType } : {}),
                  },
                }
              : {}),
            ...(style?.cornerRadius !== undefined ? { cornerRadius: style.cornerRadius } : {}),
            ...(style?.rotate !== undefined ? { rotate: style.rotate } : {}),
            layer: "decoration",
            name: this.name("shape"),
          },
        );
        return;
      }

      case "image": {
        const width = node.width ?? box.w;
        const x =
          node.align === "center"
            ? box.x + (box.w - width) / 2
            : node.align === "right"
              ? box.x + box.w - width
              : box.x;
        this.ctx.image(
          { x, y: box.y, w: width, h: box.h },
          {
            assetId: this.options.assetPrefix
              ? `${this.options.assetPrefix}${node.assetId}`
              : node.assetId,
            ...(node.sizing ? { sizing: node.sizing } : {}),
            layer: "decoration",
            name: this.name("image"),
          },
        );
        return;
      }

      case "card": {
        const inner = drawCard(this.ctx, box, {
          ...(color(this.ctx, node.fill) ? { fill: color(this.ctx, node.fill)! } : {}),
          ...(color(this.ctx, node.border) ? { border: color(this.ctx, node.border)! } : {}),
          ...(color(this.ctx, node.accent) ? { accent: color(this.ctx, node.accent)! } : {}),
          name: this.name("card"),
        });
        this.render(node.child, inner, scope);
        return;
      }

      case "stack": {
        const pad = spacing(this.ctx, node.padding);
        const gap = spacing(this.ctx, node.gap);
        const area = inset(box, pad);
        const children = node.children.filter((child) => visible(child, scope));
        if (children.length === 0) return;

        const measured = children.map((child) => this.measure(child, area.w, scope));
        const fixed = measured.reduce<number>((sum, m) => sum + (m === "fill" ? 0 : m), 0);
        const fillCount = measured.filter((m) => m === "fill").length;
        const gaps = Math.max(0, children.length - 1) * gap;
        const leftover = Math.max(0, area.h - fixed - gaps);
        const perFill = fillCount > 0 ? leftover / fillCount : 0;

        let cursor = area.y;
        children.forEach((child, i) => {
          const height = measured[i] === "fill" ? perFill : (measured[i] as number);
          this.render(child, { x: area.x, y: cursor, w: area.w, h: height }, scope);
          cursor += height + gap;
        });
        return;
      }

      case "columns": {
        const gap = spacing(this.ctx, node.gap);
        if (node.repeat && node.item) {
          const items = repeatItems(node, scope);
          if (items.length === 0) return;
          const rects = splitColumns(box, items.length, gap);
          items.forEach((item, index) => {
            this.render(node.item!, rects[index] as Rect, { ...scope, item, index });
          });
          return;
        }
        const children = (node.children ?? []).filter((child) => visible(child, scope));
        if (children.length === 0) return;
        const rects = node.weights
          ? columnsWeighted(box, node.weights.slice(0, children.length), gap)
          : splitColumns(box, children.length, gap);
        children.forEach((child, i) => {
          this.render(child, rects[i] as Rect, scope);
        });
        return;
      }

      case "absolute": {
        const origin = node.origin ?? "canvas";
        for (const child of node.children) {
          const target: Rect =
            origin === "canvas"
              ? child.box
              : {
                  x: box.x + child.box.x,
                  y: box.y + child.box.y,
                  w: child.box.w,
                  h: child.box.h,
                };
          this.render(child.node, target, scope);
        }
        return;
      }
    }
  }
}

/**
 * Declarative Pattern の layout を描画する。
 *
 * header / footer は builtin Pattern と同じ共有部品を使うので、
 * 「どの Pattern でも見出しが同じ位置に来る」という一貫性が保たれる。
 */
export function renderDeclarativeLayout(
  layout: DeclarativeLayout,
  bindings: Record<string, unknown>,
  ctx: RenderContext,
  options: DeclarativeRenderOptions = {},
): void {
  const scope: BindingScope = { bindings };
  let area: Rect = ctx.frame;

  if (layout.header) {
    const headline = resolveValueRef(layout.header.headline, scope);
    if (headline === undefined) {
      throw new DeclarativeRenderError("header.headline が解決できません（bindings が不足しています）");
    }
    const eyebrow = layout.header.eyebrow ? resolveValueRef(layout.header.eyebrow, scope) : undefined;
    const subline = layout.header.subline ? resolveValueRef(layout.header.subline, scope) : undefined;
    area = drawHeader(
      ctx,
      {
        ...(eyebrow ? { eyebrow } : {}),
        headline,
        ...(subline ? { subline } : {}),
      },
      area,
    );
  }

  if (layout.footer) {
    const text = resolveValueRef(layout.footer.text, scope);
    area = drawFooter(ctx, area, text);
  }

  new DeclarativeRenderer(ctx, options).render(layout.body, area, scope);
}
