import type { ColorToken, RenderContext } from "../../render/context.js";
import { bottom, inset, measureEm, type Rect } from "../../render/geometry.js";

/**
 * 全 Pattern が共有するレイアウト部品。
 * 「どの Pattern でも同じ位置に見出しが来る」という一貫性はここで担保します。
 * Pattern 固有のレイアウトは各 Pattern の render() が持ちます。
 */

export interface HeaderBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
}

/**
 * 見出しブロックを描画し、残りの本文領域を返す。
 * `area` を渡すとその領域内に描く（既定はスライドの作図可能領域）。
 */
export function drawHeader(ctx: RenderContext, header: HeaderBindings, area?: Rect): Rect {
  const { theme } = ctx;
  const frame = area ?? ctx.frame;
  let cursor = frame.y;

  if (header.eyebrow) {
    const h = 0.2;
    ctx.shape("rect", { x: frame.x, y: cursor + 0.045, w: 0.055, h: h - 0.08 }, {
      fill: "primary",
      layer: "decoration",
      name: "eyebrow-mark",
    });
    ctx.text(
      { x: frame.x + 0.055 + theme.spacing.sm, y: cursor, w: frame.w - 0.6, h },
      header.eyebrow,
      {
        role: "caption",
        color: "primary",
        bold: true,
        charSpacing: 1,
        valign: "middle",
        layer: "content",
        name: "eyebrow",
      },
    );
    cursor += h + theme.spacing.xs;
  }

  const headlineHeight = Math.max(
    0.42,
    ctx.measureText(header.headline, { widthIn: frame.w, role: "keyMessage", lineHeight: 1.24 }),
  );
  ctx.text({ x: frame.x, y: cursor, w: frame.w, h: headlineHeight }, header.headline, {
    role: "keyMessage",
    color: "text",
    bold: true,
    lineHeight: 1.24,
    valign: "middle",
    layer: "content",
    name: "key-message",
  });
  cursor += headlineHeight + theme.spacing.xs;

  // 見出し下の罫線: 左側だけ primary、残りは border 色
  const ruleH = 0.028;
  ctx.shape("rect", { x: frame.x, y: cursor, w: frame.w, h: ruleH }, {
    fill: "border",
    layer: "decoration",
    name: "rule-base",
  });
  ctx.shape("rect", { x: frame.x, y: cursor, w: Math.min(1.3, frame.w * 0.18), h: ruleH }, {
    fill: "primary",
    layer: "decoration",
    name: "rule-accent",
  });
  cursor += ruleH;

  if (header.subline) {
    cursor += theme.spacing.sm;
    const sublineHeight = ctx.measureText(header.subline, {
      widthIn: frame.w,
      role: "body",
      lineHeight: 1.36,
    });
    ctx.text({ x: frame.x, y: cursor, w: frame.w, h: sublineHeight }, header.subline, {
      role: "body",
      color: "mutedText",
      lineHeight: 1.36,
      layer: "content",
      name: "supporting-message",
    });
    cursor += sublineHeight;
  }

  cursor += theme.spacing.md;
  return { x: frame.x, y: cursor, w: frame.w, h: Math.max(0, bottom(frame) - cursor) };
}

/** 出典行をフレーム下端に描画し、残りの本文領域を返す */
export function drawFooter(ctx: RenderContext, body: Rect, sourceText?: string): Rect {
  if (!sourceText) return ctx.layoutBody?.(body) ?? body;
  const h = Math.max(
    0.18,
    ctx.measureText(sourceText, { widthIn: body.w, role: "caption", lineHeight: 1.25 }),
  );
  const y = bottom(body) - h;
  ctx.text({ x: body.x, y, w: body.w, h }, sourceText, {
    role: "caption",
    color: "mutedText",
    valign: "bottom",
    lineHeight: 1.25,
    layer: "content",
    name: "source",
  });
  const remaining = { x: body.x, y: body.y, w: body.w, h: Math.max(0, y - body.y - ctx.theme.spacing.sm) };
  return ctx.layoutBody?.(remaining) ?? remaining;
}

export interface CardStyle {
  fill?: ColorToken;
  border?: ColorToken;
  borderWidth?: number;
  /** 上端のアクセントバー */
  accent?: ColorToken;
  name?: string;
}

/**
 * 1 行で幅に収まる最大のフォントサイズ (pt) を求める。
 * BigNumber 系の「数字を可能な限り大きく」を Theme 非依存で実現するための計算。
 */
export function fitFontSize(
  text: string,
  widthIn: number,
  maxPt: number,
  minPt = 10,
  fillRatio = 0.94,
): number {
  const em = measureEm(text);
  if (em <= 0) return maxPt;
  const byWidth = (widthIn * 72 * fillRatio) / em;
  return Math.max(minPt, Math.min(maxPt, Math.floor(byWidth)));
}

/** ブロックを領域内で垂直中央に置いたときの上端 y を返す */
export function centerBlockY(area: Rect, blockHeight: number): number {
  return area.y + Math.max(0, (area.h - blockHeight) / 2);
}

/** カード（角丸の面）を描き、テキストを置ける内側領域を返す */
export function drawCard(ctx: RenderContext, box: Rect, style: CardStyle = {}): Rect {
  ctx.shape("roundRect", box, {
    fill: style.fill ?? "surface",
    line: style.border ? { color: style.border, width: style.borderWidth ?? 1 } : undefined,
    layer: "decoration",
    name: style.name,
  });
  if (style.accent) {
    ctx.shape("rect", { x: box.x, y: box.y, w: box.w, h: 0.05 }, {
      fill: style.accent,
      layer: "decoration",
      name: style.name ? `${style.name}-accent` : undefined,
    });
  }
  return inset(box, ctx.theme.spacing.md);
}
