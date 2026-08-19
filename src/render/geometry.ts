/** 単位はすべて inch（PptxGenJS の座標系と同じ） */
export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Point {
  x: number;
  y: number;
}

export const rect = (x: number, y: number, w: number, h: number): Rect => ({ x, y, w, h });

export const right = (r: Rect): number => r.x + r.w;
export const bottom = (r: Rect): number => r.y + r.h;
export const centerX = (r: Rect): number => r.x + r.w / 2;
export const centerY = (r: Rect): number => r.y + r.h / 2;

/** 四辺を内側へ縮める */
export function inset(r: Rect, amount: number | Partial<Record<"top" | "right" | "bottom" | "left", number>>): Rect {
  const i =
    typeof amount === "number"
      ? { top: amount, right: amount, bottom: amount, left: amount }
      : { top: 0, right: 0, bottom: 0, left: 0, ...amount };
  return {
    x: r.x + i.left,
    y: r.y + i.top,
    w: Math.max(0, r.w - i.left - i.right),
    h: Math.max(0, r.h - i.top - i.bottom),
  };
}

/** 上から height 分を切り出し、[切り出した領域, 残り] を返す */
export function splitTop(r: Rect, height: number, gap = 0): [Rect, Rect] {
  const h = Math.min(height, r.h);
  return [
    { x: r.x, y: r.y, w: r.w, h },
    { x: r.x, y: r.y + h + gap, w: r.w, h: Math.max(0, r.h - h - gap) },
  ];
}

/** 下から height 分を切り出し、[切り出した領域, 残り] を返す */
export function splitBottom(r: Rect, height: number, gap = 0): [Rect, Rect] {
  const h = Math.min(height, r.h);
  return [
    { x: r.x, y: bottom(r) - h, w: r.w, h },
    { x: r.x, y: r.y, w: r.w, h: Math.max(0, r.h - h - gap) },
  ];
}

/** 横に count 等分し、間に gap を空ける */
export function columns(r: Rect, count: number, gap: number): Rect[] {
  if (count <= 0) return [];
  const w = (r.w - gap * (count - 1)) / count;
  return Array.from({ length: count }, (_, i) => ({
    x: r.x + i * (w + gap),
    y: r.y,
    w,
    h: r.h,
  }));
}

/** 縦に count 等分し、間に gap を空ける */
export function rows(r: Rect, count: number, gap: number): Rect[] {
  if (count <= 0) return [];
  const h = (r.h - gap * (count - 1)) / count;
  return Array.from({ length: count }, (_, i) => ({
    x: r.x,
    y: r.y + i * (h + gap),
    w: r.w,
    h,
  }));
}

/** 与えられた重み比で横分割する */
export function columnsWeighted(r: Rect, weights: number[], gap: number): Rect[] {
  const total = weights.reduce((a, b) => a + b, 0);
  if (total <= 0) return [];
  const usable = r.w - gap * (weights.length - 1);
  let cursor = r.x;
  return weights.map((weight) => {
    const w = (usable * weight) / total;
    const out = { x: cursor, y: r.y, w, h: r.h };
    cursor += w + gap;
    return out;
  });
}

export function intersects(a: Rect, b: Rect, epsilon = 0.02): boolean {
  return (
    a.x < right(b) - epsilon &&
    right(a) > b.x + epsilon &&
    a.y < bottom(b) - epsilon &&
    bottom(a) > b.y + epsilon
  );
}

export function contains(outer: Rect, inner: Rect, epsilon = 0.02): boolean {
  return (
    inner.x >= outer.x - epsilon &&
    inner.y >= outer.y - epsilon &&
    right(inner) <= right(outer) + epsilon &&
    bottom(inner) <= bottom(outer) + epsilon
  );
}

export function roundRectValues(r: Rect, digits = 4): Rect {
  const f = (n: number) => Number(n.toFixed(digits));
  return { x: f(r.x), y: f(r.y), w: f(r.w), h: f(r.h) };
}

const POINTS_PER_INCH = 72;

/** 全角（CJK）判定 */
function isWide(ch: string): boolean {
  const code = ch.codePointAt(0) ?? 0;
  return (
    (code >= 0x1100 && code <= 0x115f) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7a3) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe6f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  );
}

/** 全角=1.0em / 半角=0.52em として文字列の幅を em 単位で概算する */
export function measureEm(text: string): number {
  let em = 0;
  for (const ch of text) {
    em += isWide(ch) ? 1.0 : 0.52;
  }
  return em;
}

export interface TextMetricOptions {
  /** テキストボックスの幅（inch） */
  widthIn: number;
  /** pt */
  fontSize: number;
  /** 行送り係数 */
  lineHeight?: number;
  /** 左右の内側余白 合計 (pt)。PptxGenJS の margin 相当。 */
  horizontalPaddingPt?: number;
  /** 上下の内側余白 合計 (pt) */
  verticalPaddingPt?: number;
}

/**
 * テキストの折り返し行数を概算する。
 * 正確な PowerPoint のレイアウト計算ではなく、はみ出し検出のための近似。
 */
export function estimateLineCount(text: string, opts: TextMetricOptions): number {
  const usableWidthPt = opts.widthIn * POINTS_PER_INCH - (opts.horizontalPaddingPt ?? 0);
  if (usableWidthPt <= 0) return Number.POSITIVE_INFINITY;
  const emPerLine = usableWidthPt / opts.fontSize;
  let lines = 0;
  for (const paragraph of text.split("\n")) {
    const em = measureEm(paragraph);
    lines += Math.max(1, Math.ceil(em / Math.max(emPerLine, 0.001)));
  }
  return Math.max(1, lines);
}

/** テキストが必要とする高さ（inch）を概算する */
export function estimateTextHeight(text: string, opts: TextMetricOptions): number {
  const lines = estimateLineCount(text, opts);
  const lineHeight = opts.lineHeight ?? 1.32;
  const contentPt = lines * opts.fontSize * lineHeight;
  return (contentPt + (opts.verticalPaddingPt ?? 0)) / POINTS_PER_INCH;
}

export const inchesToPoints = (inches: number): number => inches * POINTS_PER_INCH;
export const pointsToInches = (points: number): number => points / POINTS_PER_INCH;
