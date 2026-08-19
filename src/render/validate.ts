import { contains, estimateTextHeight, intersects, type Rect } from "./geometry.js";
import type { Placement } from "./surface.js";

export type LayoutIssueCode =
  | "out-of-frame"
  | "off-canvas"
  | "overlap"
  | "text-overflow"
  | "text-shrink-below-min";

/**
 * これ未満に縮小されたテキストは「生成できたが実用的に読めない」とみなす。
 * 会議室投影を想定した下限。ValidateOptions.minEffectiveFontPt で変更可。
 */
export const DEFAULT_MIN_EFFECTIVE_FONT_PT = 9;

export interface LayoutIssue {
  code: LayoutIssueCode;
  severity: "error" | "warning";
  slideId: string;
  slideNumber: number;
  patternId: string;
  message: string;
  elements: string[];
}

export interface ValidateOptions {
  slideId: string;
  slideNumber: number;
  patternId: string;
  canvas: Rect;
  frame: Rect;
  /** 重なり・はみ出し判定の許容誤差 (inch) */
  epsilon?: number;
  /** shrink 後の実効フォントサイズの下限 (pt)。既定 9pt。 */
  minEffectiveFontPt?: number;
}

/**
 * shrink（自動縮小）が有効なテキストが、実際には何 pt まで縮むかを概算する。
 * フォントを k 倍すると 1 行に入る文字数も 1/k になるため、
 * 倍率を段階的に下げながら折り返し高さを再計算して収まる倍率を探す。
 */
export function estimateEffectiveFontSize(
  text: string,
  box: Rect,
  fontSize: number,
  lineHeight: number,
  paddingPt: number,
): number {
  for (let k = 1; k >= 0.3; k -= 0.05) {
    const size = fontSize * k;
    const needed = estimateTextHeight(text, {
      widthIn: box.w,
      fontSize: size,
      lineHeight,
      horizontalPaddingPt: paddingPt * 2,
      verticalPaddingPt: paddingPt * 2,
    });
    if (needed <= box.h + 0.01) return size;
  }
  return fontSize * 0.3;
}

/**
 * 配置結果の基本検査。
 * - content レイヤ同士の重なり
 * - frame / canvas からのはみ出し
 * - テキストの推定高さがボックス高さを超えていないか
 *
 * decoration / background レイヤはカード背景など「重なることが前提」の要素なので
 * 重なり検査の対象外。ただしキャンバス外への逸脱は検査する。
 */
export function validatePlacements(
  placements: Placement[],
  opts: ValidateOptions,
): LayoutIssue[] {
  const eps = opts.epsilon ?? 0.02;
  const issues: LayoutIssue[] = [];
  const base = {
    slideId: opts.slideId,
    slideNumber: opts.slideNumber,
    patternId: opts.patternId,
  };

  for (const p of placements) {
    if (!contains(opts.canvas, p.box, eps)) {
      issues.push({
        ...base,
        code: "off-canvas",
        severity: "error",
        message: `要素がスライド外にはみ出しています (${fmt(p.box)})`,
        elements: [p.name],
      });
      continue;
    }
    if (p.layer !== "background" && !contains(opts.frame, p.box, eps)) {
      issues.push({
        ...base,
        code: "out-of-frame",
        severity: "warning",
        message: `要素が安全領域からはみ出しています (${fmt(p.box)})`,
        elements: [p.name],
      });
    }
  }

  const content = placements.filter((p) => p.layer === "content");
  for (let i = 0; i < content.length; i += 1) {
    for (let j = i + 1; j < content.length; j += 1) {
      const a = content[i]!;
      const b = content[j]!;
      if (intersects(a.box, b.box, eps)) {
        issues.push({
          ...base,
          code: "overlap",
          severity: "error",
          message: `要素同士が重なっています: ${a.name} × ${b.name}`,
          elements: [a.name, b.name],
        });
      }
    }
  }

  const minFont = opts.minEffectiveFontPt ?? DEFAULT_MIN_EFFECTIVE_FONT_PT;
  for (const p of placements) {
    if (p.kind !== "text" || !p.text) continue;

    if (p.shrink) {
      // shrink 有効なテキストは「はみ出さない」が「読めないほど縮む」可能性がある。
      // 実効フォントサイズを推定し、下限未満なら警告する。
      const effective = estimateEffectiveFontSize(
        p.text,
        p.box,
        p.fontSize ?? 12,
        p.lineHeight ?? 1.3,
        p.paddingPt ?? 0,
      );
      if (effective < minFont) {
        issues.push({
          ...base,
          code: "text-shrink-below-min",
          severity: "warning",
          message: `テキストが読める下限 (${minFont}pt) を下回る約 ${effective.toFixed(1)}pt まで縮小される可能性があります: "${truncate(p.text)}"`,
          elements: [p.name],
        });
      }
      continue;
    }

    const needed = estimateTextHeight(p.text, {
      widthIn: p.box.w,
      fontSize: p.fontSize ?? 12,
      lineHeight: p.lineHeight ?? 1.3,
      horizontalPaddingPt: (p.paddingPt ?? 0) * 2,
      verticalPaddingPt: (p.paddingPt ?? 0) * 2,
    });
    if (needed > p.box.h + eps) {
      issues.push({
        ...base,
        code: "text-overflow",
        severity: "warning",
        message: `テキストがボックスに収まらない可能性があります (必要 ${needed.toFixed(2)}in / 実際 ${p.box.h.toFixed(2)}in): "${truncate(p.text)}"`,
        elements: [p.name],
      });
    }
  }

  return issues;
}

function fmt(r: Rect): string {
  return `x=${r.x.toFixed(2)} y=${r.y.toFixed(2)} w=${r.w.toFixed(2)} h=${r.h.toFixed(2)}`;
}

function truncate(text: string, max = 32): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > max ? `${flat.slice(0, max)}…` : flat;
}
