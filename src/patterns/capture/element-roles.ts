import { hasRenderableExtent, isLineElement, type ExtractedElement } from "../../pptx/model.js";

/**
 * 「この要素をこの役割として描画できるか」の客観的な判定。
 *
 * ここは **Interpretation Validator と Pattern Compiler が共有する唯一の契約**。
 * 別々に条件を書くと「Validator は通したのに Compiler が描けない」
 * （= 解釈が黙って捨てられる）状態が生まれるので、必ずこのモジュールを使う。
 *
 * 判定するのはあくまで *描画可能性* であって、意味の妥当性ではない。
 *  - 「要素 6 は text なので label として描ける」        → ここで判定する
 *  - 「要素 6 は label らしい内容か」                      → 判定しない（Interpretation の責務）
 *
 * だから "30%" を label と解釈しても、text である以上ここは通す。
 * それが不自然かどうかを決めるのは human / heuristic / 将来の AI 側。
 */

/** 塗りで描ける図形か（線ではなく、面積を持ち、solid fill を持つ） */
export function isSolidFillShape(element: ExtractedElement): boolean {
  return (
    !isLineElement(element) &&
    hasRenderableExtent(element) &&
    element.fill?.type === "solid"
  );
}

/**
 * 線として描ける要素か。
 *
 * Phase 1 の契約どおり、水平線は h = 0、垂直線は w = 0 を許す
 * （`hasRenderableExtent` が線だけ別契約になっている）。
 */
export function isRenderableLine(element: ExtractedElement): boolean {
  return isLineElement(element) && hasRenderableExtent(element);
}

/** text として描けるか（headline / label / metric / description / index 用） */
export function canRenderAsText(element: ExtractedElement): boolean {
  return (element.text?.plain.trim().length ?? 0) > 0;
}

/** カード背景として描けるか（Compiler は card の fill / border として使う） */
export function canRenderAsBackground(element: ExtractedElement): boolean {
  return isSolidFillShape(element);
}

/** アクセントとして描けるか（Compiler は塗り図形か線として描く） */
export function canRenderAsAccent(element: ExtractedElement): boolean {
  return isSolidFillShape(element) || isRenderableLine(element);
}

/** 区切り線として描けるか（Compiler は必ず line として描く） */
export function canRenderAsDivider(element: ExtractedElement): boolean {
  return isRenderableLine(element);
}

/**
 * Template が面倒を見る装飾プレースホルダか（ページ番号・フッター・日付）。
 *
 * これらは「情報をどう見せるか」ではなく企業テンプレートの責務なので、
 * Pattern の中へ取り込まない（CLAUDE.md の Template / Pattern 責務分離）。
 * PowerPoint の placeholder type を見るだけの客観的な判定で、意味は推測しない。
 */
export const TEMPLATE_PLACEHOLDER_TYPES = ["sldNum", "ftr", "dt"] as const;

export function isTemplatePlaceholder(element: ExtractedElement): boolean {
  const type = element.placeholder?.type;
  return type !== undefined && (TEMPLATE_PLACEHOLDER_TYPES as readonly string[]).includes(type);
}

/** エラーメッセージ用の短い要素の説明（"text \"見出し…\"" / "shape(roundRect)" など） */
export function describeElement(element: ExtractedElement): string {
  const kind = isLineElement(element) ? "line" : element.kind;
  const detail =
    kind === "shape" && element.geometry ? `(${element.geometry})` : "";
  const text = element.text?.plain.replace(/\s+/g, " ").trim();
  const quoted = text ? ` "${text.length > 20 ? `${text.slice(0, 20)}…` : text}"` : "";
  return `${kind}${detail}${quoted}`;
}
