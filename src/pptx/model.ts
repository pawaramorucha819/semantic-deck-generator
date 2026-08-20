import type { Rect } from "../render/geometry.js";

/**
 * PPTX から取り出した「中立な」スライド表現。
 *
 * ここは PowerPoint の語彙をそのまま持つ層で、Semantic Deck Generator の
 * Pattern / Template とはまだ無関係。変換はこの上（importer / capture）で行う。
 */

export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;

export const emuToInch = (emu: number): number => emu / EMU_PER_INCH;
export const emuToPoint = (emu: number): number => emu / EMU_PER_POINT;

export interface ExtractedRun {
  text: string;
  /** pt */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  /** RRGGBB */
  color?: string;
  fontFace?: string;
}

export interface ExtractedText {
  /** 段落ごとの文字列 */
  paragraphs: string[];
  runs: ExtractedRun[];
  /** 改行込みの平文 */
  plain: string;
  align?: "left" | "center" | "right";
  anchor?: "top" | "middle" | "bottom";
  /** 最大のフォントサイズ (pt)。見出し推定に使う。 */
  maxFontSize?: number;
}

export interface ExtractedFill {
  type: "solid" | "none" | "unsupported";
  /** RRGGBB (type === 'solid') */
  color?: string;
  /** 非対応の塗り（グラデーション・画像など）の種類名 */
  detail?: string;
}

export interface ExtractedElement {
  id: string;
  name?: string;
  kind: "text" | "shape" | "line" | "image" | "placeholder";
  /** inch。スライド左上原点。 */
  box: Rect;
  rotation?: number;
  geometry?: string;
  text?: ExtractedText;
  fill?: ExtractedFill;
  line?: { color?: string; widthPt?: number };
  image?: { relationshipId: string; part?: string; fileName?: string; externalTarget?: string };
  placeholder?: { type?: string; index?: string };
  /** グループから展開された要素か */
  fromGroup?: boolean;
}

export interface ExtractedSlide {
  slideNumber: number;
  part: string;
  elements: ExtractedElement[];
  /** 対応していない PowerPoint 機能（silent failure を作らないため必ず記録する） */
  unsupported: string[];
  warnings: string[];
  /** スピーカーノートの平文 */
  notes?: string;
  /** レイアウト由来のタイトル文字列（あれば） */
  title?: string;
}

export interface ExtractedTheme {
  /** clrScheme（dk1 / lt1 / accent1 ... を RRGGBB で） */
  colors: Record<string, string>;
  fonts: { major?: string; minor?: string; majorEastAsian?: string; minorEastAsian?: string };
}

export interface ExtractedDeck {
  slideSize: { widthIn: number; heightIn: number };
  theme?: ExtractedTheme;
  slides: ExtractedSlide[];
  unsupported: string[];
  warnings: string[];
  /** docProps/core.xml から取れたメタデータ */
  metadata: { title?: string; creator?: string; company?: string; created?: string; modified?: string };
}
