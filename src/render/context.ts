import type { ResolvedTheme } from "../theme/theme.js";
import type { Rect } from "./geometry.js";

/**
 * RenderContext — Pattern が「描く」ために触れてよい唯一の API。
 *
 * Pattern は PptxGenJS を直接 import しません。理由は 3 つ:
 *  1. 描画バックエンドを差し替えられる（将来 SVG プレビュー等）
 *  2. 配置された矩形をすべて記録できる → はみ出し / 重なり検査が可能
 *  3. 色やフォントを *トークン* でしか指定できない → Theme 分離が構造的に守られる
 */

export type Layer = "background" | "content" | "decoration";

export type TextRole = "display" | "keyMessage" | "heading" | "body" | "caption";
export type FontRole = "heading" | "body" | "numeric";

/**
 * 色の指定。Theme のトークン名（'primary' 等）を推奨。
 * どうしてもリテラルを使う場合は '#RRGGBB' 形式（Pattern では原則禁止）。
 */
export type ColorToken = keyof ResolvedTheme["colors"] | `#${string}`;

export interface TextStyle {
  role?: TextRole;
  /** pt。指定時は role のサイズを上書き。 */
  fontSize?: number;
  font?: FontRole;
  color?: ColorToken;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  /** 行送り倍率（1.0 = シングル） */
  lineHeight?: number;
  charSpacing?: number;
  /** テキストボックス内側の余白 (pt) */
  paddingPt?: number;
  wrap?: boolean;
  /** はみ出したら PowerPoint 側で縮小する */
  shrink?: boolean;
  layer?: Layer;
  /** PowerPoint 上のオブジェクト名（デバッグ用） */
  name?: string;
}

export interface TextRun {
  text: string;
  style?: Omit<TextStyle, "layer" | "name" | "valign" | "align">;
  breakLine?: boolean;
}

export type TextContent = string | TextRun[];

export type ShapeKind =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "line"
  | "rightArrow"
  | "chevron"
  | "triangle"
  | "diamond"
  | "downArrow";

export interface ShapeStyle {
  fill?: ColorToken | "none";
  fillTransparency?: number;
  line?: {
    color: ColorToken;
    /** pt */
    width?: number;
    dashType?: "solid" | "dash" | "lgDash" | "sysDot";
    /** Native editable connector direction, used for semantic relationships. */
    endArrowType?: "triangle";
  };
  /** 0.0 - 1.0（roundRect のみ） */
  cornerRadius?: number;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  layer?: Layer;
  name?: string;
}

/**
 * 画像の描画指定。
 *
 * URL やファイルパスは受け付けない。Template / Pattern パッケージに同梱され、
 * `AssetResolver` へ登録済みの `assetId` しか参照できない
 * （= 外部 URL の無条件 fetch を構造的に禁止する）。
 */
export interface ImageSpec {
  assetId: string;
  /** 既定 'contain'（アスペクト比を保って box 内に収める） */
  sizing?: "contain" | "cover" | "stretch";
  /** 0-100 (%) */
  transparency?: number;
  layer?: Layer;
  name?: string;
}

/** assetId → 埋め込み可能な画像データ。Template / Pattern パッケージがロード時に構築する。 */
export interface ResolvedAsset {
  id: string;
  /** data URI ('image/png;base64,....')。PptxGenJS の `data` にそのまま渡せる形。 */
  data: string;
  /** 元ファイル名（デバッグ用） */
  fileName?: string;
}

export interface AssetResolver {
  resolve(assetId: string): ResolvedAsset;
  has(assetId: string): boolean;
}

export interface ChartSeries {
  name: string;
  values: number[];
}

export interface ChartSpec {
  kind: "bar" | "line" | "pie" | "doughnut";
  categories: string[];
  series: ChartSeries[];
  /** 系列色（Theme トークン） */
  colors?: ColorToken[];
  /** 'col' = 縦棒 / 'bar' = 横棒（kind='bar' のとき有効） */
  barDir?: "col" | "bar";
  showValue?: boolean;
  showLegend?: boolean;
  valueUnit?: string;
  layer?: Layer;
  name?: string;
}

export interface RenderContext {
  readonly theme: ResolvedTheme;
  /** スライド全面（inch） */
  readonly canvas: Rect;
  /** 余白を除いた作図可能領域（inch）。Pattern はこの中に収めること。 */
  readonly frame: Rect;
  readonly variantId: string;
  readonly slideNumber: number;
  readonly totalSlides: number;
  /** Optional renderer-owned composition of visible explanation and diagram body. */
  layoutBody?(area: Rect): Rect;

  text(box: Rect, content: TextContent, style?: TextStyle): void;
  shape(kind: ShapeKind, box: Rect, style?: ShapeStyle): void;
  chart(box: Rect, spec: ChartSpec): void;
  /**
   * 登録済みアセット（ロゴ等）を描画する。
   * アセットが未登録の場合は例外になる（silent failure を作らない）。
   */
  image(box: Rect, spec: ImageSpec): void;

  /** テキストが必要とする高さ (inch) を概算する。レイアウト計算用。 */
  measureText(
    content: TextContent,
    opts: { widthIn: number; role?: TextRole; fontSize?: number; lineHeight?: number; paddingPt?: number },
  ): number;

  /** role → フォントサイズ (pt) */
  fontSizeOf(role: TextRole): number;

  /**
   * Theme のトークンを淡くした色を得る（amount: 0 = そのまま, 1 = 背景色）。
   * 「3 番目の要素は 1 番目より淡く」のような *構造上の強弱* を、
   * 色そのものをハードコードせずに表現するための口。
   */
  tint(token: ColorToken, amount: number): ColorToken;
}
