/**
 * Presentation Template — 「企業としての紙面の共通構造」を担当する層。
 *
 * Theme と Template の責務は明確に違います。
 *
 *   Theme    … 色・書体・サイズ・余白などの *デザイントークン*
 *   Template … Theme + Slide Master（ロゴ・フッター・ページ番号・固定装飾）
 *              + ブランドアセット + スライドサイズ
 *
 * Pattern の中に会社ロゴやフッターを書いてはいけません。それは Template の責務です。
 * 逆に Template は「情報をどう見せるか」を知りません。それは Pattern の責務です。
 */

import type { PresentationTheme, ResolvedTheme } from "../theme/theme.js";
import { resolveTheme } from "../theme/theme.js";
import type { ColorToken, FontRole, TextRole } from "../render/context.js";
import type { Rect } from "../render/geometry.js";

export const TEMPLATE_SCHEMA_VERSION = "1.0";

/** Slide Master の役割。Renderer はスライドの種類からこの役割を選ぶ。 */
export const MASTER_ROLES = ["title", "section", "content"] as const;
export type MasterRole = (typeof MASTER_ROLES)[number];

/** テンプレートに同梱される画像アセット（ロゴなど） */
export interface TemplateAsset {
  id: string;
  /** テンプレートディレクトリからの相対パス（ディレクトリ外は参照できない） */
  path: string;
  kind: "image";
  description?: string;
}

/** 解決済みアセット（data URI 化されたもの） */
export interface LoadedTemplateAsset extends TemplateAsset {
  /** 'image/png;base64,...' */
  data: string;
}

export interface TemplateTextStyle {
  role?: TextRole;
  fontSize?: number;
  font?: FontRole;
  color?: ColorToken;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  lineHeight?: number;
  charSpacing?: number;
}

export interface TemplateShapeStyle {
  fill?: ColorToken;
  line?: { color: ColorToken; width?: number; dashType?: "solid" | "dash" | "lgDash" | "sysDot" };
  transparency?: number;
}

/**
 * Master に載せられる図形。
 * PptxGenJS の Slide Master がサポートするのは rect / line / text / image のみなので、
 * 装飾もこの範囲に限定する（角丸カード等はスライド本体＝Pattern の責務）。
 */
export type TemplateShapeKind = "rect" | "line";

/**
 * Slide Master 上の固定要素。
 *
 * 位置はスライド全面（canvas）を基準とした inch。ここに書けるのは
 * 「毎ページ同じもの」だけで、スライドごとに変わる内容は Pattern が描く。
 */
export type TemplateElement =
  | {
      kind: "text";
      box: Rect;
      /** `{{deckTitle}}` `{{deckSubtitle}}` `{{author}}` `{{organization}}` のみ展開される */
      text: string;
      style?: TemplateTextStyle;
    }
  | {
      kind: "shape";
      shape: TemplateShapeKind;
      box: Rect;
      style?: TemplateShapeStyle;
    }
  | {
      kind: "image";
      assetId: string;
      box: Rect;
      sizing?: "contain" | "cover" | "stretch";
      transparency?: number;
    };

/** ページ番号の配置（PowerPoint のページ番号フィールドとして埋め込まれる） */
export interface SlideNumberDefinition {
  box: Rect;
  style?: TemplateTextStyle;
}

export interface SlideMasterDefinition {
  id: string;
  role: MasterRole;
  name?: string;
  background?: { color: ColorToken };
  /**
   * 本文（Pattern）が描画できる領域の余白 (inch)。
   * 省略時は Theme の margins が使われる。ロゴやフッターと衝突させないための宣言。
   */
  contentMargins?: { top: number; right: number; bottom: number; left: number };
  elements?: TemplateElement[];
  slideNumber?: SlideNumberDefinition;
}

export interface TemplateMetadata {
  organization?: string;
  department?: string;
  description?: string;
  /** 機密表示など、Master のテキストから参照できる任意の語 */
  confidentialityLabel?: string;
}

/** Import 等で「対応できなかった」ものの明示的な記録（silent failure を作らない） */
export interface TemplateImportNotes {
  /** 取り込み元 */
  importedFrom?: string;
  /** 対応していない PowerPoint 機能 */
  unsupportedFeatures?: string[];
  warnings?: string[];
}

export interface PresentationTemplate {
  schemaVersion: string;
  id: string;
  version: string;
  name: string;

  /** この Template が既定で使う Theme（--theme で上書き可能） */
  theme: PresentationTheme;

  /** スライドサイズ (inch)。省略時 10 x 5.625 (16:9)。 */
  slideSize?: { widthIn: number; heightIn: number };

  masters: SlideMasterDefinition[];
  defaultMasterId: string;
  titleMasterId?: string;
  sectionMasterId?: string;

  assets?: TemplateAsset[];
  metadata?: TemplateMetadata;
  import?: TemplateImportNotes;
}

/** 既定のスライドサイズ (LAYOUT_16x9) */
export const DEFAULT_SLIDE_WIDTH_IN = 10;
export const DEFAULT_SLIDE_HEIGHT_IN = 5.625;

export interface ResolvedTemplate {
  template: PresentationTemplate;
  id: string;
  theme: ResolvedTheme;
  slideSize: { widthIn: number; heightIn: number };
  assets: LoadedTemplateAsset[];
}

export interface ResolveTemplateOptions {
  /** Template の theme を上書きする（CLI の --theme 相当） */
  theme?: PresentationTheme;
  assets?: LoadedTemplateAsset[];
}

export function resolveTemplate(
  template: PresentationTemplate,
  options: ResolveTemplateOptions = {},
): ResolvedTemplate {
  return {
    template,
    id: template.id,
    theme: resolveTheme(options.theme ?? template.theme),
    slideSize: template.slideSize ?? {
      widthIn: DEFAULT_SLIDE_WIDTH_IN,
      heightIn: DEFAULT_SLIDE_HEIGHT_IN,
    },
    assets: options.assets ?? [],
  };
}

/** 役割に対応する Master を選ぶ。宣言されていなければ default master にフォールバックする。 */
export function masterForRole(
  template: PresentationTemplate,
  role: MasterRole,
): SlideMasterDefinition {
  const explicitId =
    role === "title"
      ? template.titleMasterId
      : role === "section"
        ? template.sectionMasterId
        : template.defaultMasterId;

  const byExplicit = explicitId
    ? template.masters.find((m) => m.id === explicitId)
    : undefined;
  if (byExplicit) return byExplicit;

  const byRole = template.masters.find((m) => m.role === role);
  if (byRole) return byRole;

  const fallback = template.masters.find((m) => m.id === template.defaultMasterId);
  if (!fallback) {
    throw new Error(
      `Template '${template.id}': defaultMasterId '${template.defaultMasterId}' に対応する master がありません`,
    );
  }
  return fallback;
}

/** Master の contentMargins（無ければ Theme の margins）から本文領域を求める */
export function contentFrameOf(
  master: SlideMasterDefinition,
  theme: ResolvedTheme,
  slideSize: { widthIn: number; heightIn: number },
): Rect {
  const m = master.contentMargins ?? theme.margins;
  return {
    x: m.left,
    y: m.top,
    w: Math.max(0, slideSize.widthIn - m.left - m.right),
    h: Math.max(0, slideSize.heightIn - m.top - m.bottom),
  };
}

/** Master のテキストで展開できる変数。決定論性を保つため日付等の可変値は含めない。 */
export interface TemplateTextVariables {
  deckTitle: string;
  deckSubtitle?: string;
  author?: string;
  organization?: string;
  department?: string;
  confidentiality?: string;
}

const VARIABLE_PATTERN = /\{\{\s*([a-zA-Z][a-zA-Z0-9]*)\s*\}\}/g;

export const TEMPLATE_TEXT_VARIABLES = [
  "deckTitle",
  "deckSubtitle",
  "author",
  "organization",
  "department",
  "confidentiality",
] as const;

/**
 * Master テキストの変数展開。
 *
 * 展開できるのは `TEMPLATE_TEXT_VARIABLES` に列挙された固定語のみ。
 * 未知の変数は例外にする（テンプレートの typo が黙って空文字になるのを防ぐ）。
 */
export function expandTemplateText(text: string, vars: TemplateTextVariables): string {
  return text.replace(VARIABLE_PATTERN, (_match, name: string) => {
    if (!(TEMPLATE_TEXT_VARIABLES as readonly string[]).includes(name)) {
      throw new Error(
        `Template のテキストに未知の変数 {{${name}}} が含まれています (利用可能: ${TEMPLATE_TEXT_VARIABLES.join(", ")})`,
      );
    }
    return (vars as unknown as Record<string, string | undefined>)[name] ?? "";
  });
}
