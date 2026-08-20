import { mix, normalizeHex, type ResolvedTheme } from "../theme/theme.js";
import type { ColorToken, FontRole, TextRole } from "./context.js";

/**
 * Theme トークンの解決。
 *
 * `SlideSurface`（Pattern 用の描画面）と Template の Slide Master 生成の両方から
 * 使われる。「色・書体・サイズはトークン経由でしか指定できない」という不変条件を
 * 1 箇所に集約するための小さなモジュール。
 */

export function resolveColor(
  theme: ResolvedTheme,
  token: ColorToken | undefined,
  fallback: ColorToken = "text",
): string {
  const t = token ?? fallback;
  if (typeof t === "string" && t.startsWith("#")) return normalizeHex(t);
  const value = theme.colors[t as keyof ResolvedTheme["colors"]];
  if (!value) throw new Error(`未知のカラートークンです: ${String(t)}`);
  return normalizeHex(value);
}

/** 背景色方向へ混色する。ダークテーマでも「淡くなる」意味が保たれる。 */
export function tintToken(
  theme: ResolvedTheme,
  token: ColorToken,
  amount: number,
): ColorToken {
  return `#${mix(resolveColor(theme, token), resolveColor(theme, "background"), amount)}`;
}

export function resolveFontFace(theme: ResolvedTheme, role: FontRole = "body"): string {
  return theme.fonts[role];
}

/** テキストロールに対する既定のフォントロール */
export function defaultFontRoleFor(role: TextRole): FontRole {
  if (role === "display") return "numeric";
  if (role === "keyMessage" || role === "heading") return "heading";
  return "body";
}

/** Theme が持つカラートークン名の一覧（Declarative Pattern / Template の検証に使う） */
export const COLOR_TOKEN_NAMES = [
  "background",
  "text",
  "mutedText",
  "primary",
  "secondary",
  "accent",
  "surface",
  "border",
  "positive",
  "negative",
  "onPrimary",
] as const;

export type ColorTokenName = (typeof COLOR_TOKEN_NAMES)[number];

export const TEXT_ROLE_NAMES = [
  "display",
  "keyMessage",
  "heading",
  "body",
  "caption",
] as const;

export const FONT_ROLE_NAMES = ["heading", "body", "numeric"] as const;

export const SPACING_TOKEN_NAMES = ["xs", "sm", "md", "lg", "xl"] as const;
export type SpacingTokenName = (typeof SPACING_TOKEN_NAMES)[number];

/** spacing トークン名または inch 数値を inch へ解決する */
export function resolveSpacing(
  theme: ResolvedTheme,
  value: SpacingTokenName | number | undefined,
  fallback = 0,
): number {
  if (value === undefined) return fallback;
  if (typeof value === "number") return value;
  return theme.spacing[value];
}
