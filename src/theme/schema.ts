import { z } from "zod";
import type { PresentationTheme } from "./theme.js";

/**
 * Theme の runtime スキーマ。
 *
 * V2 では Theme が外部 JSON（企業テンプレート）から来るため、
 * ビルトイン定数と同じ制約を実行時に強制する必要がある。
 */

const hexColor = z
  .string()
  .trim()
  .regex(/^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/, "カラーコードは #RRGGBB 形式で指定してください");

const positive = z.number().finite().positive();
const nonNegative = z.number().finite().nonnegative();

export const PresentationThemeSchema: z.ZodType<PresentationTheme> = z
  .object({
    id: z.string().trim().min(1).max(80),
    name: z.string().trim().min(1).max(160),
    fonts: z
      .object({
        heading: z.string().trim().min(1).max(80),
        body: z.string().trim().min(1).max(80),
        numeric: z.string().trim().min(1).max(80).optional(),
      })
      .strict(),
    fontSizes: z
      .object({
        keyMessage: positive.max(200),
        heading: positive.max(200),
        body: positive.max(200),
        caption: positive.max(200),
        display: positive.max(400).optional(),
      })
      .strict(),
    spacing: z
      .object({
        xs: nonNegative.max(5),
        sm: nonNegative.max(5),
        md: nonNegative.max(5),
        lg: nonNegative.max(5),
        xl: nonNegative.max(5),
      })
      .strict(),
    colors: z
      .object({
        background: hexColor,
        text: hexColor,
        mutedText: hexColor,
        primary: hexColor,
        secondary: hexColor,
        accent: hexColor,
        surface: hexColor.optional(),
        border: hexColor.optional(),
        positive: hexColor.optional(),
        negative: hexColor.optional(),
        onPrimary: hexColor.optional(),
      })
      .strict(),
    margins: z
      .object({
        top: nonNegative.max(10),
        right: nonNegative.max(10),
        bottom: nonNegative.max(10),
        left: nonNegative.max(10),
      })
      .strict()
      .optional(),
    shape: z
      .object({
        cornerRadius: nonNegative.max(1),
        borderWidth: nonNegative.max(20),
      })
      .strict()
      .optional(),
  })
  .strict() as unknown as z.ZodType<PresentationTheme>;

export function parseTheme(input: unknown): PresentationTheme {
  return PresentationThemeSchema.parse(input);
}
