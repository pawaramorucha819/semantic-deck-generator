import { z } from "zod";
import { COLOR_TOKEN_NAMES, FONT_ROLE_NAMES, TEXT_ROLE_NAMES } from "../render/tokens.js";
import { PresentationThemeSchema } from "../theme/schema.js";
import {
  MASTER_ROLES,
  TEMPLATE_SCHEMA_VERSION,
  type PresentationTemplate,
} from "./template.js";

/**
 * Template の runtime スキーマ。
 *
 * Template JSON は「企業が持ち込む外部データ」なので、次を構造的に禁止する。
 *  - 任意のコード / 関数 / 式（JSON の値しか書けない）
 *  - 未知のカラートークン（Theme に存在しない色名）
 *  - 絶対パス / 親ディレクトリを含むアセットパス（ZIP / パストラバーサル対策）
 */

export const ColorTokenSchema = z.union([
  z.enum(COLOR_TOKEN_NAMES),
  z
    .string()
    .regex(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/, "色は Theme トークン名か #RRGGBB で指定してください"),
]);

export const RectSchema = z
  .object({
    x: z.number().finite().min(-100).max(100),
    y: z.number().finite().min(-100).max(100),
    w: z.number().finite().positive().max(200),
    h: z.number().finite().positive().max(200),
  })
  .strict();

/**
 * 線を載せるための矩形。
 *
 * PowerPoint の水平線は h = 0、垂直線は w = 0 で表現されるため、
 * 線に限っては 0 を許す。「長さ 0 の線」は下の superRefine で拒否する。
 */
export const LineRectSchema = z
  .object({
    x: z.number().finite().min(-100).max(100),
    y: z.number().finite().min(-100).max(100),
    w: z.number().finite().nonnegative().max(200),
    h: z.number().finite().nonnegative().max(200),
  })
  .strict();

export const TemplateTextStyleSchema = z
  .object({
    role: z.enum(TEXT_ROLE_NAMES).optional(),
    fontSize: z.number().finite().positive().max(400).optional(),
    font: z.enum(FONT_ROLE_NAMES).optional(),
    color: ColorTokenSchema.optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    valign: z.enum(["top", "middle", "bottom"]).optional(),
    lineHeight: z.number().finite().positive().max(4).optional(),
    charSpacing: z.number().finite().min(-10).max(20).optional(),
  })
  .strict();

export const TemplateShapeStyleSchema = z
  .object({
    fill: ColorTokenSchema.optional(),
    line: z
      .object({
        color: ColorTokenSchema,
        width: z.number().finite().nonnegative().max(20).optional(),
        dashType: z.enum(["solid", "dash", "lgDash", "sysDot"]).optional(),
      })
      .strict()
      .optional(),
    transparency: z.number().finite().min(0).max(100).optional(),
  })
  .strict();

/** アセットパスの安全性検証: 相対パスのみ / 親ディレクトリ参照禁止 */
export const SafeRelativePathSchema = z
  .string()
  .trim()
  .min(1)
  .max(300)
  .refine((p) => !p.startsWith("/") && !p.startsWith("\\") && !/^[a-zA-Z]:/.test(p), {
    message: "アセットパスは相対パスで指定してください",
  })
  .refine((p) => !p.split(/[\\/]/).includes(".."), {
    message: "アセットパスに '..' は使用できません",
  })
  .refine((p) => !p.includes("\0"), { message: "アセットパスが不正です" });

export const TemplateElementSchema = z.discriminatedUnion("kind", [
  z
    .object({
      kind: z.literal("text"),
      box: RectSchema,
      text: z.string().min(1).max(400),
      style: TemplateTextStyleSchema.optional(),
    })
    .strict(),
  z
    .object({
      kind: z.literal("shape"),
      shape: z.enum(["rect", "line"]),
      // shape === "line" のときだけ 0 幅 / 0 高さを許す（下の superRefine で検証）
      box: LineRectSchema,
      style: TemplateShapeStyleSchema.optional(),
    })
    .strict(),
  z
    .object({
      kind: z.literal("image"),
      assetId: z.string().trim().min(1).max(120),
      box: RectSchema,
      sizing: z.enum(["contain", "cover", "stretch"]).optional(),
      transparency: z.number().finite().min(0).max(100).optional(),
    })
    .strict(),
]);

export const SlideMasterDefinitionSchema = z
  .object({
    id: z.string().trim().min(1).max(120),
    role: z.enum(MASTER_ROLES),
    name: z.string().trim().min(1).max(160).optional(),
    background: z.object({ color: ColorTokenSchema }).strict().optional(),
    contentMargins: z
      .object({
        top: z.number().finite().nonnegative().max(20),
        right: z.number().finite().nonnegative().max(20),
        bottom: z.number().finite().nonnegative().max(20),
        left: z.number().finite().nonnegative().max(20),
      })
      .strict()
      .optional(),
    elements: z.array(TemplateElementSchema).max(40).optional(),
    slideNumber: z
      .object({ box: RectSchema, style: TemplateTextStyleSchema.optional() })
      .strict()
      .optional(),
  })
  .strict();

export const TemplateAssetSchema = z
  .object({
    id: z.string().trim().min(1).max(120),
    path: SafeRelativePathSchema,
    kind: z.literal("image"),
    description: z.string().trim().min(1).max(300).optional(),
  })
  .strict();

export const PresentationTemplateSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    id: z
      .string()
      .trim()
      .regex(/^[a-z0-9][a-z0-9-_]{0,79}$/, "template id は英小文字・数字・- _ のみ使用できます"),
    version: z
      .string()
      .trim()
      .regex(/^\d+\.\d+\.\d+$/, "version は semver (x.y.z) で指定してください"),
    name: z.string().trim().min(1).max(160),
    theme: PresentationThemeSchema,
    slideSize: z
      .object({
        widthIn: z.number().finite().positive().max(100),
        heightIn: z.number().finite().positive().max(100),
      })
      .strict()
      .optional(),
    masters: z.array(SlideMasterDefinitionSchema).min(1).max(20),
    defaultMasterId: z.string().trim().min(1).max(120),
    titleMasterId: z.string().trim().min(1).max(120).optional(),
    sectionMasterId: z.string().trim().min(1).max(120).optional(),
    assets: z.array(TemplateAssetSchema).max(50).optional(),
    metadata: z
      .object({
        organization: z.string().trim().min(1).max(160).optional(),
        department: z.string().trim().min(1).max(160).optional(),
        description: z.string().trim().min(1).max(600).optional(),
        confidentialityLabel: z.string().trim().min(1).max(120).optional(),
      })
      .strict()
      .optional(),
    import: z
      .object({
        importedFrom: z.string().trim().min(1).max(400).optional(),
        unsupportedFeatures: z.array(z.string().min(1).max(200)).max(200).optional(),
        warnings: z.array(z.string().min(1).max(400)).max(200).optional(),
      })
      .strict()
      .optional(),
  })
  .strict()
  .superRefine((template, ctx) => {
    if (template.schemaVersion !== TEMPLATE_SCHEMA_VERSION) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["schemaVersion"],
        message: `未対応の Template schemaVersion です (指定=${template.schemaVersion}, 対応=${TEMPLATE_SCHEMA_VERSION})`,
      });
    }

    const masterIds = new Set<string>();
    template.masters.forEach((master, i) => {
      if (masterIds.has(master.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["masters", i, "id"],
          message: `master.id が重複しています: ${master.id}`,
        });
      }
      masterIds.add(master.id);
    });

    for (const [key, value] of [
      ["defaultMasterId", template.defaultMasterId],
      ["titleMasterId", template.titleMasterId],
      ["sectionMasterId", template.sectionMasterId],
    ] as const) {
      if (value && !masterIds.has(value)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: [key],
          message: `${key} が存在しない master を参照しています: ${value}`,
        });
      }
    }

    const assetIds = new Set((template.assets ?? []).map((a) => a.id));
    template.masters.forEach((master, mi) => {
      (master.elements ?? []).forEach((element, ei) => {
        if (element.kind === "image" && !assetIds.has(element.assetId)) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["masters", mi, "elements", ei, "assetId"],
            message: `未登録のアセットを参照しています: ${element.assetId}`,
          });
        }

        // 大きさの契約は「線だけ別」。線は長さ 0 でなければよく、
        // それ以外の図形は従来どおり正の幅と高さを要求する。
        if (element.kind === "shape") {
          const { w, h } = element.box;
          const invalid = element.shape === "line" ? w <= 0 && h <= 0 : w <= 0 || h <= 0;
          if (invalid) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              path: ["masters", mi, "elements", ei, "box"],
              message:
                element.shape === "line"
                  ? "長さ 0 の線は配置できません（w か h のどちらかは 0 より大きい必要があります）"
                  : "図形の w / h は 0 より大きい必要があります",
            });
          }
        }
      });
    });
  });

export class TemplateValidationError extends Error {
  constructor(
    summary: string,
    readonly issues: z.ZodIssue[],
  ) {
    super(
      `${summary}\n${issues
        .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
        .join("\n")}`,
    );
    this.name = "TemplateValidationError";
  }
}

export function parseTemplate(input: unknown): PresentationTemplate {
  const result = PresentationTemplateSchema.safeParse(input);
  if (!result.success) {
    throw new TemplateValidationError(
      `Template の検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  return result.data as PresentationTemplate;
}
