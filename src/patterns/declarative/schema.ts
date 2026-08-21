import { z } from "zod";
import { EVIDENCE_TYPES, LOGICAL_RELATIONS, SLIDE_ROLES } from "../../domain/types.js";
import {
  COLOR_TOKEN_NAMES,
  FONT_ROLE_NAMES,
  SPACING_TOKEN_NAMES,
  TEXT_ROLE_NAMES,
} from "../../render/tokens.js";
import { LineRectSchema } from "../../templates/schema.js";
import { SafeRelativePathSchema } from "../../templates/schema.js";

/**
 * Declarative Visual Pattern — 「PowerPoint で人間が作った良いスライド」を
 * TypeScript を書かずに再利用可能な Pattern にするための宣言的定義。
 *
 * 設計上の絶対条件:
 *  1. **JSON から任意のコードを実行しない**。式・関数・スクリプトは書けない。
 *     参照できるのは列挙された値パス（slot / item）と条件語彙だけ。
 *  2. **PptxGenJS を直接呼ばない**。描画は必ず RenderContext を通るので、
 *     はみ出し・重なり・実効フォントサイズ検査がコード Pattern と同じように効く。
 *  3. **色・書体は Theme トークンのみ**。未知のトークンはロード時に拒否される。
 *  4. bindings は slot 定義から生成した Zod スキーマで実行時検証される。
 */

export const DECLARATIVE_PATTERN_SCHEMA_VERSION = "1.0";

// ---------------------------------------------------------------------------
// 値参照 / 条件
// ---------------------------------------------------------------------------

/** `headline` `item.label` のような安全なパス（プロトタイプ汚染系のキーは拒否） */
const FORBIDDEN_PATH_SEGMENTS = new Set(["__proto__", "constructor", "prototype"]);

export const ValuePathSchema = z
  .string()
  .trim()
  .min(1)
  .max(80)
  .regex(/^[a-zA-Z][a-zA-Z0-9_]*(\.[a-zA-Z][a-zA-Z0-9_]*){0,2}$/, "不正な値パスです")
  .refine((p) => p.split(".").every((seg) => !FORBIDDEN_PATH_SEGMENTS.has(seg)), {
    message: "使用できないプロパティ名です",
  });

export type ValueRef =
  | { literal: string }
  | { ref: string }
  | { template: string };

export const ValueRefSchema: z.ZodType<ValueRef> = z.union([
  z.object({ literal: z.string().min(1).max(400) }).strict(),
  z.object({ ref: ValuePathSchema }).strict(),
  z.object({ template: z.string().min(1).max(400) }).strict(),
]);

export type Condition =
  | { hasValue: string }
  | { equals: { ref: string; value: string } }
  | { itemCount: { eq?: number; min?: number; max?: number }; slot?: string }
  | { indexEquals: number }
  | { logicalRelation: string[] }
  | { slideRole: string[] }
  | { hasEvidenceType: string[] }
  | { keyMessageLength: { min?: number; max?: number } }
  | { not: Condition }
  | { all: Condition[] }
  | { any: Condition[] };

const countBounds = z
  .object({
    eq: z.number().int().min(0).max(50).optional(),
    min: z.number().int().min(0).max(50).optional(),
    max: z.number().int().min(0).max(50).optional(),
  })
  .strict();

export const ConditionSchema: z.ZodType<Condition> = z.lazy(() =>
  z.union([
    z.object({ hasValue: ValuePathSchema }).strict(),
    z
      .object({
        equals: z.object({ ref: ValuePathSchema, value: z.string().max(200) }).strict(),
      })
      .strict(),
    z.object({ itemCount: countBounds, slot: z.string().min(1).max(80).optional() }).strict(),
    z.object({ indexEquals: z.number().int().min(0).max(50) }).strict(),
    z.object({ logicalRelation: z.array(z.enum(LOGICAL_RELATIONS)).min(1) }).strict(),
    z.object({ slideRole: z.array(z.enum(SLIDE_ROLES)).min(1) }).strict(),
    z.object({ hasEvidenceType: z.array(z.enum(EVIDENCE_TYPES)).min(1) }).strict(),
    z
      .object({
        keyMessageLength: z
          .object({
            min: z.number().int().min(0).max(500).optional(),
            max: z.number().int().min(0).max(500).optional(),
          })
          .strict(),
      })
      .strict(),
    z.object({ not: ConditionSchema }).strict(),
    z.object({ all: z.array(ConditionSchema).min(1).max(10) }).strict(),
    z.object({ any: z.array(ConditionSchema).min(1).max(10) }).strict(),
  ]),
) as z.ZodType<Condition>;

// ---------------------------------------------------------------------------
// Slot（意味 → bindings の抽出規則）
// ---------------------------------------------------------------------------

/** テキスト slot の取得元。SlideSpec の *意味* だけを参照する。 */
export const TEXT_SLOT_SOURCES = [
  "keyMessage",
  "supportingMessage",
  "slideRoleLabel",
  "sourceText",
  "literal",
] as const;
export type TextSlotSource = (typeof TEXT_SLOT_SOURCES)[number];

/** 反復 slot の取得元。「見せる対象」は content、「根拠」は evidence。 */
export const ITEMS_SLOT_SOURCES = [
  "contentPoints",
  "contentSequence",
  "content",
  "evidence",
  "contentOrEvidence",
] as const;
export type ItemsSlotSource = (typeof ITEMS_SLOT_SOURCES)[number];

export type SlotDefinition =
  | {
      id: string;
      type: "text";
      source: TextSlotSource;
      /** source === 'literal' のときの値 */
      value?: string;
      optional?: boolean;
      maxLength?: number;
    }
  | {
      id: string;
      type: "items";
      source: ItemsSlotSource;
      min?: number;
      max?: number;
      optional?: boolean;
      /** label をこの文字数で切り詰める */
      labelMaxLength?: number;
    }
  | {
      id: string;
      type: "metric";
      /** 単一の metric evidence（複数ある場合は先頭） */
      source: "headlineMetric" | "firstMetric";
      optional?: boolean;
    };

const slotIdSchema = z
  .string()
  .trim()
  .regex(/^[a-z][a-zA-Z0-9]{0,39}$/, "slot id は英小文字始まりの英数字で指定してください");

export const SlotDefinitionSchema: z.ZodType<SlotDefinition> = z.discriminatedUnion("type", [
  z
    .object({
      id: slotIdSchema,
      type: z.literal("text"),
      source: z.enum(TEXT_SLOT_SOURCES),
      value: z.string().min(1).max(400).optional(),
      optional: z.boolean().optional(),
      maxLength: z.number().int().min(1).max(400).optional(),
    })
    .strict(),
  z
    .object({
      id: slotIdSchema,
      type: z.literal("items"),
      source: z.enum(ITEMS_SLOT_SOURCES),
      min: z.number().int().min(1).max(12).optional(),
      max: z.number().int().min(1).max(12).optional(),
      optional: z.boolean().optional(),
      labelMaxLength: z.number().int().min(4).max(200).optional(),
    })
    .strict(),
  z
    .object({
      id: slotIdSchema,
      type: z.literal("metric"),
      source: z.enum(["headlineMetric", "firstMetric"]),
      optional: z.boolean().optional(),
    })
    .strict(),
]) as z.ZodType<SlotDefinition>;

// ---------------------------------------------------------------------------
// レイアウト
// ---------------------------------------------------------------------------

export const SpacingValueSchema = z.union([
  z.enum(SPACING_TOKEN_NAMES),
  z.number().finite().min(0).max(6),
]);
export type SpacingValue = z.infer<typeof SpacingValueSchema>;

export const DeclarativeColorSchema = z.union([
  z.enum(COLOR_TOKEN_NAMES),
  z
    .object({
      /** Theme トークンを背景色方向へ淡くする（ctx.tint と同じ意味） */
      tint: z.enum(COLOR_TOKEN_NAMES),
      amount: z.number().min(0).max(1),
    })
    .strict(),
]);
export type DeclarativeColor = z.infer<typeof DeclarativeColorSchema>;

export const DeclarativeTextStyleSchema = z
  .object({
    role: z.enum(TEXT_ROLE_NAMES).optional(),
    fontSize: z.number().finite().positive().max(400).optional(),
    /** role のサイズに掛ける倍率（Theme 非依存で強弱を付けるための口） */
    fontScale: z.number().finite().min(0.4).max(4).optional(),
    font: z.enum(FONT_ROLE_NAMES).optional(),
    color: DeclarativeColorSchema.optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    valign: z.enum(["top", "middle", "bottom"]).optional(),
    lineHeight: z.number().finite().min(0.8).max(3).optional(),
    charSpacing: z.number().finite().min(-10).max(20).optional(),
    shrink: z.boolean().optional(),
  })
  .strict();
export type DeclarativeTextStyle = z.infer<typeof DeclarativeTextStyleSchema>;

export const DECLARATIVE_SHAPE_KINDS = [
  "rect",
  "roundRect",
  "ellipse",
  "line",
  "rightArrow",
  "chevron",
  "triangle",
  "diamond",
  "downArrow",
] as const;

export const DeclarativeShapeStyleSchema = z
  .object({
    fill: DeclarativeColorSchema.optional(),
    fillTransparency: z.number().min(0).max(100).optional(),
    line: z
      .object({
        color: DeclarativeColorSchema,
        width: z.number().finite().min(0).max(20).optional(),
        dashType: z.enum(["solid", "dash", "lgDash", "sysDot"]).optional(),
      })
      .strict()
      .optional(),
    cornerRadius: z.number().min(0).max(1).optional(),
    rotate: z.number().min(-360).max(360).optional(),
  })
  .strict();
export type DeclarativeShapeStyle = z.infer<typeof DeclarativeShapeStyleSchema>;

/** 高さの決め方: auto = 内容から算出 / fill = 余りを分け合う / 数値 = inch 固定 */
export const HeightSchema = z.union([
  z.literal("auto"),
  z.literal("fill"),
  z.number().finite().min(0).max(20),
]);
export type HeightSpec = z.infer<typeof HeightSchema>;

export type LayoutNode =
  | {
      type: "stack";
      gap?: SpacingValue;
      padding?: SpacingValue;
      height?: HeightSpec;
      when?: Condition;
      children: LayoutNode[];
    }
  | {
      type: "columns";
      gap?: SpacingValue;
      weights?: number[];
      height?: HeightSpec;
      /**
       * 縦方向の寄せ。既定 'stretch'（領域いっぱいに伸ばす）。
       * 'middle' などを指定すると内容の高さに縮めて配置する。
       */
      align?: "top" | "middle" | "bottom" | "stretch";
      when?: Condition;
      /** items slot を横並びで繰り返す */
      repeat?: { slot: string; max?: number };
      /** repeat 使用時のテンプレート */
      item?: LayoutNode;
      children?: LayoutNode[];
    }
  | {
      type: "card";
      fill?: DeclarativeColor;
      border?: DeclarativeColor;
      accent?: DeclarativeColor;
      height?: HeightSpec;
      when?: Condition;
      child: LayoutNode;
    }
  | {
      type: "text";
      value: ValueRef;
      style?: DeclarativeTextStyle;
      height?: HeightSpec;
      when?: Condition;
    }
  | {
      type: "shape";
      shape: (typeof DECLARATIVE_SHAPE_KINDS)[number];
      style?: DeclarativeShapeStyle;
      height?: HeightSpec;
      /** 幅 (inch)。省略時は領域いっぱい。 */
      width?: number;
      align?: "left" | "center" | "right";
      when?: Condition;
    }
  | {
      type: "image";
      assetId: string;
      height: number;
      width?: number;
      align?: "left" | "center" | "right";
      sizing?: "contain" | "cover" | "stretch";
      when?: Condition;
    }
  | { type: "spacer"; height: SpacingValue; when?: Condition }
  | {
      /** 取り込んだスライドのように絶対座標で配置する */
      type: "absolute";
      origin?: "canvas" | "area";
      height?: HeightSpec;
      when?: Condition;
      children: Array<{ box: { x: number; y: number; w: number; h: number }; node: LayoutNode }>;
    };

export const LayoutNodeSchema: z.ZodType<LayoutNode> = z.lazy(() =>
  z.discriminatedUnion("type", [
    z
      .object({
        type: z.literal("stack"),
        gap: SpacingValueSchema.optional(),
        padding: SpacingValueSchema.optional(),
        height: HeightSchema.optional(),
        when: ConditionSchema.optional(),
        children: z.array(LayoutNodeSchema).min(1).max(24),
      })
      .strict(),
    z
      .object({
        type: z.literal("columns"),
        gap: SpacingValueSchema.optional(),
        weights: z.array(z.number().finite().positive().max(100)).min(1).max(12).optional(),
        height: HeightSchema.optional(),
        align: z.enum(["top", "middle", "bottom", "stretch"]).optional(),
        when: ConditionSchema.optional(),
        repeat: z
          .object({ slot: z.string().min(1).max(80), max: z.number().int().min(1).max(12).optional() })
          .strict()
          .optional(),
        item: LayoutNodeSchema.optional(),
        children: z.array(LayoutNodeSchema).min(1).max(12).optional(),
      })
      .strict(),
    z
      .object({
        type: z.literal("card"),
        fill: DeclarativeColorSchema.optional(),
        border: DeclarativeColorSchema.optional(),
        accent: DeclarativeColorSchema.optional(),
        height: HeightSchema.optional(),
        when: ConditionSchema.optional(),
        child: LayoutNodeSchema,
      })
      .strict(),
    z
      .object({
        type: z.literal("text"),
        value: ValueRefSchema,
        style: DeclarativeTextStyleSchema.optional(),
        height: HeightSchema.optional(),
        when: ConditionSchema.optional(),
      })
      .strict(),
    z
      .object({
        type: z.literal("shape"),
        shape: z.enum(DECLARATIVE_SHAPE_KINDS),
        style: DeclarativeShapeStyleSchema.optional(),
        height: HeightSchema.optional(),
        width: z.number().finite().min(0).max(60).optional(),
        align: z.enum(["left", "center", "right"]).optional(),
        when: ConditionSchema.optional(),
      })
      .strict(),
    z
      .object({
        type: z.literal("image"),
        assetId: z.string().trim().min(1).max(120),
        height: z.number().finite().min(0).max(20),
        width: z.number().finite().min(0).max(60).optional(),
        align: z.enum(["left", "center", "right"]).optional(),
        sizing: z.enum(["contain", "cover", "stretch"]).optional(),
        when: ConditionSchema.optional(),
      })
      .strict(),
    z
      .object({
        type: z.literal("spacer"),
        height: SpacingValueSchema,
        when: ConditionSchema.optional(),
      })
      .strict(),
    z
      .object({
        type: z.literal("absolute"),
        origin: z.enum(["canvas", "area"]).optional(),
        height: HeightSchema.optional(),
        when: ConditionSchema.optional(),
        children: z
          // 水平線 (h = 0) / 垂直線 (w = 0) を置けるように 0 を許し、
          // 「線以外は正の幅と高さ」という契約は compile 時に検証する。
          .array(z.object({ box: LineRectSchema, node: LayoutNodeSchema }).strict())
          .min(1)
          .max(80),
      })
      .strict(),
  ]),
) as z.ZodType<LayoutNode>;

export interface DeclarativeLayout {
  /** 全 Pattern 共通の見出しブロック（builtin Pattern と同じ描画を使う） */
  header?: { eyebrow?: ValueRef; headline: ValueRef; subline?: ValueRef };
  /** 出典行 */
  footer?: { text: ValueRef };
  body: LayoutNode;
}

export const DeclarativeLayoutSchema: z.ZodType<DeclarativeLayout> = z
  .object({
    header: z
      .object({
        eyebrow: ValueRefSchema.optional(),
        headline: ValueRefSchema,
        subline: ValueRefSchema.optional(),
      })
      .strict()
      .optional(),
    footer: z.object({ text: ValueRefSchema }).strict().optional(),
    body: LayoutNodeSchema,
  })
  .strict() as z.ZodType<DeclarativeLayout>;

// ---------------------------------------------------------------------------
// Pattern ドキュメント本体
// ---------------------------------------------------------------------------

export const PATTERN_SCOPES = ["builtin", "company", "team", "project", "user"] as const;
export type PatternScope = (typeof PATTERN_SCOPES)[number];

/** Capture 直後は draft。人間のレビューを経て approved にすると登録できる。 */
export const PATTERN_STATUSES = ["draft", "approved"] as const;
export type PatternStatus = (typeof PATTERN_STATUSES)[number];

export interface DeclarativePatternDocument {
  schemaVersion: string;
  kind: "declarative";
  id: string;
  version: string;
  description: string;
  metadata?: {
    scope?: PatternScope;
    status?: PatternStatus;
    author?: string;
    organization?: string;
    /** パッケージディレクトリからの相対パス */
    previewImage?: string;
    tags?: string[];
    /** 取り込み元（Pattern Capture が埋める） */
    source?: {
      file?: string;
      slideNumber?: number;
      importedAt?: string;
      /** 要素の意味を誰が解釈したか（heuristic / human、将来は ai） */
      interpretationSource?: string;
      unsupportedFeatures?: string[];
      warnings?: string[];
    };
  };
  compatibility: {
    logicalRelations: string[];
    slideRoles?: string[];
    evidenceTypes?: string[];
  };
  constraints?: {
    minEvidence?: number;
    maxEvidence?: number;
    maxKeyMessageLength?: number;
    maxSupportingMessageLength?: number;
    requiresAnyEvidenceType?: string[];
  };
  slots: SlotDefinition[];
  match: {
    baseScore: number;
    rules?: Array<{ when: Condition; score: number; reason: string }>;
    /** これらの slot が埋まらなければ compatible: false */
    requireSlots?: string[];
  };
  variants?: Array<{ id: string; description?: string; layout?: DeclarativeLayout }>;
  layout: DeclarativeLayout;
  assets?: Array<{ id: string; path: string; kind: "image" }>;
}

export const DeclarativePatternDocumentSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    kind: z.literal("declarative"),
    id: z
      .string()
      .trim()
      .regex(/^[a-z0-9][a-z0-9-_]{0,79}$/, "pattern id は英小文字・数字・- _ のみ使用できます"),
    version: z
      .string()
      .trim()
      .regex(/^\d+\.\d+\.\d+$/, "version は semver (x.y.z) で指定してください"),
    description: z.string().trim().min(1).max(400),
    metadata: z
      .object({
        scope: z.enum(PATTERN_SCOPES).optional(),
        status: z.enum(PATTERN_STATUSES).optional(),
        author: z.string().trim().min(1).max(160).optional(),
        organization: z.string().trim().min(1).max(160).optional(),
        previewImage: SafeRelativePathSchema.optional(),
        tags: z.array(z.string().trim().min(1).max(40)).max(20).optional(),
        source: z
          .object({
            file: z.string().trim().min(1).max(400).optional(),
            slideNumber: z.number().int().min(1).max(10000).optional(),
            importedAt: z.string().trim().min(1).max(40).optional(),
            interpretationSource: z.string().trim().min(1).max(40).optional(),
            unsupportedFeatures: z.array(z.string().min(1).max(200)).max(200).optional(),
            warnings: z.array(z.string().min(1).max(400)).max(200).optional(),
          })
          .strict()
          .optional(),
      })
      .strict()
      .optional(),
    compatibility: z
      .object({
        logicalRelations: z.array(z.enum(LOGICAL_RELATIONS)).min(1),
        slideRoles: z.array(z.enum(SLIDE_ROLES)).min(1).optional(),
        evidenceTypes: z.array(z.enum(EVIDENCE_TYPES)).min(1).optional(),
      })
      .strict(),
    constraints: z
      .object({
        minEvidence: z.number().int().min(0).max(50).optional(),
        maxEvidence: z.number().int().min(0).max(50).optional(),
        maxKeyMessageLength: z.number().int().min(1).max(500).optional(),
        maxSupportingMessageLength: z.number().int().min(1).max(1000).optional(),
        requiresAnyEvidenceType: z.array(z.enum(EVIDENCE_TYPES)).min(1).optional(),
      })
      .strict()
      .optional(),
    slots: z.array(SlotDefinitionSchema).min(1).max(24),
    match: z
      .object({
        baseScore: z.number().min(0).max(1),
        rules: z
          .array(
            z
              .object({
                when: ConditionSchema,
                score: z.number().min(-1).max(1),
                reason: z.string().trim().min(1).max(300),
              })
              .strict(),
          )
          .max(20)
          .optional(),
        requireSlots: z.array(z.string().min(1).max(80)).max(24).optional(),
      })
      .strict(),
    variants: z
      .array(
        z
          .object({
            id: z
              .string()
              .trim()
              .regex(/^[a-z][a-zA-Z0-9-]{0,39}$/, "variant id が不正です"),
            description: z.string().trim().min(1).max(300).optional(),
            layout: DeclarativeLayoutSchema.optional(),
          })
          .strict(),
      )
      .min(1)
      .max(8)
      .optional(),
    layout: DeclarativeLayoutSchema,
    assets: z
      .array(
        z
          .object({
            id: z.string().trim().min(1).max(120),
            path: SafeRelativePathSchema,
            kind: z.literal("image"),
          })
          .strict(),
      )
      .max(30)
      .optional(),
  })
  .strict()
  .superRefine((doc, ctx) => {
    if (doc.schemaVersion !== DECLARATIVE_PATTERN_SCHEMA_VERSION) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["schemaVersion"],
        message: `未対応の Declarative Pattern schemaVersion です (指定=${doc.schemaVersion}, 対応=${DECLARATIVE_PATTERN_SCHEMA_VERSION})`,
      });
    }

    const slotIds = new Set<string>();
    doc.slots.forEach((slot, i) => {
      if (slot.type === "text" && slot.source === "literal" && !slot.value) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["slots", i, "value"],
          message: "source='literal' の slot には value が必要です",
        });
      }
      if (slotIds.has(slot.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["slots", i, "id"],
          message: `slot id が重複しています: ${slot.id}`,
        });
      }
      slotIds.add(slot.id);
    });

    for (const required of doc.match.requireSlots ?? []) {
      if (!slotIds.has(required)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["match", "requireSlots"],
          message: `存在しない slot を参照しています: ${required}`,
        });
      }
    }

    const variantIds = new Set<string>();
    (doc.variants ?? []).forEach((variant, i) => {
      if (variantIds.has(variant.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["variants", i, "id"],
          message: `variant id が重複しています: ${variant.id}`,
        });
      }
      variantIds.add(variant.id);
    });
  });

export class DeclarativePatternValidationError extends Error {
  constructor(
    summary: string,
    readonly issues: z.ZodIssue[],
  ) {
    super(
      `${summary}\n${issues
        .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
        .join("\n")}`,
    );
    this.name = "DeclarativePatternValidationError";
  }
}

export function parseDeclarativePattern(input: unknown): DeclarativePatternDocument {
  const result = DeclarativePatternDocumentSchema.safeParse(input);
  if (!result.success) {
    throw new DeclarativePatternValidationError(
      `Declarative Pattern の検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  return result.data as DeclarativePatternDocument;
}
