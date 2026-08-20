import { z } from "zod";
import { LOGICAL_RELATIONS, SLIDE_ROLES } from "../domain/types.js";
import { PATTERN_SCOPES } from "../patterns/declarative/schema.js";

/**
 * Generation Config — 「どう生成するか」の設定。
 *
 * presentation.json（意味・論理）にはテンプレート ID も AI 設定も書けないので、
 * 生成方法に関する指定はすべてこのファイルに分離する。
 *
 *   presentation.json = WHAT / WHY / LOGIC
 *   generation.config.json = HOW TO GENERATE
 */

export const GENERATION_CONFIG_SCHEMA_VERSION = "1.0";

/** off = AI を一切使わない（V1 と同じ経路） */
export const AI_MODES = ["off", "manual", "assist", "auto"] as const;
export type AiMode = (typeof AI_MODES)[number];

const relativeOrAbsolutePath = z.string().trim().min(1).max(400);

export const GenerationConfigSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),

    template: z
      .object({
        id: z.string().trim().min(1).max(120).optional(),
        /** 外部 Template ディレクトリ（config からの相対パス） */
        dirs: z.array(relativeOrAbsolutePath).max(20).optional(),
      })
      .strict()
      .optional(),

    /** Template の theme を上書きする（ビルトインテーマ ID） */
    theme: z.object({ id: z.string().trim().min(1).max(80) }).strict().optional(),

    patterns: z
      .object({
        /** ビルトイン Pattern を含めるか（既定 true） */
        builtin: z.boolean().optional(),
        sources: z
          .array(
            z
              .object({
                dir: relativeOrAbsolutePath,
                scope: z.enum(PATTERN_SCOPES).optional(),
              })
              .strict(),
          )
          .max(20)
          .optional(),
        /** draft の Pattern も読み込むか（既定 false = 人間レビュー必須） */
        allowDraft: z.boolean().optional(),
      })
      .strict()
      .optional(),

    knowledge: z
      .object({
        stores: z
          .array(
            z
              .object({
                type: z.literal("json").optional(),
                dir: relativeOrAbsolutePath,
              })
              .strict(),
          )
          .max(10)
          .optional(),
        /** 検索時に共通で使う文脈 */
        context: z
          .object({
            purpose: z.string().trim().min(1).max(160).optional(),
            audience: z.string().trim().min(1).max(160).optional(),
            organization: z.string().trim().min(1).max(160).optional(),
          })
          .strict()
          .optional(),
        /** 1 スライドあたりの参照件数（既定 3） */
        limit: z.number().int().min(1).max(20).optional(),
      })
      .strict()
      .optional(),

    ai: z
      .object({
        mode: z.enum(AI_MODES).optional(),
        provider: z
          .discriminatedUnion("type", [
            z.object({ type: z.literal("none") }).strict(),
            z
              .object({
                type: z.literal("anthropic"),
                model: z.string().trim().min(1).max(120).optional(),
                effort: z.enum(["low", "medium", "high", "xhigh", "max"]).optional(),
                /** API キーを読む環境変数名（既定 ANTHROPIC_API_KEY） */
                apiKeyEnv: z.string().trim().min(1).max(120).optional(),
              })
              .strict(),
            z
              .object({
                type: z.literal("scripted"),
                /** purpose → 応答 を書いた JSON ファイル（オフラインのデモ・CI 用） */
                responses: relativeOrAbsolutePath,
              })
              .strict(),
          ])
          .optional(),
        /** Pattern 推薦に AI を使うか（既定 true） */
        patternSelection: z.boolean().optional(),
        /** 構造化支援に AI を使うか（既定 false。generate では未使用） */
        structurePlanning: z.boolean().optional(),
      })
      .strict()
      .optional(),

    selection: z
      .object({
        /** slideId → 人間による Pattern 指定 */
        overrides: z
          .record(
            z
              .object({
                patternId: z.string().trim().min(1).max(120),
                variantId: z.string().trim().min(1).max(120).optional(),
                reason: z.string().trim().min(1).max(400).optional(),
              })
              .strict(),
          )
          .optional(),
        /** 章扉（section master）として描くスライド */
        sectionSlides: z.array(z.string().trim().min(1).max(120)).max(50).optional(),
        minimumScore: z.number().min(0).max(1).optional(),
      })
      .strict()
      .optional(),

    output: z
      .object({
        dir: relativeOrAbsolutePath.optional(),
        titleSlide: z.boolean().optional(),
        slideNumbers: z.boolean().optional(),
        candidates: z.boolean().optional(),
        renderReport: z.boolean().optional(),
        advisorReport: z.boolean().optional(),
      })
      .strict()
      .optional(),

    /** 生成物を Knowledge へ自動登録する（作った資料が知識資産になる循環） */
    recordKnowledge: z
      .object({
        /** 追記先ストア（knowledge.stores のインデックス。既定 0） */
        storeIndex: z.number().int().min(0).max(9).optional(),
        deckId: z
          .string()
          .trim()
          .regex(/^[a-zA-Z0-9][a-zA-Z0-9-_.]{0,119}$/)
          .optional(),
        purpose: z.string().trim().min(1).max(160).optional(),
        audience: z.string().trim().min(1).max(160).optional(),
        organization: z.string().trim().min(1).max(160).optional(),
      })
      .strict()
      .optional(),
  })
  .strict()
  .superRefine((config, ctx) => {
    if (config.schemaVersion !== GENERATION_CONFIG_SCHEMA_VERSION) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["schemaVersion"],
        message: `未対応の Generation Config schemaVersion です (指定=${config.schemaVersion}, 対応=${GENERATION_CONFIG_SCHEMA_VERSION})`,
      });
    }

    const mode = config.ai?.mode ?? "off";
    const providerType = config.ai?.provider?.type ?? "none";
    const usesAdvisor = config.ai?.patternSelection ?? true;

    // AI が必要なモードなのに Provider が無い場合は、黙って無効化せずエラーにする
    if ((mode === "assist" || mode === "auto") && usesAdvisor && providerType === "none") {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["ai", "provider"],
        message: `ai.mode='${mode}' には AI Provider の設定が必要です（ai.provider.type を anthropic か scripted にするか、ai.mode を off にしてください）`,
      });
    }

    if (config.recordKnowledge && (config.knowledge?.stores?.length ?? 0) === 0) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["recordKnowledge"],
        message: "recordKnowledge を使うには knowledge.stores が必要です",
      });
    }
  });

export type GenerationConfig = z.infer<typeof GenerationConfigSchema>;

export class GenerationConfigError extends Error {
  constructor(
    summary: string,
    readonly issues: z.ZodIssue[] = [],
  ) {
    super(
      issues.length > 0
        ? `${summary}\n${issues
            .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
            .join("\n")}`
        : summary,
    );
    this.name = "GenerationConfigError";
  }
}

export function parseGenerationConfig(input: unknown): GenerationConfig {
  const result = GenerationConfigSchema.safeParse(input);
  if (!result.success) {
    throw new GenerationConfigError(
      `generation.config.json の検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  return result.data;
}

/** 検索用: slideRole / logicalRelation の妥当な値（CLI のヘルプ表示に使う） */
export const KNOWLEDGE_QUERY_ENUMS = {
  slideRoles: SLIDE_ROLES,
  logicalRelations: LOGICAL_RELATIONS,
} as const;
