import { z } from "zod";
import type { SlideSpec } from "../domain/types.js";
import type { KnowledgeSearchResult } from "../knowledge/retrieval/retrieval.js";
import type { AiAdviceStatus } from "../plan/decision-trace.js";
import { meetsSelectionGate, type PatternEvaluation } from "../selector/selector.js";
import { AIUnavailableError, type AIProvider } from "./provider.js";

/**
 * AI Pattern Advisor。
 *
 * **AI に全 Pattern から自由選択させない**。
 * 決定論的 Selector が compatibility / constraints / match を通した
 * *候補集合の中から* 推薦させる。候補外の ID を返してきたら棄却する。
 *
 *   SlideSpec → 決定論的候補生成 → (候補 + 過去スライド) → AI 推薦 → 検証 → Selection Policy
 */

export interface PatternRecommendation {
  recommendedPatternId: string;
  recommendedVariantId?: string;
  /** 0.0 - 1.0 */
  confidence: number;
  reason: string;
  /** 参考にした過去スライド（`<deckId>#<slideNumber>`） */
  referenceSlides?: string[];
}

export const PatternRecommendationSchema: z.ZodType<PatternRecommendation> = z
  .object({
    recommendedPatternId: z.string().trim().min(1).max(120),
    recommendedVariantId: z.string().trim().min(1).max(120).optional(),
    confidence: z.number().min(0).max(1),
    reason: z.string().trim().min(1).max(1000),
    referenceSlides: z.array(z.string().trim().min(1).max(200)).max(20).optional(),
  })
  .strict() as unknown as z.ZodType<PatternRecommendation>;

/** 制約付きデコードに対応した Provider へ渡す JSON Schema */
export const PATTERN_RECOMMENDATION_JSON_SCHEMA: Record<string, unknown> = {
  type: "object",
  additionalProperties: false,
  required: ["recommendedPatternId", "confidence", "reason"],
  properties: {
    recommendedPatternId: { type: "string" },
    recommendedVariantId: { type: "string" },
    confidence: { type: "number", minimum: 0, maximum: 1 },
    reason: { type: "string" },
    referenceSlides: { type: "array", items: { type: "string" } },
  },
};

/** AI へ見せる候補（既に安全性検証を通ったものだけ） */
export interface AdvisableCandidate {
  patternId: string;
  variantId: string;
  variants: readonly string[];
  description: string;
  deterministicScore: number;
  reason: string;
}

export interface PatternAdviceInput {
  slide: SlideSpec;
  candidates: AdvisableCandidate[];
  references: KnowledgeSearchResult[];
}

export interface PatternAdviceResult {
  provider: string;
  status: AiAdviceStatus;
  recommendation?: PatternRecommendation;
  rejectionReason?: string;
}

const INSTRUCTIONS = `あなたはプレゼンテーション資料のレイアウト選択を支援するアシスタントです。

与えられるもの:
- slide: これから作るスライドの意味構造（主張・根拠・論理関係）
- candidates: **すでに機械的な適合検証を通過した Pattern のみ**。この中からしか選べません
- references: 社内の過去スライドで似ているもの

やること:
- candidates の中から最も適切な Pattern を 1 つ選ぶ
- なぜそれが良いのかを日本語で簡潔に述べる（過去スライドを参考にした場合はそれも書く）
- 自信度 (confidence) を 0.0-1.0 で示す

禁止事項:
- candidates に無い patternId を返すこと
- 座標・色・フォントサイズなど描画に関する指示を出すこと
- JSON 以外を出力すること`;

export class PatternAdvisor {
  constructor(private readonly provider: AIProvider) {}

  get providerId(): string {
    return this.provider.id;
  }

  get available(): boolean {
    return this.provider.available;
  }

  async advise(input: PatternAdviceInput): Promise<PatternAdviceResult> {
    if (!this.provider.available) {
      return {
        provider: this.provider.id,
        status: "unavailable",
        rejectionReason: "AI Provider が利用できません",
      };
    }
    if (input.candidates.length === 0) {
      return {
        provider: this.provider.id,
        status: "rejected",
        rejectionReason: "決定論的検証を通過した候補が 1 件もありません",
      };
    }

    let recommendation: PatternRecommendation;
    try {
      recommendation = await this.provider.generateStructured({
        purpose: "pattern-selection",
        instructions: INSTRUCTIONS,
        input: {
          slide: {
            id: input.slide.id,
            keyMessage: input.slide.keyMessage,
            supportingMessage: input.slide.supportingMessage,
            slideRole: input.slide.slideRole,
            logicalRelation: input.slide.logicalRelation,
            content: input.slide.content,
            evidence: input.slide.evidence.map((e) => ({
              id: e.id,
              type: e.type,
              statement: e.statement,
            })),
          },
          candidates: input.candidates,
          references: input.references.map((r) => ({
            slideId: r.entry.slideId,
            score: r.score,
            keyMessage: r.entry.semantic?.keyMessage,
            slideRole: r.entry.semantic?.slideRole,
            logicalRelation: r.entry.semantic?.logicalRelation,
            patternId: r.entry.visual?.patternId,
            purpose: r.entry.metadata.purpose,
            audience: r.entry.metadata.audience,
          })),
        },
        schema: PatternRecommendationSchema,
        jsonSchema: PATTERN_RECOMMENDATION_JSON_SCHEMA,
        maxOutputTokens: 1024,
      });
    } catch (error) {
      const unavailable = error instanceof AIUnavailableError;
      return {
        provider: this.provider.id,
        status: unavailable ? "unavailable" : "error",
        rejectionReason: error instanceof Error ? error.message : String(error),
      };
    }

    // --- ハードゲート: AI は候補集合の外へ出られない ---------------------------
    const candidate = input.candidates.find(
      (c) => c.patternId === recommendation.recommendedPatternId,
    );
    if (!candidate) {
      return {
        provider: this.provider.id,
        status: "rejected",
        recommendation,
        rejectionReason: `候補集合に無い patternId を推薦しました: ${recommendation.recommendedPatternId} (候補: ${input.candidates
          .map((c) => c.patternId)
          .join(", ")})`,
      };
    }
    if (
      recommendation.recommendedVariantId &&
      !candidate.variants.includes(recommendation.recommendedVariantId)
    ) {
      return {
        provider: this.provider.id,
        status: "rejected",
        recommendation,
        rejectionReason: `Pattern '${candidate.patternId}' に未実装の variant を推薦しました: ${recommendation.recommendedVariantId}`,
      };
    }

    return { provider: this.provider.id, status: "advisory", recommendation };
  }
}

export interface AdvisableCandidateOptions {
  /**
   * このスコアを下回る候補は AI へ渡さない。
   * minimumScore は「最低採用品質」なので、AI が下回る Pattern を推薦できてしまうと
   * 決定論的 Selector の契約が破れる。
   */
  minimumScore?: number;
}

/**
 * 決定論的評価 → AI に見せる候補。
 *
 * 通すのは「決定論的 Selector が採用してよいと判断した Pattern」だけ
 * （compatibility / constraints / match / minimumScore をすべて通過したもの）。
 */
export function toAdvisableCandidates(
  evaluations: readonly PatternEvaluation[],
  describe: (patternId: string) => { description: string; variants: readonly string[] },
  options: AdvisableCandidateOptions = {},
): AdvisableCandidate[] {
  return evaluations
    .filter((evaluation) => meetsSelectionGate(evaluation, options.minimumScore ?? 0))
    .map((evaluation) => {
      const info = describe(evaluation.patternId);
      return {
        patternId: evaluation.patternId,
        variantId: evaluation.variantId ?? info.variants[0] ?? "standard",
        variants: info.variants,
        description: info.description,
        deterministicScore: evaluation.score,
        reason: evaluation.reason,
      };
    });
}
