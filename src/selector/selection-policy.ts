import type { PatternAdvisor } from "../ai/pattern-advisor.js";
import { toAdvisableCandidates } from "../ai/pattern-advisor.js";
import type { SlideSpec } from "../domain/types.js";
import type {
  KnowledgeQuery,
  KnowledgeRetriever,
  KnowledgeSearchResult,
} from "../knowledge/retrieval/retrieval.js";
import type {
  AiTrace,
  DecisionTrace,
  SelectedBy,
  SelectionMode,
} from "../plan/decision-trace.js";
import { checkConstraints } from "../patterns/pattern.js";
import { defaultRegistry, type PatternRegistry } from "../patterns/registry.js";
import { evaluatePattern, selectPattern, type PatternEvaluation } from "./selector.js";

/**
 * Selection Policy — 「誰が最終的に Pattern を決めるか」を扱う層。
 *
 *   manual : AI は呼ばない。Knowledge は参考情報として集めるだけ。
 *            決定論的勝者、または人間の override が採用される。
 *   assist : AI が推薦し、その内容を Decision Trace に記録する。
 *            **採用はしない**（非対話 CLI では決定論的勝者を使う）。
 *            人間が override すれば人間が勝つ。
 *   auto   : AI の推薦を、既存の検証をすべて通ったうえで自動採用する。
 *
 * どのモードでも、AI が compatibility / constraints / bindings のゲートを
 * 突破することはできない。AI が選べるのは *決定論的に安全と分かっている候補集合の中だけ*。
 */

export class SelectionPolicyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SelectionPolicyError";
  }
}

/** 人間による明示指定（CLI / 設定ファイル / 将来の GUI から来る） */
export interface HumanOverride {
  patternId: string;
  variantId?: string;
  reason?: string;
}

export interface SelectionPolicyOptions {
  mode: SelectionMode;
  registry?: PatternRegistry;
  advisor?: PatternAdvisor;
  retriever?: KnowledgeRetriever;
  /** slideId → 人間の指定 */
  overrides?: Record<string, HumanOverride>;
  /** Knowledge から引く件数（既定 3） */
  knowledgeLimit?: number;
  /** Knowledge 検索に足す共通条件（purpose / audience など） */
  knowledgeContext?: Pick<KnowledgeQuery, "purpose" | "audience" | "organization">;
  minimumScore?: number;
  /** manual モードでも AI に参考意見を求める（既定 false） */
  adviseInManualMode?: boolean;
}

export interface SlideSelectionOutcome {
  slideId: string;
  patternId: string;
  variantId: string;
  /** 決定論的 Selector が付けたスコア（採用された Pattern のもの） */
  score: number;
  reason: string;
  fallbackUsed: boolean;
  candidates: PatternEvaluation[];
  references: KnowledgeSearchResult[];
  decisionTrace: DecisionTrace;
}

function knowledgeQueryFor(
  slide: SlideSpec,
  options: SelectionPolicyOptions,
): KnowledgeQuery {
  return {
    text: [slide.keyMessage, slide.supportingMessage ?? ""].join(" "),
    slideRole: slide.slideRole,
    logicalRelation: slide.logicalRelation,
    ...(options.knowledgeContext ?? {}),
    limit: options.knowledgeLimit ?? 3,
  };
}

/** 人間の指定を検証する。通らなければ黙って無視せず必ず例外にする。 */
function validateOverride(
  slide: SlideSpec,
  override: HumanOverride,
  registry: PatternRegistry,
): { patternId: string; variantId: string } {
  if (!registry.has(override.patternId)) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: 指定された patternId が登録されていません: ${override.patternId}`,
    );
  }
  const pattern = registry.get(override.patternId);

  if (!pattern.compatibility.logicalRelations.includes(slide.logicalRelation)) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: Pattern '${pattern.id}' は logicalRelation=${slide.logicalRelation} に対応していません`,
    );
  }
  const violations = checkConstraints(slide, pattern.constraints);
  if (violations.length > 0) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: Pattern '${pattern.id}' の constraints を満たしません: ${violations
        .map((v) => v.message)
        .join(" / ")}`,
    );
  }

  const variantId = override.variantId ?? pattern.variants[0] ?? "standard";
  if (!pattern.variants.includes(variantId)) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: Pattern '${pattern.id}' に variant '${variantId}' は実装されていません`,
    );
  }
  return { patternId: pattern.id, variantId };
}

/** 1 スライドの Pattern を決める（決定論的選択 → Knowledge → AI → 最終決定） */
export async function decideSlidePattern(
  slide: SlideSpec,
  options: SelectionPolicyOptions,
): Promise<SlideSelectionOutcome> {
  const registry = options.registry ?? defaultRegistry;

  // 1) 決定論的選択（ここが常に土台。AI の有無にかかわらず必ず走る）
  const deterministic = selectPattern(slide, {
    registry,
    ...(options.minimumScore !== undefined ? { minimumScore: options.minimumScore } : {}),
  });

  // 2) Knowledge Retrieval（AI を使わなくても人間の参考になる）
  let references: KnowledgeSearchResult[] = [];
  if (options.retriever) {
    references = await options.retriever.retrieve(knowledgeQueryFor(slide, options));
  }

  // 3) AI 推薦（manual では既定で呼ばない）
  const shouldAdvise =
    !!options.advisor &&
    (options.mode === "assist" ||
      options.mode === "auto" ||
      (options.mode === "manual" && options.adviseInManualMode === true));

  let ai: AiTrace | undefined;
  let aiSelection: { patternId: string; variantId: string } | undefined;

  if (shouldAdvise && options.advisor) {
    const advice = await options.advisor.advise({
      slide,
      candidates: toAdvisableCandidates(deterministic.candidates, (patternId) => {
        const pattern = registry.get(patternId);
        return { description: pattern.description, variants: pattern.variants };
      }),
      references,
    });

    ai = {
      provider: advice.provider,
      status: advice.status,
      ...(advice.recommendation
        ? {
            recommendedPatternId: advice.recommendation.recommendedPatternId,
            ...(advice.recommendation.recommendedVariantId
              ? { recommendedVariantId: advice.recommendation.recommendedVariantId }
              : {}),
            confidence: advice.recommendation.confidence,
            reason: advice.recommendation.reason,
            ...(advice.recommendation.referenceSlides
              ? { referenceSlides: advice.recommendation.referenceSlides }
              : {}),
          }
        : {}),
      ...(advice.rejectionReason ? { rejectionReason: advice.rejectionReason } : {}),
    };

    if (advice.status === "advisory" && advice.recommendation) {
      const pattern = registry.get(advice.recommendation.recommendedPatternId);
      const variantId =
        advice.recommendation.recommendedVariantId ?? pattern.variants[0] ?? "standard";

      // 二重の安全確認: Advisor が候補集合を検証済みでも、ここで再評価する
      const evaluation = evaluatePattern(pattern, slide);
      if (!evaluation.compatible) {
        ai.status = "rejected";
        ai.rejectionReason = `再検証で不適合と判定されました: ${evaluation.reason}`;
      } else if (!pattern.variants.includes(variantId)) {
        ai.status = "rejected";
        ai.rejectionReason = `未実装の variant です: ${variantId}`;
      } else if (options.mode === "auto") {
        aiSelection = { patternId: pattern.id, variantId };
        ai.status = "accepted";
      }
    }
  }

  // 4) 最終決定（人間の override が常に最優先）
  const override = options.overrides?.[slide.id];
  let selectedBy: SelectedBy = "deterministic";
  let selected = { patternId: deterministic.patternId, variantId: deterministic.variantId };
  let reason = deterministic.reason;

  if (override) {
    selected = validateOverride(slide, override, registry);
    selectedBy = "human";
    reason = override.reason ?? `人間が Pattern '${selected.patternId}' を明示指定`;
    if (ai?.status === "accepted") ai.status = "advisory";
  } else if (aiSelection) {
    selected = aiSelection;
    selectedBy = "ai";
    reason = ai?.reason ?? "AI 推薦を採用";
  }

  const selectedEvaluation = deterministic.candidates.find(
    (c) => c.patternId === selected.patternId,
  );

  const decisionTrace: DecisionTrace = {
    deterministic: {
      recommendedPatternId: deterministic.patternId,
      recommendedVariantId: deterministic.variantId,
      score: deterministic.score,
      fallbackUsed: deterministic.fallbackUsed,
      candidatePatternIds: deterministic.candidates
        .filter((c) => c.compatible)
        .map((c) => c.patternId),
    },
    ...(references.length > 0
      ? { knowledge: { referenceSlides: references.map((r) => r.entry.slideId) } }
      : {}),
    ...(ai ? { ai } : {}),
    decision: {
      mode: options.mode,
      selectedPatternId: selected.patternId,
      selectedVariantId: selected.variantId,
      selectedBy,
      reason,
    },
  };

  return {
    slideId: slide.id,
    patternId: selected.patternId,
    variantId: selected.variantId,
    score: selectedEvaluation?.score ?? deterministic.score,
    reason,
    // 決定論的に fallback へ落ちた場合のみ true。人間 / AI が選び直したら false。
    fallbackUsed: selectedBy === "deterministic" ? deterministic.fallbackUsed : false,
    candidates: deterministic.candidates,
    references,
    decisionTrace,
  };
}
