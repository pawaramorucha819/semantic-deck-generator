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
import { defaultRegistry, type PatternRegistry } from "../patterns/registry.js";
import { checkSelectable, selectPattern, type PatternEvaluation } from "./selector.js";

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
 * **AI も人間も、安全ゲートは突破できない。**
 * compatibility / constraints / match / minimumScore / variant 実装の確認は
 * `checkSelectable()`（決定論的 Selector と同じ判定）に一本化されており、
 * AI 推薦も human override も必ずここを通る。
 * 違うのは「誰が決めるか」だけで、「何が許されるか」は全員同じ。
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

/**
 * 人間の指定を検証する。通らなければ黙って無視せず必ず例外にする。
 *
 * Human Override は「AI より優先される」が「安全性検証を免除される」わけではない。
 * 決定論的 Selector や AI 推薦とまったく同じ `checkSelectable()` を通す。
 */
function validateOverride(
  slide: SlideSpec,
  override: HumanOverride,
  registry: PatternRegistry,
  minimumScore: number,
): { patternId: string; variantId: string } {
  if (!registry.has(override.patternId)) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: 指定された patternId が登録されていません: ${override.patternId}`,
    );
  }
  const pattern = registry.get(override.patternId);

  const selectable = checkSelectable(pattern, slide, {
    minimumScore,
    ...(override.variantId ? { variantId: override.variantId } : {}),
  });
  if (!selectable.ok) {
    throw new SelectionPolicyError(
      `slide ${slide.id}: Pattern '${pattern.id}' は指定できません: ${selectable.reason}`,
    );
  }

  return { patternId: pattern.id, variantId: selectable.variantId };
}

/** 1 スライドの Pattern を決める（決定論的選択 → Knowledge → AI → 最終決定） */
export async function decideSlidePattern(
  slide: SlideSpec,
  options: SelectionPolicyOptions,
): Promise<SlideSelectionOutcome> {
  const registry = options.registry ?? defaultRegistry;
  // 最低採用品質。決定論的選択・AI 候補・human override のすべてに同じ値を適用する。
  const minimumScore = options.minimumScore ?? 0;

  // 1) 決定論的選択（ここが常に土台。AI の有無にかかわらず必ず走る）
  const deterministic = selectPattern(slide, { registry, minimumScore });

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
      // minimumScore を満たさない Pattern は AI にも見せない
      // （見せてしまうと AI 経由で最低採用品質を迂回できてしまう）
      candidates: toAdvisableCandidates(
        deterministic.candidates,
        (patternId) => {
          const pattern = registry.get(patternId);
          return { description: pattern.description, variants: pattern.variants };
        },
        { minimumScore },
      ),
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
      const recommendedId = advice.recommendation.recommendedPatternId;

      // 二重の安全確認: Advisor 側で候補を絞っていても、採用前にここで必ず再検証する
      // （Advisor 実装を差し替えてもこのゲートは効く）。
      if (!registry.has(recommendedId)) {
        ai.status = "rejected";
        ai.rejectionReason = `未登録の patternId です: ${recommendedId}`;
      } else {
        const pattern = registry.get(recommendedId);
        const selectable = checkSelectable(pattern, slide, {
          minimumScore,
          ...(advice.recommendation.recommendedVariantId
            ? { variantId: advice.recommendation.recommendedVariantId }
            : {}),
        });

        if (!selectable.ok) {
          ai.status = "rejected";
          ai.rejectionReason = `再検証で棄却されました: ${selectable.reason}`;
        } else if (options.mode === "auto") {
          aiSelection = { patternId: pattern.id, variantId: selectable.variantId };
          ai.status = "accepted";
        }
      }
    }
  }

  // 4) 最終決定（人間の override が常に最優先）
  const override = options.overrides?.[slide.id];
  let selectedBy: SelectedBy = "deterministic";
  let selected = { patternId: deterministic.patternId, variantId: deterministic.variantId };
  let reason = deterministic.reason;

  if (override) {
    selected = validateOverride(slide, override, registry, minimumScore);
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
