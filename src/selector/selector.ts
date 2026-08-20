import type { SlideSpec } from "../domain/types.js";
import { checkConstraints, type AnyVisualPattern } from "../patterns/pattern.js";
import { defaultRegistry, PatternRegistry } from "../patterns/registry.js";

/** 1 Pattern の評価結果（デバッグのため全候補を Visual Plan に残す） */
export interface PatternEvaluation {
  patternId: string;
  version: string;
  compatible: boolean;
  /** 最終スコア (0-1)。compatible=false のときは 0。 */
  score: number;
  /** match() が返した意味フィット度 */
  matchScore: number;
  /** compatibility 宣言との適合度 */
  compatibilityScore: number;
  reason: string;
  variantId?: string;
  /** どの段階で落ちたか */
  rejectedAt?: "compatibility" | "constraints" | "match";
}

export interface PatternSelection {
  patternId: string;
  variantId: string;
  score: number;
  reason: string;
  /**
   * 適合する Pattern が 1 つも無く fallback へ退避したか。
   * fallback Pattern がスコアで勝った場合は「正当な選択」なので false。
   */
  fallbackUsed: boolean;
  /** スコア降順の全候補 */
  candidates: PatternEvaluation[];
}

/** match() と compatibility の重み。合計 1.0。 */
const MATCH_WEIGHT = 0.7;
const COMPATIBILITY_WEIGHT = 0.3;

const round3 = (n: number): number => Number(n.toFixed(3));

/**
 * compatibility 宣言に対する適合度。
 * logicalRelation は *ゲート*（外れたら不適合）、role / evidenceType は *減点* 要素。
 */
function scoreCompatibility(
  pattern: AnyVisualPattern,
  slide: SlideSpec,
): { ok: boolean; score: number; note: string } {
  const { logicalRelations, slideRoles, evidenceTypes } = pattern.compatibility;

  if (!logicalRelations.includes(slide.logicalRelation)) {
    return {
      ok: false,
      score: 0,
      note: `logicalRelation=${slide.logicalRelation} は対応外 (対応: ${logicalRelations.join(", ")})`,
    };
  }

  let score = 1;
  const notes: string[] = [];

  if (slideRoles && slideRoles.length > 0) {
    if (slideRoles.includes(slide.slideRole)) {
      notes.push(`slideRole=${slide.slideRole} が想定内`);
    } else {
      score *= 0.65;
      notes.push(`slideRole=${slide.slideRole} は想定外`);
    }
  }

  if (evidenceTypes && evidenceTypes.length > 0 && slide.evidence.length > 0) {
    const matched = slide.evidence.filter((e) => evidenceTypes.includes(e.type)).length;
    const ratio = matched / slide.evidence.length;
    score *= 0.6 + 0.4 * ratio;
    notes.push(`対応 evidenceType の割合 ${matched}/${slide.evidence.length}`);
  }

  return { ok: true, score, note: notes.join(" / ") || "compatibility 完全一致" };
}

/** 1 スライド × 1 Pattern の評価 */
export function evaluatePattern(pattern: AnyVisualPattern, slide: SlideSpec): PatternEvaluation {
  const base = { patternId: pattern.id, version: pattern.version };

  const compat = scoreCompatibility(pattern, slide);
  if (!compat.ok) {
    return {
      ...base,
      compatible: false,
      score: 0,
      matchScore: 0,
      compatibilityScore: 0,
      reason: compat.note,
      rejectedAt: "compatibility",
    };
  }

  const violations = checkConstraints(slide, pattern.constraints);
  if (violations.length > 0) {
    return {
      ...base,
      compatible: false,
      score: 0,
      matchScore: 0,
      compatibilityScore: round3(compat.score),
      reason: `constraints 違反: ${violations.map((v) => v.message).join(" / ")}`,
      rejectedAt: "constraints",
    };
  }

  const match = pattern.match(slide);
  if (!match.compatible) {
    return {
      ...base,
      compatible: false,
      score: 0,
      matchScore: 0,
      compatibilityScore: round3(compat.score),
      reason: match.reason,
      rejectedAt: "match",
    };
  }

  // スコア契約の実行時検証: Pattern 作者が範囲外の値で選択を支配できないようにする
  if (!Number.isFinite(match.score) || match.score < 0 || match.score > 1) {
    throw new Error(
      `Pattern '${pattern.id}' の match() が範囲外のスコアを返しました: ${match.score}（0〜1 の共通ルーブリックに従う契約です）`,
    );
  }

  const score = MATCH_WEIGHT * match.score + COMPATIBILITY_WEIGHT * compat.score;
  return {
    ...base,
    compatible: true,
    score: round3(score),
    matchScore: round3(match.score),
    compatibilityScore: round3(compat.score),
    reason: match.reason,
    variantId: match.variantId ?? pattern.variants[0] ?? "standard",
  };
}

export interface SelectorOptions {
  registry?: PatternRegistry;
  /** このスコアを下回る候補は採用しない（fallback へ回す） */
  minimumScore?: number;
}

/**
 * Visual Pattern Selector（決定論的スコアリング。ここに LLM は入らない）
 *
 *   compatibility → constraints → match() → スコアリング → 最高スコア
 *
 * 同点の場合は「scope 優先度が高い方（project > user > team > company > builtin）」
 * →「fallback でない方」→「レジストリ登録順」で決まるため、
 * 同じ入力からは必ず同じ Pattern が選ばれる。
 *
 * scope 優先度が効くのは *同点のとき* だけであり、
 * 意味的に不適合な Pattern が scope の力で選ばれることはない。
 */
export function selectPattern(slide: SlideSpec, options: SelectorOptions = {}): PatternSelection {
  const registry = options.registry ?? defaultRegistry;
  const minimumScore = options.minimumScore ?? 0;
  const patterns = registry.all();

  const evaluations = patterns.map((pattern, index) => ({
    index,
    pattern,
    priority: registry.priorityOf(pattern.id),
    evaluation: evaluatePattern(pattern, slide),
  }));

  const sorted = [...evaluations].sort((a, b) => {
    if (b.evaluation.score !== a.evaluation.score) return b.evaluation.score - a.evaluation.score;
    if (b.evaluation.matchScore !== a.evaluation.matchScore) {
      return b.evaluation.matchScore - a.evaluation.matchScore;
    }
    // 同点のときだけ scope 優先度（project > user > team > company > builtin）が効く
    if (b.priority !== a.priority) return b.priority - a.priority;
    const aFallback = a.pattern.fallback ? 1 : 0;
    const bFallback = b.pattern.fallback ? 1 : 0;
    if (aFallback !== bFallback) return aFallback - bFallback;
    return a.index - b.index;
  });

  const candidates = sorted.map((s) => s.evaluation);
  const winner = sorted.find(
    (s) => s.evaluation.compatible && s.evaluation.score >= minimumScore,
  );

  if (winner) {
    return {
      patternId: winner.evaluation.patternId,
      variantId: winner.evaluation.variantId ?? "standard",
      score: winner.evaluation.score,
      reason: winner.evaluation.reason,
      fallbackUsed: false,
      candidates,
    };
  }

  // どの Pattern も適合しなかった → fallback（無理に図解しない）
  const fallback = registry.fallbacks()[0];
  if (!fallback) {
    throw new Error("適合する Pattern が無く、fallback Pattern も登録されていません");
  }
  return {
    patternId: fallback.id,
    variantId: fallback.variants[0] ?? "standard",
    score: 0,
    reason: `適合する Pattern が存在しないため fallback (${fallback.id}) を使用`,
    fallbackUsed: true,
    candidates,
  };
}
