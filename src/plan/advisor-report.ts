import type { SlideSelectionOutcome } from "../selector/selection-policy.js";
import type { PatternEvaluation } from "../selector/selector.js";
import type { AiTrace, DecisionRecord, SelectionMode } from "./decision-trace.js";

/**
 * Advisor Report — 「なぜその Pattern になったのか」の *調査過程* を残す成果物。
 *
 * Visual Plan との責務分担:
 *   Visual Plan     … 再生に必要な **決定の記録**（decisionTrace）。小さく保つ。
 *   Advisor Report  … 決定に至るまでの **探索の詳細**（全候補のスコア、
 *                     参照した過去スライドの一致理由、AI の応答）。
 *
 * 同じ情報を二重管理しないため、Advisor Report は Visual Plan を置き換えない。
 * 生成のたびに書き出す任意成果物（`*.advisor-report.json`）。
 */

export const ADVISOR_REPORT_SCHEMA_VERSION = "1.0";

export interface AdvisorReportReference {
  slideId: string;
  score: number;
  matched: string[];
  keyMessage?: string;
  patternId?: string;
  purpose?: string;
  audience?: string;
}

export interface AdvisorReportSlide {
  slideId: string;
  slideNumber: number;
  deterministic: {
    recommendedPatternId: string;
    score: number;
    fallbackUsed: boolean;
    /** 全 Pattern の評価（不適合になった理由も含む） */
    candidates: PatternEvaluation[];
  };
  references: AdvisorReportReference[];
  ai?: AiTrace;
  decision: DecisionRecord;
}

export interface AdvisorReport {
  schemaVersion: string;
  mode: SelectionMode;
  /** AI Provider の識別子（未使用なら undefined） */
  provider?: string;
  /** 参照した Knowledge Store の識別子 */
  knowledgeStore?: string;
  slides: AdvisorReportSlide[];
  summary: {
    slides: number;
    aiRecommended: number;
    aiAccepted: number;
    aiRejected: number;
    humanOverridden: number;
    knowledgeReferenced: number;
  };
}

export interface BuildAdvisorReportOptions {
  mode: SelectionMode;
  provider?: string;
  knowledgeStore?: string;
}

export function buildAdvisorReport(
  outcomes: readonly SlideSelectionOutcome[],
  options: BuildAdvisorReportOptions,
): AdvisorReport {
  const slides: AdvisorReportSlide[] = outcomes.map((outcome, index) => ({
    slideId: outcome.slideId,
    slideNumber: index + 1,
    deterministic: {
      recommendedPatternId: outcome.decisionTrace.deterministic.recommendedPatternId,
      score: outcome.decisionTrace.deterministic.score,
      fallbackUsed: outcome.decisionTrace.deterministic.fallbackUsed,
      candidates: outcome.candidates,
    },
    references: outcome.references.map((reference) => ({
      slideId: reference.entry.slideId,
      score: reference.score,
      matched: reference.matched,
      ...(reference.entry.semantic?.keyMessage
        ? { keyMessage: reference.entry.semantic.keyMessage }
        : {}),
      ...(reference.entry.visual?.patternId ? { patternId: reference.entry.visual.patternId } : {}),
      ...(reference.entry.metadata.purpose ? { purpose: reference.entry.metadata.purpose } : {}),
      ...(reference.entry.metadata.audience ? { audience: reference.entry.metadata.audience } : {}),
    })),
    ...(outcome.decisionTrace.ai ? { ai: outcome.decisionTrace.ai } : {}),
    decision: outcome.decisionTrace.decision,
  }));

  return {
    schemaVersion: ADVISOR_REPORT_SCHEMA_VERSION,
    mode: options.mode,
    ...(options.provider ? { provider: options.provider } : {}),
    ...(options.knowledgeStore ? { knowledgeStore: options.knowledgeStore } : {}),
    slides,
    summary: {
      slides: slides.length,
      aiRecommended: slides.filter((s) => s.ai?.recommendedPatternId).length,
      aiAccepted: slides.filter((s) => s.ai?.status === "accepted").length,
      aiRejected: slides.filter((s) => s.ai?.status === "rejected").length,
      humanOverridden: slides.filter((s) => s.decision.selectedBy === "human").length,
      knowledgeReferenced: slides.filter((s) => s.references.length > 0).length,
    },
  };
}

export function serializeAdvisorReport(report: AdvisorReport): string {
  return `${JSON.stringify(report, null, 2)}\n`;
}
