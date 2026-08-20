import { z } from "zod";

/**
 * Decision Trace — 「なぜこの Pattern になったのか」の追跡記録。
 *
 * 企業で AI を使う以上、選択理由が後から説明できなければ運用に載らない。
 * V2 では次の 4 段を必ず記録する。
 *
 *   deterministic … 決定論的 Selector が出した候補と勝者（ハードゲート通過済み）
 *   knowledge     … 参照した過去スライド
 *   ai            … AI の推薦と、その採否
 *   decision      … 最終的に何が選ばれ、誰が選んだか
 *
 * ここに書かれるのは *決定の記録* であり、探索の詳細（検索スコアや候補全文）は
 * Advisor Report 側の責務。二重管理にしないこと。
 */

export const SELECTION_MODES = ["manual", "assist", "auto"] as const;
export type SelectionMode = (typeof SELECTION_MODES)[number];

export const SELECTED_BY = ["deterministic", "ai", "human"] as const;
export type SelectedBy = (typeof SELECTED_BY)[number];

/** AI 推薦の扱われ方 */
export const AI_ADVICE_STATUSES = [
  /** 推薦が採用された（auto モード） */
  "accepted",
  /** 推薦は記録されたが採用はしていない（assist / manual モード） */
  "advisory",
  /** 検証に落ちて棄却された（存在しない Pattern / 候補集合外 など） */
  "rejected",
  /** AI Provider が利用できない */
  "unavailable",
  /** 呼び出しに失敗した */
  "error",
] as const;
export type AiAdviceStatus = (typeof AI_ADVICE_STATUSES)[number];

export interface DeterministicTrace {
  recommendedPatternId: string;
  recommendedVariantId: string;
  score: number;
  fallbackUsed: boolean;
  /** compatibility / constraints / match をすべて通過した候補（スコア降順） */
  candidatePatternIds: string[];
}

export interface KnowledgeTrace {
  /** `<deckId>#<slideNumber>` 形式の参照 */
  referenceSlides: string[];
  storeId?: string;
}

export interface AiTrace {
  provider: string;
  status: AiAdviceStatus;
  recommendedPatternId?: string;
  recommendedVariantId?: string;
  confidence?: number;
  reason?: string;
  referenceSlides?: string[];
  /** status が rejected / error のときの理由 */
  rejectionReason?: string;
}

export interface DecisionRecord {
  mode: SelectionMode;
  selectedPatternId: string;
  selectedVariantId: string;
  selectedBy: SelectedBy;
  reason: string;
}

export interface DecisionTrace {
  deterministic: DeterministicTrace;
  knowledge?: KnowledgeTrace;
  ai?: AiTrace;
  decision: DecisionRecord;
}

export const DeterministicTraceSchema = z
  .object({
    recommendedPatternId: z.string().min(1),
    recommendedVariantId: z.string().min(1),
    score: z.number(),
    fallbackUsed: z.boolean(),
    candidatePatternIds: z.array(z.string().min(1)),
  })
  .strict();

export const KnowledgeTraceSchema = z
  .object({
    referenceSlides: z.array(z.string().min(1)).max(50),
    storeId: z.string().min(1).optional(),
  })
  .strict();

export const AiTraceSchema = z
  .object({
    provider: z.string().min(1),
    status: z.enum(AI_ADVICE_STATUSES),
    recommendedPatternId: z.string().min(1).optional(),
    recommendedVariantId: z.string().min(1).optional(),
    confidence: z.number().min(0).max(1).optional(),
    reason: z.string().max(2000).optional(),
    referenceSlides: z.array(z.string().min(1)).max(50).optional(),
    rejectionReason: z.string().max(2000).optional(),
  })
  .strict();

export const DecisionRecordSchema = z
  .object({
    mode: z.enum(SELECTION_MODES),
    selectedPatternId: z.string().min(1),
    selectedVariantId: z.string().min(1),
    selectedBy: z.enum(SELECTED_BY),
    reason: z.string().max(2000),
  })
  .strict();

export const DecisionTraceSchema: z.ZodType<DecisionTrace> = z
  .object({
    deterministic: DeterministicTraceSchema,
    knowledge: KnowledgeTraceSchema.optional(),
    ai: AiTraceSchema.optional(),
    decision: DecisionRecordSchema,
  })
  .strict() as unknown as z.ZodType<DecisionTrace>;
