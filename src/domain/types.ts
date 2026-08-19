/**
 * Domain Model — 「意味・論理」層
 *
 * このファイルに現れてよいのは *何を伝えるか / なぜそう言えるか* だけです。
 * 座標・フォント・色・グラフ種別など「どう描くか」に属する概念は
 * 絶対にこの層へ持ち込まないでください（→ src/patterns, src/render の責務）。
 */

/** スライドが資料全体の中で担う役割 */
export const SLIDE_ROLES = [
  "context",
  "problem",
  "analysis",
  "insight",
  "evidence",
  "comparison",
  "solution",
  "recommendation",
  "decision",
  "summary",
] as const;
export type SlideRole = (typeof SLIDE_ROLES)[number];

/** スライド内部の論理関係（KeyMessage と Evidence の結びつき方） */
export const LOGICAL_RELATIONS = [
  "claim_evidence",
  "comparison",
  "before_after",
  "cause_effect",
  "process",
  "timeline",
  "hierarchy",
  "composition",
  "correlation",
  "problem_solution",
  "pros_cons",
] as const;
export type LogicalRelation = (typeof LOGICAL_RELATIONS)[number];

/**
 * Evidence の種別。
 * 将来拡張する場合は (1) ここに種別名を追加し (2) 対応する `*Evidence` を定義し
 * (3) `src/domain/schema.ts` の discriminated union に 1 バリアント足すだけで済みます。
 */
export const EVIDENCE_TYPES = ["metric", "caseStudy", "text", "comparison"] as const;
export type EvidenceType = (typeof EVIDENCE_TYPES)[number];

export interface Source {
  id: string;
  title: string;
  url?: string;
  publisher?: string;
  /** ISO 8601 (YYYY-MM-DD 等) を推奨。表示形式は Renderer 側の責務。 */
  date?: string;
  /** 参照箇所（ページ・章・図表番号など）。例: "p.12", "第3章" */
  locator?: string;
  /** URL を取得した日付（ISO 8601 推奨） */
  accessedAt?: string;
}

interface EvidenceBase {
  id: string;
  /** その Evidence が主張している内容を 1 文で表したもの */
  statement: string;
  /** `SlideSpec.sources[].id` への参照（複数可） */
  sourceIds?: string[];
}

/** 単一の定量値（例: 在庫管理コストが前年比 18% 増加した） */
export interface MetricEvidence extends EvidenceBase {
  type: "metric";
  data: {
    label?: string;
    value: number;
    unit?: string;
    /** 値が良い方向か悪い方向か。色ではなく「意味」の指定。 */
    direction?: "increase" | "decrease" | "flat";
    period?: string;
    /** 比較対象となる基準値（あれば） */
    baseline?: number;
  };
}

/** 事例（他社導入実績など）。before / after を持つことが多い。 */
export interface CaseStudyEvidence extends EvidenceBase {
  type: "caseStudy";
  data: {
    organization?: string;
    before?: number;
    after?: number;
    unit?: string;
    beforeLabel?: string;
    afterLabel?: string;
    /** 変化量の説明（例: "-30%"）。省略時は before/after から導出される。 */
    changeLabel?: string;
  };
}

/** 定量化されていない事実・観察・引用 */
export interface TextEvidence extends EvidenceBase {
  type: "text";
  data?: {
    /** 短い見出し語（例: "棚卸に月40時間"）。長文の statement を要約したもの。 */
    label?: string;
    detail?: string;
  };
}

/** 2 つ以上の対象を同一軸で比較したもの */
export interface ComparisonEvidence extends EvidenceBase {
  type: "comparison";
  data: {
    unit?: string;
    items: Array<{
      label: string;
      value: number | string;
      note?: string;
    }>;
  };
}

export type Evidence = MetricEvidence | CaseStudyEvidence | TextEvidence | ComparisonEvidence;

/**
 * SemanticContent — スライドで「見せる対象」となる情報構造。
 *
 * Evidence（根拠 = 主張を裏付ける事実）とは別物です。
 * 工程・アクション・並列項目のような「主張の中身そのもの」はここに書き、
 * それを裏付ける事実がある場合は `evidenceRefs` で `evidence[]` を参照します。
 *
 * 種類を追加する場合は (1) ここに interface を足し (2) union に加え
 * (3) schema.ts の SemanticContentSchema に 1 バリアント足します。
 */
export const SEMANTIC_CONTENT_KINDS = ["sequence", "points"] as const;
export type SemanticContentKind = (typeof SEMANTIC_CONTENT_KINDS)[number];

export interface ContentItem {
  /** 短い見出し語 */
  label: string;
  /** 説明文（任意） */
  description?: string;
  /** この項目を裏付ける `evidence[].id` への参照 */
  evidenceRefs?: string[];
}

/** 順序が意味を持つ列（工程・時系列） */
export interface SequenceContent {
  kind: "sequence";
  steps: ContentItem[];
}

/** 順序が意味を持たない並列項目（論点・理由・アクション） */
export interface PointsContent {
  kind: "points";
  items: ContentItem[];
}

export type SemanticContent = SequenceContent | PointsContent;

export interface SlideSpec {
  id: string;
  /** そのスライドで最も伝えたい 1 文 */
  keyMessage: string;
  /** KeyMessage をもう一段説明する 1〜2 文 */
  supportingMessage?: string;
  slideRole: SlideRole;
  logicalRelation: LogicalRelation;
  /** 見せる対象となる情報構造（工程・並列項目など）。根拠ではない。 */
  content?: SemanticContent;
  /** 主張を裏付ける根拠。スライドに載せる情報の容器ではない。 */
  evidence: Evidence[];
  sources?: Source[];
}

export interface PresentationSpec {
  title: string;
  subtitle?: string;
  author?: string;
  slides: SlideSpec[];
}

/** `Evidence` を type で絞り込むための型ガード群 */
export const isMetric = (e: Evidence): e is MetricEvidence => e.type === "metric";
export const isCaseStudy = (e: Evidence): e is CaseStudyEvidence => e.type === "caseStudy";
export const isTextEvidence = (e: Evidence): e is TextEvidence => e.type === "text";
export const isComparison = (e: Evidence): e is ComparisonEvidence => e.type === "comparison";
