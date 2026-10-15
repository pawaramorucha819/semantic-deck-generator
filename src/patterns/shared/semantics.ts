import {
  isCaseStudy,
  isComparison,
  isMetric,
  type Evidence,
  type MetricEvidence,
  type SlideRole,
  type SlideSpec,
} from "../../domain/types.js";
import { isRelationshipContent } from "../../domain/content.js";

/**
 * SlideSpec から Pattern 用の素材を取り出すためのヘルパ群。
 * ここは「意味 → 表現」の変換に使われるので buildBindings からのみ呼ばれます。
 */

const ROLE_LABELS: Record<SlideRole, string> = {
  context: "背景",
  problem: "課題",
  analysis: "分析",
  insight: "示唆",
  evidence: "根拠",
  comparison: "比較",
  solution: "解決策",
  recommendation: "提言",
  decision: "意思決定",
  summary: "まとめ",
};

export function roleLabel(role: SlideRole): string {
  return ROLE_LABELS[role];
}

/** ロケール非依存の数値整形（同一入力 → 同一出力を保証するため toLocaleString は使わない） */
export function formatNumber(value: number): string {
  const fixed = Number.isInteger(value) ? String(value) : String(Number(value.toFixed(2)));
  const [intPart = "", decimalPart] = fixed.split(".");
  const sign = intPart.startsWith("-") ? "-" : "";
  const digits = sign ? intPart.slice(1) : intPart;
  const withCommas = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return decimalPart ? `${sign}${withCommas}.${decimalPart}` : `${sign}${withCommas}`;
}

export function formatValue(value: number | string, unit?: string): string {
  const base = typeof value === "number" ? formatNumber(value) : value;
  return unit ? `${base}${unit}` : base;
}

/** before → after の変化率を "-30%" 形式で返す */
export function formatChangeRate(before: number, after: number): string | undefined {
  if (before === 0) return undefined;
  const rate = ((after - before) / Math.abs(before)) * 100;
  const rounded = Math.round(rate * 10) / 10;
  if (rounded === 0) return "±0%";
  return `${rounded > 0 ? "+" : ""}${formatNumber(rounded)}%`;
}

/** 出典行の生成。sources が無ければ undefined。 */
export function buildSourceText(slide: SlideSpec): string | undefined {
  const sources = slide.sources ?? [];
  if (sources.length === 0) return undefined;
  const parts = sources.map((s) => {
    const bits = [s.title];
    if (s.publisher) bits.push(s.publisher);
    if (s.date) bits.push(s.date);
    return bits.join(", ");
  });
  return `出典: ${parts.join(" / ")}`;
}

/** Evidence から短い見出し語を取り出す（無ければ statement を切り詰める） */
export function evidenceLabel(evidence: Evidence, fallbackMaxLength = 22): string {
  if (evidence.type === "text" && evidence.data?.label) return evidence.data.label;
  if (evidence.type === "metric" && evidence.data.label) return evidence.data.label;
  if (evidence.type === "caseStudy" && evidence.data.organization) {
    return evidence.data.organization;
  }
  return truncate(evidence.statement, fallbackMaxLength);
}

/** Evidence の説明文（label と重複する場合は detail 側を優先） */
export function evidenceDetail(evidence: Evidence): string {
  if (evidence.type === "text" && evidence.data?.detail) return evidence.data.detail;
  return evidence.statement;
}

export function truncate(text: string, max: number): string {
  const trimmed = text.trim();
  return trimmed.length > max ? `${trimmed.slice(0, max - 1)}…` : trimmed;
}

export interface BeforeAfterFacts {
  before: { label: string; value: number | string };
  after: { label: string; value: number | string };
  change?: { label?: string; value: string };
  unit?: string;
  /** 由来（reason 文の生成に使う） */
  origin: "caseStudy" | "comparison" | "metricPair";
  sourceEvidenceIds: string[];
}

/**
 * before / after 構造を SlideSpec から抽出する。
 * 抽出できない = BeforeAfter パターンが使えない、という判定にそのまま使う。
 */
export function extractBeforeAfter(slide: SlideSpec): BeforeAfterFacts | undefined {
  // 配列に 2 値あるだけでは時間順序は分からない。矢印・変化率を使うには
  // 入力側で before_after と明示されている必要がある。
  if (slide.logicalRelation !== "before_after" || isRelationshipContent(slide.content)) {
    return undefined;
  }

  const caseStudy = slide.evidence
    .filter(isCaseStudy)
    .find((e) => e.data.before !== undefined && e.data.after !== undefined);
  if (caseStudy) {
    const { before, after, unit, beforeLabel, afterLabel, changeLabel } = caseStudy.data;
    const change = changeLabel ?? formatChangeRate(before!, after!);
    return {
      before: { label: beforeLabel ?? "変更前", value: before! },
      after: { label: afterLabel ?? "変更後", value: after! },
      change: change ? { label: "変化", value: change } : undefined,
      unit,
      origin: "caseStudy",
      sourceEvidenceIds: [caseStudy.id],
    };
  }

  const comparison = slide.evidence.filter(isComparison).find((e) => e.data.items.length === 2);
  if (comparison) {
    const [first, second] = comparison.data.items as [
      { label: string; value: number | string },
      { label: string; value: number | string },
    ];
    const change =
      typeof first.value === "number" && typeof second.value === "number"
        ? formatChangeRate(first.value, second.value)
        : undefined;
    return {
      before: { label: first.label, value: first.value },
      after: { label: second.label, value: second.value },
      change: change ? { label: "変化", value: change } : undefined,
      unit: comparison.data.unit,
      origin: "comparison",
      sourceEvidenceIds: [comparison.id],
    };
  }

  const metrics = slide.evidence.filter(isMetric);
  if (metrics.length === 2) {
    const [first, second] = metrics as [(typeof metrics)[number], (typeof metrics)[number]];
    // 異なる単位（片方だけ単位がある場合も含む）を 1 本の変化として扱わない。
    if (first.data.unit !== second.data.unit) return undefined;
    return {
      before: { label: first.data.label ?? "変更前", value: first.data.value },
      after: { label: second.data.label ?? "変更後", value: second.data.value },
      change: (() => {
        const rate = formatChangeRate(first.data.value, second.data.value);
        return rate ? { label: "変化", value: rate } : undefined;
      })(),
      unit: first.data.unit,
      origin: "metricPair",
      sourceEvidenceIds: [first.id, second.id],
    };
  }

  return undefined;
}

/**
 * 「この 1 つの数字を大きく見せればスライドが成立する」ケースを抽出する。
 * metric が 1 件だけ、というのが判定条件。
 */
export function extractHeadlineMetric(slide: SlideSpec): MetricEvidence | undefined {
  const metrics = slide.evidence.filter(isMetric);
  return metrics.length === 1 ? metrics[0] : undefined;
}

export interface ChartFacts {
  categories: string[];
  values: number[];
  seriesName: string;
  unit?: string;
  origin: "comparison" | "metricSeries";
  sourceEvidenceIds: string[];
}

/**
 * グラフ化できる系列を抽出する。
 * 「同じ軸で 2 つ以上の値が並んでいる」ことが条件で、抽出できなければグラフにしない。
 */
export function extractChartSeries(slide: SlideSpec): ChartFacts | undefined {
  // 同一軸の値だけでは part/whole・相関・分布は表現できない。
  // 対応する意味モデルと描画が揃うまでは棒グラフで代用せず abstain する。
  if (
    isRelationshipContent(slide.content) ||
    (slide.logicalRelation !== "claim_evidence" &&
      slide.logicalRelation !== "comparison" &&
      slide.logicalRelation !== "timeline")
  ) {
    return undefined;
  }

  const comparison = slide.evidence
    .filter(isComparison)
    .find((e) => e.data.items.every((item) => typeof item.value === "number"));
  if (comparison) {
    return {
      categories: comparison.data.items.map((item) => item.label),
      values: comparison.data.items.map((item) => item.value as number),
      seriesName: truncate(comparison.statement, 24),
      unit: comparison.data.unit,
      origin: "comparison",
      sourceEvidenceIds: [comparison.id],
    };
  }

  // 単位が揃った metric が 3 件以上並んでいれば系列として読める
  const metrics = slide.evidence.filter(isMetric);
  if (metrics.length >= 3) {
    const unit = metrics[0]!.data.unit;
    if (metrics.every((m) => m.data.unit === unit)) {
      return {
        categories: metrics.map((m, i) => m.data.label ?? `項目${i + 1}`),
        values: metrics.map((m) => m.data.value),
        seriesName: truncate(slide.keyMessage, 24),
        unit,
        origin: "metricSeries",
        sourceEvidenceIds: metrics.map((m) => m.id),
      };
    }
  }

  return undefined;
}
