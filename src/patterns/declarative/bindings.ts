import { z } from "zod";
import { isMetric, type ContentItem, type Evidence, type SlideSpec } from "../../domain/types.js";
import {
  buildSourceText,
  evidenceDetail,
  evidenceLabel,
  formatValue,
  roleLabel,
  truncate,
} from "../shared/semantics.js";
import type { DeclarativePatternDocument, SlotDefinition } from "./schema.js";

/**
 * Slot 定義 → bindings。
 *
 * ここが Declarative Pattern における「意味 → 表現」の変換点で、
 * コード Pattern の buildBindings と同じ位置づけ。
 * 参照できるのは SlideSpec の意味情報だけで、座標や色は一切扱わない。
 */

export interface ItemBinding {
  index: number;
  label: string;
  description?: string;
  metric?: string;
}

export interface MetricBinding {
  value: string;
  label?: string;
  unit?: string;
}

export type SlotValue = string | ItemBinding[] | MetricBinding;

export type DeclarativeBindings = Record<string, SlotValue | undefined>;

const DEFAULT_ITEM_LABEL_MAX = 40;

function itemsFromContent(slide: SlideSpec, kind: "points" | "sequence" | "any"): ContentItem[] {
  if (!slide.content) return [];
  if (kind === "points") return slide.content.kind === "points" ? slide.content.items : [];
  if (kind === "sequence") return slide.content.kind === "sequence" ? slide.content.steps : [];
  return slide.content.kind === "points" ? slide.content.items : slide.content.steps;
}

function metricOf(item: ContentItem, byId: Map<string, Evidence>): string | undefined {
  for (const ref of item.evidenceRefs ?? []) {
    const evidence = byId.get(ref);
    if (evidence && isMetric(evidence)) {
      return formatValue(evidence.data.value, evidence.data.unit);
    }
  }
  return undefined;
}

function itemsFromEvidence(evidence: Evidence[], labelMax: number): ItemBinding[] {
  return evidence.map((e, i) => ({
    index: i + 1,
    label: evidenceLabel(e, labelMax),
    ...(evidenceDetail(e) ? { description: evidenceDetail(e) } : {}),
    ...(e.type === "metric"
      ? { metric: formatValue(e.data.value, e.data.unit) }
      : e.type === "caseStudy" && e.data.changeLabel
        ? { metric: e.data.changeLabel }
        : {}),
  }));
}

function extractItems(slide: SlideSpec, slot: Extract<SlotDefinition, { type: "items" }>): ItemBinding[] {
  const labelMax = slot.labelMaxLength ?? DEFAULT_ITEM_LABEL_MAX;
  const byId = new Map(slide.evidence.map((e) => [e.id, e]));

  const fromContent = (kind: "points" | "sequence" | "any"): ItemBinding[] =>
    itemsFromContent(slide, kind).map((item, i) => ({
      index: i + 1,
      label: truncate(item.label, labelMax),
      ...(item.description ? { description: item.description } : {}),
      ...(metricOf(item, byId) ? { metric: metricOf(item, byId)! } : {}),
    }));

  let items: ItemBinding[];
  switch (slot.source) {
    case "contentPoints":
      items = fromContent("points");
      break;
    case "contentSequence":
      items = fromContent("sequence");
      break;
    case "content":
      items = fromContent("any");
      break;
    case "evidence":
      items = itemsFromEvidence(slide.evidence, labelMax);
      break;
    case "contentOrEvidence": {
      const content = fromContent("any");
      items = content.length > 0 ? content : itemsFromEvidence(slide.evidence, labelMax);
      break;
    }
  }

  const max = slot.max ?? items.length;
  return items.slice(0, max).map((item, i) => ({ ...item, index: i + 1 }));
}

function extractText(
  slide: SlideSpec,
  slot: Extract<SlotDefinition, { type: "text" }>,
): string | undefined {
  let value: string | undefined;
  switch (slot.source) {
    case "keyMessage":
      value = slide.keyMessage;
      break;
    case "supportingMessage":
      value = slide.supportingMessage;
      break;
    case "slideRoleLabel":
      value = roleLabel(slide.slideRole);
      break;
    case "sourceText":
      value = buildSourceText(slide);
      break;
    case "literal":
      value = slot.value;
      break;
  }
  if (value === undefined) return undefined;
  const trimmed = value.trim();
  if (trimmed.length === 0) return undefined;
  return slot.maxLength ? truncate(trimmed, slot.maxLength) : trimmed;
}

function extractMetric(
  slide: SlideSpec,
  slot: Extract<SlotDefinition, { type: "metric" }>,
): MetricBinding | undefined {
  const metrics = slide.evidence.filter(isMetric);
  if (metrics.length === 0) return undefined;
  if (slot.source === "headlineMetric" && metrics.length !== 1) return undefined;
  const metric = metrics[0]!;
  return {
    value: formatValue(metric.data.value, metric.data.unit),
    ...(metric.data.label ? { label: metric.data.label } : {}),
    ...(metric.data.unit ? { unit: metric.data.unit } : {}),
  };
}

export interface SlotExtractionResult {
  bindings: DeclarativeBindings;
  /** 埋まらなかった必須 slot */
  missing: string[];
  /** min/max を満たさなかった items slot */
  violations: string[];
}

/** SlideSpec から全 slot を抽出する（失敗しても例外にせず結果として返す） */
export function extractSlots(
  doc: DeclarativePatternDocument,
  slide: SlideSpec,
): SlotExtractionResult {
  const bindings: DeclarativeBindings = {};
  const missing: string[] = [];
  const violations: string[] = [];

  for (const slot of doc.slots) {
    if (slot.type === "text") {
      const value = extractText(slide, slot);
      if (value === undefined) {
        if (!slot.optional) missing.push(slot.id);
        continue;
      }
      bindings[slot.id] = value;
      continue;
    }

    if (slot.type === "metric") {
      const value = extractMetric(slide, slot);
      if (value === undefined) {
        if (!slot.optional) missing.push(slot.id);
        continue;
      }
      bindings[slot.id] = value;
      continue;
    }

    const items = extractItems(slide, slot);
    if (items.length === 0) {
      if (!slot.optional) missing.push(slot.id);
      continue;
    }
    if (slot.min !== undefined && items.length < slot.min) {
      violations.push(`slot '${slot.id}' は ${slot.min} 件以上必要ですが ${items.length} 件です`);
      continue;
    }
    bindings[slot.id] = items;
  }

  return { bindings, missing, violations };
}

const ItemBindingSchema = z
  .object({
    index: z.number().int().positive(),
    label: z.string().min(1),
    description: z.string().min(1).optional(),
    metric: z.string().min(1).optional(),
  })
  .strict();

const MetricBindingSchema = z
  .object({
    value: z.string().min(1),
    label: z.string().min(1).optional(),
    unit: z.string().min(1).optional(),
  })
  .strict();

/**
 * slot 定義から bindings の Zod スキーマを生成する。
 *
 * これにより Declarative Pattern もコード Pattern と同じく
 * 「生成時と再生時の両方で bindings が実行時検証される」契約を満たす。
 */
export function buildBindingsSchema(doc: DeclarativePatternDocument): z.ZodType<DeclarativeBindings> {
  const shape: Record<string, z.ZodTypeAny> = {};

  for (const slot of doc.slots) {
    let schema: z.ZodTypeAny;
    if (slot.type === "text") {
      schema = slot.maxLength ? z.string().min(1).max(slot.maxLength) : z.string().min(1);
    } else if (slot.type === "metric") {
      schema = MetricBindingSchema;
    } else {
      let array = z.array(ItemBindingSchema);
      if (slot.min !== undefined) array = array.min(slot.min);
      if (slot.max !== undefined) array = array.max(slot.max);
      else array = array.min(1);
      schema = array;
    }
    shape[slot.id] = slot.optional ? schema.optional() : schema;
  }

  return z.object(shape).strict() as unknown as z.ZodType<DeclarativeBindings>;
}
