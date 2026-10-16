import { z } from "zod";

/** The selected quantity is measured once, or averaged over events on one fixed corpus. */
export type PartitionAggregation =
  | { kind: "single" }
  | { kind: "mean"; eventCount: number; eventUnit: string; wholeInvariant: true };

export interface UnknownQualityCell {
  status: "unknown";
  question: string;
}

/**
 * A bounded quantitative partition with explicitly unknown quality in both parts.
 * Unit, population and quality criterion are shared by construction. No quality
 * ratio can be inferred from the known selected quantity.
 */
export interface PartitionEvidence {
  id: string;
  type: "partition";
  statement: string;
  sourceIds: string[];
  data: {
    whole: number;
    selected: number;
    unit: string;
    population: string;
    aggregation: PartitionAggregation;
    labels: { whole: string; selected: string; remainder: string };
    quality: {
      criterion: string;
      unknownScope: string;
      selected: UnknownQualityCell;
      remainder: UnknownQualityCell;
    };
  };
}

export interface PartitionAnnotation {
  text: string;
  whyNeeded: string;
}

/** Author interpretation, not a proof of correctness, recall or semantic MECE. */
export interface PartitionArgument {
  kind: "partition-with-unknowns";
  evidenceId: string;
  scope: string;
  partitionAxis: string;
  inference: string;
  visualPurpose: string;
  annotations: {
    quantity: PartitionAnnotation;
    quality: PartitionAnnotation;
    inference: PartitionAnnotation;
  };
}

const text = (max: number) => z.string().trim().min(1).max(max);

export const PartitionAggregationSchema = z.discriminatedUnion("kind", [
  z.object({ kind: z.literal("single") }).strict(),
  z.object({
    kind: z.literal("mean"),
    eventCount: z.number().int().positive().safe(),
    eventUnit: text(40),
    // A mean selected count is comparable to the corpus only when its whole
    // remains fixed for every event. This does not assert equal aggregation.
    wholeInvariant: z.literal(true),
  }).strict(),
]);

const UnknownQualityCellSchema = z.object({
  status: z.literal("unknown"),
  question: text(120),
}).strict();

/** Keep the outer object unrefined so it can join EvidenceSchema's discriminated union. */
export const PartitionEvidenceSchema = z.object({
  id: text(120),
  type: z.literal("partition"),
  statement: text(240),
  sourceIds: z.array(text(120)).min(1).max(8)
    .refine(ids => new Set(ids).size === ids.length, "partition sourceIds must be unique"),
  data: z.object({
    whole: z.number().finite().positive(),
    selected: z.number().finite().nonnegative(),
    unit: text(40),
    population: text(140),
    aggregation: PartitionAggregationSchema,
    labels: z.object({ whole: text(80), selected: text(80), remainder: text(80) }).strict(),
    quality: z.object({
      criterion: text(140),
      unknownScope: text(180),
      selected: UnknownQualityCellSchema,
      remainder: UnknownQualityCellSchema,
    }).strict(),
  }).strict().superRefine((data, ctx) => {
    if (data.selected > data.whole) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["selected"],
        message: "partition selected must not exceed the invariant whole" });
    }
  }),
}).strict();

const PartitionAnnotationSchema = z.object({ text: text(160), whyNeeded: text(160) }).strict();

export const PartitionArgumentSchema = z.object({
  kind: z.literal("partition-with-unknowns"),
  evidenceId: text(120),
  scope: text(160),
  partitionAxis: text(120),
  inference: text(240),
  visualPurpose: text(200),
  annotations: z.object({
    quantity: PartitionAnnotationSchema,
    quality: PartitionAnnotationSchema,
    inference: PartitionAnnotationSchema,
  }).strict(),
}).strict();

export interface PartitionCalculation {
  remainder: number;
  selectedFraction: number;
  remainderFraction: number;
  /** Percent values rounded to one decimal place; fractions retain full precision. */
  selectedPercent: number;
  remainderPercent: number;
  trace: {
    inputs: { whole: number; selected: number; unit: string; population: string };
    formulas: {
      remainder: "whole - selected";
      selectedFraction: "selected / whole";
      remainderFraction: "(whole - selected) / whole";
      selectedPercent: "round(selectedFraction * 100, 1)";
      remainderPercent: "round(remainderFraction * 100, 1)";
    };
    aggregation: PartitionAggregation;
  };
}

/** Runtime shape for persisted calculations; consumers must also recheck arithmetic. */
export const PartitionCalculationSchema = z.object({
  remainder: z.number().finite().nonnegative(),
  selectedFraction: z.number().finite().min(0).max(1),
  remainderFraction: z.number().finite().min(0).max(1),
  selectedPercent: z.number().finite().min(0).max(100),
  remainderPercent: z.number().finite().min(0).max(100),
  trace: z.object({
    inputs: z.object({
      whole: z.number().finite().positive(),
      selected: z.number().finite().nonnegative(),
      unit: text(40),
      population: text(140),
    }).strict(),
    formulas: z.object({
      remainder: z.literal("whole - selected"),
      selectedFraction: z.literal("selected / whole"),
      remainderFraction: z.literal("(whole - selected) / whole"),
      selectedPercent: z.literal("round(selectedFraction * 100, 1)"),
      remainderPercent: z.literal("round(remainderFraction * 100, 1)"),
    }).strict(),
    aggregation: PartitionAggregationSchema,
  }).strict(),
}).strict();

/** Pure, fixed arithmetic over validated inputs. Formula text is never evaluated. */
export function computePartition(evidence: PartitionEvidence): PartitionCalculation {
  const { whole, selected, unit, population, aggregation } = PartitionEvidenceSchema.parse(evidence).data;
  const remainder = whole - selected;
  const selectedFraction = selected / whole;
  const remainderFraction = remainder / whole;
  return {
    remainder,
    selectedFraction,
    remainderFraction,
    selectedPercent: Math.round(selectedFraction * 1000) / 10,
    remainderPercent: Math.round(remainderFraction * 1000) / 10,
    trace: {
      inputs: { whole, selected, unit, population },
      formulas: {
        remainder: "whole - selected",
        selectedFraction: "selected / whole",
        remainderFraction: "(whole - selected) / whole",
        selectedPercent: "round(selectedFraction * 100, 1)",
        remainderPercent: "round(remainderFraction * 100, 1)",
      },
      aggregation: { ...aggregation },
    },
  };
}
