import { describe, expect, it } from "vitest";
import { z } from "zod";
import {
  computePartition,
  PartitionArgumentSchema,
  PartitionCalculationSchema,
  PartitionEvidenceSchema,
  type PartitionArgument,
  type PartitionEvidence,
} from "../src/domain/partition.js";
import { EvidenceSchema, parseSlide, SlideSpecSchema } from "../src/domain/schema.js";
import { EVIDENCE_TYPES, isPartition, type SlideSpec } from "../src/domain/types.js";

function evidence(): PartitionEvidence {
  return {
    id: "corpus-partition",
    type: "partition",
    statement: "A mean of 400,000 records was selected from a fixed 7,100,000-record corpus.",
    sourceIds: ["search-report"],
    data: {
      whole: 7_100_000,
      selected: 400_000,
      unit: "records",
      population: "The same indexed corpus for every measured search",
      aggregation: { kind: "mean", eventCount: 100, eventUnit: "searches", wholeInvariant: true },
      labels: { whole: "Corpus", selected: "Selected", remainder: "Not selected" },
      quality: {
        criterion: "Relevance to the search question",
        unknownScope: "Relevant record counts in both parts have not been measured.",
        selected: { status: "unknown", question: "Which selected records are relevant?" },
        remainder: { status: "unknown", question: "Which unselected records are relevant?" },
      },
    },
  };
}

function argument(): PartitionArgument {
  return {
    kind: "partition-with-unknowns",
    evidenceId: "corpus-partition",
    scope: "The fixed corpus across the reported searches",
    partitionAxis: "Selected versus not selected for a search",
    inference: "Selection volume alone cannot establish recall because relevant counts in both parts are unknown.",
    visualPurpose: "Separate observed selection volume from unmeasured relevance.",
    annotations: {
      quantity: { text: "Counts are averaged over searches.", whyNeeded: "Explain the measurement basis." },
      quality: { text: "Relevance is unknown in both parts.", whyNeeded: "Prevent selection from implying correctness." },
      inference: { text: "Recall needs relevant counts in both parts.", whyNeeded: "Explain why volume does not establish recall." },
    },
  };
}

function slide(): SlideSpec {
  return {
    id: "partition-slide",
    keyMessage: "Selection volume does not establish recall.",
    slideRole: "analysis",
    logicalRelation: "composition",
    evidence: [evidence()],
    argument: argument(),
    sources: [{ id: "search-report", title: "Search measurement report", locator: "Table 1" }],
  };
}

describe("bounded partition evidence and argument", () => {
  it("is an unrefined evidence object usable in discriminated unions", () => {
    expect(PartitionEvidenceSchema).toBeInstanceOf(z.ZodObject);
    const union = z.discriminatedUnion("type", [PartitionEvidenceSchema]);
    const parsed: PartitionEvidence = union.parse(evidence());
    expect(parsed).toEqual(evidence());
    expect(EvidenceSchema.parse(evidence())).toEqual(evidence());
    expect(EVIDENCE_TYPES).toContain("partition");
    expect(isPartition(parsed)).toBe(true);
  });

  it.each(["composition", "claim_evidence"] as const)("accepts the bounded %s slide", logicalRelation => {
    const input = { ...slide(), logicalRelation };
    expect(parseSlide(input)).toEqual(input);
    const parsed: PartitionArgument = PartitionArgumentSchema.parse(argument());
    expect(parsed).toEqual(argument());
  });

  it("accepts a single observation, endpoints and fractional event means", () => {
    for (const selected of [0, 123.25, 7_100_000]) {
      const input = evidence();
      input.data.selected = selected;
      expect(PartitionEvidenceSchema.safeParse(input).success).toBe(true);
    }
    const input = evidence();
    input.data.aggregation = { kind: "single" };
    expect(PartitionEvidenceSchema.safeParse(input).success).toBe(true);
  });

  it.each([0, -1, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    "rejects invalid whole %s", whole => {
      const input = evidence();
      input.data.whole = whole;
      expect(PartitionEvidenceSchema.safeParse(input).success).toBe(false);
    },
  );

  it.each([-1, 7_100_001, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    "rejects invalid selected %s", selected => {
      const input = evidence();
      input.data.selected = selected;
      expect(PartitionEvidenceSchema.safeParse(input).success).toBe(false);
      expect(() => computePartition(input)).toThrow();
    },
  );

  it.each([
    { kind: "mean", eventCount: 100, eventUnit: "searches" },
    { kind: "mean", eventCount: 100, eventUnit: "searches", wholeInvariant: false },
    { kind: "mean", eventCount: 0, eventUnit: "searches", wholeInvariant: true },
    { kind: "mean", eventCount: -2, eventUnit: "searches", wholeInvariant: true },
    { kind: "mean", eventCount: 1.5, eventUnit: "searches", wholeInvariant: true },
    { kind: "mean", eventCount: Number.MAX_SAFE_INTEGER + 1, eventUnit: "searches", wholeInvariant: true },
    { kind: "mean", eventCount: 100, wholeInvariant: true },
    { kind: "mean", eventCount: 100, eventUnit: " ", wholeInvariant: true },
    { kind: "median", eventCount: 100, eventUnit: "searches", wholeInvariant: true },
    { kind: "single", eventCount: 1 },
  ])("rejects incomplete or unsupported aggregation %j", aggregation => {
    const input = { ...evidence(), data: { ...evidence().data, aggregation } };
    expect(PartitionEvidenceSchema.safeParse(input).success).toBe(false);
  });

  it.each(["remainder", "selectedFraction", "remainderFraction", "selectedPercent", "qualityRatio", "formula"])(
    "rejects supplied derived or unsupported field %s", key => {
      const input = { ...evidence(), data: { ...evidence().data, [key]: 0.5 } };
      expect(PartitionEvidenceSchema.safeParse(input).success).toBe(false);
    },
  );

  it("requires exactly two explicitly unknown quality cells and shared context", () => {
    const input = evidence();
    const quality = input.data.quality;
    const invalidQualities = [
      { ...quality, selected: { status: "known", question: "Relevant?" } },
      { ...quality, selected: { status: "unknown" } },
      { ...quality, selected: { ...quality.selected, value: 0 } },
      { ...quality, selected: { ...quality.selected, ratio: 0.8 } },
      { ...quality, remainder: { ...quality.remainder, percent: 20 } },
      { ...quality, selected: { ...quality.selected, unit: "other records" } },
      { ...quality, remainder: { ...quality.remainder, population: "another corpus" } },
      { ...quality, selected: { ...quality.selected, criterion: "another criterion" } },
      { ...quality, anotherPart: quality.selected },
      { ...quality, criterion: " " },
      { ...quality, unknownScope: " " },
      { ...quality, remainder: undefined },
    ];
    for (const invalid of invalidQualities) {
      expect(PartitionEvidenceSchema.safeParse({ ...input, data: { ...input.data, quality: invalid } }).success).toBe(false);
    }
    for (const key of ["unit", "population"] as const) {
      expect(PartitionEvidenceSchema.safeParse({ ...input, data: { ...input.data, [key]: " " } }).success).toBe(false);
    }
  });

  it("requires all concise nonempty labels and argument fields", () => {
    const input = evidence();
    for (const key of ["whole", "selected", "remainder"] as const) {
      for (const value of [undefined, " ", "x".repeat(81)]) {
        const labels = { ...input.data.labels, [key]: value };
        expect(PartitionEvidenceSchema.safeParse({ ...input, data: { ...input.data, labels } }).success).toBe(false);
      }
    }
    for (const key of ["kind", "evidenceId", "scope", "partitionAxis", "inference", "visualPurpose"] as const) {
      for (const value of [undefined, " ", "x".repeat(241)]) {
        expect(PartitionArgumentSchema.safeParse({ ...argument(), [key]: value }).success).toBe(false);
      }
    }
    const annotations = argument().annotations;
    for (const key of ["quantity", "quality", "inference"] as const) {
      for (const value of [undefined, { text: "note" }, { text: "note", whyNeeded: " " }, { text: "x".repeat(161), whyNeeded: "reason" }]) {
        expect(PartitionArgumentSchema.safeParse({ ...argument(), annotations: { ...annotations, [key]: value } }).success).toBe(false);
      }
    }
    expect(PartitionArgumentSchema.safeParse({ ...argument(), annotations: { ...annotations, extra: annotations.quantity } }).success).toBe(false);
    expect(PartitionArgumentSchema.safeParse({ ...argument(), annotations: Object.values(annotations) }).success).toBe(false);
  });

  it("requires nonempty unique sources and resolves them through existing slide validation", () => {
    for (const sourceIds of [undefined, [], [" "], ["search-report", "search-report"], ["search-report", " search-report "]]) {
      expect(PartitionEvidenceSchema.safeParse({ ...evidence(), sourceIds }).success).toBe(false);
    }
    expect(SlideSpecSchema.safeParse({ ...slide(), sources: undefined }).success).toBe(false);
    expect(SlideSpecSchema.safeParse({ ...slide(), sources: [{ id: "other", title: "Other source" }] }).success).toBe(false);
    expect(SlideSpecSchema.safeParse({ ...slide(), sources: [...slide().sources!, ...slide().sources!] }).success).toBe(false);
  });

  it("rejects missing, dangling, mismatched or multiply backed arguments", () => {
    const base = slide();
    const invalid = [
      { ...base, argument: undefined },
      { ...base, evidence: [] },
      { ...base, argument: { ...argument(), evidenceId: "other" } },
      { ...base, evidence: [evidence(), { ...evidence(), id: "partition-two" }] },
      { ...base, evidence: [evidence(), { id: "extra", type: "text", statement: "Additional fact" }] },
      { ...base, evidence: [{ id: evidence().id, type: "text", statement: "Not a partition" }] },
    ];
    for (const input of invalid) expect(SlideSpecSchema.safeParse(input).success).toBe(false);
  });

  it("rejects incompatible relation, content and legacy explanation instead of losing nodes", () => {
    const base = slide();
    const invalid = [
      { ...base, logicalRelation: "comparison" },
      { ...base, content: { kind: "points", items: [{ label: "An important node" }] } },
      { ...base, explanation: { readerQuestion: "Why?", sections: [{ heading: "Reason", body: "Important interpretation" }] } },
    ];
    for (const input of invalid) expect(SlideSpecSchema.safeParse(input).success).toBe(false);
  });

  it("preserves optional decomposition as author rationale without proving semantic MECE", () => {
    const input = { ...slide(), decomposition: { scope: "Corpus", criterion: "Selection", rationale: "Each record is selected or not." } };
    expect(parseSlide(input)).toEqual(input);
  });

  it("preserves legacy inputs without imposing new source or argument requirements", () => {
    const input = {
      id: "legacy",
      keyMessage: "Legacy facts remain valid.",
      slideRole: "analysis",
      logicalRelation: "claim_evidence",
      evidence: [
        { id: "text", type: "text", statement: "A fact", sourceIds: ["source", "source"] },
        { id: "metric", type: "metric", statement: "A measure", data: { value: 1 } },
        { id: "case", type: "caseStudy", statement: "A case", data: {} },
        { id: "comparison", type: "comparison", statement: "Comparison", data: { items: [{ label: "A", value: 1 }, { label: "B", value: 2 }] } },
      ],
      sources: [{ id: "source", title: "A" }, { id: "source", title: "B" }],
      content: { kind: "points", items: [{ id: "node", label: "A point", evidenceRefs: ["text"] }] },
      explanation: { readerQuestion: "Why?", sections: [{ heading: "Reason", body: "An explanation", nodeRefs: ["node"] }] },
    };
    expect(parseSlide(input)).toEqual(input);
  });
});

describe("fixed partition calculation", () => {
  it("derives remainder and quantity fractions, with an auditable fixed formula trace", () => {
    const input = evidence();
    const before = structuredClone(input);
    const result = computePartition(input);
    expect(result).toEqual({
      remainder: 6_700_000,
      selectedFraction: 400_000 / 7_100_000,
      remainderFraction: 6_700_000 / 7_100_000,
      selectedPercent: 5.6,
      remainderPercent: 94.4,
      trace: {
        inputs: { whole: 7_100_000, selected: 400_000, unit: "records", population: input.data.population },
        formulas: {
          remainder: "whole - selected",
          selectedFraction: "selected / whole",
          remainderFraction: "(whole - selected) / whole",
          selectedPercent: "round(selectedFraction * 100, 1)",
          remainderPercent: "round(remainderFraction * 100, 1)",
        },
        aggregation: input.data.aggregation,
      },
    });
    expect(PartitionCalculationSchema.parse(result)).toEqual(result);
    expect(computePartition(input)).toEqual(result);
    expect(input).toEqual(before);
    expect(result.trace.aggregation).not.toBe(input.data.aggregation);
    expect(JSON.stringify(result)).not.toMatch(/recall|correctness|qualityRatio|semanticMECE/);
  });

  it.each([
    [0, 7_100_000, 0, 100],
    [7_100_000, 0, 100, 0],
  ])("handles selected=%s without a division or unknown-quality workaround", (selected, remainder, selectedPercent, remainderPercent) => {
    const input = evidence();
    input.data.selected = selected;
    input.data.aggregation = { kind: "single" };
    expect(computePartition(input)).toMatchObject({ remainder, selectedPercent, remainderPercent, trace: { aggregation: { kind: "single" } } });
  });

  it("retains full fractional precision and only rounds displayed percentages", () => {
    const input = evidence();
    input.data.whole = 3;
    input.data.selected = 1;
    const result = computePartition(input);
    expect(result.selectedFraction).toBe(1 / 3);
    expect(result.remainderFraction).toBe(2 / 3);
    expect(result.selectedPercent).toBe(33.3);
    expect(result.remainderPercent).toBe(66.7);
  });

  it("validates unsafe callers rather than calculating invalid input", () => {
    const input = evidence();
    input.data.whole = 0;
    expect(() => computePartition(input)).toThrow();
    expect(() => computePartition({ ...evidence(), data: { ...evidence().data, aggregation: { kind: "mean" } } } as PartitionEvidence)).toThrow();
  });

  it("keeps the persisted calculation strict and rejects custom formulas", () => {
    const result = computePartition(evidence());
    expect(PartitionCalculationSchema.safeParse({ ...result, recall: 0.9 }).success).toBe(false);
    expect(PartitionCalculationSchema.safeParse({ ...result, selectedFraction: 2 }).success).toBe(false);
    expect(PartitionCalculationSchema.safeParse({ ...result, trace: { ...result.trace, formulas: { ...result.trace.formulas, remainder: "eval(input)" } } }).success).toBe(false);
    expect(PartitionCalculationSchema.safeParse({ ...result, trace: { ...result.trace, inputs: { ...result.trace.inputs, correctness: 1 } } }).success).toBe(false);
  });
});
