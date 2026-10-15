import { describe, expect, it } from "vitest";
import type { LogicalRelation } from "../src/domain/types.js";
import { beforeAfterPattern } from "../src/patterns/before-after/index.js";
import { chartWithMessagePattern } from "../src/patterns/chart-with-message/index.js";
import { extractBeforeAfter, extractChartSeries } from "../src/patterns/shared/semantics.js";
import { evaluatePattern, selectPattern } from "../src/selector/selector.js";
import { metricEvidence, slide } from "./helpers.js";

function comparison(logicalRelation: LogicalRelation = "comparison") {
  return slide({
    logicalRelation,
    evidence: [{
      id: "comparison",
      type: "comparison",
      statement: "AとBの売上を比較する",
      data: {
        unit: "万円",
        items: [{ label: "A", value: 100 }, { label: "B", value: 140 }],
      },
    }],
  });
}

describe("Chart intent safety", () => {
  it.each(["comparison", "claim_evidence"] as const)(
    "%s: unordered comparisons have bars without invented change or chronology",
    (relation) => {
      const input = comparison(relation);
      const bindings = chartWithMessagePattern.buildBindings(input, "standard");
      expect(bindings.chart.kind).toBe("bar");
      expect(bindings.chart.categories).toEqual(["A", "B"]);
      expect(bindings.changeText).toBeUndefined();
      expect(chartWithMessagePattern.match(input).reason).not.toContain("推移");
      expect(selectPattern(input).patternId).toBe("chart-with-message");
      expect(beforeAfterPattern.match(input).compatible).toBe(false);
    },
  );

  it("date-like labels do not supply missing timeline intent", () => {
    const input = slide({
      evidence: [{
        id: "comparison",
        type: "comparison",
        statement: "年度別の売上",
        data: { items: [{ label: "2024", value: 10 }, { label: "2025", value: 20 }] },
      }],
    });
    expect(chartWithMessagePattern.buildBindings(input, "standard").changeText).toBeUndefined();
  });

  it("only explicit timeline intent permits a first-to-last change", () => {
    const input = comparison("timeline");
    const bindings = chartWithMessagePattern.buildBindings(input, "standard");
    expect(bindings.chart.categories).toEqual(["A", "B"]);
    expect(bindings.changeText).toBe("100万円 → 140万円 (+40%)");
    expect(chartWithMessagePattern.bindingsSchema.safeParse(bindings).success).toBe(true);
  });

  it("an explicit timeline with a zero baseline does not produce a rate", () => {
    const input = slide({
      logicalRelation: "timeline",
      evidence: [{
        id: "comparison",
        type: "comparison",
        statement: "開始時点からの推移",
        data: { items: [{ label: "開始", value: 0 }, { label: "翌月", value: 20 }] },
      }],
    });
    expect(chartWithMessagePattern.buildBindings(input, "standard").changeText).toBeUndefined();
  });

  it.each(["correlation", "composition", "distribution"] as const)(
    "%s is not inferred from a one-dimensional numeric comparison",
    (relation) => {
      const input = comparison(relation);
      expect(evaluatePattern(chartWithMessagePattern, input).rejectedAt).toBe("compatibility");
      expect(chartWithMessagePattern.match(input).compatible).toBe(false);
      expect(extractChartSeries(input)).toBeUndefined();
      expect(() => chartWithMessagePattern.buildBindings(input, "standard")).toThrow();
      expect(selectPattern(input).patternId).not.toBe("chart-with-message");
    },
  );

  it("matching metric units allow comparison, but never imply a first-to-last change", () => {
    const input = slide({
      evidence: [
        metricEvidence("a", "拠点A", 10, "%"),
        metricEvidence("b", "拠点B", 20, "%"),
        metricEvidence("c", "拠点C", 30, "%"),
      ],
    });
    expect(chartWithMessagePattern.match(input).compatible).toBe(true);
    expect(chartWithMessagePattern.buildBindings(input, "standard").changeText).toBeUndefined();
  });

  it("structured content cannot be discarded in favor of an evidence chart", () => {
    const input = comparison();
    input.content = { kind: "convergence", inputs: [{ label: "A" }, { label: "B" }], outcome: { label: "C" } };
    expect(extractChartSeries(input)).toBeUndefined();
    expect(chartWithMessagePattern.match(input).compatible).toBe(false);
  });
});

describe("Before-after intent safety", () => {
  it("explicit before_after permits an ordered comparison pair", () => {
    const input = comparison("before_after");
    expect(selectPattern(input).patternId).toBe("before-after");
    expect(beforeAfterPattern.buildBindings(input, "standard").change?.value).toBe("+40%");
  });

  it("two same-unit metrics do not imply before_after", () => {
    const input = slide({
      evidence: [metricEvidence("a", "A", 10), metricEvidence("b", "B", 20)],
    });
    expect(extractBeforeAfter(input)).toBeUndefined();
    expect(beforeAfterPattern.match(input).compatible).toBe(false);
    expect(() => beforeAfterPattern.buildBindings(input, "standard")).toThrow();
  });

  it.each([
    ["万円", "年"],
    ["万円", undefined],
    [undefined, "万円"],
  ])("different metric units (%s, %s) cannot be a chronological change", (firstUnit, secondUnit) => {
    const input = slide({
      logicalRelation: "before_after",
      evidence: [
        { id: "a", type: "metric", statement: "前の値", data: { value: 10, unit: firstUnit } },
        { id: "b", type: "metric", statement: "後の値", data: { value: 20, unit: secondUnit } },
      ],
    });
    expect(extractBeforeAfter(input)).toBeUndefined();
    expect(evaluatePattern(beforeAfterPattern, input).rejectedAt).toBe("match");
    expect(() => beforeAfterPattern.buildBindings(input, "standard")).toThrow();
  });

  it("explicit before_after preserves a comparable metric pair", () => {
    const input = slide({
      logicalRelation: "before_after",
      evidence: [
        metricEvidence("a", "改善前", 40, "時間"),
        metricEvidence("b", "改善後", 12, "時間"),
      ],
    });
    const bindings = beforeAfterPattern.buildBindings(input, "standard");
    expect(bindings.before).toEqual({ label: "改善前", value: 40 });
    expect(bindings.after).toEqual({ label: "改善後", value: 12 });
    expect(bindings.unit).toBe("時間");
    expect(bindings.change?.value).toBe("-70%");
  });

  it("case-study states do not override a different slide relation", () => {
    const input = slide({
      logicalRelation: "cause_effect",
      evidence: [{
        id: "case",
        type: "caseStudy",
        statement: "施策と結果",
        data: { before: 100, after: 70 },
      }],
    });
    expect(beforeAfterPattern.match(input).compatible).toBe(false);
    expect(evaluatePattern(beforeAfterPattern, input).rejectedAt).toBe("compatibility");
  });

  it("structured content cannot be discarded in favor of a before-after pair", () => {
    const input = comparison("before_after");
    input.content = { kind: "convergence", inputs: [{ label: "A" }, { label: "B" }], outcome: { label: "C" } };
    expect(extractBeforeAfter(input)).toBeUndefined();
    expect(beforeAfterPattern.match(input).compatible).toBe(false);
  });
});
