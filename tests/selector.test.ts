import { describe, expect, it } from "vitest";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { processPattern } from "../src/patterns/process/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { evaluatePattern, selectPattern } from "../src/selector/selector.js";
import { loadExample, metricEvidence, slide, textEvidence } from "./helpers.js";

describe("Visual Pattern Selector", () => {
  it("サンプル資料では 6 枚すべてに異なる Pattern が選ばれる", async () => {
    const spec = await loadExample();
    const selected = spec.slides.map((s) => selectPattern(s).patternId);

    expect(selected).toEqual([
      "chart-with-message",
      "three-reasons",
      "process",
      "before-after",
      "big-number",
      "message-only",
    ]);
    expect(new Set(selected).size).toBe(6);
  });

  it("選択結果は patternId / variantId / score / reason を持つ", async () => {
    const spec = await loadExample();
    const selection = selectPattern(spec.slides[3]!);

    expect(selection.patternId).toBe("before-after");
    expect(selection.variantId).toBe("standard");
    expect(selection.score).toBeGreaterThan(0.9);
    expect(selection.score).toBeLessThanOrEqual(1);
    expect(selection.reason).toContain("before_after");
  });

  it("同じ入力からは常に同じ結果になる（決定論的）", async () => {
    const spec = await loadExample();
    for (const s of spec.slides) {
      const a = selectPattern(s);
      const b = selectPattern(s);
      expect(a).toEqual(b);
    }
  });

  it("logicalRelation が対応外の Pattern は compatibility 段階で落ちる", () => {
    const evaluation = evaluatePattern(processPattern, slide({ logicalRelation: "comparison" }));
    expect(evaluation.compatible).toBe(false);
    expect(evaluation.rejectedAt).toBe("compatibility");
    expect(evaluation.score).toBe(0);
  });

  it("constraints 違反は constraints 段階で落ちる", () => {
    const evaluation = evaluatePattern(
      processPattern,
      slide({ logicalRelation: "process", evidence: [textEvidence("e1", "唯一の工程")] }),
    );
    expect(evaluation.compatible).toBe(false);
    expect(evaluation.rejectedAt).toBe("constraints");
    expect(evaluation.reason).toContain("constraints 違反");
  });

  it("match() が不適合を返した場合は match 段階で落ちる", () => {
    const selection = selectPattern(
      slide({
        logicalRelation: "before_after",
        evidence: [textEvidence("e1", "定性的な変化"), textEvidence("e2", "別の変化")],
      }),
    );
    const beforeAfter = selection.candidates.find((c) => c.patternId === "before-after");
    expect(beforeAfter?.rejectedAt).toBe("match");
  });

  it("全候補がスコア降順で返る", async () => {
    const spec = await loadExample();
    const { candidates } = selectPattern(spec.slides[1]!);
    expect(candidates.length).toBe(6);
    for (let i = 1; i < candidates.length; i += 1) {
      expect(candidates[i - 1]!.score).toBeGreaterThanOrEqual(candidates[i]!.score);
    }
  });

  it("適合する Pattern が無ければ fallback へ退避する", () => {
    // message-only を含まない = fallback だけのレジストリを使い、
    // どの Pattern にも適合しないスライドを渡す
    const registry = new PatternRegistry([processPattern, messageOnlyPattern]);
    const selection = selectPattern(
      slide({ logicalRelation: "hierarchy", evidence: [textEvidence("e1", "階層")] }),
      { registry },
    );
    expect(selection.patternId).toBe("message-only");
    expect(selection.candidates.find((c) => c.patternId === "process")?.rejectedAt).toBe(
      "compatibility",
    );
  });

  it("minimumScore を下回る候補は fallback へ回す", () => {
    const s = slide({
      logicalRelation: "process",
      evidence: [textEvidence("e1", "工程1"), textEvidence("e2", "工程2")],
    });
    const strict = selectPattern(s, { minimumScore: 0.999 });
    expect(strict.fallbackUsed).toBe(true);
    expect(strict.patternId).toBe("message-only");
    expect(strict.score).toBe(0);
  });

  it("fallback Pattern がスコアで勝った場合は fallbackUsed=false", () => {
    const selection = selectPattern(
      slide({ slideRole: "recommendation", evidence: [textEvidence("e1", "次のアクション")] }),
    );
    expect(selection.patternId).toBe("message-only");
    expect(selection.fallbackUsed).toBe(false);
  });

  it("evidenceTypes が合わないと compatibilityScore が下がる", () => {
    const withCaseStudy = evaluatePattern(
      { ...processPattern, compatibility: { logicalRelations: ["process"], evidenceTypes: ["metric"] } },
      slide({
        logicalRelation: "process",
        evidence: [metricEvidence("e1", "A", 1), metricEvidence("e2", "B", 2)],
      }),
    );
    const withText = evaluatePattern(
      { ...processPattern, compatibility: { logicalRelations: ["process"], evidenceTypes: ["metric"] } },
      slide({
        logicalRelation: "process",
        evidence: [textEvidence("e1", "A"), textEvidence("e2", "B")],
      }),
    );
    expect(withCaseStudy.compatibilityScore).toBeGreaterThan(withText.compatibilityScore);
  });
});
