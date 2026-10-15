import { builtinPatterns } from "../src/patterns/registry.js";
import { describe, expect, it } from "vitest";
import { beforeAfterPattern } from "../src/patterns/before-after/index.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { processPattern } from "../src/patterns/process/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { evaluatePattern, selectPattern } from "../src/selector/selector.js";
import {
  loadExample,
  metricEvidence,
  sequenceContent,
  slide,
  textEvidence,
} from "./helpers.js";

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
      beforeAfterPattern,
      slide({ logicalRelation: "before_after", evidence: [] }),
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
    expect(candidates.length).toBe(builtinPatterns.length);
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
      content: sequenceContent("工程1", "工程2"),
      evidence: [],
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

  describe("代表的な競合スライドでの pairwise 選択", () => {
    it("before/after を持つ caseStudy では before-after が three-reasons に勝つ", () => {
      const s = slide({
        logicalRelation: "before_after",
        evidence: [
          {
            id: "e1",
            type: "caseStudy",
            statement: "導入後に30%削減された",
            data: { before: 100, after: 70 },
          },
          textEvidence("e2", "補足の事実A"),
          textEvidence("e3", "補足の事実B"),
        ],
      });
      expect(selectPattern(s).patternId).toBe("before-after");
    });

    it("単一 metric では big-number が chart-with-message に勝つ", () => {
      const s = slide({ evidence: [metricEvidence("e1", "回収期間", 1.7, "年")] });
      expect(selectPattern(s).patternId).toBe("big-number");
    });

    it("3 件の並列 text evidence では three-reasons が message-only に勝つ", () => {
      const s = slide({
        evidence: [
          textEvidence("e1", "理由A"),
          textEvidence("e2", "理由B"),
          textEvidence("e3", "理由C"),
        ],
      });
      expect(selectPattern(s).patternId).toBe("three-reasons");
    });

    it("evidence も content も無いスライドでは message-only が正当に選ばれる", () => {
      const s = slide({ slideRole: "summary", evidence: [] });
      const selection = selectPattern(s);
      expect(selection.patternId).toBe("message-only");
      expect(selection.fallbackUsed).toBe(false);
    });
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
