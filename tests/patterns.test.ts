import { describe, expect, it } from "vitest";
import type { BeforeAfterBindings } from "../src/patterns/before-after/index.js";
import { beforeAfterPattern } from "../src/patterns/before-after/index.js";
import type { MessageOnlyBindings } from "../src/patterns/message-only/index.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { checkConstraints } from "../src/patterns/pattern.js";
import type { ProcessBindings } from "../src/patterns/process/index.js";
import { processPattern } from "../src/patterns/process/index.js";
import { PatternRegistry, builtinPatterns } from "../src/patterns/registry.js";
import type { ThreeReasonsBindings } from "../src/patterns/three-reasons/index.js";
import { threeReasonsPattern } from "../src/patterns/three-reasons/index.js";
import { metricEvidence, slide, textEvidence } from "./helpers.js";

describe("Pattern レジストリ", () => {
  it("全 Pattern が必須メタデータを持つ", () => {
    for (const pattern of builtinPatterns) {
      expect(pattern.id).toMatch(/^[a-z0-9-]+$/);
      expect(pattern.version).toMatch(/^\d+\.\d+\.\d+$/);
      expect(pattern.variants.length).toBeGreaterThan(0);
      expect(pattern.variants).toContain("standard");
      expect(pattern.compatibility.logicalRelations.length).toBeGreaterThan(0);
    }
  });

  it("id の重複を検出する", () => {
    expect(() => new PatternRegistry([messageOnlyPattern, messageOnlyPattern])).toThrow(
      /重複/,
    );
  });

  it("fallback が 1 つも無いレジストリを拒否する", () => {
    expect(() => new PatternRegistry([beforeAfterPattern])).toThrow(/fallback/);
  });
});

describe("Pattern compatibility", () => {
  it("process は process / timeline にのみ対応する", () => {
    expect(processPattern.compatibility.logicalRelations).toEqual(["process", "timeline"]);
    expect(processPattern.compatibility.logicalRelations).not.toContain("claim_evidence");
  });

  it("message-only は全 logicalRelation に対応する fallback である", () => {
    expect(messageOnlyPattern.fallback).toBe(true);
    expect(messageOnlyPattern.compatibility.logicalRelations).toContain("hierarchy");
    expect(messageOnlyPattern.compatibility.logicalRelations).toContain("correlation");
  });
});

describe("Pattern constraints", () => {
  it("three-reasons は evidence 2 件未満を拒否する", () => {
    const violations = checkConstraints(
      slide({ evidence: [textEvidence("e1", "根拠1")] }),
      threeReasonsPattern.constraints,
    );
    expect(violations).toHaveLength(1);
    expect(violations[0]!.constraint).toBe("minEvidence");
  });

  it("three-reasons は evidence 5 件を拒否する", () => {
    const evidence = Array.from({ length: 5 }, (_, i) => textEvidence(`e${i}`, `根拠${i}`));
    const violations = checkConstraints(slide({ evidence }), threeReasonsPattern.constraints);
    expect(violations.map((v) => v.constraint)).toContain("maxEvidence");
  });

  it("keyMessage が長すぎる場合を拒否する", () => {
    const violations = checkConstraints(
      slide({ keyMessage: "あ".repeat(120), evidence: [textEvidence("e1", "根拠")] }),
      beforeAfterPattern.constraints,
    );
    expect(violations.map((v) => v.constraint)).toContain("maxKeyMessageLength");
  });

  it("制約を満たすなら空配列を返す", () => {
    const evidence = [textEvidence("e1", "根拠1"), textEvidence("e2", "根拠2")];
    expect(checkConstraints(slide({ evidence }), threeReasonsPattern.constraints)).toEqual([]);
  });
});

describe("BeforeAfter", () => {
  const caseStudySlide = slide({
    id: "s-ba",
    keyMessage: "導入でコストを30%削減できる",
    slideRole: "evidence",
    logicalRelation: "before_after",
    evidence: [
      {
        id: "e1",
        type: "caseStudy",
        statement: "A社では30%削減された",
        data: {
          organization: "A社",
          before: 100,
          after: 70,
          unit: "百万円",
          beforeLabel: "導入前",
          afterLabel: "導入後",
        },
      },
    ],
    sources: [{ id: "src", title: "A社事例", url: "https://example.com" }],
  });

  it("before/after を抽出できるスライドに適合する", () => {
    const result = beforeAfterPattern.match(caseStudySlide);
    expect(result.compatible).toBe(true);
    expect(result.score).toBeGreaterThan(0.8);
  });

  it("2 値を抽出できないスライドには適合しない", () => {
    const result = beforeAfterPattern.match(
      slide({ logicalRelation: "before_after", evidence: [textEvidence("e1", "定性的な変化")] }),
    );
    expect(result.compatible).toBe(false);
    expect(result.score).toBe(0);
  });

  it("metric 2 件からも before/after を組み立てられる", () => {
    const result = beforeAfterPattern.match(
      slide({
        logicalRelation: "before_after",
        evidence: [
          metricEvidence("e1", "改善前", 40, "時間"),
          metricEvidence("e2", "改善後", 12, "時間"),
        ],
      }),
    );
    expect(result.compatible).toBe(true);
  });

  it("buildBindings が Pattern 専用の構造を返す（SlideSpec を含まない）", () => {
    const bindings = beforeAfterPattern.buildBindings(
      caseStudySlide,
      "standard",
    ) as BeforeAfterBindings;

    expect(bindings).toEqual({
      eyebrow: "根拠",
      headline: "導入でコストを30%削減できる",
      subline: undefined,
      before: { label: "導入前", value: 100 },
      after: { label: "導入後", value: 70 },
      change: { label: "変化", value: "-30%" },
      unit: "百万円",
      sourceText: "出典: A社事例",
    });

    // bindings は logicalRelation / slideRole / evidence をそのまま持ち込まない
    expect(JSON.stringify(bindings)).not.toContain("before_after");
    expect(JSON.stringify(bindings)).not.toContain("caseStudy");
  });

  it("changeLabel が明示されていればそれを使う", () => {
    const bindings = beforeAfterPattern.buildBindings(
      slide({
        logicalRelation: "before_after",
        evidence: [
          {
            id: "e1",
            type: "caseStudy",
            statement: "削減された",
            data: { before: 100, after: 70, changeLabel: "▲30%" },
          },
        ],
      }),
      "standard",
    ) as BeforeAfterBindings;
    expect(bindings.change?.value).toBe("▲30%");
  });

  it("抽出できないスライドで buildBindings を呼ぶと例外になる", () => {
    expect(() =>
      beforeAfterPattern.buildBindings(slide({ evidence: [] }), "standard"),
    ).toThrow(/before\/after を抽出できません/);
  });
});

describe("ThreeReasons", () => {
  const s = slide({
    keyMessage: "3 つの理由から導入すべきである",
    evidence: [
      textEvidence("e1", "理由A", "理由Aの詳細な説明"),
      textEvidence("e2", "理由B", "理由Bの詳細な説明"),
      metricEvidence("e3", "削減率", 30, "%"),
    ],
  });

  it("evidence 3 件で高スコアになる", () => {
    const result = threeReasonsPattern.match(s);
    expect(result.compatible).toBe(true);
    expect(result.score).toBeGreaterThan(0.9);
  });

  it("buildBindings が 1 始まりの連番と説明を作る", () => {
    const bindings = threeReasonsPattern.buildBindings(s, "standard") as ThreeReasonsBindings;
    expect(bindings.reasons.map((r) => r.index)).toEqual([1, 2, 3]);
    expect(bindings.reasons[0]!.label).toBe("理由A");
    expect(bindings.reasons[0]!.description).toBe("理由Aの詳細な説明");
    expect(bindings.reasons[2]!.metric).toBe("30%");
    expect(bindings.reasons[0]!.metric).toBeUndefined();
  });
});

describe("Process", () => {
  const s = slide({
    keyMessage: "4 工程で自動化する",
    slideRole: "solution",
    logicalRelation: "process",
    evidence: [
      textEvidence("e1", "入荷登録", "入荷時にスキャンする"),
      textEvidence("e2", "在庫照合", "自動照合する"),
      textEvidence("e3", "発注判定", "発注点を判定する"),
    ],
  });

  it("process 関係で高スコアになる", () => {
    const result = processPattern.match(s);
    expect(result.compatible).toBe(true);
    expect(result.score).toBeGreaterThan(0.9);
  });

  it("buildBindings が evidence の並び順をそのまま工程順にする", () => {
    const bindings = processPattern.buildBindings(s, "standard") as ProcessBindings;
    expect(bindings.steps.map((step) => step.label)).toEqual([
      "入荷登録",
      "在庫照合",
      "発注判定",
    ]);
    expect(bindings.steps.map((step) => step.index)).toEqual([1, 2, 3]);
    expect(bindings.stepPrefix).toBe("STEP");
  });

  it("timeline では STEP 接頭辞を付けない", () => {
    const bindings = processPattern.buildBindings(
      slide({ ...s, logicalRelation: "timeline" }),
      "standard",
    ) as ProcessBindings;
    expect(bindings.stepPrefix).toBe("");
  });
});

describe("MessageOnly", () => {
  it("evidence が無いスライドでスコアが上がる", () => {
    const withNone = messageOnlyPattern.match(slide({ evidence: [] }));
    const withThree = messageOnlyPattern.match(
      slide({
        evidence: [textEvidence("e1", "A"), textEvidence("e2", "B"), textEvidence("e3", "C")],
      }),
    );
    expect(withNone.score).toBeGreaterThan(withThree.score);
  });

  it("evidence を最大 2 件まで補足として取り込む", () => {
    const bindings = messageOnlyPattern.buildBindings(
      slide({
        evidence: [textEvidence("e1", "A"), textEvidence("e2", "B"), textEvidence("e3", "C")],
      }),
      "standard",
    ) as MessageOnlyBindings;
    expect(bindings.points).toHaveLength(2);
  });
});
