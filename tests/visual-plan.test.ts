import { builtinPatterns } from "../src/patterns/registry.js";
import { describe, expect, it } from "vitest";
import { z } from "zod";
import { PatternRegistry } from "../src/patterns/registry.js";
import type { AnyVisualPattern } from "../src/patterns/pattern.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import {
  buildVisualPlan,
  parseVisualPlan,
  serializeVisualPlan,
  VISUAL_PLAN_SCHEMA_VERSION,
  VisualPlanValidationError,
} from "../src/plan/visual-plan.js";
import { loadExample, slide, textEvidence } from "./helpers.js";

describe("Visual Plan 生成", () => {
  it("スライド数ぶんの plan が作られる", async () => {
    const plan = buildVisualPlan(await loadExample());
    expect(plan.schemaVersion).toBe(VISUAL_PLAN_SCHEMA_VERSION);
    expect(plan.title).toBe("在庫管理システム導入提案");
    expect(plan.slides).toHaveLength(6);
    expect(plan.slides.map((s) => s.slideNumber)).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it("各 plan は pattern / variant / bindings を持つ", async () => {
    const plan = buildVisualPlan(await loadExample());
    for (const s of plan.slides) {
      expect(s.slideId).toMatch(/^slide-\d+$/);
      expect(typeof s.patternId).toBe("string");
      expect(s.variantId).toBe("standard");
      expect(s.patternVersion).toMatch(/^\d+\.\d+\.\d+$/);
      expect(s.bindings).toBeTypeOf("object");
      expect(s.reason.length).toBeGreaterThan(0);
    }
  });

  it("before-after の bindings が仕様どおりの形になる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const target = plan.slides.find((s) => s.patternId === "before-after")!;
    expect(target.bindings).toMatchObject({
      headline: "A社では導入後に在庫管理コストを30%削減している",
      before: { label: "導入前", value: 4800 },
      after: { label: "導入12か月後", value: 3360 },
      change: { value: "-30%" },
      unit: "万円",
    });
  });

  it("Visual Plan は JSON としてラウンドトリップできる（描画は plan だけで完結する）", async () => {
    const plan = buildVisualPlan(await loadExample());
    const json = serializeVisualPlan(plan);
    expect(json.endsWith("\n")).toBe(true);
    const parsed = JSON.parse(json);
    expect(parsed).toEqual(JSON.parse(JSON.stringify(plan)));
  });

  it("Visual Plan に座標・色・フォントは含まれない（レイアウトは Renderer の責務）", async () => {
    const json = serializeVisualPlan(buildVisualPlan(await loadExample()));
    for (const forbidden of ['"x"', '"y"', '"w"', '"h"', "fontSize", "fontFace", "#0B4F8A"]) {
      expect(json).not.toContain(forbidden);
    }
  });

  it("候補一覧をデバッグ用に保持できる / 省略もできる", async () => {
    const spec = await loadExample();
    expect(buildVisualPlan(spec).slides[0]!.candidates).toHaveLength(builtinPatterns.length);
    expect(buildVisualPlan(spec, { includeCandidates: false }).slides[0]!.candidates).toBeUndefined();
  });

  it("同じ入力からは同じ Visual Plan が得られる（決定論的）", async () => {
    const spec = await loadExample();
    expect(serializeVisualPlan(buildVisualPlan(spec))).toBe(
      serializeVisualPlan(buildVisualPlan(spec)),
    );
  });

  it("buildBindings が失敗した Pattern は fallback へ退避し notes に記録する", () => {
    const brokenPattern: AnyVisualPattern = {
      id: "broken",
      version: "0.0.1",
      description: "テスト用: 常に buildBindings が失敗する",
      variants: ["standard"],
      compatibility: { logicalRelations: ["claim_evidence"] },
      bindingsSchema: z.any(),
      match: () => ({ compatible: true, score: 1, reason: "テスト用に常に最高スコア" }),
      buildBindings: () => {
        throw new Error("意図的な失敗");
      },
      render: () => undefined,
    };

    const registry = new PatternRegistry([brokenPattern, messageOnlyPattern]);
    const plan = buildVisualPlan(
      { title: "t", slides: [slide({ evidence: [textEvidence("e1", "根拠")] })] },
      { registry },
    );

    expect(plan.slides[0]!.patternId).toBe("message-only");
    expect(plan.slides[0]!.fallbackUsed).toBe(true);
    expect(plan.slides[0]!.notes?.[0]).toContain("意図的な失敗");
  });

  it("bindingsSchema に違反する bindings を作る Pattern も fallback へ退避する", () => {
    const lyingPattern: AnyVisualPattern = {
      id: "lying",
      version: "0.0.1",
      description: "テスト用: schema と実装がずれている",
      variants: ["standard"],
      compatibility: { logicalRelations: ["claim_evidence"] },
      bindingsSchema: z.object({ headline: z.string() }).strict(),
      match: () => ({ compatible: true, score: 1, reason: "テスト用に常に最高スコア" }),
      buildBindings: () => ({ totallyWrongShape: true }),
      render: () => undefined,
    };

    const registry = new PatternRegistry([lyingPattern, messageOnlyPattern]);
    const plan = buildVisualPlan(
      { title: "t", slides: [slide({ evidence: [textEvidence("e1", "根拠")] })] },
      { registry },
    );

    expect(plan.slides[0]!.patternId).toBe("message-only");
    expect(plan.slides[0]!.fallbackUsed).toBe(true);
  });

  /**
   * variant は buildBindings より前に確定されなければならない。
   * 未実装 variant のまま bindings を作ると、Plan に記録された variantId と
   * 実際に bindings を作った variant が食い違い、再生時に別の見た目になる。
   */
  describe("variant の確定順序", () => {
    /** buildBindings に渡された variantId を bindings へ記録するテスト用 Pattern */
    function variantRecordingPattern(proposedVariant: string): AnyVisualPattern {
      return {
        id: "variant-recorder",
        version: "1.0.0",
        description: "テスト用: buildBindings に渡された variant を記録する",
        variants: ["standard"], // compact は未実装
        compatibility: { logicalRelations: ["claim_evidence"] },
        bindingsSchema: z.object({ builtForVariant: z.string() }).strict(),
        match: () => ({
          compatible: true,
          score: 1,
          reason: "テスト用に常に最高スコア",
          variantId: proposedVariant,
        }),
        buildBindings: (_slide, variantId) => ({ builtForVariant: variantId }),
        render: () => undefined,
      };
    }

    function planWith(pattern: AnyVisualPattern) {
      const registry = new PatternRegistry([pattern, messageOnlyPattern]);
      return buildVisualPlan(
        { title: "t", slides: [slide({ evidence: [textEvidence("e1", "根拠")] })] },
        { registry },
      );
    }

    it("未実装 variant が提案されても buildBindings には実装済み variant が渡る", () => {
      const plan = planWith(variantRecordingPattern("compact"));
      const planSlide = plan.slides[0]!;

      expect(planSlide.patternId).toBe("variant-recorder");
      expect(planSlide.variantId).toBe("standard");
      // Plan の variantId と bindings 生成時の variant が一致している
      expect(planSlide.bindings).toEqual({ builtForVariant: "standard" });
    });

    it("variant を補正した理由が notes に残る", () => {
      const plan = planWith(variantRecordingPattern("compact"));
      const notes = plan.slides[0]!.notes ?? [];
      expect(notes.some((n) => n.includes("compact") && n.includes("standard"))).toBe(true);
    });

    it("実装済み variant が提案された場合は補正されず notes も出ない", () => {
      const plan = planWith(variantRecordingPattern("standard"));
      const planSlide = plan.slides[0]!;
      expect(planSlide.variantId).toBe("standard");
      expect(planSlide.bindings).toEqual({ builtForVariant: "standard" });
      expect(planSlide.notes).toBeUndefined();
    });

    it("variant 補正後の Plan は parseVisualPlan / 描画を通過する（不整合が残らない）", () => {
      const pattern = variantRecordingPattern("compact");
      const registry = new PatternRegistry([pattern, messageOnlyPattern]);
      const plan = planWith(pattern);
      expect(() =>
        parseVisualPlan(JSON.parse(serializeVisualPlan(plan)), { registry }),
      ).not.toThrow();
    });

    it("fallback へ退避した場合も fallback Pattern の有効な variant で bindings を作る", () => {
      const brokenWithBadVariant: AnyVisualPattern = {
        id: "broken-variant",
        version: "1.0.0",
        description: "テスト用: variant も不正で bindings 生成も失敗する",
        variants: ["standard"],
        compatibility: { logicalRelations: ["claim_evidence"] },
        bindingsSchema: z.any(),
        match: () => ({
          compatible: true,
          score: 1,
          reason: "テスト用",
          variantId: "nonexistent",
        }),
        buildBindings: () => {
          throw new Error("意図的な失敗");
        },
        render: () => undefined,
      };

      const registry = new PatternRegistry([brokenWithBadVariant, messageOnlyPattern]);
      const plan = buildVisualPlan(
        { title: "t", slides: [slide({ evidence: [textEvidence("e1", "根拠")] })] },
        { registry },
      );
      const planSlide = plan.slides[0]!;

      expect(planSlide.patternId).toBe("message-only");
      expect(planSlide.fallbackUsed).toBe(true);
      expect(messageOnlyPattern.variants).toContain(planSlide.variantId);
      // fallback の bindings も契約に適合している
      expect(() => messageOnlyPattern.bindingsSchema.parse(planSlide.bindings)).not.toThrow();
    });
  });

  it("出典・根拠が provenance として構造化されて残る（URL を失わない）", async () => {
    const plan = buildVisualPlan(await loadExample());
    const caseStudySlide = plan.slides.find((s) => s.slideId === "slide-004")!;

    expect(caseStudySlide.provenance).toBeDefined();
    expect(caseStudySlide.provenance!.sources[0]).toMatchObject({
      id: "source-401",
      url: "https://example.com/case/a-company",
      accessedAt: "2025-02-10",
    });
    expect(caseStudySlide.provenance!.evidence[0]).toMatchObject({
      id: "evidence-401",
      type: "caseStudy",
      sourceIds: ["source-401"],
    });
  });
});

describe("parseVisualPlan（保存済み Plan の再生契約）", () => {
  it("serializeVisualPlan した Plan をそのまま再検証できる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const parsed = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));
    expect(parsed.slides.map((s) => s.patternId)).toEqual(plan.slides.map((s) => s.patternId));
  });

  it("schemaVersion が違う Plan を拒否する", async () => {
    const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
    plan.schemaVersion = "0.9";
    expect(() => parseVisualPlan(plan)).toThrow(VisualPlanValidationError);
    expect(() => parseVisualPlan(plan)).toThrow(/schemaVersion/);
  });

  it("patternVersion が現在の実装と一致しない Plan を拒否する", async () => {
    const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
    plan.slides[0].patternVersion = "0.0.1";
    expect(() => parseVisualPlan(plan)).toThrow(/バージョンが不一致/);
  });

  it("未登録の patternId を持つ Plan を拒否する", async () => {
    const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
    plan.slides[0].patternId = "nonexistent-pattern";
    expect(() => parseVisualPlan(plan)).toThrow(/未登録の patternId/);
  });

  it("改竄された bindings を拒否する（黙って違う出力にならない）", async () => {
    const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
    const target = plan.slides.find((s: { patternId: string }) => s.patternId === "before-after");
    target.bindings.before = { label: "", value: null };
    expect(() => parseVisualPlan(plan)).toThrow(/契約に違反/);
  });

  it("未実装の variantId を持つ Plan を拒否する", async () => {
    const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
    plan.slides[0].variantId = "compact";
    expect(() => parseVisualPlan(plan)).toThrow(/variant/);
  });
});
