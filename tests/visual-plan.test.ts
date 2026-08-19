import { describe, expect, it } from "vitest";
import { PatternRegistry } from "../src/patterns/registry.js";
import type { AnyVisualPattern } from "../src/patterns/pattern.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import {
  buildVisualPlan,
  serializeVisualPlan,
  VISUAL_PLAN_SCHEMA_VERSION,
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
    expect(buildVisualPlan(spec).slides[0]!.candidates).toHaveLength(6);
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
});
