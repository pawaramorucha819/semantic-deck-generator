import { describe, expect, it } from "vitest";
import { z } from "zod";
import type { AnyVisualPattern } from "../src/patterns/pattern.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan } from "../src/render/renderer.js";
import type { Placement } from "../src/render/surface.js";
import {
  estimateEffectiveFontSize,
  validatePlacements,
} from "../src/render/validate.js";
import { evaluatePattern, selectPattern } from "../src/selector/selector.js";
import { slide } from "./helpers.js";

/** テスト用 Pattern を最小構成で作る */
function fakePattern(overrides: Partial<AnyVisualPattern>): AnyVisualPattern {
  return {
    id: "fake",
    version: "0.0.1",
    description: "テスト用",
    variants: ["standard"],
    compatibility: { logicalRelations: ["claim_evidence"] },
    bindingsSchema: z.any(),
    match: () => ({ compatible: true, score: 0.5, reason: "テスト用" }),
    buildBindings: () => ({}),
    render: () => undefined,
    ...overrides,
  };
}

describe("Selector のスコア契約", () => {
  it("match() が 1 を超えるスコアを返す Pattern は実行時に拒否される", () => {
    const cheater = fakePattern({
      id: "cheater",
      match: () => ({ compatible: true, score: 1.5, reason: "選択を支配しようとする" }),
    });
    expect(() => evaluatePattern(cheater, slide())).toThrow(/範囲外のスコア/);
  });

  it("match() が負のスコアを返す Pattern も拒否される", () => {
    const negative = fakePattern({
      id: "negative",
      match: () => ({ compatible: true, score: -0.2, reason: "負のスコア" }),
    });
    expect(() => evaluatePattern(negative, slide())).toThrow(/範囲外のスコア/);
  });

  it("match() が NaN を返す Pattern も拒否される", () => {
    const nan = fakePattern({
      id: "nan",
      match: () => ({ compatible: true, score: Number.NaN, reason: "NaN" }),
    });
    expect(() => evaluatePattern(nan, slide())).toThrow(/範囲外のスコア/);
  });

  it("builtin Pattern は全域で契約内のスコアを返す（サンプル外の入力でも例外にならない)", () => {
    // 雑多なスライドで selectPattern が例外を投げないこと = 全 builtin が契約準拠
    const slides = [
      slide(),
      slide({ logicalRelation: "hierarchy" }),
      slide({ keyMessage: "あ".repeat(150), evidence: [] }),
    ];
    for (const s of slides) {
      expect(() => selectPattern(s)).not.toThrow();
    }
  });
});

describe("fallback 契約", () => {
  it("fallback (message-only) は constraints を持たない", () => {
    expect(messageOnlyPattern.constraints).toBeUndefined();
  });

  it("91 文字の keyMessage でも message-only が正規候補として選ばれ、描画できる", () => {
    // 旧実装では maxKeyMessageLength=90 で自ら失格した後、
    // 失格したはずの自分が fallback として使われていた（レビュー H2 の実測ケース）
    const longMessage = slide({
      keyMessage:
        "在庫管理システムの導入は当社の物流コスト構造を抜本的に見直す契機となるだけでなく、拠点間の在庫偏在や欠品対応の遅れといった慢性的な業務課題を同時に解消し、中期経営計画の収益性目標達成にも不可欠な施策であると結論付けられる",
      evidence: [],
    });
    expect(longMessage.keyMessage.length).toBeGreaterThan(90);

    const selection = selectPattern(longMessage);
    expect(selection.patternId).toBe("message-only");
    expect(selection.fallbackUsed).toBe(false); // 失格→退避ではなく正当な選択

    // 描画してもエラー・警告が出ない（見出しサイズが適応的に縮む）
    const plan = buildVisualPlan({ title: "stress", slides: [longMessage] });
    const { issues } = renderVisualPlan(plan, { includeTitleSlide: false });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);
  });

  it("fallback しか残らないレジストリでも必ず描画に到達できる", () => {
    const registry = new PatternRegistry([messageOnlyPattern]);
    const s = slide({ logicalRelation: "hierarchy", keyMessage: "あ".repeat(120) });
    const plan = buildVisualPlan({ title: "t", slides: [s] }, { registry });
    expect(plan.slides[0]!.patternId).toBe("message-only");
    const { issues } = renderVisualPlan(plan, {
      registry,
      includeTitleSlide: false,
    });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);
  });
});

describe("可読性検査（実効フォントサイズ）", () => {
  it("shrink で読める下限を下回るテキストを警告する", () => {
    const placements: Placement[] = [
      {
        kind: "text",
        layer: "content",
        box: { x: 1, y: 1, w: 1.2, h: 0.3 },
        name: "tiny",
        text: "非常に長い説明文が極端に小さなボックスへ押し込まれており自動縮小では読めないサイズまで縮む",
        fontSize: 13,
        lineHeight: 1.3,
        paddingPt: 0,
        shrink: true,
      },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues.map((i) => i.code)).toContain("text-shrink-below-min");
    expect(issues[0]!.severity).toBe("warning");
  });

  it("shrink でも収まる範囲なら警告しない", () => {
    const placements: Placement[] = [
      {
        kind: "text",
        layer: "content",
        box: { x: 1, y: 1, w: 4, h: 1 },
        name: "fits",
        text: "短いテキスト",
        fontSize: 13,
        lineHeight: 1.3,
        paddingPt: 0,
        shrink: true,
      },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues).toEqual([]);
  });

  it("estimateEffectiveFontSize は収まる場合に元サイズを返す", () => {
    const size = estimateEffectiveFontSize(
      "短い",
      { x: 0, y: 0, w: 3, h: 1 },
      13,
      1.3,
      0,
    );
    expect(size).toBe(13);
  });

  it("estimateEffectiveFontSize は収まらない場合に縮小後サイズを返す", () => {
    const size = estimateEffectiveFontSize(
      "とても長いテキストがとても小さな箱に入れられている状況を推定する",
      { x: 0, y: 0, w: 1.0, h: 0.25 },
      13,
      1.3,
      0,
    );
    expect(size).toBeLessThan(13);
  });
});
