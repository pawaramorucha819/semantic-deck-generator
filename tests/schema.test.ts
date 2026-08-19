import { describe, expect, it } from "vitest";
import {
  parsePresentation,
  PresentationValidationError,
  SlideSpecSchema,
} from "../src/domain/schema.js";
import { loadExampleJson } from "./helpers.js";

/** 検証対象を組み立てるための最小構成 */
function baseSpec(): Record<string, unknown> {
  return {
    title: "テスト資料",
    slides: [
      {
        id: "slide-001",
        keyMessage: "主張",
        slideRole: "analysis",
        logicalRelation: "claim_evidence",
        evidence: [],
      },
    ],
  };
}

describe("presentation.json Schema validation", () => {
  it("サンプル presentation.json を検証できる", async () => {
    const spec = parsePresentation(await loadExampleJson());
    expect(spec.title).toBe("在庫管理システム導入提案");
    expect(spec.slides).toHaveLength(6);
    expect(spec.slides.map((s) => s.id)).toEqual([
      "slide-001",
      "slide-002",
      "slide-003",
      "slide-004",
      "slide-005",
      "slide-006",
    ]);
  });

  it("最小構成を受け入れる", () => {
    expect(() => parsePresentation(baseSpec())).not.toThrow();
  });

  describe("不正入力の拒否", () => {
    it("slides が空なら拒否する", () => {
      const spec = { ...baseSpec(), slides: [] };
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("未知の slideRole を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].slideRole = "conclusion";
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("未知の logicalRelation を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].logicalRelation = "storytelling";
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("keyMessage が空文字なら拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].keyMessage = "   ";
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("slide.id の重複を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[]).push({ ...(spec.slides as any[])[0] });
      expect(() => parsePresentation(spec)).toThrow(/slide\.id が重複/);
    });

    it("未定義の sourceId 参照を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].evidence = [
        { id: "e1", type: "text", statement: "事実", sourceId: "missing-source" },
      ];
      expect(() => parsePresentation(spec)).toThrow(/未定義の sourceId/);
    });

    it("evidence.id の重複を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].evidence = [
        { id: "e1", type: "text", statement: "事実A" },
        { id: "e1", type: "text", statement: "事実B" },
      ];
      expect(() => parsePresentation(spec)).toThrow(/evidence\.id が重複/);
    });

    it("未知の evidence type を拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].evidence = [
        { id: "e1", type: "diagram", statement: "図" },
      ];
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("comparison evidence の item が 1 件なら拒否する", () => {
      const spec = baseSpec();
      (spec.slides as any[])[0].evidence = [
        { id: "e1", type: "comparison", statement: "比較", data: { items: [{ label: "A", value: 1 }] } },
      ];
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });
  });

  describe("表現・描画に属する語彙は presentation.json に書けない", () => {
    it.each([
      ["fontSize", { fontSize: 32 }],
      ["color", { color: "#FF0000" }],
      ["x/y 座標", { x: 1, y: 2 }],
      ["layout 指定", { layout: "two-column" }],
      ["patternId 指定", { patternId: "before-after" }],
      ["chartType 指定", { chartType: "bar" }],
    ])("スライドに %s を書くと検証エラーになる", (_name, extra) => {
      const spec = baseSpec();
      Object.assign((spec.slides as any[])[0], extra);
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });

    it("プレゼンテーション直下に theme を書くと検証エラーになる", () => {
      const spec = { ...baseSpec(), theme: { primary: "#000000" } };
      expect(() => parsePresentation(spec)).toThrow(PresentationValidationError);
    });
  });

  it("エラーは path 付きで整形できる", () => {
    const spec = baseSpec();
    (spec.slides as any[])[0].slideRole = "unknown";
    try {
      parsePresentation(spec);
      expect.unreachable("検証エラーになるはず");
    } catch (error) {
      expect(error).toBeInstanceOf(PresentationValidationError);
      const formatted = (error as PresentationValidationError).format();
      expect(formatted).toContain("slides.0.slideRole");
    }
  });

  it("SlideSpecSchema 単体でも検証できる", () => {
    const result = SlideSpecSchema.safeParse({
      id: "s1",
      keyMessage: "主張",
      slideRole: "summary",
      logicalRelation: "claim_evidence",
      evidence: [],
    });
    expect(result.success).toBe(true);
  });
});
