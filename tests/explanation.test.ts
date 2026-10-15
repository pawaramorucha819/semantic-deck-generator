import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { parsePresentation } from "../src/domain/schema.js";
import { contentItemEntries } from "../src/domain/content.js";
import { buildVisualPlan, parseVisualPlan, serializeVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { estimateTextHeight } from "../src/render/geometry.js";
import { defaultRegistry, PatternRegistry } from "../src/patterns/registry.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { extractTexts, readPptx, slide } from "./helpers.js";

const explanation = { readerQuestion: "この図は何を示すか？", sections: [
  { heading: "前提", body: "対象は社内の検証業務です。図は要素同士の関係を示します。" },
  { heading: "読み取り", body: "配置と矢印を読み分けます。図だけでは成果を実証できません。" },
] };

describe("visible standalone explanations", () => {
  it("preserves explanations and all six relationships through plan replay and editable PPTX", async () => {
    const spec = parsePresentation(JSON.parse(await readFile(new URL("../examples/relationships/six-relationships.json", import.meta.url), "utf8")));
    for (const s of spec.slides) {
      s.explanation = explanation;
      s.supportingMessage = undefined;
      s.keyMessage = "情報の関係を正確に読む";
      for (const { item } of contentItemEntries(s.content!)) { item.label = "要素"; delete item.description; }
    }
    const plan = buildVisualPlan(spec);
    expect(plan.slides.map(s => s.patternId)).toEqual(["three-reasons", "radial", "process", "background", "convergence", "cycle"]);
    expect(plan.slides[0]!.reason).not.toContain("横並び");
    expect(plan.slides[2]!.reason).not.toContain("横一列");
    const replay = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));
    expect(replay.slides.every(s => JSON.stringify(s.explanation) === JSON.stringify(explanation))).toBe(true);
    const result = renderVisualPlan(replay, { includeTitleSlide: false });
    expect(result.issues).toEqual([]);
    const { slideXmls } = await readPptx(await toPptxBuffer(result.pptx));
    for (const xml of slideXmls) {
      const texts = extractTexts(xml).join("\n");
      expect(texts).toContain(explanation.readerQuestion);
      explanation.sections.forEach(s => { expect(texts).toContain(s.heading); expect(texts).toContain(s.body); });
    }
    const flowLabels = result.reports[2]!.placements.filter(p => /^step-\d-label$/.test(p.name));
    expect(flowLabels).toHaveLength(3);
    expect(flowLabels[1]!.box.y).toBeGreaterThan(flowLabels[0]!.box.y);
    expect(flowLabels[1]!.box.x).toBe(flowLabels[0]!.box.x);
    expect(slideXmls[2]).toContain('prst="downArrow"');
    expect(slideXmls[0]).not.toContain('prst="downArrow"');
    expect(slideXmls[1]).toContain("radial");
    expect(slideXmls[4]).toContain("convergence");
    expect(slideXmls[5]).toContain("cycle-step-4-to-1");
    for (const report of result.reports) {
      const prose = report.placements.filter(p => p.name.startsWith("explanation-"));
      expect(prose.every(p => p.fontSize! >= 14 && p.shrink === false)).toBe(true);
      for (const p of prose) {
        const estimated = estimateTextHeight(p.text!, { widthIn: p.box.w, fontSize: p.fontSize!, lineHeight: p.lineHeight! });
        expect(p.box.h).toBeGreaterThan(estimated * 1.1);
      }
      const firstBody = prose.find(p => p.name === "explanation-body-1")!;
      const nextHeading = prose.find(p => p.name === "explanation-heading-2")!;
      expect(nextHeading.box.y - firstBody.box.y - firstBody.box.h).toBeGreaterThanOrEqual(0.119);

    }
  });

  it("accepts one substantive section and rejects its overflow", () => {
    const single = { readerQuestion: "何を示すか？", sections: [explanation.sections[0]!] };
    const spec = parsePresentation({ title: "one section", slides: [slide({ explanation: single })] });
    expect(renderVisualPlan(buildVisualPlan(spec)).issues).toEqual([]);
    const tooLong = { ...single, sections: [{ heading: "説明と限界", body: "必要な説明です。".repeat(90) }] };
    expect(() => renderVisualPlan(buildVisualPlan({ title: "dense", slides: [slide({ explanation: tooLong })] }))).toThrow(/explanation overflows/);
  });

  it("fails closed on oversized explanation instead of shrinking, clipping or hiding in notes", () => {
    const plan = buildVisualPlan({ title: "test", slides: [slide({ explanation: { ...explanation, sections: explanation.sections.map(s => ({ ...s, body: "長い説明。".repeat(180) })) } })] });
    expect(() => renderVisualPlan(plan)).toThrow(/explanation overflows/);
  });

  it("renders explanation for message-only fallback without dropping its full body", async () => {
    const text = "省略せず保持する本文".repeat(4);
    const plan = buildVisualPlan({ title: "test", slides: [slide({ explanation, evidence: [{ id: "e1", type: "text", statement: text }] })] });
    const result = renderVisualPlan(plan, { includeTitleSlide: false });
    const xml = (await readPptx(await toPptxBuffer(result.pptx))).slideXmls[0]!;
    expect(extractTexts(xml).join("\n")).toContain(text);
  });

  it("rejects unsupported custom patterns and lying capability declarations", () => {
    const registry = new PatternRegistry([{ ...messageOnlyPattern, supportsExplanatoryLayout: false }]);
    const plan = buildVisualPlan({ title: "test", slides: [slide({ explanation })] }, { registry });
    expect(() => renderVisualPlan(plan, { registry })).toThrow(/does not support visible explanation/);
    const lying = new PatternRegistry([{ ...messageOnlyPattern, render() {} }]);
    expect(() => renderVisualPlan(plan, { registry: lying })).toThrow(/did not render visible explanation/);
  });

  it("composes visible explanation alongside an editable chart", async () => {
    const plan = buildVisualPlan({ title: "Chart", slides: [slide({ explanation,
      logicalRelation: "comparison", evidence: [{ id: "e", type: "comparison", statement: "AとBを比較する",
        data: { items: [{ label: "A", value: 10 }, { label: "B", value: 20 }] } }],
    })] });
    expect(plan.slides[0]!.patternId).toBe("chart-with-message");
    const result = renderVisualPlan(plan, { includeTitleSlide: false });
    expect(result.issues).toEqual([]);
    const xml = (await readPptx(await toPptxBuffer(result.pptx))).slideXmls[0]!;
    expect(xml).toContain("<c:chart");
    expect(extractTexts(xml).join("\n")).toContain(explanation.sections[0]!.body);
  });

  it("rejects malformed explanations rather than stripping unknown fields", () => {
    expect(() => parsePresentation({ title: "t", slides: [slide({ explanation: { ...explanation, sections: [] } })] })).toThrow();
    expect(defaultRegistry.get("chart-with-message").supportsExplanatoryLayout).toBe(true);
  });
});
