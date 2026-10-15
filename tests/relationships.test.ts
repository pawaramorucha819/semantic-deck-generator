import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { parsePresentation } from "../src/domain/schema.js";
import type { SemanticContent } from "../src/domain/types.js";
import { buildVisualPlan, parseVisualPlan, serializeVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { buildStructureReview, relationshipGraph } from "../src/review/structure.js";
import { midnightTheme } from "../src/theme/theme.js";
import { extractTexts, readPptx, slide } from "./helpers.js";

const load = async () => parsePresentation(JSON.parse(await readFile(new URL("../examples/relationships/six-relationships.json", import.meta.url), "utf8")));

describe("six relationship types", () => {
  it("selects faithful relationships and preserves all body nodes in editable PPTX", async () => {
    const spec = await load();
    const plan = buildVisualPlan(spec);
    expect(plan.slides.map(s => s.patternId)).toEqual(["three-reasons", "radial", "process", "background", "convergence", "cycle"]);
    const render = renderVisualPlan(plan, { includeTitleSlide: false });
    expect(render.issues).toEqual([]);
    const { slideXmls } = await readPptx(await toPptxBuffer(render.pptx));
    spec.slides.forEach((s, i) => {
      const texts = extractTexts(slideXmls[i]!).join("\n");
      for (const node of relationshipGraph(s.content!).nodes) {
        expect(texts).toContain(node.label);
        if (node.description) expect(texts).toContain(node.description);
      }
      expect(slideXmls[i]).not.toContain("<p:pic>");
    });
    // Outward arrows, incoming arrows, and a closed loop are real editable line objects.
    expect(slideXmls[1]!.match(/<a:tailEnd type="triangle"/g)).toHaveLength(3);
    expect(slideXmls[3]).not.toContain("tailEnd");
    expect(slideXmls[4]!.match(/<a:tailEnd type="triangle"/g)).toHaveLength(2);
    expect(slideXmls[5]!.match(/<a:tailEnd type="triangle"/g)).toHaveLength(4);
    expect(slideXmls[5]).toContain("cycle-step-4-to-1");
    expect(slideXmls[1]).toContain('flipH="1"');
    expect(slideXmls[5]).toContain('flipV="1"');
  });

  it("replays unchanged, including direction, under either theme", async () => {
    const plan = buildVisualPlan(await load());
    const saved = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));
    const a = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    const b = await readPptx(await toPptxBuffer(renderVisualPlan(saved).pptx));
    expect(b.slideXmls).toEqual(a.slideXmls);
    expect(renderVisualPlan(plan, { theme: midnightTheme }).issues).toEqual([]);
  });

  it.each([2, 3, 4])("renders %i-step cycles and preserves exact closure", async count => {
    const content: SemanticContent = { kind: "cycle", steps: Array.from({ length: count }, (_, i) => ({ id: `step-${i}`, label: `工程${i + 1}`, description: "確認して進める" })) };
    const spec = { title: "反復", slides: [slide({ content, logicalRelation: "cycle" })] };
    const graph = relationshipGraph(content);
    expect(graph.edges).toHaveLength(count);
    expect(graph.edges.at(-1)).toEqual({ from: `step-${count - 1}`, to: "step-0", relation: "returns_to" });
    const render = renderVisualPlan(buildVisualPlan(spec), { includeTitleSlide: false });
    expect(render.issues).toEqual([]);
    const xml = (await readPptx(await toPptxBuffer(render.pptx))).slideXmls[0]!;
    expect(xml.match(/<a:tailEnd type="triangle"/g)).toHaveLength(count);
  });

  it("uses explicit ids and distinguishes divergence, convergence, sequence and context", async () => {
    const spec = await load();
    const review = buildStructureReview(spec, buildVisualPlan(spec));
    const graphs = review.slides.map(s => s.graph!);
    expect(graphs[0]!.edges).toEqual([]);
    expect(graphs[1]!.edges.every(e => e.from === "center" && e.relation === "distributes_to")).toBe(true);
    expect(graphs[2]!.edges).toHaveLength(2);
    expect(graphs[3]!.edges.every(e => e.from === "environment" && e.relation === "contains")).toBe(true);
    expect(graphs[4]!.edges.every(e => e.to === "outcome" && e.relation === "contributes_to")).toBe(true);
    expect(review.assessment).toBe("requires-human-review");
    expect(review.slides.every(s => s.semanticMECE === "unverified")).toBe(true);
  });

  it("does not silently truncate long process labels", () => {
    const label = "これは十四文字を超える工程ラベルをそのまま保持する";
    const plan = buildVisualPlan({ title: "工程", slides: [slide({ logicalRelation: "process", content: { kind: "sequence", steps: [{ label }, { label: "完了" }] } })] });
    expect(JSON.stringify(plan.slides[0]!.bindings)).toContain(label);
  });

  it("rejects node-local overflow even when text stays inside the slide", () => {
    const spec = { title: "密な合流", slides: [slide({ logicalRelation: "convergence", content: {
      kind: "convergence" as const,
      inputs: Array.from({ length: 4 }, (_, i) => ({ id: `input-${i}`, label: `入力${i}`, description: i === 0 ? "検証対象の日本語".repeat(5) : "短い説明" })),
      outcome: { label: "判断" },
    } })] };
    const plan = buildVisualPlan(spec);
    expect(() => renderVisualPlan(plan)).toThrow(/convergence-input-1.*収まりません.*input-0/);
    expect(JSON.stringify(plan.slides[0]!.bindings)).toContain("検証対象の日本語".repeat(5));
  });
});

describe("decomposition review", () => {
  const review = (members: string[][], universe?: string[]) => buildStructureReview({ title: "分類", slides: [slide({
    content: { kind: "points", items: members.map((m, i) => ({ id: `n${i}`, label: `分類${i}`, members: m })) },
    decomposition: { scope: "対象業務", criterion: "担当領域", rationale: "担当別に分解", ...(universe ? { universe } : {}) },
  })] }).slides[0]!;
  it("flags overlap, omissions and out-of-scope members independently", () => {
    const r = review([["A", "B"], ["B", "X"]], ["A", "B", "C"]);
    expect(r.membershipChecks).toMatchObject({ status: "review-required", overlap: [{ member: "B", nodes: ["n0", "n1"] }], missing: ["C"], outsideScope: ["X"] });
  });
  it("passes declared partitions without claiming semantic MECE", () => {
    const r = review([["A"], ["B"]], ["A", "B"]);
    expect(r.membershipChecks.status).toBe("declared-sets-pass");
    expect(r.semanticMECE).toBe("unverified");
  });
  it("reports prose-only scope as not verifiable", () => {
    expect(review([["A"], ["B"]]).membershipChecks.status).toBe("not-verifiable");
  });
});
