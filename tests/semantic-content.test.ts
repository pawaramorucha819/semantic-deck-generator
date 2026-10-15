import { describe, expect, it } from "vitest";
import {
  contentItemEntries,
  contentItems,
  isRelationshipContent,
  ContentItemSchema,
  DecompositionSchema,
  LOGICAL_RELATIONS,
  parseSlide,
  SemanticContentSchema,
  SlideSpecSchema,
  type ContentItem,
  type RelationshipContent,
  type SemanticContent,
} from "../src/domain/index.js";
import { ingestDeckFromGeneration } from "../src/knowledge/ingest/ingest.js";
import { bigNumberPattern } from "../src/patterns/big-number/index.js";
import { compileDeclarativePattern } from "../src/patterns/declarative/compile.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { processPattern } from "../src/patterns/process/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { threeReasonsPattern } from "../src/patterns/three-reasons/index.js";
import { buildVisualPlan, parseVisualPlan, serializeVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { evaluatePattern } from "../src/selector/selector.js";
import { extractTexts, metricEvidence, readPptx, slide, textEvidence } from "./helpers.js";

const KINDS = ["distribution", "convergence", "cycle", "context"] as const;
const item = (label: string): ContentItem => ({ label, description: `${label} description` });

function relationship(kind: RelationshipContent["kind"], count = 2): RelationshipContent {
  const items = Array.from({ length: count }, (_, i) => item(`Node ${i + 1}`));
  switch (kind) {
    case "distribution": return { kind, center: item("Origin"), items };
    case "convergence": return { kind, inputs: items, outcome: item("Outcome") };
    case "cycle": return { kind, steps: items };
    case "context": return { kind, environment: item("Environment"), items };
  }
}

function legacyDeclarative(source = "contentOrEvidence") {
  return compileDeclarativePattern({
    schemaVersion: "1.0",
    kind: "declarative",
    id: "legacy-items",
    version: "1.0.0",
    description: "Flat legacy item slots",
    compatibility: { logicalRelations: ["claim_evidence", ...KINDS] },
    slots: [
      { id: "headline", type: "text", source: "keyMessage" },
      { id: "items", type: "items", source, optional: true, max: 2 },
    ],
    match: { baseScore: 0.99 },
    layout: { body: { type: "text", value: { ref: "headline" }, style: { role: "body" } } },
  });
}

describe("semantic relationship domain", () => {
  it.each(KINDS)("accepts %s and its matching logical relation without visual vocabulary", (kind) => {
    const content = relationship(kind);
    expect(parseSlide(slide({ logicalRelation: kind, content })).content).toEqual(content);
    expect(LOGICAL_RELATIONS).toContain(kind);
    expect(isRelationshipContent(content)).toBe(true);
  });

  it.each(KINDS)("validates every %s evidence reference with its exact source path", (kind) => {
    const content = relationship(kind, 4);
    for (const { path } of contentItemEntries(content)) {
      const invalid = structuredClone(content);
      let node: unknown = invalid;
      for (const part of path) node = (node as Record<string | number, unknown>)[part];
      (node as ContentItem).evidenceRefs = ["missing"];
      const result = SlideSpecSchema.safeParse(slide({ content: invalid }));
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.issues.map((issue) => issue.path)).toContainEqual([
          "content", ...path, "evidenceRefs",
        ]);
      }
      expect(() => parseSlide(slide({
        content: invalid,
        evidence: [textEvidence("missing", "Supporting fact")],
      }))).not.toThrow();
    }
  });

  it.each(KINDS)("rejects oversize %s instead of truncating it", (kind) => {
    expect(SemanticContentSchema.safeParse(relationship(kind, 4)).success).toBe(true);
    expect(SemanticContentSchema.safeParse(relationship(kind, 5)).success).toBe(false);
    expect(SemanticContentSchema.safeParse(relationship(kind, 0)).success).toBe(false);
    expect(SemanticContentSchema.safeParse(relationship(kind, 1)).success).toBe(kind === "context");
  });

  it.each(KINDS)("rejects visual/unknown fields at every %s node", (kind) => {
    expect(SemanticContentSchema.safeParse({ ...relationship(kind), patternId: "visual-choice" }).success).toBe(false);
    for (const { path } of contentItemEntries(relationship(kind))) {
      const content = structuredClone(relationship(kind));
      let node: unknown = content;
      for (const part of path) node = (node as Record<string | number, unknown>)[part];
      Object.assign(node as object, { color: "primary" });
      expect(SemanticContentSchema.safeParse(content).success).toBe(false);
    }
    expect(ContentItemSchema.safeParse({ label: "Node", x: 1 }).success).toBe(false);
  });

  it("enumeration retains original roles, closure order and singleton nodes", () => {
    expect(contentItemEntries(relationship("distribution")).map((entry) => entry.path)).toEqual([
      ["center"], ["items", 0], ["items", 1],
    ]);
    expect(contentItemEntries(relationship("convergence")).map((entry) => entry.path)).toEqual([
      ["inputs", 0], ["inputs", 1], ["outcome"],
    ]);
    expect(contentItemEntries(relationship("cycle")).map((entry) => entry.path)).toEqual([
      ["steps", 0], ["steps", 1],
    ]);
    expect(contentItemEntries(relationship("context")).map((entry) => entry.path)).toEqual([
      ["environment"], ["items", 0], ["items", 1],
    ]);
    expect(contentItems()).toEqual([]);
    expect(isRelationshipContent()).toBe(false);
  });

  it("preserves six-item legacy points and sequence validation and inspection", () => {
    for (const kind of ["points", "sequence"] as const) {
      const nodes = Array.from({ length: 6 }, (_, i) => item(`Legacy ${i}`));
      const content: SemanticContent = kind === "points" ? { kind, items: nodes } : { kind, steps: nodes };
      expect(SemanticContentSchema.parse(content)).toEqual(content);
      expect(contentItems(content)).toEqual(nodes);
      expect(isRelationshipContent(content)).toBe(false);
    }
  });

  it("leaves relation/kind compatibility to the pattern gate", () => {
    expect(() => parseSlide(slide({
      content: relationship("cycle"), logicalRelation: "claim_evidence",
    }))).not.toThrow();
  });
});

describe("decomposition review contract", () => {
  const decomposition = {
    scope: "Incident response work",
    criterion: "Responsible function",
    rationale: "Separate work by the accountable team",
    exclusions: ["Product roadmap"],
    unresolvedQuestions: ["Who owns shared infrastructure?"],
  };

  it("round-trips rationale, explicit exclusions and open questions", () => {
    expect(parseSlide(slide({ decomposition })).decomposition).toEqual(decomposition);
  });

  it.each(["scope", "criterion", "rationale"])("requires nonempty %s", (key) => {
    expect(DecompositionSchema.safeParse({ ...decomposition, [key]: " " }).success).toBe(false);
    const missing = { ...decomposition } as Record<string, unknown>;
    delete missing[key];
    expect(DecompositionSchema.safeParse(missing).success).toBe(false);
  });

  it.each(["isMECE", "color", "x", "patternId"])("rejects unsupported assertion/display field %s", (field) => {
    expect(DecompositionSchema.safeParse({ ...decomposition, [field]: true }).success).toBe(false);
  });

  it("accepts stable node IDs and finite membership without asserting semantic coverage", () => {
    const content = relationship("distribution");
    contentItems(content).forEach((node, i) => { node.id = `node-${i}`; node.members = ["shared", "outside"]; });
    const input = slide({ content, decomposition: { ...decomposition, universe: ["shared", "uncovered"] } });
    expect(parseSlide(input)).toEqual(input);
  });

  it("rejects duplicate finite universe values after normalization", () => {
    expect(DecompositionSchema.safeParse({ ...decomposition, universe: ["one", " one "] }).success).toBe(false);
  });

  it.each(KINDS)("rejects duplicate supplied %s node IDs including singleton nodes", (kind) => {
    const content = relationship(kind);
    const entries = contentItemEntries(content);
    entries[0]!.item.id = "duplicate";
    entries.at(-1)!.item.id = "duplicate";
    const result = SlideSpecSchema.safeParse(slide({ content }));
    expect(result.success).toBe(false);
    if (!result.success) expect(result.error.issues[0]!.path).toEqual(["content", ...entries.at(-1)!.path, "id"]);
  });


  it.each(KINDS)("rejects supplied %s node IDs colliding with another node's default path ID", (kind) => {
    const content = relationship(kind);
    const entries = contentItemEntries(content);
    entries[0]!.item.id = entries.at(-1)!.path.join(".");
    const result = SlideSpecSchema.safeParse(slide({ content }));
    expect(result.success).toBe(false);
    if (!result.success) expect(result.error.issues[0]!.message).toContain("path ID");
  });

  it("rejects explicit points IDs shadowing another item's generated identity", () => {
    expect(() => parseSlide(slide({ content: {
      kind: "points", items: [{ id: "items.1", label: "First" }, { label: "Second" }],
    } }))).toThrow(/id が重複/);
  });

});

describe("relationship-preserving compatibility gates", () => {
  it.each(KINDS)("legacy patterns decline %s rather than use unrelated evidence", (kind) => {
    const content = relationship(kind);
    const metricSlide = slide({ content, evidence: [metricEvidence("metric", "Metric", 10)] });
    expect(evaluatePattern(bigNumberPattern, metricSlide).compatible).toBe(false);
    expect(() => bigNumberPattern.buildBindings(metricSlide, "standard")).toThrow();
    const evidenceSlide = slide({ content, evidence: [textEvidence("a", "A"), textEvidence("b", "B")] });
    expect(evaluatePattern(threeReasonsPattern, evidenceSlide).compatible).toBe(false);
    expect(() => threeReasonsPattern.buildBindings(evidenceSlide, "standard")).toThrow();
    expect(evaluatePattern(processPattern, { ...evidenceSlide, logicalRelation: "process" }).compatible).toBe(false);
  });

  it.each(["contentPoints", "contentSequence", "content", "contentOrEvidence", "evidence"])(
    "legacy declarative %s slots fail explicitly, including optional slots", (source) => {
      const pattern = legacyDeclarative(source);
      for (const kind of KINDS) {
        const input = slide({ logicalRelation: kind, content: relationship(kind), evidence: [textEvidence("a", "A")] });
        const match = pattern.match(input);
        expect(match.compatible).toBe(false);
        expect(match.reason).toContain("cannot preserve");
        expect(() => pattern.buildBindings(input, "standard")).toThrow(/cannot preserve/);
      }
    },
  );
});

describe("relationship fallback and knowledge preservation", () => {
  it.each(KINDS)("fallback preserves every %s node and description, with explicit disclosure", (kind) => {
    const content = relationship(kind, 4);
    const input = slide({ logicalRelation: kind, content });
    const bindings = messageOnlyPattern.buildBindings(input, "standard");
    expect(messageOnlyPattern.bindingsSchema.safeParse(bindings).success).toBe(true);
    expect(bindings.structureText).toContain("構造のテキスト表示");
    expect(bindings.points).toBeUndefined();
    for (const node of contentItems(content)) {
      expect(bindings.structureText).toContain(node.label);
      expect(bindings.structureText).toContain(node.description);
    }
    expect(messageOnlyPattern.match(input).reason).toContain("テキスト表示へ退避");
  });

  it("legacy fallback points and sequence retain their original summaries", () => {
    const nodes = [item("One"), item("Two"), item("Three")];
    for (const content of [{ kind: "points", items: nodes }, { kind: "sequence", steps: nodes }] as SemanticContent[]) {
      const bindings = messageOnlyPattern.buildBindings(slide({ content }), "standard");
      expect(bindings.points).toEqual(nodes.slice(0, 3).map((node) => `${node.label}：${node.description}`));
      expect(bindings).not.toHaveProperty("structureText");
    }
  });

  it("fallback round-trip renders all nodes, source notes and cycle closure without new dependencies", async () => {
    const registry = new PatternRegistry([messageOnlyPattern]);
    const input = slide({
      logicalRelation: "cycle",
      content: relationship("cycle", 4),
      sources: [{ id: "source", title: "Source report", url: "https://example.com/source", locator: "p.4" }],
      evidence: [{ ...textEvidence("fact", "Supporting fact"), sourceIds: ["source"] }],
    });
    const spec = { title: "Fallback", slides: [input] };
    const plan = buildVisualPlan(spec, { registry });
    const restored = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)), { registry });
    const rendered = renderVisualPlan(restored, { registry, includeTitleSlide: false });
    const pptx = await readPptx(await toPptxBuffer(rendered.pptx));
    const text = extractTexts(pptx.slideXmls[0]!).join("\n");
    expect(text).toContain("最終工程から最初へ戻る");
    for (const node of contentItems(input.content)) expect(text).toContain(node.description);
    expect(text).toContain("Source report");
    const notes = await pptx.zip.file("ppt/notesSlides/notesSlide1.xml")!.async("string");
    expect(notes).toContain("https://example.com/source");
    expect(notes).toContain("p.4");
    expect(rendered.issues.filter((issue) => issue.severity === "error")).toEqual([]);
  });

  it.each(KINDS)("knowledge inspection includes %s singleton and array nodes", (kind) => {
    const registry = new PatternRegistry([messageOnlyPattern]);
    const input = slide({ logicalRelation: kind, content: relationship(kind, 4) });
    const spec = { title: "Knowledge", slides: [input] };
    const plan = buildVisualPlan(spec, { registry });
    const result = ingestDeckFromGeneration(spec, plan, { deckId: "semantic-relations" });
    const entry = result.entries[0]!;
    for (const node of contentItems(input.content)) {
      expect(entry.content!.text).toContain(node.label);
      expect(entry.content!.text).toContain(node.description);
    }
    expect(entry.content!.bulletCount).toBe(contentItems(input.content).length);
  });
});


describe("saved-plan node traceability", () => {
  const nodes: ContentItem[] = [
    { id: "node-first", label: "First", evidenceRefs: ["e1"], members: ["member-first"] },
    { id: "node-second", label: "Second", evidenceRefs: ["e2"], members: [] },
  ];
  const evidence = [textEvidence("e1", "First fact"), textEvidence("e2", "Second fact")];

  it.each(["points", "sequence"] as const)("saved %s bindings retain supplied IDs, references and membership", (kind) => {
    const content: SemanticContent = kind === "points" ? { kind, items: nodes } : { kind, steps: nodes };
    const pattern = kind === "points" ? threeReasonsPattern : processPattern;
    const registry = new PatternRegistry([pattern, messageOnlyPattern]);
    const input = parseSlide(slide({ content, logicalRelation: kind === "sequence" ? "process" : "claim_evidence", evidence }));
    const plan = buildVisualPlan({ title: "Traceability", slides: [input] }, { registry });
    expect(plan.slides[0]!.patternId).toBe(pattern.id);
    const restored = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)), { registry });
    const bindings = restored.slides[0]!.bindings as { reasons?: Array<Record<string, unknown>>; steps?: Array<Record<string, unknown>> };
    const saved = bindings.reasons ?? bindings.steps!;
    nodes.forEach((node, index) => {
      expect(saved[index]).toMatchObject({ nodeId: node.id, evidenceRefs: node.evidenceRefs, members: node.members });
    });
  });

  it("metadata-free process and point bindings keep their original property sets", () => {
    const plain = nodes.map(({ label }) => ({ label }));
    const process = processPattern.buildBindings(slide({ content: { kind: "sequence", steps: plain } }), "standard");
    const reasons = threeReasonsPattern.buildBindings(slide({ content: { kind: "points", items: plain } }), "standard");
    for (const node of [...process.steps, ...reasons.reasons]) {
      expect(node).not.toHaveProperty("nodeId");
      expect(node).not.toHaveProperty("evidenceRefs");
      expect(node).not.toHaveProperty("members");
    }
  });

  it.each(KINDS)("fallback saves %s node IDs and original paths independently of text rendering", (kind) => {
    const content = relationship(kind);
    contentItemEntries(content).forEach(({ item }, index) => {
      item.id = `node-${index}`;
      item.evidenceRefs = ["e1"];
      item.members = [`member-${index}`];
    });
    const registry = new PatternRegistry([messageOnlyPattern]);
    const plan = buildVisualPlan({ title: "Fallback traceability", slides: [slide({ content, logicalRelation: kind, evidence })] }, { registry });
    const restored = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)), { registry });
    expect((restored.slides[0]!.bindings as { structureNodes: unknown }).structureNodes).toEqual(
      contentItemEntries(content).map(({ item, path }) => ({ path, nodeId: item.id, evidenceRefs: item.evidenceRefs, members: item.members })),
    );
  });

  it("binding metadata stays strict and rejects unrelated display or interpretation fields", () => {
    const process = processPattern.buildBindings(slide({ content: { kind: "sequence", steps: nodes } }), "standard");
    expect(processPattern.bindingsSchema.safeParse({ ...process, steps: process.steps.map((node) => ({ ...node, isMECE: true })) }).success).toBe(false);
    const fallback = messageOnlyPattern.buildBindings(slide({ content: relationship("context") }), "standard");
    expect(messageOnlyPattern.bindingsSchema.safeParse({ ...fallback, structureNodes: [{ path: ["environment"], color: "red" }] }).success).toBe(false);
  });
});


describe("full-content legacy fallback", () => {
  it.each(["points", "sequence"] as const)("six %s nodes survive selection, saved plans and rendered output without warnings", async (kind) => {
    const nodes = Array.from({ length: 6 }, (_, index) => ({
      id: `item-${index}`, label: `Item ${index + 1}`, description: `Detail ${index + 1}`,
      evidenceRefs: ["fact"], members: [`member-${index}`],
    }));
    const content: SemanticContent = kind === "points" ? { kind, items: nodes } : { kind, steps: nodes };
    const input = parseSlide(slide({ content, logicalRelation: kind === "sequence" ? "process" : "claim_evidence", evidence: [textEvidence("fact", "Fact")] }));
    const plan = buildVisualPlan({ title: "Complete fallback", slides: [input] });
    expect(plan.slides[0]!.patternId).toBe("message-only");
    const restored = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));
    const bindings = restored.slides[0]!.bindings as { structureText: string; structureNodes: unknown[] };
    expect(bindings.structureNodes).toEqual(contentItemEntries(content).map(({ item, path }) => ({
      path, nodeId: item.id, evidenceRefs: item.evidenceRefs, members: item.members,
    })));
    expect(bindings.structureText).not.toContain("最初へ戻る");
    if (kind === "points") expect(bindings.structureText).toContain("順序を意味しない");
    else expect(bindings.structureText).toContain("6. Item 6");
    const result = renderVisualPlan(restored, { includeTitleSlide: false });
    expect(result.issues).toEqual([]);
    const pptx = await readPptx(await toPptxBuffer(result.pptx));
    const text = extractTexts(pptx.slideXmls[0]!).join("\n");
    for (const node of nodes) {
      expect(text).toContain(node.label);
      expect(text).toContain(node.description);
    }
  });

  it.each(["points", "sequence"] as const)("short %s fallback retains its display and newly supplied metadata", (kind) => {
    const nodes = [
      { id: "a", label: "A", description: "A detail", evidenceRefs: ["e1"] },
      { id: "b", label: "B", description: "B detail", members: ["b"] },
    ];
    const content: SemanticContent = kind === "points" ? { kind, items: nodes } : { kind, steps: nodes };
    const bindings = messageOnlyPattern.buildBindings(slide({ content }), "standard");
    expect(bindings.points).toEqual(["A：A detail", "B：B detail"]);
    expect(bindings).not.toHaveProperty("structureText");
    expect(bindings.structureNodes).toEqual([
      { path: [kind === "points" ? "items" : "steps", 0], nodeId: "a", evidenceRefs: ["e1"] },
      { path: [kind === "points" ? "items" : "steps", 1], nodeId: "b", members: ["b"] },
    ]);
  });
});
