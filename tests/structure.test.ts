import { describe, expect, it } from "vitest";
import { contentItemEntries } from "../src/domain/content.js";
import { parsePresentation } from "../src/domain/schema.js";
import type { LogicalRelation, SemanticContent } from "../src/domain/types.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import {
  buildDecompositionReview, canonicalJson, compileDecomposition, draftFingerprint, fingerprint,
  parseDecompositionDraft, parseDecompositionReviewRecord, parseRawBody, resolveExactQuote,
  sourceFingerprint, validateDecompositionDraft, validateSourceSpan,
  type DecompositionDraft, type DecompositionReviewRecord, type RawBody,
} from "../src/structure/index.js";

const node = (id: string) => ({ id, label: `${id}の対象😀` });
const examples: Array<[SemanticContent, LogicalRelation, string]> = [
  [{ kind: "points", items: [node("a"), node("b")] }, "claim_evidence", "three-reasons"],
  [{ kind: "distribution", center: node("center"), items: [node("a"), node("b")] }, "distribution", "radial"],
  [{ kind: "sequence", steps: [node("a"), node("b")] }, "process", "process"],
  [{ kind: "context", environment: node("environment"), items: [node("a"), node("b")] }, "context", "background"],
  [{ kind: "convergence", inputs: [node("a"), node("b")], outcome: node("outcome") }, "convergence", "convergence"],
  [{ kind: "cycle", steps: [node("a"), node("b")] }, "cycle", "cycle"],
];

function fixture(content = examples[0]![0], logicalRelation = examples[0]![1]) {
  content = structuredClone(content);
  const raw: RawBody = {
    schemaVersion: "1.0", id: "source-body", title: "分解対象",
    body: `  ${contentItemEntries(content).map(({ item }) => item.label).join("\r\n")}\r\n`,
    source: { id: "source-1", title: "Original source", locator: "p.1" },
  };
  const first = contentItemEntries(content)[0]!.item;
  const draft: DecompositionDraft = {
    schemaVersion: "1.0", id: "candidate-1", status: "draft", sourceHash: sourceFingerprint(raw),
    author: { kind: "assistant", id: "test-assistant" },
    slide: {
      id: "slide-1", keyMessage: "対象を整理する", slideRole: "analysis", logicalRelation, content,
      decomposition: { scope: "Declared source subject", criterion: "Declared criterion", rationale: "Explicit candidate" },
      evidence: [], sources: [{ ...raw.source! }],
    },
    nodeSources: contentItemEntries(content).map(({ item }) => ({ nodeId: item.id!, spans: [resolveExactQuote(raw, item.label)], origin: "source", rationale: "Exact source excerpt" })),
    messageSources: [resolveExactQuote(raw, first.label)],
    relation: { rationale: "The relation is an explicit interpretation for review", spans: [resolveExactQuote(raw, first.label)] },
    excludedSpans: [],
  };
  return { raw, draft };
}
function approval(raw: RawBody, draft: DecompositionDraft): DecompositionReviewRecord {
  const packet = buildDecompositionReview(raw, draft);
  return {
    schemaVersion: "1.0", sourceHash: packet.sourceHash, draftHash: packet.draftHash,
    decision: "approve", actor: { kind: "fixture", id: "test-reviewer" }, rationale: "Explicit test approval; no user identity is authenticated",
    dispositions: packet.warnings.map(warning => ({ id: warning.id, decision: "accepted-limitation", reason: "Accepted only for this fixture; semantic support and completeness remain unknown" })),
  };
}

describe("bounded raw body decomposition contract", () => {
  it.each(examples)("compiles %j into the existing semantic contract and dedicated diagram", (content, relation, pattern) => {
    const { raw, draft } = fixture(content, relation);
    const compiled = compileDecomposition(raw, draft, approval(raw, draft));
    expect(parsePresentation(compiled.presentation)).toEqual(compiled.presentation);
    expect(compiled.presentation.slides[0]!.content).toEqual(content);
    expect(buildVisualPlan(compiled.presentation).slides[0]!.patternId).toBe(pattern);
    expect(compiled.lineage.semanticMECE).toBe("unverified");
    expect(compiled.lineage.authentication).toBe("self-declared");
    expect(Object.keys(compiled.presentation)).toEqual(["title", "slides"]);
    expect(compiled.presentation.slides[0]).not.toHaveProperty("nodeSources");
    expect(compiled.presentation.slides[0]).not.toHaveProperty("review");
  });

  it("preserves exact Japanese, emoji, CRLF and leading/trailing body whitespace", () => {
    const { raw, draft } = fixture();
    const packet = buildDecompositionReview(raw, draft);
    const review = approval(raw, draft);
    const original = JSON.stringify({ raw, draft, review });
    const replay = JSON.parse(original);
    const compiled = compileDecomposition(raw, draft, review);
    expect(compiled.lineage.packet.raw.body).toBe(raw.body);
    expect(packet.coverage.gaps).toEqual([]);
    expect(compileDecomposition(replay.raw, replay.draft, replay.review)).toEqual(compiled);
    expect(JSON.stringify({ raw, draft, review })).toBe(original);
    expect(parseRawBody({ ...raw, body: raw.body.replaceAll("\r\n", "\n") }).body).not.toBe(raw.body);
    expect(sourceFingerprint({ ...raw, body: raw.body.replaceAll("\r\n", "\n") })).not.toBe(sourceFingerprint(raw));
  });

  it("uses canonical stable hashes and preserves array order", () => {
    expect(canonicalJson({ b: 2, a: { z: "😀", c: 1 } })).toBe('{"a":{"c":1,"z":"😀"},"b":2}');
    expect(fingerprint({ a: 1, b: 2 })).toBe(fingerprint({ b: 2, a: 1 }));
    expect(fingerprint([1, 2])).not.toBe(fingerprint([2, 1]));
    expect(() => fingerprint(Number.NaN)).toThrow();
    expect(() => fingerprint(new Array(2))).toThrow();
    expect(() => fingerprint(new Date())).toThrow();
  });

  it("requires explicit offsets for repeated and overlapping quotes", () => {
    expect(() => resolveExactQuote("甲と甲", "甲")).toThrow(/Ambiguous/);
    expect(resolveExactQuote("甲と甲", "甲", 2)).toEqual({ start: 2, end: 3, quote: "甲" });
    expect(() => resolveExactQuote("aaa", "aa")).toThrow(/Ambiguous/);
    expect(() => resolveExactQuote("abc", "fabricated")).toThrow(/not found/);
    expect(() => resolveExactQuote("abc", "b", 0)).toThrow(/exactly match/);
  });

  it("validates JavaScript UTF-16 boundaries without splitting emoji", () => {
    const body = "前😀後";
    expect(resolveExactQuote(body, "😀")).toEqual({ start: 1, end: 3, quote: "😀" });
    expect(() => validateSourceSpan(body, { start: 1, end: 2, quote: body.slice(1, 2) })).toThrow(/surrogate/);
    expect(() => validateSourceSpan(body, { start: 2, end: 3, quote: body.slice(2, 3) })).toThrow(/surrogate/);
    expect(() => validateSourceSpan(body, { start: 1, end: 3, quote: "fake" })).toThrow(/exactly match/);
    expect(() => validateSourceSpan(body, { start: 1.5, end: 3, quote: "😀" })).toThrow();
    expect(() => validateSourceSpan(body, { start: 0, end: 99, quote: body })).toThrow();
  });

  it("rejects fabricated anchors on nodes, key message, relation and exclusions", () => {
    for (const target of ["node", "message", "relation", "excluded"] as const) {
      const { raw, draft } = fixture();
      const forged = { start: 0, end: 3, quote: "fake" };
      if (target === "node") draft.nodeSources[0]!.spans = [forged];
      if (target === "message") draft.messageSources = [forged];
      if (target === "relation") draft.relation.spans = [forged];
      if (target === "excluded") draft.excludedSpans = [{ span: forged, reason: "Supposed exclusion" }];
      expect(() => validateDecompositionDraft(raw, draft)).toThrow(/exactly match/);
    }
  });

  it("requires every explicit content id exactly once in source anchors", () => {
    const mutate = (change: (draft: DecompositionDraft) => void) => {
      const { raw, draft } = fixture(); change(draft);
      expect(() => validateDecompositionDraft(raw, draft)).toThrow();
    };
    mutate(draft => { delete contentItemEntries(draft.slide.content)[0]!.item.id; });
    mutate(draft => { contentItemEntries(draft.slide.content)[1]!.item.id = "a"; });
    mutate(draft => { draft.nodeSources[1]!.nodeId = "a"; });
    mutate(draft => { draft.nodeSources[1]!.nodeId = "not-a-node"; });
    mutate(draft => { draft.nodeSources.pop(); });
    mutate(draft => { draft.nodeSources[0]!.spans = []; });
    mutate(draft => { draft.messageSources = []; });
    mutate(draft => { draft.relation.spans = []; });
  });

  it("retains existing evidence/source reference validation and rejects duplicate source ids", () => {
    const { raw, draft } = fixture();
    contentItemEntries(draft.slide.content)[0]!.item.evidenceRefs = ["missing"];
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/missing/);
    draft.slide.evidence = [{ id: "missing", type: "text", statement: "Evidence", sourceIds: ["missing-source"] }];
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/missing-source/);
    draft.slide.evidence[0]!.sourceIds = [raw.source!.id];
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/evidence-bearing selection/);
    draft.slide.evidence = [];
    delete contentItemEntries(draft.slide.content)[0]!.item.evidenceRefs;
    expect(validateDecompositionDraft(raw, draft).slide.sources).toEqual([raw.source]);
    draft.slide.sources!.push({ id: raw.source!.id, title: "Duplicate" });
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/Duplicate source/);
  });

  it("rejects numeric evidence that can select a non-body pattern without fallback", () => {
    const { raw, draft } = fixture();
    const review = approval(raw, draft);
    draft.slide.evidence = [{ id: "metric-1", type: "metric", statement: "A numeric assertion", data: { value: 42, label: "Reported count" } }];
    const ordinaryPresentation = parsePresentation({ title: raw.title, slides: [draft.slide] });
    const ordinaryPlan = buildVisualPlan(ordinaryPresentation);
    expect(ordinaryPlan.slides[0]!.patternId).toBe("big-number");
    expect(ordinaryPlan.slides[0]!.fallbackUsed).toBe(false);
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/requires evidence=\[\]/);
    expect(() => buildDecompositionReview(raw, draft)).toThrow(/body-preservation contract/);
    expect(() => compileDecomposition(raw, draft, review)).toThrow(/body-preservation contract/);
  });

  it("rejects source laundering and preserves only the frozen raw source", () => {
    const { raw, draft } = fixture();
    const compile = () => compileDecomposition(raw, draft, approval(raw, draft));
    expect(compile().presentation.slides[0]!.sources).toEqual([raw.source]);
    draft.slide.sources![0]!.url = "https://example.com/forged";
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/frozen raw.source/);
    draft.slide.sources = [];
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/frozen raw.source/);
    delete draft.slide.sources;
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/frozen raw.source/);
    delete raw.source;
    draft.sourceHash = sourceFingerprint(raw);
    expect(validateDecompositionDraft(raw, draft).slide.sources).toBeUndefined();
    draft.slide.sources = [];
    expect(validateDecompositionDraft(raw, draft).slide.sources).toEqual([]);
    draft.slide.sources = [{ id: "invented", title: "Invented citation" }];
    expect(() => validateDecompositionDraft(raw, draft)).toThrow(/frozen raw.source/);
  });

  it("rejects topology mismatches and inputs beyond dedicated diagram capacity", () => {
    for (const [content, relation] of examples) {
      const { draft } = fixture(content, relation);
      draft.slide.logicalRelation = relation === "cycle" ? "context" : "cycle";
      expect(() => parseDecompositionDraft(draft)).toThrow(/topology/);
    }
    expect(() => parseDecompositionDraft(fixture(examples[0]![0], "comparison").draft)).toThrow(/topology/);
    expect(parseDecompositionDraft(fixture(examples[0]![0], "pros_cons").draft).slide.logicalRelation).toBe("pros_cons");
    const oversized: SemanticContent[] = [
      { kind: "points", items: [node("one")] },
      { kind: "points", items: Array.from({ length: 5 }, (_, i) => node(`${i}`)) },
      { kind: "sequence", steps: Array.from({ length: 6 }, (_, i) => node(`${i}`)) },
      { kind: "cycle", steps: Array.from({ length: 5 }, (_, i) => node(`${i}`)) },
    ];
    for (const content of oversized) {
      const relation = content.kind === "points" ? "claim_evidence" : content.kind === "sequence" ? "process" : "cycle";
      expect(() => parseDecompositionDraft(fixture(content, relation).draft)).toThrow();
    }
  });

  it("keeps source scope separate from rendered center, outcome and environment", () => {
    for (const [content, relation] of examples) {
      const { raw, draft } = fixture(content, relation);
      const packet = buildDecompositionReview(raw, draft);
      expect(packet.sourceRoot).toEqual({ id: raw.id, title: raw.title, scope: draft.slide.decomposition.scope });
      expect(packet.graph.nodes.map(n => n.id)).not.toContain(raw.id);
      expect(packet.nodeExcerpts).toHaveLength(contentItemEntries(content).length);
    }
  });

  it("does not let whole-message or relation anchors hide source coverage gaps", () => {
    const { raw, draft } = fixture();
    raw.body += "追加の未採用事項。\n";
    draft.sourceHash = sourceFingerprint(raw);
    draft.messageSources = [resolveExactQuote(raw, raw.body)];
    draft.relation.spans = [resolveExactQuote(raw, raw.body)];
    const packet = buildDecompositionReview(raw, draft);
    expect(packet.coverage.gaps).toHaveLength(1);
    const gap = packet.coverage.gaps[0]!;
    expect(gap.quote).toBe("追加の未採用事項。");
    expect(packet.warnings.map(w => w.id)).toContain(`gap:${gap.start}:${gap.end}`);
    const review = approval(raw, draft);
    review.dispositions = review.dispositions.filter(d => !d.id.startsWith("gap:"));
    expect(() => compileDecomposition(raw, draft, review)).toThrow(/gap:/);
    const compiled = compileDecomposition(raw, draft, approval(raw, draft));
    expect(compiled.lineage.semanticMECE).toBe("unverified");
    expect(compiled.lineage.packet.coverage.gaps).toHaveLength(1);
    draft.excludedSpans.push({ span: gap, reason: "Explicitly outside this slide's declared scope" });
    expect(buildDecompositionReview(raw, draft).coverage.gaps).toEqual([]);
  });

  it("keeps interpreted labels and declared unknowns unverified after review", () => {
    const { raw, draft } = fixture();
    contentItemEntries(draft.slide.content)[0]!.item.label = "Reviewer interpretation";
    draft.nodeSources[0]!.origin = "interpretation";
    draft.slide.decomposition.unresolvedQuestions = ["Does the declared scope cover the intended problem?"];
    const packet = buildDecompositionReview(raw, draft);
    expect(packet.warnings.map(w => w.id)).toEqual(["semantic-support", "semantic-mece", "question:1"]);
    const compiled = compileDecomposition(raw, draft, approval(raw, draft));
    expect(compiled.presentation.slides[0]!.decomposition!.unresolvedQuestions).toEqual(draft.slide.decomposition.unresolvedQuestions);
    expect(compiled.lineage.semanticMECE).toBe("unverified");
  });

  it("requires an explicit approve record bound to exact source and draft fingerprints", () => {
    const { raw, draft } = fixture();
    const review = approval(raw, draft);
    expect(() => compileDecomposition(raw, draft, undefined)).toThrow();
    expect(() => compileDecomposition(raw, draft, { ...review, decision: "request-changes" })).toThrow(/approve/);
    expect(() => compileDecomposition({ ...raw, body: `${raw.body}new` }, draft, review)).toThrow(/sourceHash/);
    const changedSource = { ...raw, title: "Changed title" };
    expect(() => compileDecomposition(changedSource, { ...draft, sourceHash: sourceFingerprint(changedSource) }, review)).toThrow(/sourceHash/);
    const changedDraft = { ...draft, id: "candidate-2" };
    expect(draftFingerprint(changedDraft)).not.toBe(draftFingerprint(draft));
    expect(() => compileDecomposition(raw, changedDraft, review)).toThrow(/draftHash/);
    draft.slide.keyMessage = "A materially changed assertion";
    expect(() => compileDecomposition(raw, draft, review)).toThrow(/draftHash/);
  });

  it("rejects missing, duplicated, unknown or unexplained warning dispositions", () => {
    const { raw, draft } = fixture();
    const review = approval(raw, draft);
    expect(() => compileDecomposition(raw, draft, { ...review, dispositions: [] })).toThrow(/semantic-support/);
    expect(() => compileDecomposition(raw, draft, { ...review, dispositions: [...review.dispositions, review.dispositions[0]] })).toThrow(/Duplicate/);
    expect(() => compileDecomposition(raw, draft, { ...review, dispositions: [...review.dispositions, { ...review.dispositions[0], id: "imaginary" }] })).toThrow(/Unknown/);
    expect(() => parseDecompositionReviewRecord({ ...review, dispositions: [{ ...review.dispositions[0], reason: " " }] })).toThrow();
    expect(() => parseDecompositionReviewRecord({ ...review, dispositions: [{ ...review.dispositions[0], decision: "resolved" }] })).toThrow();
  });

  it("is strict, bounded, and never accepts rendering or approval fields in a draft", () => {
    const { raw, draft } = fixture();
    expect(() => parseRawBody({ ...raw, body: "😀".repeat(10_001) })).toThrow();
    expect(() => parseRawBody({ ...raw, body: " \r\n" })).toThrow();
    expect(() => parseRawBody({ ...raw, extra: "ignored?" })).toThrow();
    expect(() => parseDecompositionDraft({ ...draft, status: "approved" })).toThrow();
    expect(() => parseDecompositionDraft({ ...draft, slide: { ...draft.slide, x: 1 } })).toThrow();
    expect(() => parseDecompositionDraft({ ...draft, nodeSources: [{ ...draft.nodeSources[0], provenanceVerified: true }] })).toThrow();
  });
});
