import { contentItemEntries } from "../domain/content.js";
import { relationshipGraph } from "../review/structure.js";
import { draftFingerprint, sourceFingerprint, validateDecompositionDraft } from "./provenance.js";
import { parseRawBody, type SourceSpan } from "./schema.js";

export interface DecompositionWarning {
  id: string;
  kind: "semantic-support" | "semantic-mece" | "coverage-gap" | "unresolved-question";
  message: string;
  span?: SourceSpan;
}

/** Only node anchors and explicit exclusions count toward body coverage. */
function coverageGaps(body: string, spans: SourceSpan[]): SourceSpan[] {
  const sorted = [...spans].sort((a, b) => a.start - b.start || a.end - b.end);
  const gaps: SourceSpan[] = [];
  function addGap(start: number, end: number) {
    // Trim whitespace from the gap boundaries, preserving exact offsets and source text.
    while (start < end && body[start]!.trim() === "") start++;
    while (end > start && body[end - 1]!.trim() === "") end--;
    if (start < end) gaps.push({ start, end, quote: body.slice(start, end) });
  }
  let cursor = 0;
  for (const span of sorted) {
    if (span.start > cursor) addGap(cursor, span.start);
    cursor = Math.max(cursor, span.end);
  }
  addGap(cursor, body.length);
  return gaps;
}

/** A deterministic review packet; it does not approve a draft or certify semantic MECE. */
export function buildDecompositionReview(rawInput: unknown, draftInput: unknown) {
  const raw = parseRawBody(rawInput);
  const draft = validateDecompositionDraft(raw, draftInput);
  const gaps = coverageGaps(raw.body, [
    ...draft.nodeSources.flatMap(node => node.spans),
    ...draft.excludedSpans.map(excluded => excluded.span),
  ]);
  const warnings: DecompositionWarning[] = [
    { id: "semantic-support", kind: "semantic-support", message: "Exact source quotes prove occurrence only. Node labels, key message, relation and interpretations still require semantic review." },
    { id: "semantic-mece", kind: "semantic-mece", message: "Semantic mutual exclusivity and completeness remain unverified, including the chosen source scope and decomposition criterion." },
    ...gaps.map(span => ({ id: `gap:${span.start}:${span.end}`, kind: "coverage-gap" as const, message: "This source text is neither anchored to a node nor explicitly excluded.", span })),
    ...(draft.slide.decomposition.unresolvedQuestions ?? []).map((question, i) => ({ id: `question:${i + 1}`, kind: "unresolved-question" as const, message: question })),
  ];
  const sources = new Map(draft.nodeSources.map(source => [source.nodeId, source]));
  return {
    schemaVersion: "1.0" as const,
    assessment: "requires-explicit-review" as const,
    authentication: "self-declared" as const,
    semanticMECE: "unverified" as const,
    sourceHash: sourceFingerprint(raw),
    draftHash: draftFingerprint(draft),
    raw,
    draft,
    // This is the decomposition subject, never a synthetic rendered center or environment.
    sourceRoot: { id: raw.id, title: raw.title, scope: draft.slide.decomposition.scope },
    graph: relationshipGraph(draft.slide.content),
    nodeExcerpts: contentItemEntries(draft.slide.content).map(({ item, path }) => ({
      ...sources.get(item.id!)!, inputPath: path.join("."), label: item.label,
    })),
    coverage: { basis: "node-spans-and-explicit-exclusions" as const, gaps, excludedSpans: draft.excludedSpans },
    warnings,
  };
}
export type DecompositionReviewPacket = ReturnType<typeof buildDecompositionReview>;
