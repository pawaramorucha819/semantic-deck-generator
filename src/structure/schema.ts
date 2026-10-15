import { z } from "zod";
import { contentItemEntries } from "../domain/content.js";
import { DecompositionSchema, SemanticContentSchema, SlideSpecSchema, SourceSchema } from "../domain/schema.js";
import type { LogicalRelation, SemanticContent } from "../domain/types.js";

export const MAX_RAW_BODY_LENGTH = 20_000;
const text = (max = 1200) => z.string().trim().min(1).max(max);
const hash = z.string().regex(/^[a-f0-9]{64}$/, "Expected a lowercase SHA-256 fingerprint");

/** Offsets are JavaScript UTF-16 code units, with an exclusive end. Quotes are never trimmed. */
export const SourceSpanSchema = z.object({
  start: z.number().int().min(0).max(MAX_RAW_BODY_LENGTH),
  end: z.number().int().min(1).max(MAX_RAW_BODY_LENGTH),
  quote: z.string().min(1).max(MAX_RAW_BODY_LENGTH).refine(value => value.trim().length > 0, "A source quote must contain non-whitespace text"),
}).strict().refine(span => span.end > span.start, "Source span end must be greater than start");
export type SourceSpan = z.infer<typeof SourceSpanSchema>;

/** Local input only: source.url is metadata and is never fetched. Body bytes/text are not normalized. */
export const RawBodySchema = z.object({
  schemaVersion: z.literal("1.0"),
  id: text(120),
  title: text(200),
  body: z.string().min(1).max(MAX_RAW_BODY_LENGTH).refine(value => value.trim().length > 0, "Raw body must contain non-whitespace text"),
  source: SourceSchema.optional(),
}).strict();
export type RawBody = z.infer<typeof RawBodySchema>;

/** Authorship is a declaration, never an authenticated identity or user approval. */
export const DecompositionActorSchema = z.object({
  kind: z.enum(["human", "assistant", "fixture"]),
  id: text(120),
}).strict();
export type DecompositionActor = z.infer<typeof DecompositionActorSchema>;

const allowedRelations: Record<SemanticContent["kind"], readonly LogicalRelation[]> = {
  points: ["claim_evidence", "pros_cons"],
  sequence: ["process", "timeline"],
  distribution: ["distribution"],
  convergence: ["convergence"],
  cycle: ["cycle"],
  context: ["context"],
};

// Reuse the existing semantic contract and its reference checks; only tighten the upstream boundary.
const DraftSlideSchema = SlideSpecSchema.innerType().extend({
  content: SemanticContentSchema,
  decomposition: DecompositionSchema,
  evidence: SlideSpecSchema.innerType().shape.evidence.max(0, "This bounded body-decomposition route requires evidence=[]; evidence-bearing selection needs a future body-preservation contract"),
}).superRefine((slide, ctx) => {
  const existing = SlideSpecSchema.safeParse(slide);
  if (!existing.success) existing.error.issues.forEach(issue => ctx.addIssue(issue));
  if (!allowedRelations[slide.content.kind].includes(slide.logicalRelation)) {
    ctx.addIssue({ code: "custom", path: ["logicalRelation"], message: `logicalRelation ${slide.logicalRelation} does not match ${slide.content.kind} topology` });
  }
  contentItemEntries(slide.content).forEach(({ item, path }) => {
    if (!item.id) ctx.addIssue({ code: "custom", path: ["content", ...path, "id"], message: "Every draft content node requires an explicit stable id" });
  });
  if (slide.content.kind === "points" && (slide.content.items.length < 2 || slide.content.items.length > 4)) {
    ctx.addIssue({ code: "custom", path: ["content", "items"], message: "The dedicated points diagram supports 2–4 nodes; split the source explicitly instead of dropping nodes" });
  }
  if (slide.content.kind === "sequence" && slide.content.steps.length > 5) {
    ctx.addIssue({ code: "custom", path: ["content", "steps"], message: "The dedicated sequence diagram supports 2–5 nodes; split the source explicitly instead of dropping nodes" });
  }
  const sources = new Set<string>();
  (slide.sources ?? []).forEach((source, i) => {
    if (sources.has(source.id)) ctx.addIssue({ code: "custom", path: ["sources", i, "id"], message: `Duplicate source id: ${source.id}` });
    sources.add(source.id);
  });
});

const spans = () => z.array(SourceSpanSchema).min(1).max(64);
export const DecompositionDraftSchema = z.object({
  schemaVersion: z.literal("1.0"),
  id: text(120),
  status: z.literal("draft"),
  sourceHash: hash,
  author: DecompositionActorSchema,
  slide: DraftSlideSchema,
  nodeSources: z.array(z.object({
    nodeId: text(120),
    spans: spans(),
    origin: z.enum(["source", "interpretation"]),
    rationale: text(),
  }).strict()).min(1).max(6),
  messageSources: spans(),
  relation: z.object({ rationale: text(), spans: spans() }).strict(),
  excludedSpans: z.array(z.object({ span: SourceSpanSchema, reason: text() }).strict()).max(64),
}).strict().superRefine((draft, ctx) => {
  const expected = new Set(contentItemEntries(draft.slide.content).map(({ item }) => item.id));
  const seen = new Set<string>();
  draft.nodeSources.forEach((source, i) => {
    if (!expected.has(source.nodeId)) ctx.addIssue({ code: "custom", path: ["nodeSources", i, "nodeId"], message: `Dangling node source id: ${source.nodeId}` });
    if (seen.has(source.nodeId)) ctx.addIssue({ code: "custom", path: ["nodeSources", i, "nodeId"], message: `Duplicate node source id: ${source.nodeId}` });
    seen.add(source.nodeId);
  });
  for (const nodeId of expected) {
    if (nodeId && !seen.has(nodeId)) ctx.addIssue({ code: "custom", path: ["nodeSources"], message: `Missing source anchors for node: ${nodeId}` });
  }
});
export type DecompositionDraft = z.infer<typeof DecompositionDraftSchema>;

export const DecompositionReviewRecordSchema = z.object({
  schemaVersion: z.literal("1.0"),
  sourceHash: hash,
  draftHash: hash,
  decision: z.enum(["approve", "request-changes"]),
  actor: DecompositionActorSchema,
  rationale: text(4000),
  dispositions: z.array(z.object({
    id: text(120),
    decision: z.enum(["accepted-limitation", "out-of-scope"]),
    reason: text(),
  }).strict()).max(MAX_RAW_BODY_LENGTH + 20),
}).strict();
export type DecompositionReviewRecord = z.infer<typeof DecompositionReviewRecordSchema>;

export const parseRawBody = (input: unknown): RawBody => RawBodySchema.parse(input);
export const parseDecompositionDraft = (input: unknown): DecompositionDraft => DecompositionDraftSchema.parse(input);
export const parseDecompositionReviewRecord = (input: unknown): DecompositionReviewRecord => DecompositionReviewRecordSchema.parse(input);
