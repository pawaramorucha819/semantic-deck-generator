import { z } from "zod";
import { parsePresentation } from "../domain/schema.js";
import { computePartition } from "../domain/partition.js";
import { fingerprint } from "../structure/provenance.js";

export const ARGUMENT_CHECKS = ["source-scope", "aggregation", "claim-support", "unknown-values", "reader-comprehension"] as const;
export const ArgumentReviewSchema = z.object({
  schemaVersion: z.literal("1.0"), inputHash: z.string().regex(/^[a-f0-9]{64}$/),
  actor: z.object({ kind: z.enum(["human", "assistant"]), id: z.string().trim().min(1) }).strict(),
  decision: z.enum(["accept", "request-changes"]),
  findings: z.array(z.object({ check: z.enum(ARGUMENT_CHECKS), assessment: z.string().trim().min(10).max(1000) }).strict()).length(ARGUMENT_CHECKS.length),
}).strict();
export type ArgumentReview = z.infer<typeof ArgumentReviewSchema>;

/** Bind an authored semantic review to all claims, data, annotations and source metadata. */
export function buildArgumentReviewPacket(input: unknown) {
  const spec = parsePresentation(input);
  const slides = spec.slides.filter(s => s.argument).map(s => ({
    slideId: s.id, keyMessage: s.keyMessage, supportingMessage: s.supportingMessage, argument: s.argument,
    evidence: s.evidence, sources: s.sources,
    calculations: s.evidence.filter(e => e.type === "partition").map(e => ({ evidenceId: e.id, ...computePartition(e) })),
  }));
  return { schemaVersion: "1.0" as const, inputHash: fingerprint(spec), slides,
    status: "unreviewed" as const, semanticMECE: "unverified" as const,
    limitation: "機械検査は根拠の真偽・主張の妥当性・読者理解を保証しません。レビュー主体は自己申告です。",
    reviewTemplate: { schemaVersion: "1.0", inputHash: fingerprint(spec), actor: { kind: "human", id: "" }, decision: "request-changes",
      findings: ARGUMENT_CHECKS.map(check => ({ check, assessment: "" })) } };
}
export function validateArgumentReview(input: unknown, review: unknown): ArgumentReview {
  const packet = buildArgumentReviewPacket(input);
  if (!packet.slides.length) throw new Error("No explicit argument to review");
  const accepted = ArgumentReviewSchema.parse(review);
  if (accepted.inputHash !== packet.inputHash) throw new Error("Argument review is stale: claim, data, explanation or source changed");
  if (accepted.decision !== "accept") throw new Error("Argument review requests changes");
  if (new Set(accepted.findings.map(f => f.check)).size !== ARGUMENT_CHECKS.length) throw new Error("Every argument review check needs one distinct finding");
  return accepted;
}
