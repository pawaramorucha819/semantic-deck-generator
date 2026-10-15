import { parsePresentation } from "../domain/schema.js";
import { parseDecompositionReviewRecord } from "./schema.js";
import { buildDecompositionReview } from "./review.js";

/** Compile only the existing PresentationSpec; full review/provenance stays in companion lineage. */
export function compileDecomposition(rawInput: unknown, draftInput: unknown, reviewInput: unknown) {
  const packet = buildDecompositionReview(rawInput, draftInput);
  const review = parseDecompositionReviewRecord(reviewInput);
  if (review.decision !== "approve") throw new Error("Compilation requires an explicit approve review decision");
  if (review.sourceHash !== packet.sourceHash) throw new Error("Review sourceHash is stale or does not match the raw source");
  if (review.draftHash !== packet.draftHash) throw new Error("Review draftHash is stale or does not match the exact draft revision");
  const expected = new Set(packet.warnings.map(warning => warning.id));
  const seen = new Set<string>();
  for (const disposition of review.dispositions) {
    if (seen.has(disposition.id)) throw new Error(`Duplicate review disposition: ${disposition.id}`);
    if (!expected.has(disposition.id)) throw new Error(`Unknown review disposition: ${disposition.id}`);
    seen.add(disposition.id);
  }
  const missing = [...expected].filter(id => !seen.has(id));
  if (missing.length) throw new Error(`Review must explicitly acknowledge every warning: ${missing.join(", ")}`);
  const presentation = parsePresentation({ title: packet.raw.title, slides: [packet.draft.slide] });
  return {
    presentation,
    lineage: {
      schemaVersion: "1.0" as const,
      sourceHash: packet.sourceHash,
      draftHash: packet.draftHash,
      authentication: "self-declared" as const,
      semanticMECE: "unverified" as const,
      review,
      packet,
    },
  };
}
export type CompiledDecomposition = ReturnType<typeof compileDecomposition>;
