import { createHash } from "node:crypto";
import {
  parseDecompositionDraft, parseRawBody, SourceSpanSchema,
  type DecompositionDraft, type RawBody, type SourceSpan,
} from "./schema.js";

/** Canonical JSON: sorted object keys, preserved array order and exact string content. */
export function canonicalJson(value: unknown): string {
  const ancestors = new Set<object>();
  function encode(item: unknown): string {
    if (item === null || typeof item === "string" || typeof item === "boolean") return JSON.stringify(item);
    if (typeof item === "number" && Number.isFinite(item)) return JSON.stringify(item);
    if (typeof item !== "object" || item === null) throw new Error("Fingerprint requires a finite JSON value");
    if (ancestors.has(item)) throw new Error("Fingerprint cannot encode circular values");
    if (!Array.isArray(item) && Object.getPrototypeOf(item) !== Object.prototype && Object.getPrototypeOf(item) !== null) {
      throw new Error("Fingerprint requires plain JSON objects");
    }
    ancestors.add(item);
    const encoded = Array.isArray(item)
      ? `[${Array.from(item, encode).join(",")}]`
      : `{${Object.keys(item).sort().filter(key => (item as Record<string, unknown>)[key] !== undefined)
        .map(key => `${JSON.stringify(key)}:${encode((item as Record<string, unknown>)[key])}`).join(",")}}`;
    ancestors.delete(item);
    return encoded;
  }
  return encode(value);
}
export function fingerprint(value: unknown): string {
  return createHash("sha256").update(canonicalJson(value), "utf8").digest("hex");
}
export function sourceFingerprint(raw: unknown): string { return fingerprint(parseRawBody(raw)); }
export function draftFingerprint(draft: unknown): string { return fingerprint(parseDecompositionDraft(draft)); }

function splitsSurrogatePair(body: string, offset: number): boolean {
  const before = body.charCodeAt(offset - 1);
  const after = body.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
}

export function validateSourceSpan(body: string, input: unknown): SourceSpan {
  const span = SourceSpanSchema.parse(input);
  if (span.end > body.length || body.slice(span.start, span.end) !== span.quote) {
    throw new Error(`Source span [${span.start},${span.end}) quote does not exactly match the raw body`);
  }
  if (splitsSurrogatePair(body, span.start) || splitsSurrogatePair(body, span.end)) {
    throw new Error(`Source span [${span.start},${span.end}) splits a UTF-16 surrogate pair`);
  }
  return span;
}

/** A duplicate exact quote requires an explicit start offset; the first match is never guessed. */
export function resolveExactQuote(rawOrBody: RawBody | string, quote: string, start?: number): SourceSpan {
  const body = typeof rawOrBody === "string" ? rawOrBody : parseRawBody(rawOrBody).body;
  if (!quote.length || !quote.trim().length) throw new Error("An exact quote must contain non-whitespace text");
  if (start !== undefined) return validateSourceSpan(body, { start, end: start + quote.length, quote });
  const found = body.indexOf(quote);
  if (found < 0) throw new Error("Exact quote was not found in the raw body");
  if (body.indexOf(quote, found + 1) >= 0) throw new Error("Ambiguous exact quote: provide an explicit start offset");
  return validateSourceSpan(body, { start: found, end: found + quote.length, quote });
}

/** Validate cross-file identity and every declared quotation, without inferring semantic support. */
export function validateDecompositionDraft(rawInput: unknown, draftInput: unknown): DecompositionDraft {
  const raw = parseRawBody(rawInput);
  const draft = parseDecompositionDraft(draftInput);
  if (draft.sourceHash !== sourceFingerprint(raw)) throw new Error("Draft sourceHash does not match the raw source");
  const frozenSources = raw.source ? [raw.source] : [];
  if (fingerprint(draft.slide.sources ?? []) !== fingerprint(frozenSources)) {
    throw new Error("Draft slide.sources must exactly match the frozen raw.source metadata; new or modified sources are not allowed");
  }
  const anchors = [
    ...draft.nodeSources.flatMap(node => node.spans),
    ...draft.messageSources,
    ...draft.relation.spans,
    ...draft.excludedSpans.map(excluded => excluded.span),
  ];
  anchors.forEach(span => validateSourceSpan(raw.body, span));
  return draft;
}
