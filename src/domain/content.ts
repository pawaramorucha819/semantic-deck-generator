import type { ContentItem, RelationshipContent, SemanticContent } from "./types.js";

export interface ContentItemEntry {
  item: ContentItem;
  /** Original path relative to content, for diagnostics and provenance. */
  path: Array<string | number>;
}

/** Enumerate every node without reinterpreting its relationship as a list. */
export function contentItemEntries(content?: SemanticContent): ContentItemEntry[] {
  if (!content) return [];
  const entries = (items: ContentItem[], key: string): ContentItemEntry[] =>
    items.map((item, index) => ({ item, path: [key, index] }));
  switch (content.kind) {
    case "points":
      return entries(content.items, "items");
    case "sequence":
    case "cycle":
      return entries(content.steps, "steps");
    case "distribution":
      return [{ item: content.center, path: ["center"] }, ...entries(content.items, "items")];
    case "convergence":
      return [...entries(content.inputs, "inputs"), { item: content.outcome, path: ["outcome"] }];
    case "context":
      return [{ item: content.environment, path: ["environment"] }, ...entries(content.items, "items")];
  }
}

/** Node enumeration for inspection/search only, not an expression conversion. */
export function contentItems(content?: SemanticContent): ContentItem[] {
  return contentItemEntries(content).map(({ item }) => item);
}

export function isRelationshipContent(content?: SemanticContent): content is RelationshipContent {
  return content !== undefined && content.kind !== "points" && content.kind !== "sequence";
}
