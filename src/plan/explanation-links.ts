import { contentItemEntries } from "../domain/content.js";
import type { SlideSpec } from "../domain/types.js";

/** Resolve semantic IDs to redundant visible labels in the expression layer.
 * The source remains untouched. Numbers are visual keys, never causal or order claims.
 */
export function resolveExplanationLinks(source: SlideSpec): SlideSpec {
  if (!source.explanation?.sections.some(section => section.nodeRefs?.length)) return source;
  const slide = structuredClone(source);
  const labels = new Map<string, string>();
  contentItemEntries(slide.content).forEach(({ item }, index) => {
    if (!item.id) return;
    item.label = `[${index + 1}] ${item.label}`;
    labels.set(item.id, item.label);
  });
  slide.explanation!.sections = slide.explanation!.sections.map(section => {
    const target = section.nodeRefs?.map(ref => {
      const label = labels.get(ref);
      if (!label) throw new Error(`Unknown explanation node reference: ${ref}`);
      return label;
    }).join(" ／ ");
    // Preserve semantic references beside resolved display text for traceability.
    return { heading: target ? (section.nodeRefs?.length === 1 && target.endsWith(`] ${section.heading}`) ? target : `${target}｜${section.heading}`) : `全体｜${section.heading}`, body: section.body, ...(section.nodeRefs ? { nodeRefs: [...section.nodeRefs] } : {}) };
  });
  return slide;
}
