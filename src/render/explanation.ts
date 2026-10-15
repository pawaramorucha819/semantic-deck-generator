import type { SlideExplanation } from "../domain/types.js";
import type { RenderContext } from "./context.js";
import type { Rect } from "./geometry.js";

/** Compose visible explanatory prose beside the intact diagram. Never shrink or truncate. */
export function drawExplanation(ctx: RenderContext, area: Rect, explanation: SlideExplanation): Rect {
  const gap = Math.max(0.25, ctx.theme.spacing.lg);
  const diagramWidth = (area.w - gap) * 0.40;
  const prose = { x: area.x + diagramWidth + gap, y: area.y, w: area.w - diagramWidth - gap, h: area.h };
  const bodySize = Math.max(14, ctx.theme.fontSizes.body);
  const headingSize = Math.max(14, ctx.theme.fontSizes.heading);
  const blocks = [
    { text: explanation.readerQuestion, size: headingSize, bold: true, name: "explanation-question", gap: 0.12 },
    ...explanation.sections.flatMap((section, index) => [
      { text: section.heading, size: headingSize, bold: true, name: `explanation-heading-${index + 1}`, gap: 0.03 },
      { text: section.body, size: bodySize, bold: false, name: `explanation-body-${index + 1}`, gap: 0.12 },
    ]),
  // Font shaping and CJK punctuation wrapping differ between renderers. Reserve
  // explicit slack instead of trusting estimated text height as a hard bound.
  ].map(block => ({ ...block, h: ctx.measureText(block.text, { widthIn: prose.w, fontSize: block.size, lineHeight: 1.25 }) * 1.12 + 0.03 }));
  const height = blocks.reduce((sum, block) => sum + block.h + block.gap, 0) - blocks.at(-1)!.gap;
  if (prose.w <= 0 || area.h <= 0 || height > prose.h + 0.001) {
    throw new Error(`Visible explanation overflows (${height.toFixed(2)}in required, ${prose.h.toFixed(2)}in available); split the slide or shorten the explanation without removing necessary context`);
  }
  let y = prose.y;
  for (const block of blocks) {
    ctx.text({ x: prose.x, y, w: prose.w, h: block.h }, block.text, {
      fontSize: block.size, bold: block.bold, color: block.bold ? "primary" : "text",
      lineHeight: 1.25, shrink: false, name: block.name,
    });
    y += block.h + block.gap;
  }
  return { ...area, w: diagramWidth };
}
