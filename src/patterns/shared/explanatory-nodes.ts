import type { RenderContext } from "../../render/context.js";
import { rows, inset, type Rect } from "../../render/geometry.js";

/** Vertical cards keep diagram labels readable beside explanatory prose. */
export function drawExplanatoryNodes(
  ctx: RenderContext,
  body: Rect,
  nodes: Array<{ label: string; description?: string; metric?: string; name: string }>,
  ordered: boolean,
): void {
  const gap = ordered ? 0.24 : 0.16;
  const boxes = rows(body, nodes.length, gap);
  nodes.forEach((node, index) => {
    const box = boxes[index]!;
    ctx.shape("rect", box, { fill: "surface", line: { color: "primary", width: 1 }, layer: "decoration", name: `${node.name}-box` });
    const inner = inset(box, 0.09);
    const pieces = [
      { text: node.label, size: Math.max(15, ctx.theme.fontSizes.heading), bold: true, suffix: "label" },
      ...(node.metric ? [{ text: node.metric, size: Math.max(14, ctx.theme.fontSizes.body), bold: true, suffix: "metric" }] : []),
      ...(node.description ? [{ text: node.description, size: Math.max(14, ctx.theme.fontSizes.body), bold: false, suffix: "description" }] : []),
    ].map(p => ({ ...p, h: ctx.measureText(p.text, { widthIn: inner.w, fontSize: p.size, lineHeight: 1.2 }) * 1.1 + 0.02 }));
    const total = pieces.reduce((sum, p) => sum + p.h, 0) + (pieces.length - 1) * 0.04;
    if (total > inner.h || inner.w <= 0) throw new Error(`[${node.name}] Vertical explanatory diagram text overflows; split the slide or shorten labels`);
    let y = inner.y + (inner.h - total) / 2;
    pieces.forEach(p => {
      ctx.text({ x: inner.x, y, w: inner.w, h: p.h }, p.text, {
        fontSize: p.size, bold: p.bold, align: "center", lineHeight: 1.2, shrink: false,
        color: p.bold ? "primary" : "text", name: `${node.name}-${p.suffix}`,
      });
      y += p.h + 0.04;
    });
    if (ordered && index < boxes.length - 1) {
      ctx.shape("downArrow", { x: body.x + body.w / 2 - 0.075, y: box.y + box.h + 0.035, w: 0.15, h: gap - 0.07 }, {
        fill: "primary", layer: "decoration", name: `${node.name}-to-next`,
      });
    }
  });
}
