import { z } from "zod";
import type { ContentItem, SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { centerX, centerY, columns, inset, rows, type Rect } from "../../render/geometry.js";
import type { VisualPattern } from "../pattern.js";
import { drawFooter, drawHeader } from "../shared/layout.js";
import { buildSourceText, roleLabel } from "../shared/semantics.js";

// Each kind expresses a different semantic relationship, not a cosmetic variant.
const nodeSchema = z.object({
  label: z.string().min(1), description: z.string().min(1).optional(),
  nodeId: z.string().min(1).optional(), evidenceRefs: z.array(z.string().min(1)).optional(),
  members: z.array(z.string().min(1)).optional(),
}).strict();
const common = {
  headline: z.string().min(1), eyebrow: z.string().optional(), subline: z.string().optional(),
  sourceText: z.string().optional(),
};
const radialSchema = z.object({ ...common, center: nodeSchema, items: z.array(nodeSchema).min(2).max(4) }).strict();
const convergenceSchema = z.object({ ...common, inputs: z.array(nodeSchema).min(2).max(4), outcome: nodeSchema }).strict();
const cycleSchema = z.object({ ...common, steps: z.array(nodeSchema).min(2).max(4) }).strict();
const backgroundSchema = z.object({ ...common, environment: nodeSchema, items: z.array(nodeSchema).min(1).max(4) }).strict();
type Node = z.infer<typeof nodeSchema>;

function node(item: ContentItem): Node {
  return { label: item.label, ...(item.description ? { description: item.description } : {}),
    ...(item.id ? { nodeId: item.id } : {}), ...(item.evidenceRefs ? { evidenceRefs: item.evidenceRefs } : {}),
    ...(item.members ? { members: item.members } : {}),
  };
}
function header(slide: SlideSpec) {
  return { headline: slide.keyMessage, eyebrow: roleLabel(slide.slideRole), subline: slide.supportingMessage, sourceText: buildSourceText(slide) };
}
function area(b: z.infer<typeof radialSchema> | z.infer<typeof convergenceSchema> | z.infer<typeof cycleSchema> | z.infer<typeof backgroundSchema>, ctx: RenderContext) {
  return drawFooter(ctx, drawHeader(ctx, b), b.sourceText);
}
function drawNode(ctx: RenderContext, item: Node, box: Rect, name: string, primary = false) {
  ctx.shape("rect", box, { fill: primary ? "primary" : "background", line: { color: "primary", width: ctx.theme.shape.borderWidth }, name: `${name}-box` });
  const inner = inset(box, ctx.theme.spacing.sm);
  const labelH = ctx.measureText(item.label, { widthIn: inner.w, role: "heading", lineHeight: 1.2 });
  const detailH = item.description ? ctx.measureText(item.description, { widthIn: inner.w, role: "body", lineHeight: 1.25 }) : 0;
  const gap = item.description ? ctx.theme.spacing.xs : 0;
  if (inner.w <= 0 || labelH + gap + detailH > inner.h + 0.001) {
    throw new Error(`[${name}] ノード内の文字が収まりません (${item.nodeId ?? item.label})。内容を短くするかスライドを分けてください。必要高さ ${(labelH + gap + detailH).toFixed(2)}in / 枠 ${inner.h.toFixed(2)}in`);
  }
  const y = inner.y + Math.max(0, (inner.h - labelH - gap - detailH) / 2);
  ctx.text({ ...inner, y, h: labelH }, item.label, { role: "heading", color: primary ? "onPrimary" : "text", bold: true, align: "center", lineHeight: 1.2, name: `${name}-label` });
  if (item.description) ctx.text({ ...inner, y: y + labelH + gap, h: detailH }, item.description, { role: "body", color: primary ? "onPrimary" : "mutedText", align: "center", lineHeight: 1.25, name: `${name}-description` });
}

/** Clip directed connectors to node boundaries. Native line arrows remain editable. */
function connect(ctx: RenderContext, from: Rect, to: Rect, name: string, arrow = true) {
  const dx = centerX(to) - centerX(from), dy = centerY(to) - centerY(from);
  const trim = (box: Rect) => 1 / Math.max(Math.abs(dx) / (box.w / 2), Math.abs(dy) / (box.h / 2));
  const a = trim(from), b = trim(to);
  const x1 = centerX(from) + dx * a, y1 = centerY(from) + dy * a;
  const x2 = centerX(to) - dx * b, y2 = centerY(to) - dy * b;
  ctx.shape("line", { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) }, {
    line: { color: "primary", width: ctx.theme.shape.borderWidth * 1.5, ...(arrow ? { endArrowType: "triangle" as const } : {}) },
    flipH: x2 < x1, flipV: y2 < y1, name,
  });
}

export const radialPattern: VisualPattern<z.infer<typeof radialSchema>> = {
  id: "radial", version: "1.0.0", description: "一つの中心から複数の対象へ向かう関係を放射図にする。", variants: ["standard"],
  compatibility: { logicalRelations: ["distribution"] }, bindingsSchema: radialSchema,
  match: s => ({ compatible: s.content?.kind === "distribution", score: s.content?.kind === "distribution" ? 1 : 0, reason: "distribution の中心と送り先を外向き矢印で保持" }),
  buildBindings(s) { if (s.content?.kind !== "distribution") throw new Error("radial requires distribution"); return { ...header(s), center: node(s.content.center), items: s.content.items.map(node) }; },
  render(b, ctx) {
    const r = area(b, ctx), w = r.w * 0.29, h = r.h * 0.34;
    const center = { x: centerX(r) - r.w * 0.115, y: centerY(r) - h / 2, w: r.w * 0.23, h };
    const left = r.x, right = r.x + r.w - w, top = r.y, lower = r.y + r.h - h, middle = centerY(r) - h / 2;
    const positions = b.items.length === 2 ? [[left, middle], [right, middle]] : b.items.length === 3 ? [[left, middle], [right, top], [right, lower]] : [[left, top], [right, top], [right, lower], [left, lower]];
    const boxes = positions.map(([x, y]) => ({ x: x!, y: y!, w, h }));
    boxes.forEach((box, i) => connect(ctx, center, box, `radial-center-to-${i + 1}`));
    drawNode(ctx, b.center, center, "radial-center", true);
    boxes.forEach((box, i) => drawNode(ctx, b.items[i]!, box, `radial-target-${i + 1}`));
  },
};

export const convergencePattern: VisualPattern<z.infer<typeof convergenceSchema>> = {
  id: "convergence", version: "1.0.0", description: "複数の入力が一つの成果へ合流する関係を描く。", variants: ["standard"],
  compatibility: { logicalRelations: ["convergence"] }, bindingsSchema: convergenceSchema,
  match: s => ({ compatible: s.content?.kind === "convergence", score: s.content?.kind === "convergence" ? 1 : 0, reason: "convergence の入力から共通の成果への方向を保持" }),
  buildBindings(s) { if (s.content?.kind !== "convergence") throw new Error("convergence requires convergence content"); return { ...header(s), inputs: s.content.inputs.map(node), outcome: node(s.content.outcome) }; },
  render(b, ctx) {
    const r = area(b, ctx);
    const inputs = rows({ ...r, w: r.w * 0.43 }, b.inputs.length, ctx.theme.spacing.md);
    const outcome = { x: r.x + r.w * 0.64, y: r.y + r.h * 0.27, w: r.w * 0.36, h: r.h * 0.46 };
    inputs.forEach((box, i) => connect(ctx, box, outcome, `convergence-input-${i + 1}-to-outcome`));
    inputs.forEach((box, i) => drawNode(ctx, b.inputs[i]!, box, `convergence-input-${i + 1}`));
    drawNode(ctx, b.outcome, outcome, "convergence-outcome", true);
  },
};

export const cyclePattern: VisualPattern<z.infer<typeof cycleSchema>> = {
  id: "cycle", version: "1.0.0", description: "順序を持つ工程を、最終工程から最初へ戻る閉路として描く。", variants: ["standard"],
  compatibility: { logicalRelations: ["cycle"] }, bindingsSchema: cycleSchema,
  match: s => ({ compatible: s.content?.kind === "cycle", score: s.content?.kind === "cycle" ? 1 : 0, reason: "cycle の順序と最後から最初への戻りを保持" }),
  buildBindings(s) { if (s.content?.kind !== "cycle") throw new Error("cycle requires cycle content"); return { ...header(s), steps: s.content.steps.map(node) }; },
  render(b, ctx) {
    const r = area(b, ctx), w = r.w * 0.34, h = r.h * 0.34;
    const left = r.x, right = r.x + r.w - w, middle = centerX(r) - w / 2, bottom = r.y + r.h - h;
    const positions = b.steps.length === 2 ? [[left, r.y], [right, bottom]] : b.steps.length === 3 ? [[middle, r.y], [right, bottom], [left, bottom]] : [[left, r.y], [right, r.y], [right, bottom], [left, bottom]];
    const boxes = positions.map(([x, y]) => ({ x: x!, y: y!, w, h }));
    // Two-node cycles use separate horizontal/vertical paths to avoid coincident opposing arrows.
    if (boxes.length === 2) {
      const a = boxes[0]!, b2 = boxes[1]!;
      const bend1 = { x: centerX(b2), y: centerY(a), w: 0.001, h: 0.001 };
      const bend2 = { x: centerX(a), y: centerY(b2), w: 0.001, h: 0.001 };
      connect(ctx, a, bend1, "cycle-step-1-via", false); connect(ctx, bend1, b2, "cycle-step-1-to-2");
      connect(ctx, b2, bend2, "cycle-step-2-via", false); connect(ctx, bend2, a, "cycle-step-2-to-1");
    } else boxes.forEach((box, i) => connect(ctx, box, boxes[(i + 1) % boxes.length]!, `cycle-step-${i + 1}-to-${(i + 1) % boxes.length + 1}`));
    boxes.forEach((box, i) => drawNode(ctx, { ...b.steps[i]!, label: `${i + 1}. ${b.steps[i]!.label}` }, box, `cycle-step-${i + 1}`));
  },
};

export const backgroundPattern: VisualPattern<z.infer<typeof backgroundSchema>> = {
  id: "background", version: "1.0.0", description: "共通の環境・基盤と、その上にある要素を描く。因果矢印は追加しない。", variants: ["standard"],
  compatibility: { logicalRelations: ["context"] }, bindingsSchema: backgroundSchema,
  match: s => ({ compatible: s.content?.kind === "context", score: s.content?.kind === "context" ? 1 : 0, reason: "context の共通基盤への所属を面で保持。因果や順番は推定しない" }),
  buildBindings(s) { if (s.content?.kind !== "context") throw new Error("background requires context content"); return { ...header(s), environment: node(s.content.environment), items: s.content.items.map(node) }; },
  render(b, ctx) {
    const r = area(b, ctx);
    ctx.shape("rect", r, { fill: "surface", line: { color: "border" }, name: "background-shared-environment", layer: "background" });
    const inside = inset(r, ctx.theme.spacing.md);
    const baseH = inside.h * 0.3;
    const base = { ...inside, y: inside.y + inside.h - baseH, h: baseH };
    const top = { ...inside, h: inside.h - baseH - ctx.theme.spacing.md };
    const boxes = columns(top, b.items.length, ctx.theme.spacing.md);
    boxes.forEach((box, i) => drawNode(ctx, b.items[i]!, box, `background-member-${i + 1}`));
    drawNode(ctx, b.environment, base, "background-environment", true);
  },
};

export const relationshipPatterns = [radialPattern, convergencePattern, cyclePattern, backgroundPattern];
