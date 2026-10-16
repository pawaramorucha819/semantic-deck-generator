import { z } from "zod";
import { SourceSchema } from "../../domain/schema.js";
import { PartitionEvidenceSchema, PartitionArgumentSchema, PartitionCalculationSchema, computePartition } from "../../domain/partition.js";
import type { VisualPattern } from "../pattern.js";
import { drawHeader } from "../shared/layout.js";
import { formatNumber } from "../shared/semantics.js";
import type { TextRole } from "../../render/context.js";

/** Saved plans carry both original semantic design and verifiable, fixed arithmetic. */
export const PartitionBindingsSchema = z.object({
  headline: z.string().min(1), subline: z.string().min(1).optional(),
  evidence: PartitionEvidenceSchema,
  argument: PartitionArgumentSchema,
  calculation: PartitionCalculationSchema,
  sources: z.array(SourceSchema).min(1).max(12),
}).strict().superRefine((b, ctx) => {
  try {
    const expected = computePartition(b.evidence);
    for (const key of ["remainder", "selectedFraction", "remainderFraction", "selectedPercent", "remainderPercent"] as const) {
      if (b.calculation[key] !== expected[key]) ctx.addIssue({ code: "custom", path: ["calculation", key], message: "計算結果が入力値と一致しません" });
    }
    // Parse the same strict shape before comparing to avoid depending on JSON property order.
    if (JSON.stringify(PartitionCalculationSchema.parse(b.calculation)) !== JSON.stringify(PartitionCalculationSchema.parse(expected))) {
      ctx.addIssue({ code: "custom", path: ["calculation"], message: "計算の来歴が入力値と一致しません" });
    }
  } catch {
    ctx.addIssue({ code: "custom", path: ["evidence"], message: "分割の入力値が不正です" });
  }
  const ids = new Set(b.sources.map(s => s.id));
  if (ids.size !== b.sources.length || b.evidence.sourceIds.some(id => !ids.has(id)) || b.argument.evidenceId !== b.evidence.id) {
    ctx.addIssue({ code: "custom", message: "根拠・出典参照が一致しません" });
  }
  for (const source of b.sources) if (source.url && !/^https?:\/\//.test(source.url)) {
    ctx.addIssue({ code: "custom", message: "出典リンクは HTTP(S) が必要です" });
  }
});
export type PartitionBindings = z.infer<typeof PartitionBindingsSchema>;

export const partitionWithUnknownsPattern: VisualPattern<PartitionBindings> = {
  id: "partition-with-unknowns", version: "1.0.0",
  description: "全体と既知部分の構成比、および数値未確認の別分類を区別して示す",
  variants: ["standard"], supportedArgumentKinds: ["partition-with-unknowns"],
  compatibility: { logicalRelations: ["composition", "claim_evidence"], evidenceTypes: ["partition"] },
  bindingsSchema: PartitionBindingsSchema,
  match(slide) {
    const compatible = slide.argument?.kind === "partition-with-unknowns" && slide.evidence.length === 1 && slide.evidence[0]?.type === "partition" && !slide.content && !slide.explanation;
    return { compatible, score: compatible ? 1 : 0, reason: compatible ? "母数・採否の部分・未知の品質内訳と主張の関係が明示されている" : "明示的な分割根拠と論拠設計が必要" };
  },
  buildBindings(slide) {
    const evidence = slide.evidence[0];
    if (evidence?.type !== "partition" || !slide.argument) throw new Error("Partition argument required");
    return { headline: slide.keyMessage, subline: slide.supportingMessage, evidence, argument: slide.argument,
      calculation: computePartition(evidence), sources: slide.sources ?? [] };
  },
  render(b, ctx) {
    const area = drawHeader(ctx, { headline: b.headline, subline: b.subline });
    const d = b.evidence.data, c = b.calculation, a = b.argument;
    const gap = ctx.theme.spacing.sm;
    const measure = (text: string, w = area.w, role: TextRole = "body") => ctx.measureText(text, { widthIn: w, role, lineHeight: 1.22 });
    const sourceLines = b.sources.map(s => ({ text: `出典: ${s.title}${s.date ? ` (${s.date})` : ""}${s.accessedAt ? ` / 確認 ${s.accessedAt}` : ""}`, url: s.url }));
    const sourceHeight = sourceLines.reduce((h, s) => h + measure(s.text, area.w, "caption") + .03, 0);
    const bottom = area.y + area.h - sourceHeight - gap;
    let y = area.y;
    function text(text: string, role: TextRole, name: string, bold = false) {
      const h = measure(text, area.w, role);
      ctx.text({ x: area.x, y, w: area.w, h }, text, { role, bold, lineHeight: 1.22, name });
      y += h + .025;
    }
    const mean = d.aggregation.kind === "mean";
    const aggregate = d.aggregation.kind === "mean" ? `${d.aggregation.eventCount}${d.aggregation.eventUnit}の平均に基づく模式図` : "単一の対象集合の分割";
    text(`${d.population} / ${aggregate}`, "caption", "population-aggregation");
    text(`${d.labels.whole} ${formatNumber(d.whole)}${d.unit}（報告値）`, "body", "whole-label", true);
    text(a.annotations.quantity.text, "caption", "quantity-annotation");
    const colW = (area.w - .36) / 2;
    const xs = [area.x, area.x + colW + .36];
    const labels = [
      `${d.labels.selected}\n${mean ? "平均 " : ""}${formatNumber(d.selected)}${d.unit}（報告値） / 約${c.selectedPercent.toFixed(1)}%（計算）`,
      `${d.labels.remainder}\n${mean ? "平均差分 約" : "差分 "}${formatNumber(c.remainder)}${d.unit}（計算） / 約${c.remainderPercent.toFixed(1)}%（計算）`,
    ];
    const lh = Math.max(...labels.map(t => measure(t, colW, "body")));
    labels.forEach((t, i) => ctx.text({x: xs[i]!, y, w: colW, h: lh}, t, { role: "body", bold: true, lineHeight: 1.22, name: `partition-label-${i}` }));
    y += lh + .10;
    // Only inclusion quantity is encoded as area. Unknown relevance never changes widths/colors.
    const bandY = y, bandH = .20, boundary = area.x + area.w * c.selectedFraction;
    if (c.selectedFraction > 0) ctx.shape("rect", { x: area.x, y, w: area.w * c.selectedFraction, h: bandH }, { fill: "primary", layer: "decoration", name: "selected-quantity-band" });
    if (c.remainderFraction > 0) ctx.shape("rect", { x: boundary, y, w: area.w * c.remainderFraction, h: bandH }, { fill: "border", layer: "decoration", name: "remainder-quantity-band" });
    // Non-causal reference lines from each real segment to its explicitly named quality question.
    y += bandH + .22;
    const centers = [area.x + area.w * c.selectedFraction / 2, boundary + area.w * c.remainderFraction / 2];
    centers.forEach((x, i) => {
      ctx.shape("line", { x, y: bandY + bandH, w: 0, h: .10 }, { line: { color: "mutedText", width: .7 }, layer: "decoration", name: `reference-stem-${i}` });
      const target = xs[i]! + colW / 2;
      ctx.shape("line", { x: Math.min(x, target), y: bandY + bandH + .10, w: Math.abs(target-x), h: 0 }, { line: {color:"mutedText",width:.7}, layer:"decoration",name:`reference-link-${i}` });
      ctx.shape("line", { x: target, y: bandY + bandH + .10, w: 0, h: .10 }, { line: {color:"mutedText",width:.7},layer:"decoration",name:`reference-end-${i}` });
    });
    const questions = [
      `「${d.labels.selected}」の内訳\n${d.selected === 0 ? "対象がないため内訳なし" : `${d.quality.selected.question}：未確認`}`,
      `「${d.labels.remainder}」の内訳\n${c.remainder === 0 ? "対象がないため内訳なし" : `${d.quality.remainder.question}：未確認`}`,
    ];
    const qh = Math.max(...questions.map(t => measure(t,colW,"body")));
    questions.forEach((t,i) => ctx.text({x:xs[i]!,y,w:colW,h:qh},t,{role:"body",lineHeight:1.22,name:`unknown-quality-${i}`}));
    y += qh + .04;
    text(`${d.quality.criterion}。${a.annotations.quality.text}`, "caption", "quality-annotation");
    text(d.quality.unknownScope, "caption", "unknown-scope");
    text(a.annotations.inference.text, "body", "inference-definition");
    text(a.inference, "body", "argument-inference", true);
    if (y > bottom + .001) throw new Error(`[partition-with-unknowns] content exceeds available height by ${(y-bottom).toFixed(2)} in; shorten text or use a larger template. No content was omitted.`);
    y = bottom + gap;
    for (const s of sourceLines) {
      const h = measure(s.text, area.w, "caption");
      ctx.text({x:area.x,y,w:area.w,h},s.text,{role:"caption",color:"mutedText",lineHeight:1.22,hyperlink:s.url,name:"source-link"});
      y += h + .03;
    }
  },
};
