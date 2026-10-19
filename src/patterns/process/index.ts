import { z } from "zod";
import { drawExplanatoryNodes } from "../shared/explanatory-nodes.js";
import type { SequenceContent, SlideSpec } from "../../domain/types.js";
import type { RenderContext, TextRun } from "../../render/context.js";
import { columns, inset, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { centerBlockY, drawFooter, drawHeader } from "../shared/layout.js";
import { buildSourceText, roleLabel } from "../shared/semantics.js";

export interface ProcessStepBinding {
  /** Traceability metadata; rendering does not reinterpret these fields. */
  nodeId?: string;
  evidenceRefs?: string[];
  members?: string[];
  index: number;
  label: string;
  description?: string;
}

export interface ProcessBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  steps: ProcessStepBinding[];
  /** ステップ番号の接頭辞（例: "STEP"） */
  stepPrefix: string;
  sourceText?: string;
}

export const ProcessBindingsSchema: z.ZodType<ProcessBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    steps: z
      .array(
        z
          .object({
            nodeId: z.string().min(1).max(120).optional(),
            evidenceRefs: z.array(z.string().min(1).max(120)).min(1).max(8).optional(),
            members: z.array(z.string().min(1).max(120)).max(1000).optional(),
            index: z.number().int().positive(),
            label: z.string().min(1),
            description: z.string().min(1).optional(),
          })
          .strict(),
      )
      .min(2)
      .max(5),
    stepPrefix: z.string(),
    sourceText: z.string().min(1).optional(),
  })
  .strict();

const MAX_STEPS = 5;

function sequenceOf(slide: SlideSpec): SequenceContent | undefined {
  return slide.content?.kind === "sequence" ? slide.content : undefined;
}

/**
 * Process — 工程・順序を左から右への矢印で表現する。
 *
 * 工程は `content.sequence`（見せる対象の情報構造）から読む。
 * evidence は「根拠」であり、工程の格納場所として使ってはいけない。
 */
export const processPattern: VisualPattern<ProcessBindings> = {
  supportsExplanatoryLayout: true,
  id: "process",
  version: "1.1.0",
  description: "content.sequence の工程を左から右へのシェブロンで表現する。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: ["process", "timeline"],
  },

  constraints: {
    maxKeyMessageLength: 70,
  },

  bindingsSchema: ProcessBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    const sequence = sequenceOf(slide);
    if (!sequence) {
      return {
        compatible: false,
        score: 0,
        reason: "工程を表す content (kind=sequence) が定義されていない",
      };
    }

    const count = sequence.steps.length;
    if (count > MAX_STEPS) {
      return {
        compatible: false,
        score: 0,
        reason: `${count} 工程は対応数の範囲外（最大 ${MAX_STEPS} 工程）`,
      };
    }

    const reasons: string[] = [];
    let score = 0.6;

    if (slide.logicalRelation === "process") {
      score += 0.3;
      reasons.push("logicalRelation が process で順序構造を持つ");
    } else {
      score += 0.18;
      reasons.push("logicalRelation が timeline で時系列を順序として表現できる");
    }

    if (count >= 3 && count <= 4) {
      score += 0.08;
      reasons.push(slide.explanation ? `${count} 工程の順序を説明文の横の縦配置で保持する（収まりは描画時に検査）` : `${count} 工程で横一列に収まる`);
    }

    const longest = Math.max(...sequence.steps.map((s) => s.label.length));
    if (longest > 16) {
      score -= 0.08;
      reasons.push(slide.explanation ? "工程名が長くカード内の収まりを描画時に検査する" : "工程名が長くシェブロン内が窮屈になる可能性");
    }

    return {
      compatible: true,
      score: Math.max(0, Math.min(1, score)),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): ProcessBindings {
    const sequence = sequenceOf(slide);
    if (!sequence) {
      throw new Error(
        `[process] buildBindings: content.sequence がありません (slide=${slide.id})`,
      );
    }
    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      stepPrefix: slide.logicalRelation === "timeline" ? "" : "STEP",
      steps: sequence.steps.map((step, i) => ({
        ...(step.id !== undefined ? { nodeId: step.id } : {}),
        ...(step.evidenceRefs !== undefined ? { evidenceRefs: [...step.evidenceRefs] } : {}),
        ...(step.members !== undefined ? { members: [...step.members] } : {}),
        index: i + 1,
        label: step.label,
        description: step.description,
      })),
      sourceText: buildSourceText(slide),
    };
  },

  render(b: ProcessBindings, ctx: RenderContext): void {
    const { theme } = ctx;
    const afterHeader = drawHeader(ctx, {
      eyebrow: b.eyebrow,
      headline: b.headline,
      subline: b.subline,
    });
    const body = drawFooter(ctx, afterHeader, b.sourceText);
    if (ctx.layoutBody) {
      drawExplanatoryNodes(ctx, body, b.steps.map(item => ({ ...item, label: item.label.startsWith("[") ? item.label : `${b.stepPrefix ? b.stepPrefix + " " : ""}${item.index}　${item.label}`, name: `step-${item.index}` })), true);
      return;
    }

    const chevronHeight = Math.min(0.78, body.h * 0.32);
    const cellW = (body.w - theme.spacing.md * (b.steps.length - 1)) / b.steps.length;
    const descriptionHeight = Math.max(
      0,
      ...b.steps.map((step) =>
        step.description
          ? ctx.measureText(step.description, { widthIn: cellW, role: "body", lineHeight: 1.4 })
          : 0,
      ),
    );

    // シェブロン + 説明文をひとかたまりとして本文領域の中央に置く
    const blockH = Math.min(
      body.h,
      chevronHeight + (descriptionHeight > 0 ? theme.spacing.md + descriptionHeight : 0),
    );
    const block: Rect = { x: body.x, y: centerBlockY(body, blockH), w: body.w, h: blockH };

    const [chevronRow, descriptionRow] = splitTop(block, chevronHeight, theme.spacing.md);
    const chevronCells = columns(chevronRow, b.steps.length, theme.spacing.xs);
    const descriptionCells = columns(descriptionRow, b.steps.length, theme.spacing.md);

    b.steps.forEach((step, i) => {
      const cell = chevronCells[i] as Rect;
      ctx.shape("chevron", cell, {
        fill: "primary",
        layer: "decoration",
        name: `step-${step.index}-chevron`,
      });

      // シェブロンは左右が斜めなので、テキストは高さの 40% 分内側へ寄せる
      const textBox = inset(cell, {
        left: cell.h * 0.42,
        right: cell.h * 0.42,
        top: theme.spacing.xs,
        bottom: theme.spacing.xs,
      });

      const runs: TextRun[] = [];
      const stepLabel = b.stepPrefix ? `${b.stepPrefix} ${step.index}` : String(step.index);
      runs.push({
        text: stepLabel,
        style: { fontSize: theme.fontSizes.caption, bold: false, font: "numeric" },
        breakLine: true,
      });
      runs.push({
        text: step.label,
        style: { fontSize: theme.fontSizes.heading, bold: true, font: "heading" },
      });

      ctx.text(textBox, runs, {
        color: "onPrimary",
        align: "center",
        valign: "middle",
        lineHeight: 1.24,
        shrink: true,
        layer: "content",
        name: `step-${step.index}-label`,
      });

      const description = step.description;
      if (description) {
        const descCell = descriptionCells[i] as Rect;
        ctx.text(descCell, description, {
          role: "body",
          color: "mutedText",
          lineHeight: 1.4,
          valign: "top",
          shrink: true,
          layer: "content",
          name: `step-${step.index}-description`,
        });
      }
    });
  },
};
