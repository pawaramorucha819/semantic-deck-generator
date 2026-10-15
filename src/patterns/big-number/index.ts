import { z } from "zod";
import { isRelationshipContent } from "../../domain/content.js";
import type { SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { columnsWeighted, splitBottom, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { centerBlockY, drawCard, drawFooter, drawHeader, fitFontSize } from "../shared/layout.js";
import {
  buildSourceText,
  extractHeadlineMetric,
  formatValue,
  roleLabel,
} from "../shared/semantics.js";

export interface BigNumberBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  /** 単位まで含めて整形済みの値 */
  value: string;
  /** 数値が何を表すか */
  caption?: string;
  /** 期間などの補足 */
  period?: string;
  /** 数値の読み取り方 */
  note?: string;
  sourceText?: string;
}

export const BigNumberBindingsSchema: z.ZodType<BigNumberBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    value: z.string().min(1),
    caption: z.string().min(1).optional(),
    period: z.string().min(1).optional(),
    note: z.string().min(1).optional(),
    sourceText: z.string().min(1).optional(),
  })
  .strict();

/**
 * BigNumber — 1 つの数字だけでスライドが成立するケース。
 * 「metric が 1 件しかない」＝ 図解する構造が無いが数字は強い、という意味構造に対応する。
 */
export const bigNumberPattern: VisualPattern<BigNumberBindings> = {
  id: "big-number",
  version: "1.0.0",
  description: "1 つの重要な数値を巨大に見せ、右側に読み取り方を添える。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: [
      "claim_evidence",
      "cause_effect",
      "correlation",
      "composition",
      "problem_solution",
    ],
    evidenceTypes: ["metric"],
  },

  constraints: {
    minEvidence: 1,
    maxEvidence: 2,
    maxKeyMessageLength: 70,
    requiresAnyEvidenceType: ["metric"],
  },

  bindingsSchema: BigNumberBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    if (isRelationshipContent(slide.content)) {
      return { compatible: false, score: 0, reason: `content.${slide.content.kind} の構造は単一の数値へ置き換えられない` };
    }
    const metric = extractHeadlineMetric(slide);
    if (!metric) {
      return {
        compatible: false,
        score: 0,
        reason: "主役にできる単一の metric evidence が存在しない",
      };
    }

    const reasons = ["単一の metric evidence を主役の数値として提示できる"];
    let score = 0.82;

    if (slide.evidence.length === 1) {
      score += 0.1;
      reasons.push("他に図示すべき evidence が無い");
    }
    if (metric.data.label) {
      score += 0.04;
      reasons.push("数値のラベルが定義されている");
    }

    return {
      compatible: true,
      score: Math.min(1, score),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): BigNumberBindings {
    if (isRelationshipContent(slide.content)) {
      throw new Error(`[big-number] buildBindings: content.${slide.content.kind} の関係構造を保持できません (slide=${slide.id})`);
    }
    const metric = extractHeadlineMetric(slide);
    if (!metric) {
      throw new Error(`[big-number] buildBindings: metric を抽出できません (slide=${slide.id})`);
    }
    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      value: formatValue(metric.data.value, metric.data.unit),
      caption: metric.data.label,
      period: metric.data.period,
      note: metric.statement,
      sourceText: buildSourceText(slide),
    };
  },

  render(b: BigNumberBindings, ctx: RenderContext): void {
    const { theme } = ctx;
    const afterHeader = drawHeader(ctx, {
      eyebrow: b.eyebrow,
      headline: b.headline,
      subline: b.subline,
    });
    const body = drawFooter(ctx, afterHeader, b.sourceText);

    const blockH = Math.min(body.h, 2.3);
    const block: Rect = { x: body.x, y: centerBlockY(body, blockH), w: body.w, h: blockH };

    const hasNote = Boolean(b.note);
    const [numberCol, noteCol] = hasNote
      ? (columnsWeighted(block, [1, 1], theme.spacing.lg) as [Rect, Rect])
      : ([block, undefined] as [Rect, Rect | undefined]);

    const inner = drawCard(ctx, numberCol, {
      fill: "surface",
      border: "border",
      accent: "primary",
      name: "big-number-card",
    });

    let valueArea = inner;
    if (b.caption) {
      const [captionRow, rest] = splitTop(inner, 0.3, theme.spacing.xs);
      ctx.text(captionRow, b.caption, {
        role: "heading",
        color: "mutedText",
        bold: true,
        align: "center",
        valign: "top",
        layer: "content",
        name: "big-number-caption",
      });
      valueArea = rest;
    }

    if (b.period) {
      const [periodRow, rest] = splitBottom(valueArea, 0.24, theme.spacing.xs);
      ctx.text(periodRow, b.period, {
        role: "caption",
        color: "mutedText",
        align: "center",
        valign: "bottom",
        shrink: true,
        layer: "content",
        name: "big-number-period",
      });
      valueArea = rest;
    }

    const valueSize = Math.min(
      fitFontSize(b.value, valueArea.w, Math.round(theme.fontSizes.display * 1.15)),
      Math.floor(valueArea.h * 72 * 0.7),
    );
    ctx.text(valueArea, b.value, {
      role: "display",
      fontSize: valueSize,
      font: "numeric",
      color: "primary",
      bold: true,
      align: "center",
      valign: "middle",
      shrink: true,
      layer: "content",
      name: "big-number-value",
    });

    if (b.note && noteCol) {
      ctx.shape("rect", { x: noteCol.x, y: noteCol.y + 0.06, w: 0.05, h: noteCol.h - 0.12 }, {
        fill: ctx.tint("primary", 0.55),
        layer: "decoration",
        name: "note-rule",
      });
      const textBox: Rect = {
        x: noteCol.x + 0.05 + theme.spacing.md,
        y: noteCol.y,
        w: noteCol.w - 0.05 - theme.spacing.md,
        h: noteCol.h,
      };
      ctx.text(textBox, b.note, {
        role: "heading",
        color: "text",
        lineHeight: 1.5,
        valign: "middle",
        shrink: true,
        layer: "content",
        name: "big-number-note",
      });
    }
  },
};
