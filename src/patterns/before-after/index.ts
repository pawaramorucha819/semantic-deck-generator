import { z } from "zod";
import type { SlideSpec } from "../../domain/types.js";
import type { ColorToken, RenderContext } from "../../render/context.js";
import { columnsWeighted, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { centerBlockY, drawCard, drawFooter, drawHeader, fitFontSize } from "../shared/layout.js";
import {
  buildSourceText,
  extractBeforeAfter,
  formatValue,
  roleLabel,
} from "../shared/semantics.js";

export interface BeforeAfterBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;

  before: {
    label: string;
    value: string | number;
  };

  after: {
    label: string;
    value: string | number;
  };

  change?: {
    label?: string;
    value: string | number;
  };

  unit?: string;
  sourceText?: string;
}

const ValuePairSchema = z
  .object({
    label: z.string().min(1),
    value: z.union([z.string().min(1), z.number().finite()]),
  })
  .strict();

export const BeforeAfterBindingsSchema: z.ZodType<BeforeAfterBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    before: ValuePairSchema,
    after: ValuePairSchema,
    change: z
      .object({
        label: z.string().min(1).optional(),
        value: z.union([z.string().min(1), z.number().finite()]),
      })
      .strict()
      .optional(),
    unit: z.string().min(1).optional(),
    sourceText: z.string().min(1).optional(),
  })
  .strict();

/**
 * BeforeAfter — 2 つの状態の変化を左右対比で見せる。
 *
 * before_after が明示され、比較可能な 2 値を実際に取り出せる場合にのみ使う。
 * 項目数や配列順序だけから時間順序を作らない。
 */
export const beforeAfterPattern: VisualPattern<BeforeAfterBindings> = {
  supportsExplanatoryLayout: true,
  id: "before-after",
  version: "1.0.0",
  description: "導入前後・状態変化を 2 枚のカードと矢印で対比する。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: ["before_after"],
    evidenceTypes: ["caseStudy", "comparison", "metric"],
  },

  constraints: {
    minEvidence: 1,
    maxEvidence: 4,
    maxKeyMessageLength: 70,
  },

  bindingsSchema: BeforeAfterBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    const facts = extractBeforeAfter(slide);
    if (!facts) {
      return {
        compatible: false,
        score: 0,
        reason: "明示的な before_after 関係と比較可能な 2 値が必要",
      };
    }

    let score = 0.9;
    const reasons: string[] = ["logicalRelation が before_after"];

    if (facts.origin === "caseStudy") {
      score += 0.08;
      reasons.push("caseStudy に before/after が定義されている");
    } else if (facts.origin === "comparison") {
      score += 0.05;
      reasons.push("comparison evidence の 2 項目を対比できる");
    } else {
      reasons.push("2 件の metric を対比できる");
    }

    if (facts.change) {
      score += 0.04;
      reasons.push("変化量を提示できる");
    }

    return {
      compatible: true,
      score: Math.min(1, score),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): BeforeAfterBindings {
    const facts = extractBeforeAfter(slide);
    if (!facts) {
      throw new Error(
        `[before-after] buildBindings: before/after を抽出できません (slide=${slide.id})`,
      );
    }
    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      before: facts.before,
      after: facts.after,
      change: facts.change,
      unit: facts.unit,
      sourceText: buildSourceText(slide),
    };
  },

  render(b: BeforeAfterBindings, ctx: RenderContext): void {
    const { theme } = ctx;
    const afterHeader = drawHeader(ctx, {
      eyebrow: b.eyebrow,
      headline: b.headline,
      subline: b.subline,
    });
    const body = drawFooter(ctx, afterHeader, b.sourceText);

    // カードは本文領域いっぱいに引き伸ばさず、上限高さで中央に置く
    const cardH = Math.min(body.h, 2.5);
    const cardArea: Rect = { x: body.x, y: centerBlockY(body, cardH), w: body.w, h: cardH };

    const [beforeCol, arrowCol, afterCol] = columnsWeighted(
      cardArea,
      [1, 0.42, 1],
      theme.spacing.md,
    ) as [Rect, Rect, Rect];

    drawValueCard(ctx, beforeCol, {
      label: b.before.label,
      value: formatValue(b.before.value, b.unit),
      accent: "mutedText",
      valueColor: "text",
      fill: ctx.tint("mutedText", 0.88),
      name: "before",
    });

    drawValueCard(ctx, afterCol, {
      label: b.after.label,
      value: formatValue(b.after.value, b.unit),
      accent: "primary",
      valueColor: "primary",
      fill: "surface",
      name: "after",
    });

    // 中央: 矢印 + 変化量
    const arrowH = Math.min(0.7, arrowCol.h * 0.34);
    const arrowY = arrowCol.y + arrowCol.h / 2 - arrowH / 2;
    const arrowW = Math.min(arrowCol.w * 0.86, 0.92);
    ctx.shape(
      "rightArrow",
      { x: arrowCol.x + (arrowCol.w - arrowW) / 2, y: arrowY, w: arrowW, h: arrowH },
      { fill: ctx.tint("primary", 0.25), layer: "decoration", name: "transition-arrow" },
    );

    if (b.change) {
      const changeText = String(b.change.value);
      const changeSize = Math.round(theme.fontSizes.heading * 1.25);
      const h = ctx.measureText(changeText, {
        widthIn: arrowCol.w,
        fontSize: changeSize,
        lineHeight: 1.3,
      });
      const y = arrowY - h - theme.spacing.sm;
      ctx.text({ x: arrowCol.x, y, w: arrowCol.w, h }, changeText, {
        role: "heading",
        fontSize: changeSize,
        font: "numeric",
        color: "accent",
        bold: true,
        align: "center",
        valign: "bottom",
        layer: "content",
        name: "change-value",
      });

      if (b.change.label) {
        ctx.text(
          { x: arrowCol.x, y: arrowY + arrowH + theme.spacing.xs, w: arrowCol.w, h: 0.22 },
          b.change.label,
          {
            role: "caption",
            color: "mutedText",
            align: "center",
            valign: "top",
            layer: "content",
            name: "change-label",
          },
        );
      }
    }
  },
};

interface ValueCardOptions {
  label: string;
  value: string;
  accent: ColorToken;
  valueColor: ColorToken;
  fill: ColorToken;
  name: string;
}

function drawValueCard(ctx: RenderContext, box: Rect, opts: ValueCardOptions): void {
  const inner = drawCard(ctx, box, {
    fill: opts.fill,
    border: "border",
    accent: opts.accent,
    name: `${opts.name}-card`,
  });

  const [labelRow, valueRow] = splitTop(inner, 0.3, ctx.theme.spacing.sm);
  ctx.text(labelRow, opts.label, {
    role: "heading",
    color: "mutedText",
    bold: true,
    align: "center",
    valign: "top",
    layer: "content",
    name: `${opts.name}-label`,
  });

  const valueSize = Math.min(
    fitFontSize(opts.value, valueRow.w, ctx.theme.fontSizes.display),
    Math.floor(valueRow.h * 72 * 0.62),
  );
  ctx.text(valueRow, opts.value, {
    role: "display",
    fontSize: valueSize,
    font: "numeric",
    color: opts.valueColor,
    bold: true,
    align: "center",
    valign: "middle",
    shrink: true,
    layer: "content",
    name: `${opts.name}-value`,
  });
}
