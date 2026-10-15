import { z } from "zod";
import type { SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { columnsWeighted, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { drawFooter, drawHeader } from "../shared/layout.js";
import {
  buildSourceText,
  extractChartSeries,
  formatChangeRate,
  formatValue,
  roleLabel,
} from "../shared/semantics.js";

export interface ChartWithMessageBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  chart: {
    kind: "bar";
    direction: "col" | "bar";
    categories: string[];
    series: Array<{ name: string; values: number[] }>;
    unit?: string;
  };
  /** グラフから読み取ってほしい一文 */
  readingText?: string;
  /** 明示された時系列の端から端への変化（あれば） */
  changeText?: string;
  sourceText?: string;
}

export const ChartWithMessageBindingsSchema: z.ZodType<ChartWithMessageBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    chart: z
      .object({
        kind: z.literal("bar"),
        direction: z.enum(["col", "bar"]),
        categories: z.array(z.string().min(1)).min(2).max(12),
        series: z
          .array(
            z
              .object({
                name: z.string().min(1),
                values: z.array(z.number().finite()),
              })
              .strict(),
          )
          .min(1)
          .max(4),
        unit: z.string().min(1).optional(),
      })
      .strict()
      .superRefine((chart, ctx) => {
        chart.series.forEach((series, i) => {
          if (series.values.length !== chart.categories.length) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              path: ["series", i, "values"],
              message: `values の件数 (${series.values.length}) が categories の件数 (${chart.categories.length}) と一致しません`,
            });
          }
        });
      }),
    readingText: z.string().min(1).optional(),
    changeText: z.string().min(1).optional(),
    sourceText: z.string().min(1).optional(),
  })
  .strict();

/**
 * ChartWithMessage — 同じ軸に並んだ複数の値を、PowerPoint ネイティブの
 * （＝編集可能な）グラフとして描き、右側に読み取り方を添える。
 */
export const chartWithMessagePattern: VisualPattern<ChartWithMessageBindings> = {
  id: "chart-with-message",
  version: "1.0.0",
  description: "編集可能なネイティブグラフ + 読み取り文。同一軸に並ぶ 2 値以上が必要。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: [
      "claim_evidence",
      "comparison",
      "timeline",
    ],
    evidenceTypes: ["comparison", "metric"],
  },

  constraints: {
    minEvidence: 1,
    maxKeyMessageLength: 70,
  },

  bindingsSchema: ChartWithMessageBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    const facts = extractChartSeries(slide);
    if (!facts) {
      return {
        compatible: false,
        score: 0,
        reason: "同一軸に並ぶ数値系列を evidence から抽出できない",
      };
    }

    const reasons: string[] = [];
    let score = 0.72;

    if (facts.categories.length >= 3) {
      score += 0.16;
      reasons.push(`${facts.categories.length} 項目の${slide.logicalRelation === "timeline" ? "推移" : "比較"}を読み取れる`);
    } else {
      score += 0.06;
      reasons.push("2 項目の比較として読み取れる");
    }

    if (facts.origin === "comparison") {
      score += 0.06;
      reasons.push("comparison evidence がそのまま系列になる");
    } else {
      reasons.push("単位の揃った metric を系列として扱う");
    }

    if (slide.logicalRelation === "comparison" || slide.logicalRelation === "timeline") {
      score += 0.06;
      reasons.push(`logicalRelation=${slide.logicalRelation} がグラフ表現に一致`);
    }

    return {
      compatible: true,
      score: Math.min(1, score),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): ChartWithMessageBindings {
    const facts = extractChartSeries(slide);
    if (!facts) {
      throw new Error(
        `[chart-with-message] buildBindings: 数値系列を抽出できません (slide=${slide.id})`,
      );
    }

    const first = facts.values[0];
    const last = facts.values[facts.values.length - 1];
    const change =
      slide.logicalRelation === "timeline" &&
      facts.values.length >= 2 && first !== undefined && last !== undefined
        ? formatChangeRate(first, last)
        : undefined;

    const readingSource = slide.evidence.find((e) => facts.sourceEvidenceIds.includes(e.id));

    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      chart: {
        kind: "bar",
        direction: "col",
        categories: facts.categories,
        series: [{ name: facts.seriesName, values: facts.values }],
        unit: facts.unit,
      },
      readingText: readingSource?.statement,
      changeText: change
        ? `${formatValue(first!, facts.unit)} → ${formatValue(last!, facts.unit)} (${change})`
        : undefined,
      sourceText: buildSourceText(slide),
    };
  },

  render(b: ChartWithMessageBindings, ctx: RenderContext): void {
    const { theme } = ctx;
    const afterHeader = drawHeader(ctx, {
      eyebrow: b.eyebrow,
      headline: b.headline,
      subline: b.subline,
    });
    const body = drawFooter(ctx, afterHeader, b.sourceText);

    const hasSideText = Boolean(b.readingText || b.changeText);
    const [chartCol, sideCol] = hasSideText
      ? (columnsWeighted(body, [1.7, 1], theme.spacing.lg) as [Rect, Rect])
      : ([body, undefined] as [Rect, Rect | undefined]);

    ctx.chart(chartCol, {
      kind: "bar",
      barDir: b.chart.direction,
      categories: b.chart.categories,
      series: b.chart.series,
      colors: ["primary"],
      showValue: true,
      showLegend: false,
      valueUnit: b.chart.unit,
      layer: "content",
      name: "chart",
    });

    if (!sideCol) return;

    let cursor = sideCol.y;

    if (b.changeText) {
      const h = 0.5;
      ctx.shape("roundRect", { x: sideCol.x, y: cursor, w: sideCol.w, h }, {
        fill: ctx.tint("primary", 0.88),
        layer: "decoration",
        name: "change-badge",
      });
      ctx.text({ x: sideCol.x, y: cursor, w: sideCol.w, h }, b.changeText, {
        role: "heading",
        font: "numeric",
        color: "primary",
        bold: true,
        align: "center",
        valign: "middle",
        shrink: true,
        layer: "content",
        name: "change-text",
      });
      cursor += h + theme.spacing.md;
    }

    if (b.readingText) {
      ctx.text(
        { x: sideCol.x, y: cursor, w: sideCol.w, h: Math.max(0, sideCol.y + sideCol.h - cursor) },
        b.readingText,
        {
          role: "body",
          color: "text",
          lineHeight: 1.5,
          valign: "top",
          shrink: true,
          layer: "content",
          name: "reading-text",
        },
      );
    }
  },
};
