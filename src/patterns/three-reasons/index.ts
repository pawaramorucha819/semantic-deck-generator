import { z } from "zod";
import { isMetric, type ContentItem, type Evidence, type SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { columns, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { centerBlockY, drawCard, drawFooter, drawHeader } from "../shared/layout.js";
import {
  buildSourceText,
  evidenceDetail,
  evidenceLabel,
  formatValue,
  roleLabel,
} from "../shared/semantics.js";

export interface ReasonBinding {
  index: number;
  label: string;
  description?: string;
  /** 定量 evidence の場合に表示する数値 */
  metric?: string;
}

export interface ThreeReasonsBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  reasons: ReasonBinding[];
  sourceText?: string;
}

export const ThreeReasonsBindingsSchema: z.ZodType<ThreeReasonsBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    reasons: z
      .array(
        z
          .object({
            index: z.number().int().positive(),
            label: z.string().min(1),
            description: z.string().min(1).optional(),
            metric: z.string().min(1).optional(),
          })
          .strict(),
      )
      .min(2)
      .max(4),
    sourceText: z.string().min(1).optional(),
  })
  .strict();

type ReasonSource =
  | { origin: "points"; items: ContentItem[] }
  | { origin: "evidence"; items: Evidence[] };

/** 並列カードの素材。content.points があればそれを優先し、無ければ evidence を使う。 */
function reasonSource(slide: SlideSpec): ReasonSource | undefined {
  if (slide.content?.kind === "points") {
    const items = slide.content.items;
    if (items.length >= 2 && items.length <= 4) return { origin: "points", items };
    return undefined;
  }
  if (slide.evidence.length >= 2 && slide.evidence.length <= 4) {
    return { origin: "evidence", items: slide.evidence };
  }
  return undefined;
}

/**
 * ThreeReasons — KeyMessage を複数（2〜4）の並列項目で支える。
 *
 * 素材は 2 系統ある:
 *  - content.points … 論点・要因など「見せる対象」の並列構造
 *  - evidence       … 主張を裏付ける複数の根拠（claim_evidence の典型形）
 */
export const threeReasonsPattern: VisualPattern<ThreeReasonsBindings> = {
  id: "three-reasons",
  version: "1.1.0",
  description: "KeyMessage を 2〜4 件の並列カード（content.points または evidence）で支える。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: ["claim_evidence", "problem_solution", "cause_effect", "composition", "pros_cons"],
  },

  constraints: {
    maxKeyMessageLength: 70,
  },

  bindingsSchema: ThreeReasonsBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    const source = reasonSource(slide);
    if (!source) {
      return {
        compatible: false,
        score: 0,
        reason: "並列表示できる 2〜4 件の content.points または evidence が存在しない",
      };
    }

    const count = source.items.length;
    const reasons: string[] = [];
    let score = 0.55;

    if (count === 3) {
      score += 0.3;
      reasons.push(`${source.origin} が 3 件で横並び 3 分割に最適`);
    } else if (count === 2 || count === 4) {
      score += 0.22;
      reasons.push(`${source.origin} が ${count} 件で並列表示できる`);
    }

    if (slide.logicalRelation === "claim_evidence") {
      score += 0.1;
      reasons.push("logicalRelation が claim_evidence で主張と根拠の並列構造に一致");
    } else {
      reasons.push(`logicalRelation=${slide.logicalRelation} を並列構造として扱う`);
    }

    // 各項目が短くまとまっているほどカード表示に向く
    const longest = Math.max(
      ...source.items.map((item) =>
        source.origin === "points"
          ? ((item as ContentItem).description ?? (item as ContentItem).label).length
          : evidenceDetail(item as Evidence).length,
      ),
    );
    if (longest <= 60) {
      score += 0.05;
      reasons.push("各項目が短くカードに収まる");
    } else if (longest > 120) {
      score -= 0.12;
      reasons.push("文章が長くカード表示にはやや不利");
    }

    return {
      compatible: true,
      score: Math.max(0, Math.min(1, score)),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): ThreeReasonsBindings {
    const source = reasonSource(slide);
    if (!source) {
      throw new Error(
        `[three-reasons] buildBindings: 並列表示できる素材がありません (slide=${slide.id})`,
      );
    }

    const evidenceById = new Map(slide.evidence.map((e) => [e.id, e]));

    const reasons: ReasonBinding[] =
      source.origin === "points"
        ? source.items.map((item, i) => {
            // evidenceRefs で参照された metric があれば数値として見せる
            const metric = (item.evidenceRefs ?? [])
              .map((id) => evidenceById.get(id))
              .filter((e): e is Evidence => e !== undefined)
              .filter(isMetric)
              .map((e) => formatValue(e.data.value, e.data.unit))[0];
            return {
              index: i + 1,
              label: item.label,
              description: item.description,
              metric,
            };
          })
        : source.items.map((e, i) => ({
            index: i + 1,
            label: evidenceLabel(e, 20),
            description: evidenceDetail(e),
            metric:
              e.type === "metric"
                ? formatValue(e.data.value, e.data.unit)
                : e.type === "caseStudy" && e.data.changeLabel
                  ? e.data.changeLabel
                  : undefined,
          }));

    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      reasons,
      sourceText: buildSourceText(slide),
    };
  },

  render(b: ThreeReasonsBindings, ctx: RenderContext): void {
    const { theme } = ctx;
    const afterHeader = drawHeader(ctx, {
      eyebrow: b.eyebrow,
      headline: b.headline,
      subline: b.subline,
    });
    const body = drawFooter(ctx, afterHeader, b.sourceText);

    const cols = columns(body, b.reasons.length, theme.spacing.md);
    const pad = theme.spacing.md;
    const badge = 0.42;
    const innerW = (cols[0]?.w ?? body.w) - pad * 2;

    // 数値は heading より一段大きく見せるので、行の高さも実サイズから求める
    const metricSize = Math.round(theme.fontSizes.heading * 1.3);
    const metricHeightOf = (metric: string, widthIn: number): number =>
      ctx.measureText(metric, { widthIn, fontSize: metricSize });

    // カード高さは中身から決める（本文領域いっぱいに間延びさせない）
    const neededHeights = b.reasons.map((reason) => {
      const labelH = Math.max(
        0.3,
        ctx.measureText(reason.label, { widthIn: innerW, role: "heading", lineHeight: 1.28 }),
      );
      const metricH = reason.metric ? metricHeightOf(reason.metric, innerW) + theme.spacing.xs : 0;
      const descH = reason.description
        ? ctx.measureText(reason.description, {
            widthIn: innerW,
            role: "body",
            lineHeight: 1.38,
          })
        : 0;
      return badge + theme.spacing.sm + labelH + theme.spacing.sm + metricH + descH + pad * 2;
    });
    const cardH = Math.min(body.h, Math.max(...neededHeights) + theme.spacing.sm);
    const cardY = centerBlockY(body, cardH);

    b.reasons.forEach((reason, i) => {
      const col = cols[i] as Rect;
      const inner = drawCard(ctx, { x: col.x, y: cardY, w: col.w, h: cardH }, {
        fill: "surface",
        border: "border",
        accent: i === 0 ? "primary" : ctx.tint("primary", i * 0.28),
        name: `reason-${reason.index}`,
      });
      ctx.shape("ellipse", { x: inner.x, y: inner.y, w: badge, h: badge }, {
        fill: "primary",
        layer: "decoration",
        name: `reason-${reason.index}-badge`,
      });
      ctx.text(
        { x: inner.x, y: inner.y, w: badge, h: badge },
        String(reason.index),
        {
          role: "body",
          font: "numeric",
          color: "onPrimary",
          bold: true,
          align: "center",
          valign: "middle",
          layer: "content",
          name: `reason-${reason.index}-number`,
        },
      );

      let rest: Rect = {
        x: inner.x,
        y: inner.y + badge + theme.spacing.sm,
        w: inner.w,
        h: Math.max(0, inner.h - badge - theme.spacing.sm),
      };

      const labelH = Math.max(
        0.3,
        ctx.measureText(reason.label, { widthIn: rest.w, role: "heading", lineHeight: 1.28 }),
      );
      const [labelRow, afterLabel] = splitTop(rest, labelH, theme.spacing.sm);
      ctx.text(labelRow, reason.label, {
        role: "heading",
        color: "text",
        bold: true,
        lineHeight: 1.28,
        layer: "content",
        name: `reason-${reason.index}-label`,
      });
      rest = afterLabel;

      if (reason.metric) {
        const [metricRow, afterMetric] = splitTop(
          rest,
          metricHeightOf(reason.metric, rest.w),
          theme.spacing.xs,
        );
        ctx.text(metricRow, reason.metric, {
          role: "heading",
          fontSize: metricSize,
          font: "numeric",
          color: "primary",
          bold: true,
          valign: "middle",
          layer: "content",
          name: `reason-${reason.index}-metric`,
        });
        rest = afterMetric;
      }

      if (reason.description) {
        ctx.text(rest, reason.description, {
          role: "body",
          color: "mutedText",
          lineHeight: 1.38,
          shrink: true,
          layer: "content",
          name: `reason-${reason.index}-description`,
        });
      }
    });
  },
};
