import type { SlideSpec } from "../../domain/types.js";
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
  description: string;
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

/**
 * ThreeReasons — KeyMessage を複数（2〜4）の根拠で支える。
 * 「主張 → 根拠が並列に複数」という論理構造をそのまま横並びのカードにする。
 */
export const threeReasonsPattern: VisualPattern<ThreeReasonsBindings> = {
  id: "three-reasons",
  version: "1.0.0",
  description: "KeyMessage を 2〜4 件の並列な根拠カードで支える。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: ["claim_evidence", "problem_solution", "cause_effect", "composition", "pros_cons"],
  },

  constraints: {
    minEvidence: 2,
    maxEvidence: 4,
    maxKeyMessageLength: 70,
  },

  match(slide: SlideSpec): MatchResult {
    const count = slide.evidence.length;
    const reasons: string[] = [];
    let score = 0.55;

    if (count === 3) {
      score += 0.3;
      reasons.push("evidence が 3 件で横並び 3 分割に最適");
    } else if (count === 2 || count === 4) {
      score += 0.22;
      reasons.push(`evidence が ${count} 件で並列表示できる`);
    }

    if (slide.logicalRelation === "claim_evidence") {
      score += 0.1;
      reasons.push("logicalRelation が claim_evidence で主張と根拠の並列構造に一致");
    } else {
      reasons.push(`logicalRelation=${slide.logicalRelation} を並列根拠として扱う`);
    }

    // 各 evidence が短くまとまっているほどカード表示に向く
    const longest = Math.max(...slide.evidence.map((e) => evidenceDetail(e).length));
    if (longest <= 60) {
      score += 0.05;
      reasons.push("各 evidence が短くカードに収まる");
    } else if (longest > 120) {
      score -= 0.12;
      reasons.push("evidence の文章が長くカード表示にはやや不利");
    }

    return {
      compatible: true,
      score: Math.max(0, Math.min(1, score)),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): ThreeReasonsBindings {
    const reasons: ReasonBinding[] = slide.evidence.map((e, i) => ({
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

    // カード高さは中身から決める（本文領域いっぱいに間延びさせない）
    const neededHeights = b.reasons.map((reason) => {
      const labelH = Math.max(
        0.3,
        ctx.measureText(reason.label, { widthIn: innerW, role: "heading", lineHeight: 1.28 }),
      );
      const metricH = reason.metric ? 0.36 + theme.spacing.xs : 0;
      const descH = ctx.measureText(reason.description, {
        widthIn: innerW,
        role: "body",
        lineHeight: 1.38,
      });
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
        const [metricRow, afterMetric] = splitTop(rest, 0.36, theme.spacing.xs);
        ctx.text(metricRow, reason.metric, {
          role: "heading",
          fontSize: Math.round(theme.fontSizes.heading * 1.3),
          font: "numeric",
          color: "primary",
          bold: true,
          valign: "middle",
          layer: "content",
          name: `reason-${reason.index}-metric`,
        });
        rest = afterMetric;
      }

      ctx.text(rest, reason.description, {
        role: "body",
        color: "mutedText",
        lineHeight: 1.38,
        shrink: true,
        layer: "content",
        name: `reason-${reason.index}-description`,
      });
    });
  },
};
