import type { SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { splitBottom, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { buildSourceText, evidenceDetail, roleLabel, truncate } from "../shared/semantics.js";

export interface MessageOnlyBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  /** 補足として並べる短い一文（最大 2 件） */
  points?: string[];
  sourceText?: string;
}

const STRONG_ROLES = new Set(["recommendation", "decision", "summary", "insight", "context"]);

/**
 * MessageOnly — KeyMessage を最大の主役として見せる。
 * 全 logicalRelation に対応する fallback Pattern でもある。
 */
export const messageOnlyPattern: VisualPattern<MessageOnlyBindings> = {
  id: "message-only",
  version: "1.0.0",
  description: "KeyMessage を大きく主役として見せる。適合する Pattern がない場合の受け皿。",
  variants: ["standard"],
  fallback: true,

  compatibility: {
    logicalRelations: [
      "claim_evidence",
      "comparison",
      "before_after",
      "cause_effect",
      "process",
      "timeline",
      "hierarchy",
      "composition",
      "correlation",
      "problem_solution",
      "pros_cons",
    ],
  },

  constraints: {
    maxKeyMessageLength: 90,
  },

  match(slide: SlideSpec): MatchResult {
    let score = 0.3;
    const reasons: string[] = [];

    if (slide.evidence.length === 0) {
      score += 0.28;
      reasons.push("図示すべき evidence が存在しない");
    } else if (slide.evidence.length === 1) {
      score += 0.16;
      reasons.push("evidence が 1 件のみで構造化の余地が小さい");
    }

    if (STRONG_ROLES.has(slide.slideRole)) {
      score += 0.12;
      reasons.push(`slideRole=${slide.slideRole} はメッセージ自体を主役にできる`);
    }

    if (slide.keyMessage.length <= 40) {
      score += 0.06;
      reasons.push("keyMessage が短く大きく見せられる");
    }

    if (reasons.length === 0) reasons.push("常に適用可能な既定表現");

    return {
      compatible: true,
      score: Math.min(1, score),
      reason: reasons.join(" / "),
    };
  },

  buildBindings(slide: SlideSpec): MessageOnlyBindings {
    const points = slide.evidence
      .slice(0, 2)
      .map((e) => truncate(evidenceDetail(e), 60))
      .filter((p) => p.length > 0);

    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      points: points.length > 0 ? points : undefined,
      sourceText: buildSourceText(slide),
    };
  },

  render(b: MessageOnlyBindings, ctx: RenderContext): void {
    const { theme, frame } = ctx;

    // 出典行を先に確保
    let body: Rect = frame;
    if (b.sourceText) {
      const h = Math.max(
        0.18,
        ctx.measureText(b.sourceText, { widthIn: frame.w, role: "caption" }),
      );
      const [footer, rest] = splitBottom(frame, h, theme.spacing.sm);
      ctx.text(footer, b.sourceText, {
        role: "caption",
        color: "mutedText",
        valign: "bottom",
        layer: "content",
        name: "source",
      });
      body = rest;
    }

    if (b.eyebrow) {
      const [eyebrowRow, rest] = splitTop(body, 0.24, theme.spacing.sm);
      ctx.shape("rect", { x: eyebrowRow.x, y: eyebrowRow.y + 0.05, w: 0.06, h: 0.14 }, {
        fill: "primary",
        layer: "decoration",
        name: "eyebrow-mark",
      });
      ctx.text(
        { x: eyebrowRow.x + 0.06 + theme.spacing.sm, y: eyebrowRow.y, w: eyebrowRow.w - 0.5, h: 0.24 },
        b.eyebrow,
        {
          role: "caption",
          color: "primary",
          bold: true,
          charSpacing: 1,
          valign: "middle",
          layer: "content",
          name: "eyebrow",
        },
      );
      body = rest;
    }

    const barWidth = 0.09;
    const textX = body.x + barWidth + theme.spacing.lg;
    // 主役メッセージは端まで伸ばさず、読みやすい幅で折り返す
    const textW = Math.min(body.w - barWidth - theme.spacing.lg, 7.7);

    const headlineSize = Math.round(theme.fontSizes.keyMessage * 1.45);
    const headlineH = ctx.measureText(b.headline, {
      widthIn: textW,
      fontSize: headlineSize,
      lineHeight: 1.26,
    });

    const sublineSize = theme.fontSizes.heading;
    const sublineH = b.subline
      ? ctx.measureText(b.subline, { widthIn: textW, fontSize: sublineSize, lineHeight: 1.4 })
      : 0;

    const pointSize = theme.fontSizes.body;
    const pointHeights = (b.points ?? []).map((p) =>
      ctx.measureText(p, { widthIn: textW - 0.24, fontSize: pointSize, lineHeight: 1.36 }),
    );

    const gapAfterHeadline = b.subline ? theme.spacing.lg : 0;
    const gapBeforePoints = pointHeights.length > 0 ? theme.spacing.lg : 0;
    const pointsTotal =
      pointHeights.reduce((a, h) => a + h, 0) +
      Math.max(0, pointHeights.length - 1) * theme.spacing.sm;

    const totalH = headlineH + gapAfterHeadline + sublineH + gapBeforePoints + pointsTotal;
    const startY = Math.max(body.y, body.y + (body.h - totalH) / 2);

    ctx.shape("rect", { x: body.x, y: startY, w: barWidth, h: Math.min(totalH, body.h) }, {
      fill: "primary",
      layer: "decoration",
      name: "accent-bar",
    });

    let cursor = startY;
    ctx.text({ x: textX, y: cursor, w: textW, h: headlineH }, b.headline, {
      role: "keyMessage",
      fontSize: headlineSize,
      color: "text",
      bold: true,
      lineHeight: 1.26,
      layer: "content",
      name: "key-message",
    });
    cursor += headlineH + gapAfterHeadline;

    if (b.subline) {
      ctx.text({ x: textX, y: cursor, w: textW, h: sublineH }, b.subline, {
        role: "heading",
        fontSize: sublineSize,
        color: "mutedText",
        bold: false,
        lineHeight: 1.4,
        layer: "content",
        name: "supporting-message",
      });
      cursor += sublineH;
    }

    cursor += gapBeforePoints;
    (b.points ?? []).forEach((point, i) => {
      const h = pointHeights[i]!;
      ctx.shape("ellipse", { x: textX + 0.02, y: cursor + h / 2 - 0.045, w: 0.09, h: 0.09 }, {
        fill: "secondary",
        layer: "decoration",
        name: `point-marker-${i + 1}`,
      });
      ctx.text({ x: textX + 0.24, y: cursor, w: textW - 0.24, h }, point, {
        role: "body",
        color: "text",
        lineHeight: 1.36,
        layer: "content",
        name: `point-${i + 1}`,
      });
      cursor += h + theme.spacing.sm;
    });
  },
};
