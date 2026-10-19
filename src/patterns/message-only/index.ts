import { drawExplanatoryNodes } from "../shared/explanatory-nodes.js";
import { z } from "zod";
import { contentItemEntries, isRelationshipContent } from "../../domain/content.js";
import { LOGICAL_RELATIONS, type ContentItem, type SemanticContent, type SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import { splitBottom, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { drawHeader, drawFooter } from "../shared/layout.js";
import { buildSourceText, evidenceDetail, roleLabel, truncate } from "../shared/semantics.js";

export interface StructureNodeBinding {
  /** Original path relative to content; retained only for traceability. */
  path: Array<string | number>;
  nodeId?: string;
  evidenceRefs?: string[];
  members?: string[];
}

export interface MessageOnlyBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  /** 補足として並べる短い一文（最大 3 件） */
  points?: string[];
  sourceText?: string;
  /** Explicit text representation of semantic content; contains all nodes without truncation. */
  structureText?: string;
  structureNodes?: StructureNodeBinding[];
  explanatoryItems?: Array<{ label: string; description?: string }>;
}

export const MessageOnlyBindingsSchema: z.ZodType<MessageOnlyBindings> = z
  .object({
    eyebrow: z.string().min(1).optional(),
    headline: z.string().min(1),
    subline: z.string().min(1).optional(),
    points: z.array(z.string().min(1)).min(1).max(3).optional(),
    sourceText: z.string().min(1).optional(),
    structureText: z.string().min(1).optional(),
    explanatoryItems: z.array(z.object({ label: z.string().min(1), description: z.string().min(1).optional() }).strict()).min(1).max(6).optional(),
    structureNodes: z.array(z.object({
      path: z.array(z.union([z.string().min(1), z.number().int().nonnegative()])).min(1).max(2),
      nodeId: z.string().min(1).max(120).optional(),
      evidenceRefs: z.array(z.string().min(1).max(120)).min(1).max(8).optional(),
      members: z.array(z.string().min(1).max(120)).max(1000).optional(),
    }).strict()).min(1).max(6).optional(),
  })
  .strict();

const STRONG_ROLES = new Set(["recommendation", "decision", "summary", "insight", "context"]);

/**
 * MessageOnly — KeyMessage を最大の主役として見せる。
 *
 * 全 logicalRelation に対応する fallback Pattern。
 * fallback 契約により constraints を持たず、どんなに長い keyMessage でも
 * render() 側がフォントサイズを段階的に下げて必ず収める。
 */
export const messageOnlyPattern: VisualPattern<MessageOnlyBindings> = {
  supportsExplanatoryLayout: true,
  id: "message-only",
  version: "1.1.0",
  description: "KeyMessage を大きく主役として見せる。適合する Pattern がない場合の受け皿。",
  variants: ["standard"],
  fallback: true,

  compatibility: {
    logicalRelations: [...LOGICAL_RELATIONS],
  },

  bindingsSchema: MessageOnlyBindingsSchema,

  match(slide: SlideSpec): MatchResult {
    let score = 0.3;
    const reasons: string[] = [];

    if (slide.evidence.length === 0 && !slide.content) {
      score += 0.28;
      reasons.push("図示すべき evidence / content が存在しない");
    } else if (slide.evidence.length <= 1 && !slide.content) {
      score += 0.16;
      reasons.push("evidence が 1 件のみで構造化の余地が小さい");
    } else if (slide.content?.kind === "points" && slide.content.items.length === 1) {
      score += 0.24;
      reasons.push("補足 1 件のみでメッセージを主役にできる");
    }

    if (STRONG_ROLES.has(slide.slideRole)) {
      score += 0.12;
      reasons.push(`slideRole=${slide.slideRole} はメッセージ自体を主役にできる`);
    }

    if (slide.keyMessage.length <= 40) {
      score += 0.06;
      reasons.push("keyMessage が短く大きく見せられる");
    }

    if (slide.content && (slide.explanation || requiresFullContentText(slide.content))) {
      reasons.push(`content.${slide.content.kind} は図解ではなく全項目を保持した構造のテキスト表示へ退避する。長文は縮小され、可読性は Render Report で検査する`);
    }
    if (reasons.length === 0) reasons.push("常に適用可能な既定表現");

    return {
      compatible: true,
      score: Math.min(1, score),
      reason: reasons.join(" / "),
    };
  },

  buildBindings(slide: SlideSpec): MessageOnlyBindings {
    if (slide.content && (slide.explanation || requiresFullContentText(slide.content))) {
      return {
        eyebrow: roleLabel(slide.slideRole),
        headline: slide.keyMessage,
        subline: slide.supportingMessage,
        structureText: describeContent(slide.content),
        ...(slide.explanation && slide.content.kind === "points" ? { explanatoryItems: slide.content.items.map(({ label, description }) => ({ label, ...(description ? { description } : {}) })) } : {}),
        structureNodes: structureNodes(slide.content),
        sourceText: buildSourceText(slide),
      };
    }
    // 補足行: content があれば content を、無ければ evidence の要約を使う
    const contentItems =
      slide.content?.kind === "points"
        ? slide.content.items
        : slide.content?.kind === "sequence"
          ? slide.content.steps
          : [];

    const points =
      contentItems.length > 0
        ? contentItems
            .map((item) =>
              item.description ? `${item.label}：${item.description}` : item.label,
            )
        : slide.evidence
            .slice(0, slide.explanation ? undefined : 2)
            .map((e) => slide.explanation ? evidenceDetail(e) : truncate(evidenceDetail(e), 60))
            .filter((p) => p.length > 0);

    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      points: points.length > 0 ? points : undefined,
      ...(slide.content && contentItems.some((item) => item.id !== undefined || item.evidenceRefs !== undefined || item.members !== undefined)
        ? { structureNodes: structureNodes(slide.content) } : {}),
      sourceText: buildSourceText(slide),
    };
  },

  render(b: MessageOnlyBindings, ctx: RenderContext): void {
    if (ctx.layoutBody) {
      const body = drawFooter(ctx, drawHeader(ctx, b), b.sourceText);
      if (b.explanatoryItems) {
        drawExplanatoryNodes(ctx, body, b.explanatoryItems.map((item, i) => ({ ...item, name: `item-${i + 1}` })), false);
        return;
      }
      const text = b.structureText ?? (b.points ?? []).join("\n\n");
      if (text) {
        const fontSize = Math.max(14, ctx.theme.fontSizes.body);
        if (ctx.measureText(text, { widthIn: body.w, fontSize, lineHeight: 1.3 }) > body.h) {
          throw new Error("Explanatory fallback body overflows; split the slide");
        }
        ctx.text(body, text, { fontSize, lineHeight: 1.3, shrink: false, name: "explanatory-body" });
      }
      return;
    }
    if (b.structureText) {
      renderStructureText(b, ctx);
      return;
    }
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
    const fixedH = gapAfterHeadline + sublineH + gapBeforePoints + pointsTotal;

    // fallback 契約: どんなに長い keyMessage でも必ず収める。
    // 見出しサイズを段階的に下げ、本文領域に収まる最大サイズを選ぶ。
    const maxSize = Math.round(theme.fontSizes.keyMessage * 1.45);
    const minSize = theme.fontSizes.heading;
    let headlineSize = maxSize;
    let headlineH = ctx.measureText(b.headline, {
      widthIn: textW,
      fontSize: headlineSize,
      lineHeight: 1.26,
    });
    while (headlineH + fixedH > body.h && headlineSize - 2 >= minSize) {
      headlineSize -= 2;
      headlineH = ctx.measureText(b.headline, {
        widthIn: textW,
        fontSize: headlineSize,
        lineHeight: 1.26,
      });
    }

    const totalH = headlineH + fixedH;
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

function requiresFullContentText(content?: SemanticContent): boolean {
  return isRelationshipContent(content) || contentItemEntries(content).length > 3;
}

function structureNodes(content: SemanticContent): StructureNodeBinding[] {
  return contentItemEntries(content).map(({ item, path }) => ({
    path: [...path],
    ...(item.id !== undefined ? { nodeId: item.id } : {}),
    ...(item.evidenceRefs !== undefined ? { evidenceRefs: [...item.evidenceRefs] } : {}),
    ...(item.members !== undefined ? { members: [...item.members] } : {}),
  }));
}

/** Full-content emergency expression; small legacy point summaries retain their original layout. */
function describeContent(content: SemanticContent): string {
  const describe = (item: ContentItem): string =>
    item.description ? `${item.label}：${item.description}` : item.label;
  const list = (items: ContentItem[], ordered = false): string =>
    items.map((item, index) => `${ordered ? `${index + 1}.` : "•"} ${describe(item)}`).join("\n");
  switch (content.kind) {
    case "points":
      return `構造のテキスト表示：並列項目（順序を意味しない）\n${list(content.items)}`;
    case "sequence":
      return `構造のテキスト表示：順序（この順番で進む）\n${list(content.steps, true)}`;
    case "distribution":
      return `構造のテキスト表示：分配（1 対多）\n起点：${describe(content.center)}\n分配先（すべて同じ起点から）：\n${list(content.items)}`;
    case "convergence":
      return `構造のテキスト表示：集約（多対 1）\n入力（すべて同じ到達点へ）：\n${list(content.inputs)}\n到達点：${describe(content.outcome)}`;
    case "cycle":
      return `構造のテキスト表示：循環（この順番で進み、最終工程から最初へ戻る）\n${list(content.steps, true)}`;
    case "context":
      return `構造のテキスト表示：共有環境\n環境：${describe(content.environment)}\nこの環境を共有する項目（順序・因果は意味しない）：\n${list(content.items)}`;
  }
}

/** Bounded boxes keep every character; unusually dense text is flagged by layout validation. */
function renderStructureText(b: MessageOnlyBindings, ctx: RenderContext): void {
  const { frame, theme } = ctx;
  const gap = Math.min(theme.spacing.sm, frame.h * 0.02);
  let remaining = frame;
  if (b.sourceText) {
    const [footer, rest] = splitBottom(remaining, frame.h * 0.1, gap);
    ctx.text(footer, b.sourceText, { role: "caption", color: "mutedText", shrink: true, name: "source" });
    remaining = rest;
  }
  const [headline, afterHeadline] = splitTop(remaining, frame.h * 0.18, gap);
  ctx.text(headline, b.headline, { role: "keyMessage", color: "text", bold: true, shrink: true, name: "key-message" });
  remaining = afterHeadline;
  if (b.subline) {
    const [subline, rest] = splitTop(remaining, frame.h * 0.15, gap);
    ctx.text(subline, b.subline, { role: "body", color: "mutedText", shrink: true, name: "supporting-message" });
    remaining = rest;
  }
  ctx.text(remaining, b.structureText!, {
    role: "body", color: "text", lineHeight: 1.25, shrink: true, name: "relationship-text-fallback",
  });
}
