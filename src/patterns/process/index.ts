import type { SlideSpec } from "../../domain/types.js";
import type { RenderContext, TextRun } from "../../render/context.js";
import { columns, inset, splitTop, type Rect } from "../../render/geometry.js";
import type { MatchResult, VisualPattern } from "../pattern.js";
import { centerBlockY, drawFooter, drawHeader } from "../shared/layout.js";
import {
  buildSourceText,
  evidenceDetail,
  evidenceLabel,
  roleLabel,
} from "../shared/semantics.js";

export interface ProcessStepBinding {
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

/**
 * Process — 工程・順序を左から右への矢印で表現する。
 *
 * 「evidence の配列順が意味を持つ」のは logicalRelation が process / timeline の
 * ときだけ、という前提に立っている。この前提が Pattern 適用条件そのもの。
 */
export const processPattern: VisualPattern<ProcessBindings> = {
  id: "process",
  version: "1.0.0",
  description: "工程・順序を左から右へのシェブロンで表現する。",
  variants: ["standard"],

  compatibility: {
    logicalRelations: ["process", "timeline"],
  },

  constraints: {
    minEvidence: 2,
    maxEvidence: 5,
    maxKeyMessageLength: 70,
  },

  match(slide: SlideSpec): MatchResult {
    const count = slide.evidence.length;
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
      reasons.push(`${count} 工程で横一列に収まる`);
    }

    const longest = Math.max(...slide.evidence.map((e) => evidenceLabel(e, 20).length));
    if (longest > 16) {
      score -= 0.08;
      reasons.push("工程名が長くシェブロン内が窮屈になる可能性");
    }

    return {
      compatible: true,
      score: Math.max(0, Math.min(1, score)),
      reason: reasons.join(" / "),
      variantId: "standard",
    };
  },

  buildBindings(slide: SlideSpec): ProcessBindings {
    return {
      eyebrow: roleLabel(slide.slideRole),
      headline: slide.keyMessage,
      subline: slide.supportingMessage,
      stepPrefix: slide.logicalRelation === "timeline" ? "" : "STEP",
      steps: slide.evidence.map((e, i) => {
        const label = evidenceLabel(e, 14);
        const description = evidenceDetail(e);
        return {
          index: i + 1,
          label,
          description: description === label ? undefined : description,
        };
      }),
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
