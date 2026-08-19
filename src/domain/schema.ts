import { z } from "zod";
import {
  EVIDENCE_TYPES,
  LOGICAL_RELATIONS,
  SLIDE_ROLES,
  type PresentationSpec,
  type SlideSpec,
} from "./types.js";

/**
 * presentation.json のスキーマ。
 *
 * すべて `.strict()` にしてあります。これは単なる厳しさではなく設計の強制です:
 * `fontSize` / `color` / `x` / `y` のような「表現・描画」の語彙が
 * presentation.json に紛れ込んだ瞬間に validation エラーになります。
 */

const nonEmpty = (label: string, max = 400) =>
  z.string().trim().min(1, `${label} は空にできません`).max(max, `${label} が長すぎます`);

export const SourceSchema = z
  .object({
    id: nonEmpty("source.id", 120),
    title: nonEmpty("source.title", 300),
    url: z.string().url("source.url は URL 形式である必要があります").optional(),
    publisher: z.string().trim().min(1).max(200).optional(),
    date: z.string().trim().min(1).max(40).optional(),
    locator: z.string().trim().min(1).max(200).optional(),
    accessedAt: z.string().trim().min(1).max(40).optional(),
  })
  .strict();

const evidenceBase = {
  id: nonEmpty("evidence.id", 120),
  statement: nonEmpty("evidence.statement", 400),
  sourceIds: z.array(z.string().trim().min(1).max(120)).min(1).max(8).optional(),
};

export const MetricEvidenceSchema = z
  .object({
    ...evidenceBase,
    type: z.literal("metric"),
    data: z
      .object({
        label: z.string().trim().min(1).max(120).optional(),
        value: z.number().finite(),
        unit: z.string().trim().max(40).optional(),
        direction: z.enum(["increase", "decrease", "flat"]).optional(),
        period: z.string().trim().max(80).optional(),
        baseline: z.number().finite().optional(),
      })
      .strict(),
  })
  .strict();

export const CaseStudyEvidenceSchema = z
  .object({
    ...evidenceBase,
    type: z.literal("caseStudy"),
    data: z
      .object({
        organization: z.string().trim().min(1).max(120).optional(),
        before: z.number().finite().optional(),
        after: z.number().finite().optional(),
        unit: z.string().trim().max(40).optional(),
        beforeLabel: z.string().trim().min(1).max(80).optional(),
        afterLabel: z.string().trim().min(1).max(80).optional(),
        changeLabel: z.string().trim().min(1).max(80).optional(),
      })
      .strict(),
  })
  .strict();

export const TextEvidenceSchema = z
  .object({
    ...evidenceBase,
    type: z.literal("text"),
    data: z
      .object({
        label: z.string().trim().min(1).max(120).optional(),
        detail: z.string().trim().min(1).max(400).optional(),
      })
      .strict()
      .optional(),
  })
  .strict();

export const ComparisonEvidenceSchema = z
  .object({
    ...evidenceBase,
    type: z.literal("comparison"),
    data: z
      .object({
        unit: z.string().trim().max(40).optional(),
        items: z
          .array(
            z
              .object({
                label: nonEmpty("comparison.items[].label", 120),
                value: z.union([z.number().finite(), z.string().trim().min(1).max(80)]),
                note: z.string().trim().min(1).max(200).optional(),
              })
              .strict(),
          )
          .min(2, "comparison evidence には 2 件以上の item が必要です")
          .max(6, "comparison evidence の item は 6 件までです"),
      })
      .strict(),
  })
  .strict();

export const EvidenceSchema = z.discriminatedUnion("type", [
  MetricEvidenceSchema,
  CaseStudyEvidenceSchema,
  TextEvidenceSchema,
  ComparisonEvidenceSchema,
]);

const ContentItemSchema = z
  .object({
    label: nonEmpty("content item の label", 120),
    description: z.string().trim().min(1).max(400).optional(),
    evidenceRefs: z.array(z.string().trim().min(1).max(120)).min(1).max(8).optional(),
  })
  .strict();

export const SemanticContentSchema = z.discriminatedUnion("kind", [
  z
    .object({
      kind: z.literal("sequence"),
      steps: z
        .array(ContentItemSchema)
        .min(2, "sequence には 2 件以上の step が必要です")
        .max(6, "sequence の step は 6 件までです"),
    })
    .strict(),
  z
    .object({
      kind: z.literal("points"),
      items: z
        .array(ContentItemSchema)
        .min(1, "points には 1 件以上の item が必要です")
        .max(6, "points の item は 6 件までです"),
    })
    .strict(),
]);

export const SlideSpecSchema = z
  .object({
    id: nonEmpty("slide.id", 120),
    keyMessage: nonEmpty("slide.keyMessage", 200),
    supportingMessage: z.string().trim().min(1).max(400).optional(),
    slideRole: z.enum(SLIDE_ROLES),
    logicalRelation: z.enum(LOGICAL_RELATIONS),
    content: SemanticContentSchema.optional(),
    evidence: z.array(EvidenceSchema).max(12, "evidence は 12 件までです"),
    sources: z.array(SourceSchema).max(12).optional(),
  })
  .strict()
  .superRefine((slide, ctx) => {
    const evidenceIds = new Set<string>();
    slide.evidence.forEach((e, i) => {
      if (evidenceIds.has(e.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["evidence", i, "id"],
          message: `evidence.id が重複しています: ${e.id}`,
        });
      }
      evidenceIds.add(e.id);
    });

    const sourceIdSet = new Set((slide.sources ?? []).map((s) => s.id));
    slide.evidence.forEach((e, i) => {
      for (const ref of e.sourceIds ?? []) {
        if (!sourceIdSet.has(ref)) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["evidence", i, "sourceIds"],
            message: `未定義の sourceId を参照しています: ${ref}`,
          });
        }
      }
    });

    // content の evidenceRefs は evidence[].id を参照していなければならない
    if (slide.content) {
      const items = slide.content.kind === "sequence" ? slide.content.steps : slide.content.items;
      const itemsKey = slide.content.kind === "sequence" ? "steps" : "items";
      items.forEach((item, i) => {
        for (const ref of item.evidenceRefs ?? []) {
          if (!evidenceIds.has(ref)) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              path: ["content", itemsKey, i, "evidenceRefs"],
              message: `未定義の evidence を参照しています: ${ref}`,
            });
          }
        }
      });
    }
  });

export const PresentationSpecSchema = z
  .object({
    title: nonEmpty("title", 200),
    subtitle: z.string().trim().min(1).max(300).optional(),
    author: z.string().trim().min(1).max(120).optional(),
    slides: z.array(SlideSpecSchema).min(1, "slides は 1 枚以上必要です"),
  })
  .strict()
  .superRefine((spec, ctx) => {
    const seen = new Set<string>();
    spec.slides.forEach((slide, i) => {
      if (seen.has(slide.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["slides", i, "id"],
          message: `slide.id が重複しています: ${slide.id}`,
        });
      }
      seen.add(slide.id);
    });
  });

function formatIssues(issues: z.ZodIssue[]): string {
  return issues
    .map((issue) => `  - ${issue.path.join(".") || "(root)"}: ${issue.message}`)
    .join("\n");
}

export class PresentationValidationError extends Error {
  constructor(
    summary: string,
    readonly issues: z.ZodIssue[],
  ) {
    super(`${summary}\n${formatIssues(issues)}`);
    this.name = "PresentationValidationError";
  }

  /** path 付きの箇条書き（message にも同じ内容が含まれる） */
  format(): string {
    return formatIssues(this.issues);
  }
}

/** 検証済みの `PresentationSpec` を返す。失敗時は `PresentationValidationError`。 */
export function parsePresentation(input: unknown): PresentationSpec {
  const result = PresentationSpecSchema.safeParse(input);
  if (!result.success) {
    throw new PresentationValidationError(
      `presentation.json のスキーマ検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  return result.data as PresentationSpec;
}

export function parseSlide(input: unknown): SlideSpec {
  const result = SlideSpecSchema.safeParse(input);
  if (!result.success) {
    throw new PresentationValidationError(
      `slide の検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  return result.data as SlideSpec;
}

/** Evidence の種別名一覧（外部公開用） */
export const evidenceTypeNames = EVIDENCE_TYPES;
