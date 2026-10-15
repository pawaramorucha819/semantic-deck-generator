import { z } from "zod";
import { contentItemEntries } from "./content.js";
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

export const ContentItemSchema = z
  .object({
    id: nonEmpty("content item の id", 120).optional(),
    members: z.array(nonEmpty("content item の members[]", 120)).max(1000).optional(),
    label: nonEmpty("content item の label", 120),
    description: z.string().trim().min(1).max(400).optional(),
    evidenceRefs: z.array(z.string().trim().min(1).max(120)).min(1).max(8).optional(),
  })
  .strict();

export const SequenceContentSchema = z
  .object({
    kind: z.literal("sequence"),
    steps: z.array(ContentItemSchema).min(2, "sequence には 2 件以上の step が必要です")
      .max(6, "sequence の step は 6 件までです"),
  })
  .strict();

export const PointsContentSchema = z
  .object({
    kind: z.literal("points"),
    items: z.array(ContentItemSchema).min(1, "points には 1 件以上の item が必要です")
      .max(6, "points の item は 6 件までです"),
  })
  .strict();

export const DistributionContentSchema = z
  .object({
    kind: z.literal("distribution"),
    center: ContentItemSchema,
    items: z.array(ContentItemSchema).min(2).max(4),
  })
  .strict();

export const ConvergenceContentSchema = z
  .object({
    kind: z.literal("convergence"),
    inputs: z.array(ContentItemSchema).min(2).max(4),
    outcome: ContentItemSchema,
  })
  .strict();

export const CycleContentSchema = z
  .object({
    kind: z.literal("cycle"),
    steps: z.array(ContentItemSchema).min(2).max(4),
  })
  .strict();

export const ContextContentSchema = z
  .object({
    kind: z.literal("context"),
    environment: ContentItemSchema,
    items: z.array(ContentItemSchema).min(1).max(4),
  })
  .strict();

export const SemanticContentSchema = z.discriminatedUnion("kind", [
  SequenceContentSchema,
  PointsContentSchema,
  DistributionContentSchema,
  ConvergenceContentSchema,
  CycleContentSchema,
  ContextContentSchema,
]);

export const DecompositionSchema = z
  .object({
    scope: nonEmpty("decomposition.scope", 400),
    criterion: nonEmpty("decomposition.criterion", 400),
    rationale: nonEmpty("decomposition.rationale", 1200),
    exclusions: z.array(nonEmpty("decomposition.exclusions[]", 400)).max(12).optional(),
    unresolvedQuestions: z.array(nonEmpty("decomposition.unresolvedQuestions[]", 400)).max(12).optional(),
    universe: z.array(nonEmpty("decomposition.universe[]", 120)).min(1).max(1000)
      .refine((members) => new Set(members).size === members.length, "decomposition.universe は重複できません")
      .optional(),
  })
  .strict();

export const SlideExplanationSchema = z.object({
  readerQuestion: nonEmpty("explanation.readerQuestion", 160),
  sections: z.array(z.object({
    heading: nonEmpty("explanation.heading", 80),
    body: nonEmpty("explanation.body", 1200),
    nodeRefs: z.array(nonEmpty("explanation.nodeRefs[]", 120)).min(1).max(12).refine(refs => new Set(refs).size === refs.length, "説明の参照は重複できません").optional(),
  }).strict()).min(1).max(3),
}).strict();

export const SlideSpecSchema = z
  .object({
    id: nonEmpty("slide.id", 120),
    keyMessage: nonEmpty("slide.keyMessage", 200),
    supportingMessage: z.string().trim().min(1).max(400).optional(),
    slideRole: z.enum(SLIDE_ROLES),
    logicalRelation: z.enum(LOGICAL_RELATIONS),
    content: SemanticContentSchema.optional(),
    decomposition: DecompositionSchema.optional(),
    explanation: SlideExplanationSchema.optional(),
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

    const explicitNodeIds = new Set(contentItemEntries(slide.content).map(({ item }) => item.id).filter(Boolean));
    slide.explanation?.sections.forEach((section, index) => {
      for (const ref of section.nodeRefs ?? []) {
        if (!explicitNodeIds.has(ref)) ctx.addIssue({ code: z.ZodIssueCode.custom,
          path: ["explanation", "sections", index, "nodeRefs"], message: `未定義の説明対象 node ID: ${ref}` });
      }
    });

    // content の evidenceRefs は evidence[].id を参照していなければならない
    if (slide.content) {
      const contentIds = new Set<string>();
      contentItemEntries(slide.content).forEach(({ item, path }) => {
        const effectiveId = item.id ?? path.join(".");
        if (contentIds.has(effectiveId)) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["content", ...path, "id"],
            message: `content item の id が重複しています（省略時の path ID を含む）: ${effectiveId}`,
          });
        }
        contentIds.add(effectiveId);
        for (const ref of item.evidenceRefs ?? []) {
          if (!evidenceIds.has(ref)) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              path: ["content", ...path, "evidenceRefs"],
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
