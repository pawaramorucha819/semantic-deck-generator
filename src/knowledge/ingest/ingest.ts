import { basename } from "node:path";
import type { LogicalRelation, PresentationSpec, SlideRole } from "../../domain/types.js";
import type { VisualPlan } from "../../plan/visual-plan.js";
import { readDeck } from "../../pptx/extract.js";
import type { ExtractedSlide } from "../../pptx/model.js";
import { PptxPackage } from "../../pptx/package.js";
import { tokenize } from "../retrieval/retrieval.js";
import {
  KNOWLEDGE_SCHEMA_VERSION,
  slideIdOf,
  type DeckDensity,
  type DeckKnowledgeEntry,
} from "../schema.js";

/**
 * Deck Ingestor — 過去資料を Knowledge へ取り込む。
 *
 * 入口は 2 つある。
 *  1. `ingestDeckFromPptx`      … 過去の .pptx（意味は不明なのでテキストと構造だけ）
 *  2. `ingestDeckFromGeneration`… 自分で生成したデッキ（意味も Pattern も分かる）
 *
 * 2 があることで「生成 → 知識が貯まる → 次の生成が良くなる」という循環が閉じる。
 */

export interface DeckMetadataInput {
  title?: string;
  organization?: string;
  department?: string;
  purpose?: string;
  audience?: string;
  createdAt?: string;
  tags?: string[];
}

/** スライド単位の人手アノテーション（意味は PPTX からは決まらないため） */
export interface SlideAnnotation {
  keyMessage?: string;
  slideRole?: SlideRole;
  logicalRelation?: LogicalRelation;
  keywords?: string[];
  patternId?: string;
}

export interface IngestReport {
  deckId: string;
  source: string;
  slides: number;
  annotated: number;
  unsupportedFeatures: string[];
  warnings: string[];
}

export interface IngestResult {
  entries: DeckKnowledgeEntry[];
  report: IngestReport;
}

export interface IngestPptxOptions {
  deckId: string;
  metadata?: DeckMetadataInput;
  /** スライド番号 → 人手アノテーション */
  annotations?: Record<number, SlideAnnotation>;
  /** 記録する取り込み日時（決定論性のため呼び出し側が渡す） */
  importedAt?: string;
  sourceLabel?: string;
  /** キーワードとして残す上位語数（既定 12） */
  maxKeywords?: number;
}

function densityOf(slide: ExtractedSlide, textLength: number): DeckDensity {
  const count = slide.elements.length;
  if (count <= 5 && textLength <= 200) return "sparse";
  if (count >= 12 || textLength >= 600) return "dense";
  return "medium";
}

/** 出現頻度の高い順にキーワードを取り出す（決定論的） */
export function extractKeywords(text: string, max = 12): string[] {
  const counts = new Map<string, number>();
  for (const token of tokenize(text)) {
    if (token.length < 2) continue;
    counts.set(token, (counts.get(token) ?? 0) + 1);
  }
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, max)
    .map(([token]) => token);
}

/** 過去の .pptx を取り込む */
export async function ingestDeckFromPptx(
  input: Buffer | string,
  options: IngestPptxOptions,
): Promise<IngestResult> {
  const pkg = await PptxPackage.open(input);
  const deck = await readDeck(pkg, { includeNotes: true });
  const sourceLabel =
    options.sourceLabel ?? (typeof input === "string" ? basename(input) : "(buffer)");

  const entries: DeckKnowledgeEntry[] = [];
  let annotated = 0;

  for (const slide of deck.slides) {
    const texts = slide.elements
      .map((e) => e.text?.plain ?? "")
      .filter((t) => t.trim().length > 0);
    const text = texts.join("\n");
    const annotation = options.annotations?.[slide.slideNumber];
    if (annotation) annotated += 1;

    // 見出しはタイトルプレースホルダ、無ければ最大フォントのテキスト
    const largest = [...slide.elements]
      .filter((e) => e.text?.plain)
      .sort((a, b) => (b.text?.maxFontSize ?? 0) - (a.text?.maxFontSize ?? 0))[0];
    const keyMessage =
      annotation?.keyMessage ?? slide.title ?? largest?.text?.plain.split("\n")[0];

    const counts = {
      text: slide.elements.filter((e) => e.kind === "text" || e.kind === "placeholder").length,
      shape: slide.elements.filter((e) => e.kind === "shape" || e.kind === "line").length,
      image: slide.elements.filter((e) => e.kind === "image").length,
    };

    entries.push({
      schemaVersion: KNOWLEDGE_SCHEMA_VERSION,
      deckId: options.deckId,
      slideId: slideIdOf(options.deckId, slide.slideNumber),
      metadata: {
        ...(options.metadata?.title ?? deck.metadata.title
          ? { title: options.metadata?.title ?? deck.metadata.title! }
          : {}),
        ...(options.metadata?.organization ?? deck.metadata.company
          ? { organization: options.metadata?.organization ?? deck.metadata.company! }
          : {}),
        ...(options.metadata?.department ? { department: options.metadata.department } : {}),
        ...(options.metadata?.purpose ? { purpose: options.metadata.purpose } : {}),
        ...(options.metadata?.audience ? { audience: options.metadata.audience } : {}),
        ...(options.metadata?.createdAt ?? deck.metadata.created
          ? { createdAt: options.metadata?.createdAt ?? deck.metadata.created! }
          : {}),
        ...(options.metadata?.tags?.length ? { tags: options.metadata.tags } : {}),
      },
      semantic: {
        ...(keyMessage ? { keyMessage: keyMessage.slice(0, 400) } : {}),
        ...(annotation?.slideRole ? { slideRole: annotation.slideRole } : {}),
        ...(annotation?.logicalRelation ? { logicalRelation: annotation.logicalRelation } : {}),
        keywords: annotation?.keywords ?? extractKeywords(text, options.maxKeywords ?? 12),
      },
      visual: {
        ...(annotation?.patternId ? { patternId: annotation.patternId } : {}),
        density: densityOf(slide, text.length),
        elementCounts: counts,
      },
      content: { text: text.slice(0, 20000), bulletCount: texts.length },
      source: {
        file: sourceLabel,
        slideNumber: slide.slideNumber,
        ...(options.importedAt ? { importedAt: options.importedAt } : {}),
        origin: "pptx",
      },
    });
  }

  return {
    entries,
    report: {
      deckId: options.deckId,
      source: sourceLabel,
      slides: entries.length,
      annotated,
      unsupportedFeatures: [
        ...new Set([...deck.unsupported, ...deck.slides.flatMap((s) => s.unsupported)]),
      ],
      warnings: [
        ...new Set([...deck.warnings, ...deck.slides.flatMap((s) => s.warnings)]),
        ...(annotated === 0
          ? [
              "slideRole / logicalRelation は PPTX からは決まらないため未設定です（annotations で補えます）",
            ]
          : []),
      ],
    },
  };
}

export interface IngestGenerationOptions {
  deckId: string;
  metadata?: DeckMetadataInput;
  importedAt?: string;
  sourceLabel?: string;
}

/**
 * 自分で生成したデッキを取り込む。
 *
 * presentation.json（意味）と Visual Plan（表現の決定）が揃っているので、
 * slideRole / logicalRelation / patternId まで含めた完全なエントリになる。
 * これが「作った資料がそのまま知識資産になる」経路。
 */
export function ingestDeckFromGeneration(
  spec: PresentationSpec,
  plan: VisualPlan,
  options: IngestGenerationOptions,
): IngestResult {
  const sourceLabel = options.sourceLabel ?? `${options.deckId}.pptx`;
  const specById = new Map(spec.slides.map((s) => [s.id, s]));

  const entries: DeckKnowledgeEntry[] = plan.slides.map((planSlide) => {
    const slide = specById.get(planSlide.slideId);
    const contentItems = slide?.content
      ? slide.content.kind === "sequence"
        ? slide.content.steps
        : slide.content.items
      : [];
    const text = [
      slide?.keyMessage ?? "",
      slide?.supportingMessage ?? "",
      ...contentItems.map((i) => `${i.label} ${i.description ?? ""}`),
      ...(slide?.evidence ?? []).map((e) => e.statement),
    ]
      .join("\n")
      .trim();

    const elementCount = contentItems.length + (slide?.evidence.length ?? 0);
    const density: DeckDensity =
      elementCount <= 2 ? "sparse" : elementCount >= 6 ? "dense" : "medium";

    return {
      schemaVersion: KNOWLEDGE_SCHEMA_VERSION,
      deckId: options.deckId,
      slideId: slideIdOf(options.deckId, planSlide.slideNumber),
      metadata: {
        title: options.metadata?.title ?? plan.title,
        ...(options.metadata?.organization ? { organization: options.metadata.organization } : {}),
        ...(options.metadata?.department ? { department: options.metadata.department } : {}),
        ...(options.metadata?.purpose ? { purpose: options.metadata.purpose } : {}),
        ...(options.metadata?.audience ? { audience: options.metadata.audience } : {}),
        ...(options.metadata?.createdAt ? { createdAt: options.metadata.createdAt } : {}),
        ...(options.metadata?.tags?.length ? { tags: options.metadata.tags } : {}),
      },
      semantic: {
        ...(slide?.keyMessage ? { keyMessage: slide.keyMessage } : {}),
        ...(slide?.slideRole ? { slideRole: slide.slideRole } : {}),
        ...(slide?.logicalRelation ? { logicalRelation: slide.logicalRelation } : {}),
        keywords: extractKeywords(text),
      },
      visual: {
        patternId: planSlide.patternId,
        variantId: planSlide.variantId,
        ...(plan.templateId ? { templateId: plan.templateId } : {}),
        density,
      },
      content: { text: text.slice(0, 20000), bulletCount: contentItems.length },
      source: {
        file: sourceLabel,
        slideNumber: planSlide.slideNumber,
        ...(options.importedAt ? { importedAt: options.importedAt } : {}),
        origin: "generation",
      },
    };
  });

  return {
    entries,
    report: {
      deckId: options.deckId,
      source: sourceLabel,
      slides: entries.length,
      annotated: entries.length,
      unsupportedFeatures: [],
      warnings: [],
    },
  };
}
