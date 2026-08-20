import { z } from "zod";
import { LOGICAL_RELATIONS, SLIDE_ROLES } from "../domain/types.js";
import type { LogicalRelation, SlideRole } from "../domain/types.js";
import { DECK_ID_PATTERN } from "./deck-id.js";

/**
 * Deck Knowledge Base — 過去資料を「1 スライド 1 エントリ」で蓄積する。
 *
 * ここに入るのは *過去の事実* であり、presentation.json（これから作る資料の意味）とは
 * 別物。Knowledge は生成の入力を汚さず、あくまで **参照・推薦の材料** として使う。
 */

export const KNOWLEDGE_SCHEMA_VERSION = "1.0";

/** 情報密度。レイアウト選択の参考に使う粗い指標。 */
export const DECK_DENSITIES = ["sparse", "medium", "dense"] as const;
export type DeckDensity = (typeof DECK_DENSITIES)[number];

export interface DeckKnowledgeEntry {
  schemaVersion: string;
  deckId: string;
  /** `<deckId>#<slideNumber>` 形式の一意な ID */
  slideId: string;

  metadata: {
    title?: string;
    organization?: string;
    department?: string;
    /** 資料の目的（例: "投資判断", "月次報告"） */
    purpose?: string;
    /** 想定読者（例: "経営会議", "現場リーダー"） */
    audience?: string;
    /** ISO 8601 推奨 */
    createdAt?: string;
    tags?: string[];
  };

  semantic?: {
    keyMessage?: string;
    slideRole?: SlideRole;
    logicalRelation?: LogicalRelation;
    keywords?: string[];
  };

  visual?: {
    patternId?: string;
    variantId?: string;
    templateId?: string;
    density?: DeckDensity;
    /** 要素数の内訳（レイアウトの重さの目安） */
    elementCounts?: { text?: number; shape?: number; image?: number; chart?: number };
  };

  /** 検索用の平文（PPTX から取り込んだ場合はスライド上の全テキスト） */
  content?: {
    text: string;
    bulletCount?: number;
  };

  source: {
    file: string;
    slideNumber: number;
    importedAt?: string;
    /** 生成物から取り込んだ場合の出所 */
    origin?: "pptx" | "generation";
  };
}

const shortText = (max: number) => z.string().trim().min(1).max(max);

export const DeckKnowledgeEntrySchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    deckId: z
      .string()
      .trim()
      .regex(DECK_ID_PATTERN, "deckId は英数字と - _ . のみ使用でき、先頭は英数字である必要があります"),
    slideId: shortText(200),
    metadata: z
      .object({
        title: shortText(300).optional(),
        organization: shortText(160).optional(),
        department: shortText(160).optional(),
        purpose: shortText(160).optional(),
        audience: shortText(160).optional(),
        createdAt: shortText(40).optional(),
        tags: z.array(shortText(60)).max(30).optional(),
      })
      .strict(),
    semantic: z
      .object({
        keyMessage: shortText(400).optional(),
        slideRole: z.enum(SLIDE_ROLES).optional(),
        logicalRelation: z.enum(LOGICAL_RELATIONS).optional(),
        keywords: z.array(shortText(60)).max(60).optional(),
      })
      .strict()
      .optional(),
    visual: z
      .object({
        patternId: shortText(120).optional(),
        variantId: shortText(120).optional(),
        templateId: shortText(120).optional(),
        density: z.enum(DECK_DENSITIES).optional(),
        elementCounts: z
          .object({
            text: z.number().int().min(0).max(10000).optional(),
            shape: z.number().int().min(0).max(10000).optional(),
            image: z.number().int().min(0).max(10000).optional(),
            chart: z.number().int().min(0).max(10000).optional(),
          })
          .strict()
          .optional(),
      })
      .strict()
      .optional(),
    content: z
      .object({
        text: z.string().max(20000),
        bulletCount: z.number().int().min(0).max(1000).optional(),
      })
      .strict()
      .optional(),
    source: z
      .object({
        file: shortText(400),
        slideNumber: z.number().int().min(1).max(10000),
        importedAt: shortText(40).optional(),
        origin: z.enum(["pptx", "generation"]).optional(),
      })
      .strict(),
  })
  .strict();

export class KnowledgeValidationError extends Error {
  constructor(
    summary: string,
    readonly issues: z.ZodIssue[],
  ) {
    super(
      `${summary}\n${issues
        .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
        .join("\n")}`,
    );
    this.name = "KnowledgeValidationError";
  }
}

export function parseKnowledgeEntry(input: unknown): DeckKnowledgeEntry {
  const result = DeckKnowledgeEntrySchema.safeParse(input);
  if (!result.success) {
    throw new KnowledgeValidationError(
      `Knowledge エントリの検証に失敗しました (${result.error.issues.length} 件)`,
      result.error.issues,
    );
  }
  const entry = result.data as DeckKnowledgeEntry;
  if (entry.schemaVersion !== KNOWLEDGE_SCHEMA_VERSION) {
    throw new KnowledgeValidationError(
      `未対応の Knowledge schemaVersion です (指定=${entry.schemaVersion}, 対応=${KNOWLEDGE_SCHEMA_VERSION})`,
      [],
    );
  }
  return entry;
}

/** 保存ファイル（decks/<deckId>.json）の形 */
export const DeckKnowledgeFileSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    deckId: z.string().trim().min(1).max(120),
    entries: z.array(DeckKnowledgeEntrySchema),
  })
  .strict();

/** 索引（index.json）の形 */
export const KnowledgeIndexSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    decks: z.array(
      z
        .object({
          deckId: z.string().trim().min(1).max(120),
          title: z.string().trim().min(1).max(300).optional(),
          slideCount: z.number().int().min(0).max(10000),
          file: z.string().trim().min(1).max(400),
          organization: z.string().trim().min(1).max(160).optional(),
          purpose: z.string().trim().min(1).max(160).optional(),
          audience: z.string().trim().min(1).max(160).optional(),
        })
        .strict(),
    ),
  })
  .strict();

export type KnowledgeIndex = z.infer<typeof KnowledgeIndexSchema>;

export function slideIdOf(deckId: string, slideNumber: number): string {
  return `${deckId}#${slideNumber}`;
}
