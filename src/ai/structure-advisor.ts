import { z } from "zod";
import { parsePresentation } from "../domain/schema.js";
import type { PresentationSpec } from "../domain/types.js";
import { LOGICAL_RELATIONS, SLIDE_ROLES } from "../domain/types.js";
import type { KnowledgeSearchResult } from "../knowledge/retrieval/retrieval.js";
import { AIUnavailableError, type AIProvider } from "./provider.js";

/**
 * AI Structure Advisor — 上流（意味の構造化）を支援する拡張ポイント。
 *
 * 重要なのは出口の契約:
 *   AI の出力は **必ず既存の `parsePresentation()` を通る**。
 *   AI 専用の別 Semantic Model を作って Renderer へ流すことはしない。
 *
 * つまり AI が座標や色を書こうとしても、presentation.json のスキーマが
 * `.strict()` なので検証で落ちる。
 */

export interface StructureAdviceInput {
  /** 作りたい資料の説明・素材（自由文） */
  request: string;
  /** 参考にする過去スライド */
  references?: KnowledgeSearchResult[];
  maxSlides?: number;
  /** 資料タイトルの希望 */
  title?: string;
}

/**
 * AI から受け取る「素の」構造。
 * ここで受けたあと `parsePresentation()` に通すので、二重に検証されることになる。
 */
const DraftSchema = z.object({}).passthrough();

const INSTRUCTIONS = `あなたは資料の論理構造を設計するアシスタントです。
presentation.json（意味と論理だけを持つ中間表現）を出力してください。

必ず守ること:
- 出力は presentation.json のスキーマに厳密に従う JSON オブジェクト 1 つだけ
- 各スライドは keyMessage（最も伝えたい 1 文）を必ず持つ
- slideRole は次のいずれか: ${SLIDE_ROLES.join(", ")}
- logicalRelation は次のいずれか: ${LOGICAL_RELATIONS.join(", ")}
- 「見せる対象」（工程・並列項目）は content に、「主張を裏付ける事実」は evidence に書く
- evidence[].sourceIds は同じスライドの sources[].id だけを参照する

絶対に書いてはいけないもの（スキーマ検証で必ず失敗します）:
- patternId / templateId / variantId
- x / y / w / h / color / fontSize などの描画情報
- スキーマに無いキー`;

export const STRUCTURE_ADVICE_JSON_SCHEMA: Record<string, unknown> = {
  type: "object",
  required: ["title", "slides"],
  properties: {
    title: { type: "string" },
    subtitle: { type: "string" },
    author: { type: "string" },
    slides: {
      type: "array",
      items: {
        type: "object",
        required: ["id", "keyMessage", "slideRole", "logicalRelation", "evidence"],
        properties: {
          id: { type: "string" },
          keyMessage: { type: "string" },
          supportingMessage: { type: "string" },
          slideRole: { type: "string", enum: [...SLIDE_ROLES] },
          logicalRelation: { type: "string", enum: [...LOGICAL_RELATIONS] },
          content: { type: "object" },
          evidence: { type: "array" },
          sources: { type: "array" },
        },
      },
    },
  },
};

export interface StructureAdviceResult {
  provider: string;
  spec: PresentationSpec;
  /** 参考にした過去スライド */
  referenceSlides: string[];
}

export class StructureAdvisor {
  constructor(private readonly provider: AIProvider) {}

  get available(): boolean {
    return this.provider.available;
  }

  /**
   * 過去資料を参考に presentation.json の下書きを作る。
   *
   * 失敗（Provider 不在 / 検証エラー）は握りつぶさず必ず例外にする。
   * 「AI を使ったふりをして空の資料を返す」ことはしない。
   */
  async advise(input: StructureAdviceInput): Promise<StructureAdviceResult> {
    if (!this.provider.available) {
      throw new AIUnavailableError(
        "Structure Advisor には AI Provider が必要です（--ai-mode off では利用できません）",
      );
    }

    const draft = await this.provider.generateStructured({
      purpose: "structure-planning",
      instructions: INSTRUCTIONS,
      input: {
        request: input.request,
        ...(input.title ? { preferredTitle: input.title } : {}),
        maxSlides: input.maxSlides ?? 8,
        references: (input.references ?? []).map((r) => ({
          slideId: r.entry.slideId,
          keyMessage: r.entry.semantic?.keyMessage,
          slideRole: r.entry.semantic?.slideRole,
          logicalRelation: r.entry.semantic?.logicalRelation,
          purpose: r.entry.metadata.purpose,
          audience: r.entry.metadata.audience,
          text: r.entry.content?.text?.slice(0, 500),
        })),
      },
      schema: DraftSchema,
      jsonSchema: STRUCTURE_ADVICE_JSON_SCHEMA,
      maxOutputTokens: 8192,
    });

    // ここが最終ゲート。AI 専用モデルは存在せず、必ず既存スキーマを通る。
    const spec = parsePresentation(draft);

    return {
      provider: this.provider.id,
      spec,
      referenceSlides: (input.references ?? []).map((r) => r.entry.slideId),
    };
  }
}

/** 単発利用のための便宜関数 */
export async function adviseStructure(
  provider: AIProvider,
  input: StructureAdviceInput,
): Promise<StructureAdviceResult> {
  return new StructureAdvisor(provider).advise(input);
}
