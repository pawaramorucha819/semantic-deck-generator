import type { PatternAdvisor } from "./ai/pattern-advisor.js";
import type { SlideSpec } from "./domain/types.js";
import type {
  KnowledgeQuery,
  KnowledgeRetriever,
  KnowledgeSearchResult,
} from "./knowledge/retrieval/retrieval.js";
import type { DeckKnowledgeStore } from "./knowledge/store/store.js";
import type { PatternRegistry } from "./patterns/registry.js";
import type { SelectionMode } from "./plan/decision-trace.js";
import {
  decideSlidePattern,
  type SlideSelectionOutcome,
} from "./selector/selection-policy.js";

/**
 * よく使う操作の短い入口。
 *
 * ここには「別名を付けただけ」の薄い関数しか置きません。
 * 実体は各モジュールにあり、細かい制御が要るときはそちらを直接使ってください。
 */

export { ingestDeckFromPptx as ingestDeck } from "./knowledge/ingest/ingest.js";
export { loadTemplateFromDirectory as loadTemplate } from "./templates/loader.js";
export { loadPatternPackage } from "./patterns/declarative/loader.js";

/** 過去スライドを検索する（`store.search` の別名） */
export async function searchKnowledge(
  store: DeckKnowledgeStore,
  query: KnowledgeQuery,
): Promise<KnowledgeSearchResult[]> {
  return store.search(query);
}

export interface AdvisePatternOptions {
  registry?: PatternRegistry;
  advisor?: PatternAdvisor;
  retriever?: KnowledgeRetriever;
  /** 既定 'assist'（推薦は記録するが採用はしない） */
  mode?: SelectionMode;
}

/**
 * 1 スライドについて Pattern を決める（決定論的選択 → Knowledge → AI → 最終決定）。
 *
 * 戻り値には Decision Trace が含まれるので、「なぜそうなったか」をそのまま説明できます。
 */
export async function advisePattern(
  slide: SlideSpec,
  options: AdvisePatternOptions = {},
): Promise<SlideSelectionOutcome> {
  return decideSlidePattern(slide, {
    mode: options.mode ?? "assist",
    ...(options.registry ? { registry: options.registry } : {}),
    ...(options.advisor ? { advisor: options.advisor } : {}),
    ...(options.retriever ? { retriever: options.retriever } : {}),
  });
}
