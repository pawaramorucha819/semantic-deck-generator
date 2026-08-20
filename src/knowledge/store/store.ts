import {
  searchEntries,
  type KnowledgeQuery,
  type KnowledgeSearchResult,
} from "../retrieval/retrieval.js";
import { parseKnowledgeEntry, type DeckKnowledgeEntry } from "../schema.js";

/**
 * Deck Knowledge Store の抽象。
 *
 * V2 の既定はローカル JSON（Git 管理できる）。
 * 将来 PostgreSQL / pgvector / Elasticsearch / Vector DB へ差し替えるときは
 * このインターフェースを実装すればよく、Selector や AI 層は影響を受けない。
 */
export interface DeckKnowledgeStore {
  readonly id: string;
  add(entries: readonly DeckKnowledgeEntry[]): Promise<void>;
  all(): Promise<DeckKnowledgeEntry[]>;
  get(slideId: string): Promise<DeckKnowledgeEntry | undefined>;
  removeDeck(deckId: string): Promise<void>;
  search(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]>;
  stats(): Promise<{ decks: number; slides: number }>;
}

/** テスト・一時利用向けのインメモリ実装 */
export class InMemoryKnowledgeStore implements DeckKnowledgeStore {
  readonly id = "memory";
  private readonly bySlideId = new Map<string, DeckKnowledgeEntry>();

  constructor(entries: readonly DeckKnowledgeEntry[] = []) {
    for (const entry of entries) this.bySlideId.set(entry.slideId, parseKnowledgeEntry(entry));
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  async add(entries: readonly DeckKnowledgeEntry[]): Promise<void> {
    for (const entry of entries) {
      this.bySlideId.set(entry.slideId, parseKnowledgeEntry(entry));
    }
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  async all(): Promise<DeckKnowledgeEntry[]> {
    return [...this.bySlideId.values()].sort((a, b) => a.slideId.localeCompare(b.slideId));
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  async get(slideId: string): Promise<DeckKnowledgeEntry | undefined> {
    return this.bySlideId.get(slideId);
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  async removeDeck(deckId: string): Promise<void> {
    for (const [slideId, entry] of this.bySlideId) {
      if (entry.deckId === deckId) this.bySlideId.delete(slideId);
    }
  }

  async search(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]> {
    return searchEntries(await this.all(), query);
  }

  async stats(): Promise<{ decks: number; slides: number }> {
    const entries = await this.all();
    return { decks: new Set(entries.map((e) => e.deckId)).size, slides: entries.length };
  }
}

/** 複数ストアを 1 つに見せる（company / project の知識を束ねる用途） */
export class CompositeKnowledgeStore implements DeckKnowledgeStore {
  readonly id: string;

  constructor(private readonly stores: readonly DeckKnowledgeStore[]) {
    this.id = `composite(${stores.map((s) => s.id).join(",")})`;
  }

  async add(entries: readonly DeckKnowledgeEntry[]): Promise<void> {
    const first = this.stores[0];
    if (!first) throw new Error("追加先のストアがありません");
    await first.add(entries);
  }

  async all(): Promise<DeckKnowledgeEntry[]> {
    const seen = new Map<string, DeckKnowledgeEntry>();
    for (const store of this.stores) {
      for (const entry of await store.all()) {
        if (!seen.has(entry.slideId)) seen.set(entry.slideId, entry);
      }
    }
    return [...seen.values()].sort((a, b) => a.slideId.localeCompare(b.slideId));
  }

  async get(slideId: string): Promise<DeckKnowledgeEntry | undefined> {
    for (const store of this.stores) {
      const found = await store.get(slideId);
      if (found) return found;
    }
    return undefined;
  }

  async removeDeck(deckId: string): Promise<void> {
    for (const store of this.stores) await store.removeDeck(deckId);
  }

  async search(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]> {
    return searchEntries(await this.all(), query);
  }

  async stats(): Promise<{ decks: number; slides: number }> {
    const entries = await this.all();
    return { decks: new Set(entries.map((e) => e.deckId)).size, slides: entries.length };
  }
}
