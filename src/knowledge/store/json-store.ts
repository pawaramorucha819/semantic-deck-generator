import { mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import {
  searchEntries,
  type KnowledgeQuery,
  type KnowledgeSearchResult,
} from "../retrieval/retrieval.js";
import {
  DeckKnowledgeFileSchema,
  KNOWLEDGE_SCHEMA_VERSION,
  KnowledgeIndexSchema,
  KnowledgeValidationError,
  parseKnowledgeEntry,
  type DeckKnowledgeEntry,
  type KnowledgeIndex,
} from "../schema.js";
import type { DeckKnowledgeStore } from "./store.js";

/**
 * ローカル JSON ストア（V2 の既定）。
 *
 *   knowledge/
 *     index.json        … デッキ一覧
 *     decks/<deckId>.json … 1 デッキ分のエントリ
 *
 * Git で差分が読めることを優先した形式。
 * 外部 Vector DB は必須にしない（要件 19）。
 */

const DECK_ID_PATTERN = /^[a-zA-Z0-9][a-zA-Z0-9-_.]{0,119}$/;

export class JsonKnowledgeStore implements DeckKnowledgeStore {
  readonly id: string;
  private readonly rootDir: string;
  private cache: Map<string, DeckKnowledgeEntry[]> | undefined;

  constructor(rootDir: string) {
    this.rootDir = resolve(rootDir);
    this.id = `json:${this.rootDir}`;
  }

  private deckFile(deckId: string): string {
    if (!DECK_ID_PATTERN.test(deckId)) {
      throw new KnowledgeValidationError(`deckId が不正です: ${deckId}`, []);
    }
    return join(this.rootDir, "decks", `${deckId}.json`);
  }

  private async load(): Promise<Map<string, DeckKnowledgeEntry[]>> {
    if (this.cache) return this.cache;
    const byDeck = new Map<string, DeckKnowledgeEntry[]>();
    const decksDir = join(this.rootDir, "decks");

    let files: string[] = [];
    try {
      files = (await readdir(decksDir)).filter((f) => f.endsWith(".json")).sort();
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }

    for (const file of files) {
      const raw = JSON.parse(await readFile(join(decksDir, file), "utf8"));
      const parsed = DeckKnowledgeFileSchema.safeParse(raw);
      if (!parsed.success) {
        throw new KnowledgeValidationError(
          `Knowledge ファイルが不正です: ${file}`,
          parsed.error.issues,
        );
      }
      if (parsed.data.schemaVersion !== KNOWLEDGE_SCHEMA_VERSION) {
        throw new KnowledgeValidationError(
          `未対応の Knowledge schemaVersion です: ${file} (${parsed.data.schemaVersion})`,
          [],
        );
      }
      byDeck.set(
        parsed.data.deckId,
        parsed.data.entries.map((e) => parseKnowledgeEntry(e)),
      );
    }

    this.cache = byDeck;
    return byDeck;
  }

  private async persist(byDeck: Map<string, DeckKnowledgeEntry[]>): Promise<void> {
    await mkdir(join(this.rootDir, "decks"), { recursive: true });

    for (const [deckId, entries] of byDeck) {
      const sorted = [...entries].sort((a, b) => a.source.slideNumber - b.source.slideNumber);
      await writeFile(
        this.deckFile(deckId),
        `${JSON.stringify(
          { schemaVersion: KNOWLEDGE_SCHEMA_VERSION, deckId, entries: sorted },
          null,
          2,
        )}\n`,
        "utf8",
      );
    }

    const index: KnowledgeIndex = {
      schemaVersion: KNOWLEDGE_SCHEMA_VERSION,
      decks: [...byDeck.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([deckId, entries]) => {
          const first = entries[0];
          return {
            deckId,
            ...(first?.metadata.title ? { title: first.metadata.title } : {}),
            slideCount: entries.length,
            file: `decks/${deckId}.json`,
            ...(first?.metadata.organization ? { organization: first.metadata.organization } : {}),
            ...(first?.metadata.purpose ? { purpose: first.metadata.purpose } : {}),
            ...(first?.metadata.audience ? { audience: first.metadata.audience } : {}),
          };
        }),
    };
    KnowledgeIndexSchema.parse(index);
    await writeFile(
      join(this.rootDir, "index.json"),
      `${JSON.stringify(index, null, 2)}\n`,
      "utf8",
    );
  }

  async add(entries: readonly DeckKnowledgeEntry[]): Promise<void> {
    const byDeck = await this.load();
    for (const raw of entries) {
      const entry = parseKnowledgeEntry(raw);
      const list = byDeck.get(entry.deckId) ?? [];
      const existing = list.findIndex((e) => e.slideId === entry.slideId);
      if (existing >= 0) list[existing] = entry;
      else list.push(entry);
      byDeck.set(entry.deckId, list);
    }
    await this.persist(byDeck);
  }

  async all(): Promise<DeckKnowledgeEntry[]> {
    const byDeck = await this.load();
    return [...byDeck.values()]
      .flat()
      .sort((a, b) => a.slideId.localeCompare(b.slideId));
  }

  async get(slideId: string): Promise<DeckKnowledgeEntry | undefined> {
    return (await this.all()).find((e) => e.slideId === slideId);
  }

  async removeDeck(deckId: string): Promise<void> {
    const byDeck = await this.load();
    if (!byDeck.delete(deckId)) return;
    await this.persist(byDeck);
    // 消したデッキのファイルは空配列で残さず、索引から外れた状態にする
    await writeFile(
      this.deckFile(deckId),
      `${JSON.stringify({ schemaVersion: KNOWLEDGE_SCHEMA_VERSION, deckId, entries: [] }, null, 2)}\n`,
      "utf8",
    );
  }

  async search(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]> {
    return searchEntries(await this.all(), query);
  }

  async stats(): Promise<{ decks: number; slides: number }> {
    const byDeck = await this.load();
    return {
      decks: [...byDeck.values()].filter((entries) => entries.length > 0).length,
      slides: [...byDeck.values()].flat().length,
    };
  }

  /** ディスク上の変更を読み直す */
  invalidate(): void {
    this.cache = undefined;
  }
}
