import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import PptxGenJS from "pptxgenjs";
import { afterAll, describe, expect, it } from "vitest";
import {
  extractKeywords,
  ingestDeckFromGeneration,
  ingestDeckFromPptx,
} from "../src/knowledge/ingest/ingest.js";
import {
  DeterministicRetriever,
  searchEntries,
  tokenize,
} from "../src/knowledge/retrieval/retrieval.js";
import { JsonKnowledgeStore } from "../src/knowledge/store/json-store.js";
import {
  CompositeKnowledgeStore,
  InMemoryKnowledgeStore,
} from "../src/knowledge/store/store.js";
import {
  KNOWLEDGE_SCHEMA_VERSION,
  KnowledgeValidationError,
  parseKnowledgeEntry,
  slideIdOf,
  type DeckKnowledgeEntry,
} from "../src/knowledge/schema.js";
import { generate } from "../src/pipeline.js";
import { loadExample } from "./helpers.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-knowledge-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

function entry(overrides: Partial<DeckKnowledgeEntry> = {}): DeckKnowledgeEntry {
  return {
    schemaVersion: KNOWLEDGE_SCHEMA_VERSION,
    deckId: "fy2025-strategy",
    slideId: "fy2025-strategy#1",
    metadata: { title: "FY2025 戦略", organization: "ACME", purpose: "投資判断", audience: "経営会議" },
    semantic: {
      keyMessage: "在庫管理コストの増加が経営課題である",
      slideRole: "problem",
      logicalRelation: "claim_evidence",
      keywords: ["在庫", "コスト"],
    },
    visual: { patternId: "chart-with-message", density: "medium" },
    content: { text: "在庫管理コストは3年間で年平均12%増加している" },
    source: { file: "fy2025-strategy.pptx", slideNumber: 1, origin: "pptx" },
    ...overrides,
  };
}

async function buildPastDeck(): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";

  const s1 = pptx.addSlide();
  s1.addText("在庫管理コストの増加が経営課題となっている", {
    x: 0.6, y: 0.5, w: 8.8, h: 0.6, fontSize: 22, bold: true,
  });
  s1.addText("2022年度 3,800万円 → 2024年度 4,800万円", {
    x: 0.6, y: 1.4, w: 8.8, h: 0.4, fontSize: 14,
  });

  const s2 = pptx.addSlide();
  s2.addText("投資回収は2年以内に見込める", { x: 0.6, y: 0.5, w: 8.8, h: 0.6, fontSize: 22, bold: true });
  s2.addText("初期投資2,400万円に対し年間1,440万円の削減効果", {
    x: 0.6, y: 1.4, w: 8.8, h: 0.4, fontSize: 14,
  });
  s2.addNotes("経営会議向けの投資判断スライド");

  return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
}

describe("Knowledge スキーマ", () => {
  it("正しいエントリを受け入れる", () => {
    expect(parseKnowledgeEntry(entry()).slideId).toBe("fy2025-strategy#1");
  });

  it("未知のキーを拒否する", () => {
    expect(() => parseKnowledgeEntry({ ...entry(), patternId: "x" })).toThrow(
      KnowledgeValidationError,
    );
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() => parseKnowledgeEntry({ ...entry(), schemaVersion: "9.9" })).toThrow(
      /schemaVersion/,
    );
  });

  it("不正な slideRole を拒否する", () => {
    expect(() =>
      parseKnowledgeEntry({ ...entry(), semantic: { slideRole: "nonsense" } }),
    ).toThrow(KnowledgeValidationError);
  });

  it("slideIdOf は deckId とスライド番号から一意な ID を作る", () => {
    expect(slideIdOf("deck-a", 3)).toBe("deck-a#3");
  });
});

describe("Ingest", () => {
  it("過去 PPTX を 1 スライド 1 エントリで取り込む", async () => {
    const result = await ingestDeckFromPptx(await buildPastDeck(), {
      deckId: "past-deck",
      metadata: { organization: "ACME", purpose: "投資判断", audience: "経営会議" },
      sourceLabel: "past-deck.pptx",
      importedAt: "2026-01-01",
    });

    expect(result.entries).toHaveLength(2);
    expect(result.entries[0]!.slideId).toBe("past-deck#1");
    expect(result.entries[0]!.semantic?.keyMessage).toContain("在庫管理コスト");
    expect(result.entries[0]!.semantic?.keywords?.length).toBeGreaterThan(0);
    expect(result.entries[0]!.content?.text).toContain("3,800万円");
    expect(result.entries[0]!.source).toMatchObject({
      file: "past-deck.pptx",
      slideNumber: 1,
      origin: "pptx",
    });
  });

  it("PPTX からは意味（役割・論理関係）が決まらないことを警告する", async () => {
    const result = await ingestDeckFromPptx(await buildPastDeck(), { deckId: "past-deck" });
    expect(result.report.warnings.join()).toContain("slideRole / logicalRelation");
    expect(result.entries[0]!.semantic?.slideRole).toBeUndefined();
  });

  it("アノテーションで意味を補える", async () => {
    const result = await ingestDeckFromPptx(await buildPastDeck(), {
      deckId: "past-deck",
      annotations: {
        1: { slideRole: "problem", logicalRelation: "claim_evidence" },
        2: { slideRole: "recommendation", logicalRelation: "cause_effect", patternId: "big-number" },
      },
    });
    expect(result.entries[0]!.semantic?.slideRole).toBe("problem");
    expect(result.entries[1]!.visual?.patternId).toBe("big-number");
    expect(result.report.annotated).toBe(2);
  });

  it("自分で生成したデッキは意味も Pattern も含めて取り込める（知識が循環する）", async () => {
    const spec = await loadExample();
    const generated = await generate(spec);
    const result = ingestDeckFromGeneration(spec, generated.plan, {
      deckId: "inventory-2025",
      metadata: { purpose: "投資判断", audience: "経営会議" },
    });

    expect(result.entries).toHaveLength(6);
    const first = result.entries[0]!;
    expect(first.semantic?.slideRole).toBe("problem");
    expect(first.semantic?.logicalRelation).toBe("claim_evidence");
    expect(first.visual?.patternId).toBe("chart-with-message");
    expect(first.source.origin).toBe("generation");
  });

  it("同じ入力からは同じエントリになる（決定論的）", async () => {
    const buffer = await buildPastDeck();
    const a = await ingestDeckFromPptx(buffer, { deckId: "d" });
    const b = await ingestDeckFromPptx(buffer, { deckId: "d" });
    expect(b.entries).toEqual(a.entries);
  });

  it("キーワード抽出は頻度順で決定論的", () => {
    const keywords = extractKeywords("在庫 在庫 コスト cost cost cost", 3);
    expect(keywords[0]).toBe("cost");
    expect(keywords).toEqual(extractKeywords("在庫 在庫 コスト cost cost cost", 3));
  });
});

describe("Store", () => {
  it("インメモリストアで追加・取得・削除できる", async () => {
    const store = new InMemoryKnowledgeStore();
    await store.add([entry(), entry({ slideId: "fy2025-strategy#2", source: { file: "f", slideNumber: 2 } })]);
    expect(await store.stats()).toEqual({ decks: 1, slides: 2 });
    expect((await store.get("fy2025-strategy#2"))?.slideId).toBe("fy2025-strategy#2");
    await store.removeDeck("fy2025-strategy");
    expect(await store.stats()).toEqual({ decks: 0, slides: 0 });
  });

  it("JSON ストアはローカルに永続化され、読み直せる", async () => {
    const dir = await makeTempDir();
    const store = new JsonKnowledgeStore(dir);
    await store.add([
      entry(),
      entry({ slideId: "fy2025-strategy#2", source: { file: "f", slideNumber: 2 } }),
    ]);

    const index = JSON.parse(await readFile(join(dir, "index.json"), "utf8"));
    expect(index.decks[0]).toMatchObject({ deckId: "fy2025-strategy", slideCount: 2 });

    const reopened = new JsonKnowledgeStore(dir);
    expect(await reopened.stats()).toEqual({ decks: 1, slides: 2 });
    expect((await reopened.all()).map((e) => e.slideId)).toEqual([
      "fy2025-strategy#1",
      "fy2025-strategy#2",
    ]);
  });

  it("同じ slideId を追加すると上書きされる（重複しない）", async () => {
    const dir = await makeTempDir();
    const store = new JsonKnowledgeStore(dir);
    await store.add([entry()]);
    await store.add([entry({ semantic: { keyMessage: "更新後のメッセージ" } })]);
    const all = await store.all();
    expect(all).toHaveLength(1);
    expect(all[0]!.semantic?.keyMessage).toBe("更新後のメッセージ");
  });

  it("壊れた Knowledge ファイルは読み込み時に拒否される", async () => {
    const dir = await makeTempDir();
    const store = new JsonKnowledgeStore(dir);
    await store.add([entry()]);

    const { writeFile } = await import("node:fs/promises");
    await writeFile(
      join(dir, "decks", "fy2025-strategy.json"),
      JSON.stringify({ schemaVersion: "1.0", deckId: "fy2025-strategy", entries: [{ bad: true }] }),
      "utf8",
    );
    const reopened = new JsonKnowledgeStore(dir);
    await expect(reopened.all()).rejects.toThrow(KnowledgeValidationError);
  });

  it("不正な deckId でファイルを書き出そうとしない（パス操作対策）", async () => {
    const dir = await makeTempDir();
    const store = new JsonKnowledgeStore(dir);
    await expect(store.add([entry({ deckId: "../../escape" })])).rejects.toThrow();
  });

  it("複数ストアを束ねて検索できる", async () => {
    const company = new InMemoryKnowledgeStore([entry()]);
    const project = new InMemoryKnowledgeStore([
      entry({ deckId: "project-x", slideId: "project-x#1", source: { file: "p", slideNumber: 1 } }),
    ]);
    const composite = new CompositeKnowledgeStore([project, company]);
    expect(await composite.stats()).toEqual({ decks: 2, slides: 2 });
  });
});

describe("決定論的 Retrieval（AI 無しで動く）", () => {
  const entries: DeckKnowledgeEntry[] = [
    entry(),
    entry({
      slideId: "fy2025-strategy#2",
      semantic: {
        keyMessage: "投資回収は2年以内に見込める",
        slideRole: "recommendation",
        logicalRelation: "cause_effect",
        keywords: ["投資", "回収"],
      },
      visual: { patternId: "big-number" },
      content: { text: "初期投資2,400万円に対し年間1,440万円の削減効果" },
      source: { file: "fy2025-strategy.pptx", slideNumber: 2 },
    }),
    entry({
      deckId: "fy2026-budget",
      slideId: "fy2026-budget#9",
      metadata: { title: "FY2026 予算", organization: "ACME", purpose: "予算承認", audience: "部門長" },
      semantic: {
        keyMessage: "3つの施策で費用を圧縮する",
        slideRole: "recommendation",
        logicalRelation: "claim_evidence",
        keywords: ["施策", "費用"],
      },
      visual: { patternId: "three-reasons" },
      content: { text: "3つの施策により費用を圧縮できる見込みである" },
      source: { file: "fy2026-budget.pptx", slideNumber: 9 },
    }),
  ];

  it("slideRole と logicalRelation で絞り込める", () => {
    const results = searchEntries(entries, {
      slideRole: "recommendation",
      logicalRelation: "claim_evidence",
    });
    expect(results[0]!.entry.slideId).toBe("fy2026-budget#9");
    expect(results[0]!.matched).toContain("slideRole=recommendation");
  });

  it("strict 指定でメタデータ条件を必須にできる", () => {
    const loose = searchEntries(entries, { slideRole: "recommendation", purpose: "予算承認" });
    const strict = searchEntries(entries, {
      slideRole: "recommendation",
      purpose: "予算承認",
      strict: true,
    });
    expect(loose.length).toBeGreaterThan(strict.length);
    expect(strict.every((r) => r.entry.metadata.purpose === "予算承認")).toBe(true);
  });

  it("日本語のテキスト類似度で似たスライドを引ける（Embedding API 不要）", () => {
    const results = searchEntries(entries, { text: "投資の回収期間はどれくらいか" });
    expect(results[0]!.entry.slideId).toBe("fy2025-strategy#2");
    expect(results[0]!.matched.join()).toContain("textSimilarity");
  });

  it("patternId で過去の利用実績を引ける", () => {
    const results = searchEntries(entries, { patternId: "three-reasons" });
    expect(results[0]!.entry.visual?.patternId).toBe("three-reasons");
  });

  it("同じクエリからは常に同じ順序が返る（決定論的）", () => {
    const query = { slideRole: "recommendation" as const, text: "費用" };
    expect(searchEntries(entries, query)).toEqual(searchEntries(entries, query));
  });

  it("limit で件数を制御できる", () => {
    expect(searchEntries(entries, { organization: "ACME", limit: 2 })).toHaveLength(2);
  });

  it("一致が無ければ空配列（無理に返さない）", () => {
    expect(searchEntries(entries, { slideRole: "timeline" as never, strict: true })).toEqual([]);
  });

  it("Retriever 経由でも同じ結果になる（将来 Embedding へ差し替え可能）", async () => {
    const store = new InMemoryKnowledgeStore(entries);
    const retriever = new DeterministicRetriever(() => store.all());
    const viaRetriever = await retriever.retrieve({ slideRole: "recommendation" });
    const viaStore = await store.search({ slideRole: "recommendation" });
    expect(viaRetriever.map((r) => r.entry.slideId)).toEqual(viaStore.map((r) => r.entry.slideId));
    expect(retriever.id).toBe("deterministic");
  });

  it("トークナイザは日英混在を扱える", () => {
    const tokens = tokenize("在庫 cost 12%");
    expect(tokens).toContain("cost");
    expect(tokens.some((t) => t.length === 2 && /[一-龯ぁ-ヿ]/.test(t))).toBe(true);
  });
});
