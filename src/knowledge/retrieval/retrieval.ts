import type { LogicalRelation, SlideRole } from "../../domain/types.js";
import type { DeckKnowledgeEntry } from "../schema.js";

/**
 * Retrieval — 過去スライドの検索。
 *
 * V2 の既定は **外部 API を一切使わない決定論的スコアリング**。
 * メタデータ一致 + トークン重なりだけで動くので、
 * 「Embedding API が無いと過去資料を引けない」状態にはならない。
 *
 * 将来 Embedding 検索へ差し替えられるよう、入口は `KnowledgeRetriever` に揃えてある。
 */

export interface KnowledgeQuery {
  /** 自由文（KeyMessage など）。トークン重なりで効く。 */
  text?: string;
  keywords?: string[];
  slideRole?: SlideRole;
  logicalRelation?: LogicalRelation;
  purpose?: string;
  audience?: string;
  organization?: string;
  patternId?: string;
  templateId?: string;
  /** 上位何件を返すか（既定 5） */
  limit?: number;
  /**
   * true にすると、指定したメタデータ条件に一致しないエントリを除外する
   * （既定 false = 加点のみ）。
   */
  strict?: boolean;
}

export interface KnowledgeSearchResult {
  entry: DeckKnowledgeEntry;
  /** 0.0 - 1.0 */
  score: number;
  /** なぜ一致したか（説明可能性のため必ず残す） */
  matched: string[];
}

export interface KnowledgeRetriever {
  readonly id: string;
  retrieve(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]>;
}

/** メタデータ一致の重み。合計 0.6、残り 0.4 をテキスト類似度に割り当てる。 */
const WEIGHTS = {
  slideRole: 0.16,
  logicalRelation: 0.16,
  purpose: 0.08,
  audience: 0.08,
  patternId: 0.06,
  templateId: 0.03,
  organization: 0.03,
  text: 0.4,
} as const;

const CJK = /[぀-ヿ㐀-䶿一-鿿豈-﫿]/;

/**
 * 日本語と英語が混在する文字列のトークン化。
 *  - ASCII 系: 小文字化した単語
 *  - CJK: 文字バイグラム（形態素解析なしで実用的な重なりが取れる）
 */
export function tokenize(text: string): string[] {
  const tokens: string[] = [];
  const normalized = text.normalize("NFKC").toLowerCase();
  let ascii = "";
  let cjk = "";

  const flushAscii = (): void => {
    if (ascii.length >= 2) tokens.push(ascii);
    ascii = "";
  };
  const flushCjk = (): void => {
    if (cjk.length === 1) tokens.push(cjk);
    for (let i = 0; i + 1 < cjk.length; i += 1) tokens.push(cjk.slice(i, i + 2));
    cjk = "";
  };

  for (const ch of normalized) {
    if (CJK.test(ch)) {
      flushAscii();
      cjk += ch;
      continue;
    }
    if (/[a-z0-9]/.test(ch)) {
      flushCjk();
      ascii += ch;
      continue;
    }
    flushAscii();
    flushCjk();
  }
  flushAscii();
  flushCjk();
  return tokens;
}

function entryText(entry: DeckKnowledgeEntry): string {
  return [
    entry.semantic?.keyMessage ?? "",
    (entry.semantic?.keywords ?? []).join(" "),
    entry.metadata.title ?? "",
    entry.content?.text ?? "",
  ].join(" ");
}

/**
 * トークンの重なり具合（0-1）。
 * 語彙頻度の逆数で重み付けし、どのスライドにも出る語の影響を抑える。
 */
export function textSimilarity(
  queryTokens: string[],
  entryTokens: string[],
  documentFrequency: Map<string, number>,
  totalDocuments: number,
): number {
  if (queryTokens.length === 0 || entryTokens.length === 0) return 0;
  const entrySet = new Set(entryTokens);
  let matched = 0;
  let total = 0;
  for (const token of new Set(queryTokens)) {
    const df = documentFrequency.get(token) ?? 0;
    const weight = Math.log(1 + totalDocuments / (1 + df));
    total += weight;
    if (entrySet.has(token)) matched += weight;
  }
  return total > 0 ? matched / total : 0;
}

const round3 = (n: number): number => Number(n.toFixed(3));

/**
 * 決定論的スコアリング。
 * 同じ入力からは常に同じ順序が返る（同点は slideId 昇順で決める）。
 */
export function searchEntries(
  entries: readonly DeckKnowledgeEntry[],
  query: KnowledgeQuery,
): KnowledgeSearchResult[] {
  const tokenized = entries.map((entry) => tokenize(entryText(entry)));
  const documentFrequency = new Map<string, number>();
  for (const tokens of tokenized) {
    for (const token of new Set(tokens)) {
      documentFrequency.set(token, (documentFrequency.get(token) ?? 0) + 1);
    }
  }

  const queryTokens = tokenize([query.text ?? "", ...(query.keywords ?? [])].join(" "));

  const results: KnowledgeSearchResult[] = [];
  entries.forEach((entry, index) => {
    const matched: string[] = [];
    let score = 0;
    let excluded = false;

    const check = (
      condition: boolean | undefined,
      requested: unknown,
      weight: number,
      label: string,
    ): void => {
      if (requested === undefined) return;
      if (condition) {
        score += weight;
        matched.push(label);
      } else if (query.strict) {
        excluded = true;
      }
    };

    check(
      entry.semantic?.slideRole === query.slideRole,
      query.slideRole,
      WEIGHTS.slideRole,
      `slideRole=${query.slideRole}`,
    );
    check(
      entry.semantic?.logicalRelation === query.logicalRelation,
      query.logicalRelation,
      WEIGHTS.logicalRelation,
      `logicalRelation=${query.logicalRelation}`,
    );
    check(
      entry.metadata.purpose === query.purpose,
      query.purpose,
      WEIGHTS.purpose,
      `purpose=${query.purpose}`,
    );
    check(
      entry.metadata.audience === query.audience,
      query.audience,
      WEIGHTS.audience,
      `audience=${query.audience}`,
    );
    check(
      entry.visual?.patternId === query.patternId,
      query.patternId,
      WEIGHTS.patternId,
      `patternId=${query.patternId}`,
    );
    check(
      entry.visual?.templateId === query.templateId,
      query.templateId,
      WEIGHTS.templateId,
      `templateId=${query.templateId}`,
    );
    check(
      entry.metadata.organization === query.organization,
      query.organization,
      WEIGHTS.organization,
      `organization=${query.organization}`,
    );

    if (excluded) return;

    if (queryTokens.length > 0) {
      const similarity = textSimilarity(
        queryTokens,
        tokenized[index]!,
        documentFrequency,
        entries.length,
      );
      if (similarity > 0) {
        score += similarity * WEIGHTS.text;
        matched.push(`textSimilarity=${similarity.toFixed(2)}`);
      }
    }

    if (score <= 0) return;
    results.push({ entry, score: round3(Math.min(1, score)), matched });
  });

  results.sort((a, b) => b.score - a.score || a.entry.slideId.localeCompare(b.entry.slideId));
  return results.slice(0, query.limit ?? 5);
}

/** ストアを包んで `KnowledgeRetriever` にする（既定の決定論的実装） */
export class DeterministicRetriever implements KnowledgeRetriever {
  readonly id = "deterministic";

  constructor(private readonly load: () => Promise<readonly DeckKnowledgeEntry[]>) {}

  async retrieve(query: KnowledgeQuery): Promise<KnowledgeSearchResult[]> {
    return searchEntries(await this.load(), query);
  }
}
