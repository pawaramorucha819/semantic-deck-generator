import { createHash } from "node:crypto";

/**
 * deckId の生成と検証。
 *
 * 企業利用では「経営会議_生成AI戦略.pptx」のような日本語のファイル名・資料名が普通なので、
 * ASCII 前提の置換だけで deckId を作ると
 *  - 先頭が '-' になる / ほぼ全部 '-' になる
 *  - Knowledge スキーマの deckId 制約を満たさない
 *  - 別々の資料が同じ ID に潰れる
 * といった問題が起きる。ここで一箇所に集約して安全化する。
 *
 * 方針: 依存を増やさず、決定論的で、必ず有効な ID を返す。
 * ローマ字変換のような重い正規化はしない（落ちた情報はハッシュで区別する）。
 */

/** deckId として許される形。Knowledge スキーマと JSON ストアが共有する。 */
export const DECK_ID_PATTERN = /^[a-zA-Z0-9][a-zA-Z0-9-_.]{0,119}$/;

export const MAX_DECK_ID_LENGTH = 120;

/** 情報が落ちたときに付ける決定論的ハッシュの長さ */
const HASH_LENGTH = 8;
const MIN_DECK_ID_LENGTH = 16;
const FALLBACK_PREFIX = "deck";

export function isValidDeckId(value: string): boolean {
  return DECK_ID_PATTERN.test(value);
}

/** 同じ入力からは常に同じ値になる短いハッシュ（timestamp / random は使わない） */
function shortHash(input: string): string {
  return createHash("sha256").update(input, "utf8").digest("hex").slice(0, HASH_LENGTH);
}

/** 末尾の区切り文字を落としながら切り詰める */
function truncate(value: string, maxLength: number): string {
  if (value.length <= maxLength) return value;
  return value.slice(0, maxLength).replace(/[-_.]+$/, "");
}

export interface CreateSafeDeckIdOptions {
  /** 英数字が 1 文字も残らなかった場合に使う接頭辞（既定 'deck'） */
  fallbackPrefix?: string;
  /** 生成する ID の最大長（既定 120。16 未満は 16 に切り上げ） */
  maxLength?: number;
}

/**
 * 任意の文字列（ファイル名・資料タイトル）から安全な deckId を作る。
 *
 * 保証すること:
 *  - 戻り値は必ず `DECK_ID_PATTERN` を満たす（空にならない / 先頭は英数字）
 *  - 同じ入力からは必ず同じ ID になる（決定論的）
 *  - 日本語だけの入力でも有効な ID になる
 *  - ASCII 英数字だけで構成された名前はそのまま残る
 *
 * 例:
 *   "FY2026-strategy"     → "FY2026-strategy"     （情報が落ちないのでそのまま）
 *   "AI戦略_日本版"        → "AI-<hash>"           （残った ASCII + 衝突回避のハッシュ）
 *   "経営会議_生成AI戦略"  → "AI-<hash>"
 *   "決算説明会"           → "deck-<hash>"         （英数字が残らない場合）
 */
export function createSafeDeckId(input: string, options: CreateSafeDeckIdOptions = {}): string {
  const maxLength = Math.min(
    MAX_DECK_ID_LENGTH,
    Math.max(MIN_DECK_ID_LENGTH, options.maxLength ?? MAX_DECK_ID_LENGTH),
  );

  // 全角英数字（ＦＹ２０２６）は半角へ寄せてから扱う
  const normalized = input.normalize("NFKC").trim();
  const hash = shortHash(normalized);

  // 非 ASCII（日本語など）を含む場合、slug 化で意味のある情報が落ちる。
  // 別々の資料が同じ ID に潰れないよう決定論的ハッシュを付ける。
  const lossy = /[^\x20-\x7E]/.test(normalized);

  const slug = normalized
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^[-_.]+/, "")
    .replace(/[-_.]+$/, "");

  const candidate = !/[a-zA-Z0-9]/.test(slug)
    ? `${sanitizePrefix(options.fallbackPrefix)}-${hash}`
    : lossy
      ? `${truncate(slug, maxLength - HASH_LENGTH - 1)}-${hash}`
      : truncate(slug, maxLength);

  // 最後の安全網: どんな入力でも必ず有効な ID を返す
  return isValidDeckId(candidate) ? candidate : `${FALLBACK_PREFIX}-${hash}`;
}

function sanitizePrefix(prefix: string | undefined): string {
  if (!prefix) return FALLBACK_PREFIX;
  const cleaned = prefix
    .normalize("NFKC")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^[-_.]+/, "")
    .replace(/[-_.]+$/, "");
  return /^[a-zA-Z0-9]/.test(cleaned) ? cleaned.slice(0, 40) : FALLBACK_PREFIX;
}
