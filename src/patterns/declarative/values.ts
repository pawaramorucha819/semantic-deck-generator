import type { SlideSpec } from "../../domain/types.js";
import type { Condition, ValueRef } from "./schema.js";

/**
 * Declarative Pattern の「値の解決」と「条件の評価」。
 *
 * ここが安全性の要。JSON から実行できるのは
 *  - 列挙された条件語彙
 *  - 安全なパスによる値参照
 * だけで、任意の式・関数は一切評価しない。
 */

export interface BindingScope {
  /** slot 名 → 値 */
  bindings: Record<string, unknown>;
  /** repeat の中で解決される現在の要素 */
  item?: Record<string, unknown>;
  /** repeat の中の 0 始まりインデックス */
  index?: number;
}

export class DeclarativeValueError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "DeclarativeValueError";
  }
}

const FORBIDDEN = new Set(["__proto__", "constructor", "prototype"]);

/** 安全なパス解決。存在しないパスは undefined、危険なキーは例外。 */
export function lookupPath(scope: BindingScope, path: string): unknown {
  const segments = path.split(".");
  if (segments.some((s) => FORBIDDEN.has(s))) {
    throw new DeclarativeValueError(`使用できないプロパティ名です: ${path}`);
  }
  const [head, ...rest] = segments as [string, ...string[]];

  let current: unknown;
  if (head === "item") {
    current = scope.item;
  } else if (head === "index") {
    current = scope.index === undefined ? undefined : scope.index + 1;
  } else {
    current = scope.bindings[head];
  }

  for (const segment of rest) {
    if (current === null || current === undefined) return undefined;
    if (typeof current !== "object") return undefined;
    current = (current as Record<string, unknown>)[segment];
  }
  return current;
}

function toText(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "string") return value.length > 0 ? value : undefined;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "true" : undefined;
  return undefined;
}

const TEMPLATE_PATTERN = /\{\{\s*([a-zA-Z][a-zA-Z0-9_.]*)\s*\}\}/g;

/**
 * 値参照を文字列へ解決する。
 * 解決できない場合は undefined を返す（呼び出し側が `when` や optional で扱う）。
 */
export function resolveValueRef(ref: ValueRef, scope: BindingScope): string | undefined {
  if ("literal" in ref) return ref.literal;
  if ("ref" in ref) return toText(lookupPath(scope, ref.ref));

  let missing = false;
  const text = ref.template.replace(TEMPLATE_PATTERN, (_m, path: string) => {
    const value = toText(lookupPath(scope, path));
    if (value === undefined) missing = true;
    return value ?? "";
  });
  if (missing) return undefined;
  return text.length > 0 ? text : undefined;
}

/** レイアウト条件（bindings スコープに対する評価） */
export function evaluateCondition(condition: Condition, scope: BindingScope): boolean {
  if ("hasValue" in condition) {
    return toText(lookupPath(scope, condition.hasValue)) !== undefined;
  }
  if ("equals" in condition) {
    return toText(lookupPath(scope, condition.equals.ref)) === condition.equals.value;
  }
  if ("indexEquals" in condition) {
    return scope.index === condition.indexEquals;
  }
  if ("itemCount" in condition) {
    const value = condition.slot
      ? lookupPath(scope, condition.slot)
      : firstArray(scope.bindings);
    const count = Array.isArray(value) ? value.length : 0;
    const { eq, min, max } = condition.itemCount;
    if (eq !== undefined && count !== eq) return false;
    if (min !== undefined && count < min) return false;
    if (max !== undefined && count > max) return false;
    return true;
  }
  if ("not" in condition) return !evaluateCondition(condition.not, scope);
  if ("all" in condition) return condition.all.every((c) => evaluateCondition(c, scope));
  if ("any" in condition) return condition.any.some((c) => evaluateCondition(c, scope));

  // スライドを必要とする条件は match 評価専用。レイアウト評価では常に true。
  return true;
}

function firstArray(bindings: Record<string, unknown>): unknown[] | undefined {
  for (const value of Object.values(bindings)) {
    if (Array.isArray(value)) return value;
  }
  return undefined;
}

/** match 評価用のコンテキスト（意味構造も参照できる） */
export interface MatchScope extends BindingScope {
  slide: SlideSpec;
}

/** match 条件（SlideSpec + bindings に対する評価） */
export function evaluateMatchCondition(condition: Condition, scope: MatchScope): boolean {
  if ("logicalRelation" in condition) {
    return condition.logicalRelation.includes(scope.slide.logicalRelation);
  }
  if ("slideRole" in condition) {
    return condition.slideRole.includes(scope.slide.slideRole);
  }
  if ("hasEvidenceType" in condition) {
    return scope.slide.evidence.some((e) => condition.hasEvidenceType.includes(e.type));
  }
  if ("keyMessageLength" in condition) {
    const length = scope.slide.keyMessage.length;
    const { min, max } = condition.keyMessageLength;
    if (min !== undefined && length < min) return false;
    if (max !== undefined && length > max) return false;
    return true;
  }
  if ("not" in condition) return !evaluateMatchCondition(condition.not, scope);
  if ("all" in condition) return condition.all.every((c) => evaluateMatchCondition(c, scope));
  if ("any" in condition) return condition.any.some((c) => evaluateMatchCondition(c, scope));
  return evaluateCondition(condition, scope);
}
