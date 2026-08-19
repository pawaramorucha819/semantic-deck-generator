import { LOGICAL_RELATIONS } from "../domain/types.js";
import { beforeAfterPattern } from "./before-after/index.js";
import { bigNumberPattern } from "./big-number/index.js";
import { chartWithMessagePattern } from "./chart-with-message/index.js";
import { messageOnlyPattern } from "./message-only/index.js";
import type { AnyVisualPattern } from "./pattern.js";
import { processPattern } from "./process/index.js";
import { threeReasonsPattern } from "./three-reasons/index.js";

/**
 * Pattern レジストリ。
 * Pattern を追加するときは実装ファイルを作り、この配列へ 1 行足すだけ。
 * Selector / Renderer / CLI のいずれも変更不要。
 *
 * 配列の順序はスコア同点時のタイブレークにしか影響しない（決定論性のため）。
 */
export const builtinPatterns: AnyVisualPattern[] = [
  beforeAfterPattern,
  threeReasonsPattern,
  processPattern,
  bigNumberPattern,
  chartWithMessagePattern,
  messageOnlyPattern,
];

export class PatternRegistry {
  private readonly byId = new Map<string, AnyVisualPattern>();

  constructor(patterns: AnyVisualPattern[] = builtinPatterns) {
    for (const pattern of patterns) {
      if (this.byId.has(pattern.id)) {
        throw new Error(`Pattern id が重複しています: ${pattern.id}`);
      }
      if (pattern.fallback) {
        // fallback は「常に描画できる緊急表現」なので、自分を失格させうる条件を持てない
        if (pattern.constraints) {
          throw new Error(
            `fallback Pattern '${pattern.id}' は constraints を持てません（どんなスライドでも描画できる必要があります）`,
          );
        }
        const missing = LOGICAL_RELATIONS.filter(
          (r) => !pattern.compatibility.logicalRelations.includes(r),
        );
        if (missing.length > 0) {
          throw new Error(
            `fallback Pattern '${pattern.id}' は全 logicalRelation に対応する必要があります (不足: ${missing.join(", ")})`,
          );
        }
      }
      this.byId.set(pattern.id, pattern);
    }
    if (this.all().every((p) => !p.fallback)) {
      throw new Error("fallback: true の Pattern が 1 つも登録されていません");
    }
  }

  all(): AnyVisualPattern[] {
    return [...this.byId.values()];
  }

  get(id: string): AnyVisualPattern {
    const pattern = this.byId.get(id);
    if (!pattern) {
      throw new Error(
        `未登録の patternId です: ${id} (登録済み: ${[...this.byId.keys()].join(", ")})`,
      );
    }
    return pattern;
  }

  has(id: string): boolean {
    return this.byId.has(id);
  }

  fallbacks(): AnyVisualPattern[] {
    return this.all().filter((p) => p.fallback);
  }
}

export const defaultRegistry = new PatternRegistry();
