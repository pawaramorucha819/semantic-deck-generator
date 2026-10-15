import { LOGICAL_RELATIONS } from "../domain/types.js";
import type { LoadedTemplateAsset } from "../templates/template.js";
import { beforeAfterPattern } from "./before-after/index.js";
import { bigNumberPattern } from "./big-number/index.js";
import { chartWithMessagePattern } from "./chart-with-message/index.js";
import type { LoadedPatternPackage } from "./declarative/loader.js";
import { loadPatternPackages, type LoadPatternPackageOptions } from "./declarative/loader.js";
import type { PatternScope } from "./declarative/schema.js";
import { messageOnlyPattern } from "./message-only/index.js";
import type { AnyVisualPattern } from "./pattern.js";
import { processPattern } from "./process/index.js";
import { relationshipPatterns } from "./relationships/index.js";
import { threeReasonsPattern } from "./three-reasons/index.js";

/**
 * Pattern レジストリ。
 *
 * コード Pattern（builtin）と Declarative Pattern（company / team / project / user）を
 * 同じ入口から扱う。scope は「どこから来た Pattern か」であり、
 * 同スコア時のタイブレークにのみ使われる。
 *
 * **重要**: scope の優先度で *意味的に不適合な Pattern* が選ばれることはない。
 * compatibility / constraints / match のハードゲートは scope より常に優先される
 * （src/selector/selector.ts を参照）。
 */

export type { PatternScope };

/** scope の優先度。大きいほど優先（同点時のみ効く）。 */
export const SCOPE_PRIORITY: Record<PatternScope, number> = {
  builtin: 0,
  company: 10,
  team: 20,
  user: 30,
  project: 40,
};

export interface PatternRegistration {
  pattern: AnyVisualPattern;
  scope?: PatternScope;
  /** 読み込み元（ディレクトリなど。デバッグ・一覧表示用） */
  source?: string;
  /** 明示的な優先度。省略時は scope から決まる。 */
  priority?: number;
  /** Declarative Pattern パッケージの付随情報 */
  assets?: LoadedTemplateAsset[];
  previewImagePath?: string;
}

export interface PatternRegistryEntry {
  pattern: AnyVisualPattern;
  scope: PatternScope;
  priority: number;
  source?: string;
  assets: LoadedTemplateAsset[];
  previewImagePath?: string;
  /** 登録順（同点タイブレークの最終手段。決定論性のため） */
  index: number;
}

/**
 * ビルトイン Pattern。
 * 配列の順序はスコア同点時のタイブレークにしか影響しない（決定論性のため）。
 */
export const builtinPatterns: AnyVisualPattern[] = [
  ...relationshipPatterns,
  beforeAfterPattern,
  threeReasonsPattern,
  processPattern,
  bigNumberPattern,
  chartWithMessagePattern,
  messageOnlyPattern,
];

function toRegistration(input: AnyVisualPattern | PatternRegistration): PatternRegistration {
  return "pattern" in input ? input : { pattern: input };
}

export class PatternRegistry {
  private readonly byId = new Map<string, PatternRegistryEntry>();
  private counter = 0;

  constructor(patterns: Array<AnyVisualPattern | PatternRegistration> = builtinPatterns) {
    for (const item of patterns) this.register(toRegistration(item));
    // 初期セットには必ず fallback が要る（後から register で追加する場合を除く）
    if (patterns.length > 0 && this.all().every((p) => !p.fallback)) {
      throw new Error("fallback: true の Pattern が 1 つも登録されていません");
    }
  }

  register(input: AnyVisualPattern | PatternRegistration): this {
    const registration = toRegistration(input);
    const pattern = registration.pattern;

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

    const scope = registration.scope ?? "builtin";
    this.byId.set(pattern.id, {
      pattern,
      scope,
      priority: registration.priority ?? SCOPE_PRIORITY[scope],
      ...(registration.source ? { source: registration.source } : {}),
      assets: registration.assets ?? [],
      ...(registration.previewImagePath ? { previewImagePath: registration.previewImagePath } : {}),
      index: this.counter++,
    });
    return this;
  }

  /** Declarative Pattern パッケージを登録する */
  registerPackage(pkg: LoadedPatternPackage): this {
    return this.register({
      pattern: pkg.pattern,
      scope: pkg.scope,
      source: pkg.dir,
      assets: pkg.assets,
      ...(pkg.previewImagePath ? { previewImagePath: pkg.previewImagePath } : {}),
    });
  }

  /** ディレクトリ配下の Declarative Pattern をまとめて登録する */
  async registerDirectory(
    rootDir: string,
    options: LoadPatternPackageOptions = {},
  ): Promise<LoadedPatternPackage[]> {
    const packages = await loadPatternPackages(rootDir, options);
    for (const pkg of packages) this.registerPackage(pkg);
    return packages;
  }

  all(): AnyVisualPattern[] {
    return this.entries().map((e) => e.pattern);
  }

  /** 登録順のエントリ（scope / priority 付き） */
  entries(): PatternRegistryEntry[] {
    return [...this.byId.values()].sort((a, b) => a.index - b.index);
  }

  entry(id: string): PatternRegistryEntry {
    const entry = this.byId.get(id);
    if (!entry) {
      throw new Error(
        `未登録の patternId です: ${id} (登録済み: ${[...this.byId.keys()].join(", ")})`,
      );
    }
    return entry;
  }

  get(id: string): AnyVisualPattern {
    return this.entry(id).pattern;
  }

  has(id: string): boolean {
    return this.byId.has(id);
  }

  scopeOf(id: string): PatternScope {
    return this.entry(id).scope;
  }

  priorityOf(id: string): number {
    return this.entry(id).priority;
  }

  fallbacks(): AnyVisualPattern[] {
    return this.all().filter((p) => p.fallback);
  }

  /** 全 Declarative Pattern パッケージのアセット（名前空間化済み） */
  assets(): LoadedTemplateAsset[] {
    return this.entries().flatMap((e) => e.assets);
  }

  /** 一覧表示用の要約 */
  list(): Array<{
    id: string;
    version: string;
    scope: PatternScope;
    priority: number;
    kind: "code" | "declarative";
    description: string;
    variants: readonly string[];
    source?: string;
    previewImagePath?: string;
  }> {
    return this.entries().map((entry) => ({
      id: entry.pattern.id,
      version: entry.pattern.version,
      scope: entry.scope,
      priority: entry.priority,
      kind: "document" in entry.pattern ? ("declarative" as const) : ("code" as const),
      description: entry.pattern.description,
      variants: entry.pattern.variants,
      ...(entry.source ? { source: entry.source } : {}),
      ...(entry.previewImagePath ? { previewImagePath: entry.previewImagePath } : {}),
    }));
  }
}

/** builtin だけを含む既定レジストリ */
export const defaultRegistry = new PatternRegistry();
