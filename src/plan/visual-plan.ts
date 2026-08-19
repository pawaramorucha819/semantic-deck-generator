import { z } from "zod";
import { SourceSchema } from "../domain/schema.js";
import type { PresentationSpec, SlideSpec, Source } from "../domain/types.js";
import type { AnyVisualPattern } from "../patterns/pattern.js";
import { defaultRegistry, PatternRegistry } from "../patterns/registry.js";
import { selectPattern, type PatternEvaluation } from "../selector/selector.js";

export const VISUAL_PLAN_SCHEMA_VERSION = "1.1";

/**
 * Provenance — この 1 枚が依拠している出典と根拠の構造化された記録。
 *
 * 表示用の短い出典行（bindings.sourceText）とは別に保持する。
 * Renderer はこれをスピーカーノートへ完全な形（URL・参照箇所込み）で書き出し、
 * 将来の引用検証・AI レビュー・外部データ更新の入力になる。
 */
export interface SlideProvenance {
  sources: Source[];
  evidence: Array<{
    id: string;
    type: string;
    statement: string;
    sourceIds?: string[];
  }>;
}

export interface VisualPlanSlide {
  slideId: string;
  slideNumber: number;
  patternId: string;
  patternVersion: string;
  variantId: string;
  /** Selector が付けた最終スコア */
  score: number;
  /** なぜこの Pattern が選ばれたか */
  reason: string;
  fallbackUsed: boolean;
  /**
   * Pattern 固有の描画用データ。Renderer が見るのはここだけ。
   * 形は各 Pattern の `bindingsSchema` が定義し、生成時・再生時の両方で検証される。
   */
  bindings: unknown;
  /** 出典・根拠の構造化された記録（→ スピーカーノート / 検証用） */
  provenance?: SlideProvenance;
  /** デバッグ用: 全 Pattern の評価結果（スコア降順） */
  candidates?: PatternEvaluation[];
  /** buildBindings 失敗などの注記 */
  notes?: string[];
}

export interface VisualPlan {
  schemaVersion: string;
  title: string;
  subtitle?: string;
  author?: string;
  themeId: string;
  slides: VisualPlanSlide[];
}

export interface BuildVisualPlanOptions {
  registry?: PatternRegistry;
  themeId?: string;
  /** 候補一覧を Visual Plan に含めるか（既定 true） */
  includeCandidates?: boolean;
  minimumScore?: number;
}

/**
 * presentation.json（意味） → Visual Plan（表現）
 *
 * この関数を通過した時点で「どの Pattern をどの props で呼ぶか」が確定し、
 * 以降 Renderer は SlideSpec を一切参照しない。
 * 生成された bindings はその場で Pattern の bindingsSchema により検証される。
 */
export function buildVisualPlan(
  spec: PresentationSpec,
  options: BuildVisualPlanOptions = {},
): VisualPlan {
  const registry = options.registry ?? defaultRegistry;
  const includeCandidates = options.includeCandidates ?? true;

  const slides = spec.slides.map((slide, index) =>
    planSlide(slide, index + 1, registry, includeCandidates, options.minimumScore),
  );

  return {
    schemaVersion: VISUAL_PLAN_SCHEMA_VERSION,
    title: spec.title,
    subtitle: spec.subtitle,
    author: spec.author,
    themeId: options.themeId ?? "default",
    slides,
  };
}

function buildProvenance(slide: SlideSpec): SlideProvenance | undefined {
  const sources = slide.sources ?? [];
  if (sources.length === 0 && slide.evidence.length === 0) return undefined;
  return {
    sources,
    evidence: slide.evidence.map((e) => ({
      id: e.id,
      type: e.type,
      statement: e.statement,
      ...(e.sourceIds ? { sourceIds: e.sourceIds } : {}),
    })),
  };
}

/**
 * Pattern が実際に実装している variant を確定する。
 *
 * buildBindings より **前** に呼ぶこと。未実装の variant のまま bindings を作ると、
 * Plan に記録される variantId と bindings が食い違う（= 再生時に別の見た目になる）。
 */
function resolveVariant(
  pattern: AnyVisualPattern,
  requested: string,
  notes: string[],
): string {
  if (pattern.variants.includes(requested)) return requested;
  const fallbackVariant = pattern.variants[0] ?? "standard";
  notes.push(
    `Pattern '${pattern.id}' に未実装の variantId '${requested}' が指定されたため '${fallbackVariant}' を使用`,
  );
  return fallbackVariant;
}

function planSlide(
  slide: SlideSpec,
  slideNumber: number,
  registry: PatternRegistry,
  includeCandidates: boolean,
  minimumScore: number | undefined,
): VisualPlanSlide {
  const selection = selectPattern(slide, { registry, minimumScore });
  const notes: string[] = [];

  let patternId = selection.patternId;
  let fallbackUsed = selection.fallbackUsed;
  let pattern = registry.get(patternId);

  // variant を先に確定してから bindings を作る。
  // これにより buildBindings には必ず実装済みの variant だけが渡り、
  // Plan の variantId と bindings が常に一致する。
  let variantId = resolveVariant(pattern, selection.variantId, notes);
  let bindings: unknown;

  try {
    // 生成した bindings をその場で Pattern の契約（bindingsSchema）に通す。
    // Pattern 実装のバグ（schema と実装のずれ）をここで検出する。
    bindings = pattern.bindingsSchema.parse(pattern.buildBindings(slide, variantId));
  } catch (error) {
    // Pattern 側の想定外（match は通ったが bindings を作れない/契約違反）→ fallback へ退避
    notes.push(
      `${patternId} の bindings 生成に失敗したため fallback へ切り替えました: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    const fallback = registry.fallbacks()[0];
    if (!fallback) throw error;
    pattern = fallback;
    patternId = fallback.id;
    fallbackUsed = true;
    // fallback 側でも先に variant を確定してから bindings を作る
    variantId = resolveVariant(fallback, fallback.variants[0] ?? "standard", notes);
    bindings = pattern.bindingsSchema.parse(pattern.buildBindings(slide, variantId));
  }

  const provenance = buildProvenance(slide);

  return {
    slideId: slide.id,
    slideNumber,
    patternId,
    patternVersion: pattern.version,
    variantId,
    score: selection.score,
    reason: selection.reason,
    fallbackUsed,
    bindings,
    ...(provenance ? { provenance } : {}),
    ...(includeCandidates ? { candidates: selection.candidates } : {}),
    ...(notes.length > 0 ? { notes } : {}),
  };
}

/** Visual Plan を安定した JSON 文字列にする（差分レビュー用に整形） */
export function serializeVisualPlan(plan: VisualPlan): string {
  return `${JSON.stringify(plan, null, 2)}\n`;
}

// ---------------------------------------------------------------------------
// 保存済み Visual Plan の再生（parseVisualPlan）
// ---------------------------------------------------------------------------

const PatternEvaluationShape = z
  .object({
    patternId: z.string(),
    version: z.string(),
    compatible: z.boolean(),
    score: z.number(),
    matchScore: z.number(),
    compatibilityScore: z.number(),
    reason: z.string(),
    variantId: z.string().optional(),
    rejectedAt: z.enum(["compatibility", "constraints", "match"]).optional(),
  })
  .strict();

const VisualPlanSlideShape = z
  .object({
    slideId: z.string().min(1),
    slideNumber: z.number().int().positive(),
    patternId: z.string().min(1),
    patternVersion: z.string().min(1),
    variantId: z.string().min(1),
    score: z.number(),
    reason: z.string(),
    fallbackUsed: z.boolean(),
    bindings: z.unknown(),
    provenance: z
      .object({
        sources: z.array(SourceSchema),
        evidence: z.array(
          z
            .object({
              id: z.string().min(1),
              type: z.string().min(1),
              statement: z.string().min(1),
              sourceIds: z.array(z.string().min(1)).optional(),
            })
            .strict(),
        ),
      })
      .strict()
      .optional(),
    candidates: z.array(PatternEvaluationShape).optional(),
    notes: z.array(z.string()).optional(),
  })
  .strict();

const VisualPlanShape = z
  .object({
    schemaVersion: z.string(),
    title: z.string().min(1),
    subtitle: z.string().optional(),
    author: z.string().optional(),
    themeId: z.string().min(1),
    slides: z.array(VisualPlanSlideShape).min(1),
  })
  .strict();

export class VisualPlanValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "VisualPlanValidationError";
  }
}

export interface ParseVisualPlanOptions {
  registry?: PatternRegistry;
}

/**
 * 保存済み Visual Plan（JSON）を検証して再生可能な状態にする。
 *
 * 検証内容:
 *  1. 構造（VisualPlanShape）
 *  2. schemaVersion が現行と一致すること
 *  3. patternId が現在のレジストリに存在すること
 *  4. patternVersion が現在の Pattern 実装と一致すること（不一致は再生成を要求）
 *  5. variantId が実装済みであること
 *  6. bindings が各 Pattern の bindingsSchema に適合すること
 *
 * これに通らない Plan を「黙って違う出力」にすることはない。必ず例外になる。
 */
export function parseVisualPlan(
  input: unknown,
  options: ParseVisualPlanOptions = {},
): VisualPlan {
  const registry = options.registry ?? defaultRegistry;

  const structure = VisualPlanShape.safeParse(input);
  if (!structure.success) {
    const issues = structure.error.issues
      .slice(0, 5)
      .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
      .join("\n");
    throw new VisualPlanValidationError(`Visual Plan の構造が不正です:\n${issues}`);
  }
  const plan = structure.data;

  if (plan.schemaVersion !== VISUAL_PLAN_SCHEMA_VERSION) {
    throw new VisualPlanValidationError(
      `Visual Plan の schemaVersion が不一致です (plan=${plan.schemaVersion}, 現行=${VISUAL_PLAN_SCHEMA_VERSION})。presentation.json から再生成してください`,
    );
  }

  const slides = plan.slides.map((slide) => {
    if (!registry.has(slide.patternId)) {
      throw new VisualPlanValidationError(
        `slide ${slide.slideId}: 未登録の patternId です: ${slide.patternId}`,
      );
    }
    const pattern = registry.get(slide.patternId);

    if (pattern.version !== slide.patternVersion) {
      throw new VisualPlanValidationError(
        `slide ${slide.slideId}: Pattern '${slide.patternId}' のバージョンが不一致です (plan=${slide.patternVersion}, 実装=${pattern.version})。presentation.json から再生成してください`,
      );
    }
    if (!pattern.variants.includes(slide.variantId)) {
      throw new VisualPlanValidationError(
        `slide ${slide.slideId}: Pattern '${slide.patternId}' に variant '${slide.variantId}' は実装されていません`,
      );
    }

    const bindings = pattern.bindingsSchema.safeParse(slide.bindings);
    if (!bindings.success) {
      const issues = bindings.error.issues
        .slice(0, 5)
        .map((i) => `  - bindings.${i.path.join(".")}: ${i.message}`)
        .join("\n");
      throw new VisualPlanValidationError(
        `slide ${slide.slideId}: bindings が Pattern '${slide.patternId}' の契約に違反しています:\n${issues}`,
      );
    }

    return { ...slide, bindings: bindings.data } as VisualPlanSlide;
  });

  return { ...plan, slides };
}
