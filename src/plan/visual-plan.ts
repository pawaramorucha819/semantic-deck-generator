import { z } from "zod";
import { SourceSchema, SlideExplanationSchema } from "../domain/schema.js";
import type { PresentationSpec, SlideSpec, SlideExplanation, Source } from "../domain/types.js";
import type { AnyVisualPattern } from "../patterns/pattern.js";
import { defaultRegistry, PatternRegistry } from "../patterns/registry.js";
import { selectPattern, type PatternEvaluation } from "../selector/selector.js";
import {
  decideSlidePattern,
  type SelectionPolicyOptions,
  type SlideSelectionOutcome,
} from "../selector/selection-policy.js";
import type { MasterRole } from "../templates/template.js";
import { MASTER_ROLES } from "../templates/template.js";
import { DecisionTraceSchema, type DecisionTrace } from "./decision-trace.js";

export const VISUAL_PLAN_SCHEMA_VERSION = "2.0";

/**
 * 読み込みを受け付ける Visual Plan の schemaVersion。
 *
 * V2 で追加されたフィールド（templateId / masterRole / decisionTrace）はすべて任意なので、
 * V1.1 で保存された Plan は「同じ描画結果になる」ことが保証できる範囲でのみ受け入れる。
 * 受け入れた場合は `upgradedFrom` に元のバージョンを残し、黙って読み替えたことにしない。
 */
export const SUPPORTED_VISUAL_PLAN_SCHEMA_VERSIONS = ["1.1", "2.0"] as const;

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
  /** Visible reading content, composed outside the selected diagram. */
  explanation?: SlideExplanation;
  /** 出典・根拠の構造化された記録（→ スピーカーノート / 検証用） */
  provenance?: SlideProvenance;
  /** デバッグ用: 全 Pattern の評価結果（スコア降順） */
  candidates?: PatternEvaluation[];
  /** buildBindings 失敗などの注記 */
  notes?: string[];
  /**
   * このスライドを描く Slide Master の役割（Template 使用時のみ意味を持つ）。
   * 省略時は 'content'。
   */
  masterRole?: MasterRole;
  /**
   * なぜこの Pattern になったのかの追跡記録（V2）。
   * deterministic / knowledge / ai / decision の 4 段構成。
   */
  decisionTrace?: DecisionTrace;
}

export interface VisualPlan {
  schemaVersion: string;
  title: string;
  subtitle?: string;
  author?: string;
  themeId: string;
  /** 企業テンプレート ID（未指定なら Template 非使用 = V1 と同じ紙面） */
  templateId?: string;
  slides: VisualPlanSlide[];
  /** 旧バージョンから読み込んだ場合の元バージョン */
  upgradedFrom?: string;
}

export interface BuildVisualPlanOptions {
  registry?: PatternRegistry;
  themeId?: string;
  /** 企業テンプレート ID（Renderer が Slide Master を解決するのに使う） */
  templateId?: string;
  /**
   * 章扉（section master）として描くスライドの id。
   * 「どのスライドを章扉にするか」は意味ではなく表現の決定なので
   * presentation.json ではなくこちら（生成設定）で指定する。
   */
  sectionSlideIds?: readonly string[];
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
  const sectionIds = new Set(options.sectionSlideIds ?? []);

  const slides = spec.slides.map((slide, index) => {
    const planned = planSlide(slide, index + 1, registry, includeCandidates, options.minimumScore);
    return sectionIds.has(slide.id)
      ? ({ ...planned, masterRole: "section" as MasterRole })
      : planned;
  });

  return {
    schemaVersion: VISUAL_PLAN_SCHEMA_VERSION,
    title: spec.title,
    subtitle: spec.subtitle,
    author: spec.author,
    themeId: options.themeId ?? "default",
    ...(options.templateId ? { templateId: options.templateId } : {}),
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

/** Pattern が確定した状態（決定論的選択、または Selection Policy の結果） */
interface SlideDecision {
  patternId: string;
  variantId: string;
  score: number;
  reason: string;
  fallbackUsed: boolean;
  candidates: PatternEvaluation[];
  decisionTrace?: DecisionTrace;
}

/**
 * 決定済みの Pattern から Visual Plan の 1 枚を組み立てる。
 *
 * bindings 生成の失敗時に fallback へ退避する挙動は、
 * 決定論モードでも AI モードでも同じ（AI だから検証を飛ばす、ということはない）。
 */
function planSlideFromDecision(
  slide: SlideSpec,
  slideNumber: number,
  registry: PatternRegistry,
  includeCandidates: boolean,
  decision: SlideDecision,
): VisualPlanSlide {
  const notes: string[] = [];

  let patternId = decision.patternId;
  let fallbackUsed = decision.fallbackUsed;
  let pattern = registry.get(patternId);

  // variant を先に確定してから bindings を作る。
  // これにより buildBindings には必ず実装済みの variant だけが渡り、
  // Plan の variantId と bindings が常に一致する。
  let variantId = resolveVariant(pattern, decision.variantId, notes);
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

  // 決定と実際の描画がずれた場合（fallback 退避）は Decision Trace にも残す
  const decisionTrace =
    decision.decisionTrace && patternId !== decision.patternId
      ? {
          ...decision.decisionTrace,
          decision: {
            ...decision.decisionTrace.decision,
            selectedPatternId: patternId,
            selectedVariantId: variantId,
            reason: `${decision.decisionTrace.decision.reason} / bindings 生成に失敗したため fallback へ退避`,
          },
        }
      : decision.decisionTrace;

  return {
    slideId: slide.id,
    slideNumber,
    patternId,
    patternVersion: pattern.version,
    variantId,
    score: decision.score,
    reason: decision.reason,
    fallbackUsed,
    bindings,
    ...(slide.explanation ? { explanation: SlideExplanationSchema.parse(slide.explanation) } : {}),
    ...(provenance ? { provenance } : {}),
    ...(includeCandidates ? { candidates: decision.candidates } : {}),
    ...(notes.length > 0 ? { notes } : {}),
    ...(decisionTrace ? { decisionTrace } : {}),
  };
}

function planSlide(
  slide: SlideSpec,
  slideNumber: number,
  registry: PatternRegistry,
  includeCandidates: boolean,
  minimumScore: number | undefined,
): VisualPlanSlide {
  const selection = selectPattern(slide, { registry, minimumScore });
  return planSlideFromDecision(slide, slideNumber, registry, includeCandidates, {
    patternId: selection.patternId,
    variantId: selection.variantId,
    score: selection.score,
    reason: selection.reason,
    fallbackUsed: selection.fallbackUsed,
    candidates: selection.candidates,
  });
}

export interface BuildVisualPlanWithPolicyOptions extends BuildVisualPlanOptions {
  /** manual / assist / auto と、AI・Knowledge の接続 */
  policy: SelectionPolicyOptions;
}

export interface PolicyPlanResult {
  plan: VisualPlan;
  /** スライドごとの決定過程（Advisor Report の材料） */
  outcomes: SlideSelectionOutcome[];
}

/**
 * Selection Policy を通した Visual Plan の生成（AI / Knowledge を使う経路）。
 *
 * 決定が確定したあとの処理は `buildVisualPlan` とまったく同じで、
 * bindings 検証・fallback 退避・Renderer への受け渡しは共通。
 * **Visual Plan が出来上がった時点で AI の関与は終わり**、以降は決定論的。
 */
export async function buildVisualPlanWithPolicy(
  spec: PresentationSpec,
  options: BuildVisualPlanWithPolicyOptions,
): Promise<PolicyPlanResult> {
  const registry = options.registry ?? defaultRegistry;
  const includeCandidates = options.includeCandidates ?? true;
  const sectionIds = new Set(options.sectionSlideIds ?? []);

  const outcomes: SlideSelectionOutcome[] = [];
  const slides: VisualPlanSlide[] = [];

  for (const [index, slide] of spec.slides.entries()) {
    const outcome = await decideSlidePattern(slide, {
      ...options.policy,
      registry,
      ...(options.minimumScore !== undefined ? { minimumScore: options.minimumScore } : {}),
    });
    outcomes.push(outcome);

    const planned = planSlideFromDecision(slide, index + 1, registry, includeCandidates, {
      patternId: outcome.patternId,
      variantId: outcome.variantId,
      score: outcome.score,
      reason: outcome.reason,
      fallbackUsed: outcome.fallbackUsed,
      candidates: outcome.candidates,
      decisionTrace: outcome.decisionTrace,
    });
    slides.push(sectionIds.has(slide.id) ? { ...planned, masterRole: "section" } : planned);
  }

  return {
    plan: {
      schemaVersion: VISUAL_PLAN_SCHEMA_VERSION,
      title: spec.title,
      subtitle: spec.subtitle,
      author: spec.author,
      themeId: options.themeId ?? "default",
      ...(options.templateId ? { templateId: options.templateId } : {}),
      slides,
    },
    outcomes,
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
    explanation: SlideExplanationSchema.optional(),
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
    masterRole: z.enum(MASTER_ROLES).optional(),
    decisionTrace: DecisionTraceSchema.optional(),
  })
  .strict();

const VisualPlanShape = z
  .object({
    schemaVersion: z.string(),
    title: z.string().min(1),
    subtitle: z.string().optional(),
    author: z.string().optional(),
    themeId: z.string().min(1),
    templateId: z.string().min(1).optional(),
    slides: z.array(VisualPlanSlideShape).min(1),
    upgradedFrom: z.string().min(1).optional(),
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

  if (!(SUPPORTED_VISUAL_PLAN_SCHEMA_VERSIONS as readonly string[]).includes(plan.schemaVersion)) {
    throw new VisualPlanValidationError(
      `Visual Plan の schemaVersion が非対応です (plan=${plan.schemaVersion}, 対応=${SUPPORTED_VISUAL_PLAN_SCHEMA_VERSIONS.join(" / ")})。presentation.json から再生成してください`,
    );
  }
  // 旧バージョンの Plan は「V2 で追加されたフィールドがすべて任意」であることを根拠に
  // 受け入れる。読み替えたことは upgradedFrom として必ず残す。
  const upgradedFrom =
    plan.schemaVersion === VISUAL_PLAN_SCHEMA_VERSION ? undefined : plan.schemaVersion;

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

    if (slide.explanation && !pattern.supportsExplanatoryLayout) {
      throw new VisualPlanValidationError(`slide ${slide.slideId}: Pattern '${slide.patternId}' does not support visible explanation; choose a supported pattern`);
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

  return {
    ...plan,
    schemaVersion: VISUAL_PLAN_SCHEMA_VERSION,
    ...(upgradedFrom ? { upgradedFrom } : {}),
    slides,
  };
}
