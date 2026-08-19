import type { PresentationSpec, SlideSpec } from "../domain/types.js";
import { defaultRegistry, PatternRegistry } from "../patterns/registry.js";
import { selectPattern, type PatternEvaluation } from "../selector/selector.js";

export const VISUAL_PLAN_SCHEMA_VERSION = "1.0";

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
  /** Pattern 固有の描画用データ。Renderer が見るのはここだけ。 */
  bindings: unknown;
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
 * この関数を通過した時点で「どう見せるか」は完全に確定し、
 * 以降 Renderer は SlideSpec を一切参照しない。
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
  let variantId = selection.variantId;
  let fallbackUsed = selection.fallbackUsed;
  let pattern = registry.get(patternId);
  let bindings: unknown;

  try {
    bindings = pattern.buildBindings(slide, variantId);
  } catch (error) {
    // Pattern 側の想定外（match は通ったが bindings を作れない）→ fallback へ退避
    notes.push(
      `${patternId}.buildBindings が失敗したため fallback へ切り替えました: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    const fallback = registry.fallbacks()[0];
    if (!fallback) throw error;
    pattern = fallback;
    patternId = fallback.id;
    variantId = fallback.variants[0] ?? "standard";
    fallbackUsed = true;
    bindings = pattern.buildBindings(slide, variantId);
  }

  if (!pattern.variants.includes(variantId)) {
    notes.push(`未実装の variantId '${variantId}' が指定されたため '${pattern.variants[0]}' を使用`);
    variantId = pattern.variants[0] ?? "standard";
  }

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
    ...(includeCandidates ? { candidates: selection.candidates } : {}),
    ...(notes.length > 0 ? { notes } : {}),
  };
}

/** Visual Plan を安定した JSON 文字列にする（差分レビュー用に整形） */
export function serializeVisualPlan(plan: VisualPlan): string {
  return `${JSON.stringify(plan, null, 2)}\n`;
}
