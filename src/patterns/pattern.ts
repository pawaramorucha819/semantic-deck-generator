import type { EvidenceType, LogicalRelation, SlideRole, SlideSpec } from "../domain/types.js";
import type { RenderContext } from "../render/context.js";

/** `match()` の結果 */
export interface MatchResult {
  /** false の場合、この Pattern はこのスライドに適用不可 */
  compatible: boolean;
  /** 0.0 - 1.0。意味構造とのフィット度。 */
  score: number;
  /** なぜその判断になったのか（Visual Plan にそのまま出力される） */
  reason: string;
  /** Pattern 側から Variant を提案する場合 */
  variantId?: string;
}

/**
 * Pattern の「使用条件」。
 * compatibility が *意味の型* による粗いゲートなのに対し、
 * constraints は *データ量* による定量的なゲート。
 */
export interface PatternConstraints {
  minEvidence?: number;
  maxEvidence?: number;
  maxKeyMessageLength?: number;
  maxSupportingMessageLength?: number;
  /** これらの型の evidence が最低 1 件必要 */
  requiresAnyEvidenceType?: EvidenceType[];
}

export interface PatternCompatibility {
  logicalRelations: LogicalRelation[];
  slideRoles?: SlideRole[];
  evidenceTypes?: EvidenceType[];
}

export interface VisualPattern<TBindings = unknown> {
  readonly id: string;
  readonly version: string;
  /** 人間向けの説明（README / デバッグ出力用） */
  readonly description: string;
  /** 実装済み Variant。先頭が既定。 */
  readonly variants: readonly string[];
  /**
   * true の場合、他に適合する Pattern がないときの受け皿として使われる。
   * fallback Pattern は必ず 1 つ以上必要。
   */
  readonly fallback?: boolean;

  readonly compatibility: PatternCompatibility;
  readonly constraints?: PatternConstraints;

  /** 意味構造とのフィット度を判定する。座標計算はここでは行わない。 */
  match(slide: SlideSpec): MatchResult;

  /**
   * SlideSpec → この Pattern 専用の Bindings。
   * ここが「意味 → 表現」の変換点。render() は SlideSpec を見てはいけない。
   */
  buildBindings(slide: SlideSpec, variantId: string): TBindings;

  /** Bindings → 描画。SlideSpec には一切アクセスできない。 */
  render(bindings: TBindings, context: RenderContext): void;
}

/** 型を失わずに Pattern をレジストリへ格納するための消去型 */
export type AnyVisualPattern = VisualPattern<any>;

export interface ConstraintViolation {
  constraint: keyof PatternConstraints;
  message: string;
}

/** constraints を評価する。空配列なら合格。 */
export function checkConstraints(
  slide: SlideSpec,
  constraints: PatternConstraints | undefined,
): ConstraintViolation[] {
  if (!constraints) return [];
  const violations: ConstraintViolation[] = [];
  const count = slide.evidence.length;

  if (constraints.minEvidence !== undefined && count < constraints.minEvidence) {
    violations.push({
      constraint: "minEvidence",
      message: `evidence が ${constraints.minEvidence} 件以上必要ですが ${count} 件です`,
    });
  }
  if (constraints.maxEvidence !== undefined && count > constraints.maxEvidence) {
    violations.push({
      constraint: "maxEvidence",
      message: `evidence は ${constraints.maxEvidence} 件までですが ${count} 件です`,
    });
  }
  if (
    constraints.maxKeyMessageLength !== undefined &&
    slide.keyMessage.length > constraints.maxKeyMessageLength
  ) {
    violations.push({
      constraint: "maxKeyMessageLength",
      message: `keyMessage は ${constraints.maxKeyMessageLength} 文字までですが ${slide.keyMessage.length} 文字です`,
    });
  }
  if (
    constraints.maxSupportingMessageLength !== undefined &&
    (slide.supportingMessage?.length ?? 0) > constraints.maxSupportingMessageLength
  ) {
    violations.push({
      constraint: "maxSupportingMessageLength",
      message: `supportingMessage は ${constraints.maxSupportingMessageLength} 文字までですが ${slide.supportingMessage?.length} 文字です`,
    });
  }
  if (constraints.requiresAnyEvidenceType?.length) {
    const wanted = new Set<EvidenceType>(constraints.requiresAnyEvidenceType);
    if (!slide.evidence.some((e) => wanted.has(e.type))) {
      violations.push({
        constraint: "requiresAnyEvidenceType",
        message: `次のいずれかの evidence が必要です: ${constraints.requiresAnyEvidenceType.join(" / ")}`,
      });
    }
  }
  return violations;
}
