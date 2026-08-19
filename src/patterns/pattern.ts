import type { z } from "zod";
import type { EvidenceType, LogicalRelation, SlideRole, SlideSpec } from "../domain/types.js";
import type { RenderContext } from "../render/context.js";

/**
 * `match()` の結果。
 *
 * score は Pattern 間で比較されるため、次の共通ルーブリックに従うこと
 * （範囲外の値は Selector が実行時に拒否する）:
 *
 *   0.90 - 1.00  意味構造がこの Pattern の理想形（専用の logicalRelation + 完全なデータ）
 *   0.70 - 0.90  適用可能で有力（構造は合うがデータが理想形ではない）
 *   0.40 - 0.70  適用は可能だが、より適切な表現が他にありうる
 *   0.00 - 0.40  ほぼ fallback 向け（意味を図解できない）
 *
 * 独自の初期点・加点体系を発明せず、このアンカーに沿って設計すること。
 */
export interface MatchResult {
  /** false の場合、この Pattern はこのスライドに適用不可 */
  compatible: boolean;
  /** 0.0 - 1.0。意味構造とのフィット度（上記ルーブリック準拠、実行時検証あり）。 */
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
   *
   * fallback Pattern の契約（PatternRegistry が登録時に強制する）:
   *  - constraints を持ってはならない（どんなスライドでも必ず描画できること）
   *  - 全 logicalRelation に対応していること
   *  - match() が compatible: false を返してはならない
   */
  readonly fallback?: boolean;

  readonly compatibility: PatternCompatibility;
  readonly constraints?: PatternConstraints;

  /**
   * この Pattern の Bindings のランタイムスキーマ。
   *
   * レジストリ格納時に型情報（TBindings）は消えるため、これが実行時の契約になる。
   * buildVisualPlan は生成した Bindings をこのスキーマで検証し、
   * Renderer / parseVisualPlan は保存済み Visual Plan の再生時にこれで検証する。
   */
  readonly bindingsSchema: z.ZodType<TBindings>;

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

/**
 * 型を失わずに Pattern をレジストリへ格納するための消去型。
 * 静的型は消えるが、`bindingsSchema` による実行時検証が契約を守る。
 */
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
