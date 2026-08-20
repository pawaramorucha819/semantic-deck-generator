import type { z } from "zod";
import type { SlideSpec } from "../../domain/types.js";
import type { RenderContext } from "../../render/context.js";
import type { MatchResult, PatternCompatibility, PatternConstraints, VisualPattern } from "../pattern.js";
import { buildBindingsSchema, extractSlots, type DeclarativeBindings } from "./bindings.js";
import { renderDeclarativeLayout } from "./renderer.js";
import {
  DeclarativePatternValidationError,
  parseDeclarativePattern,
  type Condition,
  type DeclarativeLayout,
  type DeclarativePatternDocument,
  type LayoutNode,
  type ValueRef,
} from "./schema.js";
import { evaluateMatchCondition } from "./values.js";

/**
 * Declarative Pattern Document → VisualPattern。
 *
 * コンパイルと言っても *コードは生成しない*。JSON をそのまま解釈する
 * クロージャを作るだけなので、任意コード実行の経路は存在しない。
 *
 * コード Pattern との違いは実装手段だけで、契約（match / buildBindings /
 * render / bindingsSchema / version）はまったく同じ。
 */

export interface CompiledDeclarativePattern extends VisualPattern<DeclarativeBindings> {
  /** 元の宣言（Registry の一覧表示 / Preview 情報に使う） */
  readonly document: DeclarativePatternDocument;
}

/** Pattern パッケージのアセット ID を名前空間化する（Template のアセットと衝突させない） */
export function patternAssetId(patternId: string, assetId: string): string {
  return `pattern:${patternId}:${assetId}`;
}

const clamp01 = (n: number): number => Math.max(0, Math.min(1, n));

// ---------------------------------------------------------------------------
// レイアウトの静的検証（ロード時に typo を弾く）
// ---------------------------------------------------------------------------

interface LayoutValidationContext {
  slotIds: Set<string>;
  assetIds: Set<string>;
  issues: string[];
}

const TEMPLATE_PATTERN = /\{\{\s*([a-zA-Z][a-zA-Z0-9_.]*)\s*\}\}/g;

function checkPath(path: string, inRepeat: boolean, ctx: LayoutValidationContext, where: string): void {
  const head = path.split(".")[0]!;
  if (head === "item" || head === "index") {
    if (!inRepeat) {
      ctx.issues.push(`${where}: '${path}' は repeat の中でしか参照できません`);
    }
    return;
  }
  if (!ctx.slotIds.has(head)) {
    ctx.issues.push(
      `${where}: 未定義の slot を参照しています: '${path}' (定義済み: ${[...ctx.slotIds].join(", ")})`,
    );
  }
}

function checkValueRef(ref: ValueRef, inRepeat: boolean, ctx: LayoutValidationContext, where: string): void {
  if ("ref" in ref) {
    checkPath(ref.ref, inRepeat, ctx, where);
    return;
  }
  if ("template" in ref) {
    const matches = [...ref.template.matchAll(TEMPLATE_PATTERN)];
    if (matches.length === 0) {
      ctx.issues.push(`${where}: template に {{...}} の参照がありません`);
    }
    for (const match of matches) checkPath(match[1]!, inRepeat, ctx, where);
  }
}

function checkCondition(condition: Condition, inRepeat: boolean, ctx: LayoutValidationContext, where: string): void {
  if ("hasValue" in condition) checkPath(condition.hasValue, inRepeat, ctx, where);
  else if ("equals" in condition) checkPath(condition.equals.ref, inRepeat, ctx, where);
  else if ("itemCount" in condition && condition.slot) checkPath(condition.slot, inRepeat, ctx, where);
  else if ("indexEquals" in condition && !inRepeat) {
    ctx.issues.push(`${where}: indexEquals は repeat の中でしか使えません`);
  } else if ("not" in condition) checkCondition(condition.not, inRepeat, ctx, where);
  else if ("all" in condition) condition.all.forEach((c) => checkCondition(c, inRepeat, ctx, where));
  else if ("any" in condition) condition.any.forEach((c) => checkCondition(c, inRepeat, ctx, where));
}

function checkNode(node: LayoutNode, inRepeat: boolean, ctx: LayoutValidationContext, path: string): void {
  if (node.when) checkCondition(node.when, inRepeat, ctx, path);

  switch (node.type) {
    case "text":
      checkValueRef(node.value, inRepeat, ctx, `${path}.value`);
      return;
    case "image":
      if (!ctx.assetIds.has(node.assetId)) {
        ctx.issues.push(`${path}: 未登録のアセットを参照しています: ${node.assetId}`);
      }
      return;
    case "stack":
      node.children.forEach((child, i) => checkNode(child, inRepeat, ctx, `${path}.children[${i}]`));
      return;
    case "columns": {
      if (node.repeat) {
        if (!node.item) {
          ctx.issues.push(`${path}: repeat を使う columns には item が必要です`);
          return;
        }
        if (node.children) {
          ctx.issues.push(`${path}: columns で repeat と children は併用できません`);
        }
        if (!ctx.slotIds.has(node.repeat.slot)) {
          ctx.issues.push(`${path}.repeat: 未定義の slot です: ${node.repeat.slot}`);
        }
        checkNode(node.item, true, ctx, `${path}.item`);
        return;
      }
      if (!node.children || node.children.length === 0) {
        ctx.issues.push(`${path}: columns には repeat+item か children のどちらかが必要です`);
        return;
      }
      node.children.forEach((child, i) => checkNode(child, inRepeat, ctx, `${path}.children[${i}]`));
      return;
    }
    case "card":
      checkNode(node.child, inRepeat, ctx, `${path}.child`);
      return;
    case "absolute":
      node.children.forEach((child, i) => {
        checkBoxExtent(child.box, child.node, ctx, `${path}.children[${i}].box`);
        checkNode(child.node, inRepeat, ctx, `${path}.children[${i}].node`);
      });
      return;
    default:
      return;
  }
}

/**
 * absolute 配置の矩形の大きさを検証する。
 *
 * **線だけ契約が違う**。PowerPoint の水平線は h = 0、垂直線は w = 0 なので、
 * 線に限り「長さが 0 でない」ことだけを求める。それ以外の要素は
 * 従来どおり正の幅と高さを要求する（特殊ケースを全要素へ広げない）。
 */
function checkBoxExtent(
  box: { w: number; h: number },
  node: LayoutNode,
  ctx: LayoutValidationContext,
  where: string,
): void {
  const isLine = node.type === "shape" && node.shape === "line";
  if (isLine) {
    if (box.w <= 0 && box.h <= 0) {
      ctx.issues.push(`${where}: 長さ 0 の線は配置できません（w か h のどちらかが必要です）`);
    }
    return;
  }
  if (box.w <= 0 || box.h <= 0) {
    ctx.issues.push(`${where}: w / h は 0 より大きい必要があります (w=${box.w}, h=${box.h})`);
  }
}

function checkLayout(layout: DeclarativeLayout, ctx: LayoutValidationContext, where: string): void {
  if (layout.header) {
    checkValueRef(layout.header.headline, false, ctx, `${where}.header.headline`);
    if (layout.header.eyebrow) checkValueRef(layout.header.eyebrow, false, ctx, `${where}.header.eyebrow`);
    if (layout.header.subline) checkValueRef(layout.header.subline, false, ctx, `${where}.header.subline`);
  }
  if (layout.footer) checkValueRef(layout.footer.text, false, ctx, `${where}.footer.text`);
  checkNode(layout.body, false, ctx, `${where}.body`);
}

/**
 * 宣言の整合性検査。
 * Zod では表現しきれない「slot 参照の妥当性」「repeat の構造」をここで弾く。
 */
export function validateDeclarativePattern(doc: DeclarativePatternDocument): string[] {
  const ctx: LayoutValidationContext = {
    slotIds: new Set(doc.slots.map((s) => s.id)),
    assetIds: new Set((doc.assets ?? []).map((a) => a.id)),
    issues: [],
  };
  checkLayout(doc.layout, ctx, "layout");
  for (const variant of doc.variants ?? []) {
    if (variant.layout) checkLayout(variant.layout, ctx, `variants.${variant.id}.layout`);
  }
  return ctx.issues;
}

// ---------------------------------------------------------------------------
// コンパイル
// ---------------------------------------------------------------------------

export interface CompileOptions {
  /**
   * draft（人間レビュー前）の Pattern も受け入れるか。既定 false。
   * Pattern Capture が生成した直後の下書きを誤って本番登録しないためのゲート。
   */
  allowDraft?: boolean;
}

export function compileDeclarativePattern(
  input: unknown,
  options: CompileOptions = {},
): CompiledDeclarativePattern {
  const doc = parseDeclarativePattern(input);

  const status = doc.metadata?.status ?? "approved";
  if (status === "draft" && !options.allowDraft) {
    throw new DeclarativePatternValidationError(
      `Declarative Pattern '${doc.id}' は draft です。人間のレビュー後に metadata.status を 'approved' にしてください`,
      [],
    );
  }

  const issues = validateDeclarativePattern(doc);
  if (issues.length > 0) {
    throw new DeclarativePatternValidationError(
      `Declarative Pattern '${doc.id}' の宣言が不整合です (${issues.length} 件)`,
      issues.map((message) => ({ code: "custom", path: [], message }) as z.ZodIssue),
    );
  }

  const bindingsSchema = buildBindingsSchema(doc);
  const variants = (doc.variants ?? [{ id: "standard" }]).map((v) => v.id);
  const layoutByVariant = new Map<string, DeclarativeLayout>();
  for (const variant of doc.variants ?? []) {
    layoutByVariant.set(variant.id, variant.layout ?? doc.layout);
  }

  const compatibility: PatternCompatibility = {
    logicalRelations: doc.compatibility.logicalRelations as PatternCompatibility["logicalRelations"],
    ...(doc.compatibility.slideRoles
      ? { slideRoles: doc.compatibility.slideRoles as PatternCompatibility["slideRoles"] }
      : {}),
    ...(doc.compatibility.evidenceTypes
      ? { evidenceTypes: doc.compatibility.evidenceTypes as PatternCompatibility["evidenceTypes"] }
      : {}),
  };

  const constraints = doc.constraints as PatternConstraints | undefined;
  const assetPrefix = `pattern:${doc.id}:`;

  const pattern: CompiledDeclarativePattern = {
    document: doc,
    id: doc.id,
    version: doc.version,
    description: doc.description,
    variants,
    compatibility,
    ...(constraints && Object.keys(constraints).length > 0 ? { constraints } : {}),
    bindingsSchema,

    match(slide: SlideSpec): MatchResult {
      const extraction = extractSlots(doc, slide);
      if (extraction.missing.length > 0) {
        return {
          compatible: false,
          score: 0,
          reason: `必要な情報が揃っていません (未取得の slot: ${extraction.missing.join(", ")})`,
        };
      }
      if (extraction.violations.length > 0) {
        return { compatible: false, score: 0, reason: extraction.violations.join(" / ") };
      }
      for (const required of doc.match.requireSlots ?? []) {
        if (extraction.bindings[required] === undefined) {
          return {
            compatible: false,
            score: 0,
            reason: `必須 slot '${required}' が埋まりません`,
          };
        }
      }

      const scope = { bindings: extraction.bindings, slide };
      const reasons: string[] = [];
      let score = doc.match.baseScore;
      for (const rule of doc.match.rules ?? []) {
        if (evaluateMatchCondition(rule.when, scope)) {
          score += rule.score;
          reasons.push(rule.reason);
        }
      }
      if (reasons.length === 0) reasons.push(doc.description);

      return {
        compatible: true,
        score: clamp01(score),
        reason: reasons.join(" / "),
        variantId: variants[0] ?? "standard",
      };
    },

    buildBindings(slide: SlideSpec): DeclarativeBindings {
      const extraction = extractSlots(doc, slide);
      if (extraction.missing.length > 0 || extraction.violations.length > 0) {
        throw new Error(
          `[${doc.id}] buildBindings: slot を抽出できません (${[
            ...extraction.missing.map((m) => `missing:${m}`),
            ...extraction.violations,
          ].join(" / ")})`,
        );
      }
      return extraction.bindings;
    },

    render(bindings: DeclarativeBindings, ctx: RenderContext): void {
      const layout = layoutByVariant.get(ctx.variantId) ?? doc.layout;
      renderDeclarativeLayout(layout, bindings as Record<string, unknown>, ctx, {
        assetPrefix,
        namePrefix: doc.id,
      });
    },
  };

  return pattern;
}
