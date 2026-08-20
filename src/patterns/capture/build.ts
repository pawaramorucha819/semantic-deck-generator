import type { ExtractedElement } from "../../pptx/model.js";
import type { Rect } from "../../render/geometry.js";
import { nearestColorToken, TEXT_ROLE_NAMES } from "../../render/tokens.js";
import { resolveTheme, type PresentationTheme, type ResolvedTheme } from "../../theme/theme.js";
import type {
  DeclarativeColor,
  DeclarativePatternDocument,
  DeclarativeTextStyle,
  LayoutNode,
  SlotDefinition,
  ValueRef,
} from "../declarative/schema.js";
import { DECLARATIVE_PATTERN_SCHEMA_VERSION } from "../declarative/schema.js";
import type { SlotInferenceResult } from "./infer.js";

/**
 * 推定結果 + 抽出要素 → Declarative Pattern の下書き。
 *
 * 取り込んだスライドの *幾何をそのまま残す* ため absolute レイアウトで組む。
 * 一方で色とフォントサイズは必ず Theme トークン / role へ寄せる
 * （そうしないと「Pattern に色をハードコードしない」という不変条件が壊れる）。
 */

export interface BuildCaptureOptions {
  id: string;
  description?: string;
  scope?: DeclarativePatternDocument["metadata"] extends infer M
    ? M extends { scope?: infer S }
      ? S
      : never
    : never;
  organization?: string;
  author?: string;
  /** 色・サイズを寄せる先の Theme */
  theme: PresentationTheme;
  slideSize: { widthIn: number; heightIn: number };
  source: { file: string; slideNumber: number; importedAt?: string };
  unsupportedFeatures?: string[];
  warnings?: string[];
  /** 画像アセット（assetId → 相対パス） */
  assets?: Array<{ id: string; path: string }>;
  /** 画像要素インデックス → assetId */
  imageAssetIds?: Map<number, string>;
}

export interface BuildCaptureResult {
  document: DeclarativePatternDocument;
  warnings: string[];
}

const round = (n: number): number => Number(n.toFixed(3));

/** 実サイズ (pt) に最も近い Theme のテキストロールと倍率を求める */
export function nearestTextRole(
  theme: ResolvedTheme,
  sizePt: number,
): { role: (typeof TEXT_ROLE_NAMES)[number]; fontScale?: number } {
  let best: { role: (typeof TEXT_ROLE_NAMES)[number]; diff: number } | undefined;
  for (const role of TEXT_ROLE_NAMES) {
    const diff = Math.abs(theme.fontSizes[role] - sizePt);
    if (!best || diff < best.diff) best = { role, diff };
  }
  const role = best!.role;
  const base = theme.fontSizes[role];
  const scale = base > 0 ? sizePt / base : 1;
  // 15% 以内なら role のサイズをそのまま使う（Theme 変更に追従させる）
  if (Math.abs(scale - 1) <= 0.15) return { role };
  return { role, fontScale: Number(Math.min(4, Math.max(0.4, scale)).toFixed(2)) };
}

function colorTokenOf(
  theme: ResolvedTheme,
  hex: string | undefined,
  fallback: DeclarativeColor,
  warnings: string[],
  label: string,
): DeclarativeColor {
  if (!hex) return fallback;
  const nearest = nearestColorToken(theme, hex);
  if (!nearest.exact) {
    warnings.push(`${label}: #${hex} を Theme トークン '${nearest.token}' に寄せました`);
  }
  return nearest.token;
}

function textStyleOf(
  theme: ResolvedTheme,
  element: ExtractedElement,
  warnings: string[],
  defaults: Partial<DeclarativeTextStyle> = {},
): DeclarativeTextStyle {
  const size = element.text?.maxFontSize ?? theme.fontSizes.body;
  const { role, fontScale } = nearestTextRole(theme, size);
  const color = element.text?.runs.find((r) => r.color)?.color;
  return {
    role,
    ...(fontScale ? { fontScale } : {}),
    color: colorTokenOf(theme, color, defaults.color ?? "text", warnings, "文字色"),
    ...(element.text?.runs.some((r) => r.bold) ? { bold: true } : {}),
    ...(element.text?.align ? { align: element.text.align } : {}),
    ...(element.text?.anchor ? { valign: element.text.anchor } : {}),
    ...defaults,
  };
}

function absoluteChild(box: Rect, node: LayoutNode): { box: Rect; node: LayoutNode } {
  return {
    box: { x: round(box.x), y: round(box.y), w: round(Math.max(0.05, box.w)), h: round(Math.max(0.05, box.h)) },
    node,
  };
}

function boundingBox(boxes: Rect[]): Rect {
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  const right = Math.max(...boxes.map((b) => b.x + b.w));
  const bottom = Math.max(...boxes.map((b) => b.y + b.h));
  return { x, y, w: right - x, h: bottom - y };
}

const NUMERIC_PATTERN = /^[\s0-9.,+\-%¥$€£億万千人年月日時間件社%]+$/;

function looksNumeric(text: string): boolean {
  return /\d/.test(text) && NUMERIC_PATTERN.test(text.trim());
}

/** 推定結果から Declarative Pattern の下書きを組み立てる */
export function buildCapturedPattern(
  elements: ExtractedElement[],
  inference: SlotInferenceResult,
  options: BuildCaptureOptions,
): BuildCaptureResult {
  const theme = resolveTheme(options.theme);
  const warnings = [...(options.warnings ?? [])];
  const slots: SlotDefinition[] = [];
  const children: Array<{ box: Rect; node: LayoutNode }> = [];

  const textOf = (index: number | undefined): ExtractedElement | undefined =>
    index === undefined ? undefined : elements[index];

  // --- 見出し・補足・役割ラベル・出典 --------------------------------------
  const headline = textOf(inference.headlineIndex);
  if (headline) {
    slots.push({ id: "headline", type: "text", source: "keyMessage" });
    children.push(
      absoluteChild(headline.box, {
        type: "text",
        value: { ref: "headline" },
        style: textStyleOf(theme, headline, warnings, { color: "text" }),
      }),
    );
  } else {
    // 見出しが無いと Pattern として成立しないので、最低限の slot を用意する
    slots.push({ id: "headline", type: "text", source: "keyMessage" });
    warnings.push("見出しを推定できなかったため、headline slot の配置は手で指定してください");
    children.push(
      absoluteChild(
        { x: 0.6, y: 0.4, w: options.slideSize.widthIn - 1.2, h: 0.6 },
        { type: "text", value: { ref: "headline" }, style: { role: "keyMessage", bold: true } },
      ),
    );
  }

  const eyebrow = textOf(inference.eyebrowIndex);
  if (eyebrow) {
    slots.push({ id: "eyebrow", type: "text", source: "slideRoleLabel", optional: true });
    children.push(
      absoluteChild(eyebrow.box, {
        type: "text",
        value: { ref: "eyebrow" },
        style: textStyleOf(theme, eyebrow, warnings, { color: "primary" }),
        when: { hasValue: "eyebrow" },
      }),
    );
  }

  const subline = textOf(inference.sublineIndex);
  if (subline) {
    slots.push({ id: "subline", type: "text", source: "supportingMessage", optional: true });
    children.push(
      absoluteChild(subline.box, {
        type: "text",
        value: { ref: "subline" },
        style: textStyleOf(theme, subline, warnings, { color: "mutedText" }),
        when: { hasValue: "subline" },
      }),
    );
  }

  const source = textOf(inference.sourceIndex);
  if (source) {
    slots.push({ id: "sourceText", type: "text", source: "sourceText", optional: true });
    children.push(
      absoluteChild(source.box, {
        type: "text",
        value: { ref: "sourceText" },
        style: textStyleOf(theme, source, warnings, { color: "mutedText" }),
        when: { hasValue: "sourceText" },
      }),
    );
  }

  // --- 固定デザイン（ロゴ・装飾・固定ラベル） --------------------------------
  for (const index of inference.staticIndexes) {
    const element = elements[index];
    if (!element) continue;

    if (element.kind === "image") {
      const assetId = options.imageAssetIds?.get(index);
      if (!assetId) {
        warnings.push(`画像を解決できなかったため除外しました: ${element.name ?? element.id}`);
        continue;
      }
      children.push(
        absoluteChild(element.box, {
          type: "image",
          assetId,
          height: round(element.box.h),
          width: round(element.box.w),
          sizing: "contain",
        }),
      );
      continue;
    }

    if (element.text && element.text.plain.trim().length > 0) {
      children.push(
        absoluteChild(element.box, {
          type: "text",
          value: { literal: element.text.plain.replace(/\n/g, " ").slice(0, 400) },
          style: textStyleOf(theme, element, warnings, {}),
        }),
      );
      continue;
    }

    if (element.kind === "line") {
      children.push(
        absoluteChild(element.box, {
          type: "shape",
          shape: "line",
          style: {
            line: {
              color: colorTokenOf(theme, element.line?.color, "border", warnings, "線の色"),
              ...(element.line?.widthPt ? { width: round(element.line.widthPt) } : {}),
            },
          },
        }),
      );
      continue;
    }

    if (element.fill?.type === "solid") {
      children.push(
        absoluteChild(element.box, {
          type: "shape",
          shape: element.geometry === "roundRect" ? "roundRect" : "rect",
          style: { fill: colorTokenOf(theme, element.fill.color, "surface", warnings, "図形の色") },
        }),
      );
      continue;
    }

    if (element.fill?.type === "unsupported") {
      warnings.push(`${element.fill.detail ?? "非対応の塗り"} の図形を除外しました`);
    }
  }

  // --- 反復セル -------------------------------------------------------------
  const usableCells = inference.cells.filter((cell) => cell.textIndexes.length > 0);
  if (usableCells.length >= 2) {
    const cellBoxes = usableCells.map((cell) => {
      const boxes = [
        ...(cell.backgroundIndex !== undefined && elements[cell.backgroundIndex]
          ? [elements[cell.backgroundIndex]!.box]
          : []),
        ...cell.textIndexes.map((i) => elements[i]!.box),
      ];
      return boundingBox(boxes);
    });
    const region = boundingBox(cellBoxes);
    const gaps = cellBoxes
      .slice(1)
      .map((box, i) => box.x - (cellBoxes[i]!.x + cellBoxes[i]!.w))
      .filter((g) => g > 0);
    const gap = gaps.length > 0 ? round(gaps.reduce((a, b) => a + b, 0) / gaps.length) : 0.2;

    const count = usableCells.length;
    slots.push({
      id: "items",
      type: "items",
      source: "contentOrEvidence",
      min: count,
      max: count,
      labelMaxLength: 40,
    });

    const template = usableCells[0]!;
    const templateTexts = template.textIndexes
      .map((i) => elements[i])
      .filter((e): e is ExtractedElement => !!e);

    const stackChildren: LayoutNode[] = [];
    let usedLabel = false;
    let usedMetric = false;
    templateTexts.forEach((element, position) => {
      const plain = element.text?.plain ?? "";
      let ref: ValueRef;
      let defaults: Partial<DeclarativeTextStyle> = {};
      if (!usedLabel) {
        ref = { ref: "item.label" };
        defaults = { bold: true, color: "text" };
        usedLabel = true;
      } else if (!usedMetric && looksNumeric(plain)) {
        ref = { ref: "item.metric" };
        defaults = { bold: true, color: "primary" };
        usedMetric = true;
      } else {
        ref = { ref: "item.description" };
        defaults = { color: "mutedText" };
      }
      stackChildren.push({
        type: "text",
        value: ref,
        style: textStyleOf(theme, element, warnings, defaults),
        ...(position === templateTexts.length - 1 ? { height: "fill" as const } : {}),
        ...("ref" in ref && ref.ref !== "item.label" ? { when: { hasValue: ref.ref } } : {}),
      });
    });

    const background =
      template.backgroundIndex !== undefined ? elements[template.backgroundIndex] : undefined;
    const stack: LayoutNode = { type: "stack", gap: "sm", children: stackChildren };
    const cellNode: LayoutNode = background
      ? {
          type: "card",
          fill: colorTokenOf(theme, background.fill?.color, "surface", warnings, "カードの色"),
          ...(background.line?.color
            ? { border: colorTokenOf(theme, background.line.color, "border", warnings, "カードの枠線") }
            : {}),
          child: stack,
        }
      : stack;

    children.push(
      absoluteChild(region, {
        type: "columns",
        gap,
        repeat: { slot: "items" },
        item: cellNode,
      }),
    );
  }

  const document: DeclarativePatternDocument = {
    schemaVersion: DECLARATIVE_PATTERN_SCHEMA_VERSION,
    kind: "declarative",
    id: options.id,
    version: "1.0.0",
    description:
      options.description ??
      `${options.source.file} のスライド ${options.source.slideNumber} から取り込んだ Pattern`,
    metadata: {
      scope: (options.scope as never) ?? "company",
      // Capture の成果物は必ず下書き。人間のレビューを経て approved にする。
      status: "draft",
      ...(options.author ? { author: options.author } : {}),
      ...(options.organization ? { organization: options.organization } : {}),
      source: {
        file: options.source.file,
        slideNumber: options.source.slideNumber,
        ...(options.source.importedAt ? { importedAt: options.source.importedAt } : {}),
        ...(options.unsupportedFeatures?.length
          ? { unsupportedFeatures: [...new Set(options.unsupportedFeatures)] }
          : {}),
        ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {}),
      },
    },
    compatibility: {
      // 取り込み時点では意味が分からないので広めに宣言し、人間が絞る前提にする
      logicalRelations: ["claim_evidence"],
    },
    slots,
    match: {
      baseScore: 0.5,
      ...(slots.some((s) => s.id === "items") ? { requireSlots: ["items"] } : {}),
    },
    layout: {
      body: { type: "absolute", origin: "canvas", children },
    },
    ...(options.assets && options.assets.length > 0
      ? { assets: options.assets.map((a) => ({ id: a.id, path: a.path, kind: "image" as const })) }
      : {}),
  };

  return { document, warnings };
}
