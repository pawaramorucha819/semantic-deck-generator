import { isLineElement, type ExtractedElement } from "../../pptx/model.js";
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
import type { CaptureInterpretation, InterpretedCell } from "./interpretation.js";

/**
 * Pattern Compiler — 解釈 + 抽出要素 → Declarative Pattern の下書き。
 *
 * **ここは意味判断をしない**。「この文字列は数字だから metric だろう」のような推定は
 * すべて Interpretation 層（interpretation.ts）の責務で、Compiler は
 * `CaptureInterpretation` に書かれた役割をそのまま描画構造へ写すだけ。
 * こうしておくと、解釈の出どころを heuristic → human → AI と差し替えても
 * Compiler は変更不要になる。
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

/** absolute レイアウトの 1 要素 */
interface AbsoluteChild {
  box: Rect;
  node: LayoutNode;
}

/**
 * 元 PPTX の重なり順（Z-order）を保つための一時表現。
 *
 * `sourceOrder` は抽出時の要素インデックス（= PowerPoint の描画順）。
 * 最後にこの順で並べ替えてから absolute.children にするので、
 * 「PowerPoint で後ろにあったものは後ろ、前にあったものは前」が保たれる。
 */
interface LayeredCapturedNode {
  sourceOrder: number;
  child: AbsoluteChild;
}

/** 線は w か h が 0 になりうるので、最小サイズへ丸めない */
function absoluteChild(box: Rect, node: LayoutNode): AbsoluteChild {
  const isLine = node.type === "shape" && node.shape === "line";
  const size = (value: number): number => round(isLine ? Math.max(0, value) : Math.max(0.05, value));
  return {
    box: { x: round(box.x), y: round(box.y), w: size(box.w), h: size(box.h) },
    node,
  };
}

/** セルを構成する要素の矩形（幾何の基準） */
function cellBoxesOf(elements: readonly ExtractedElement[], cell: InterpretedCell): Rect[] {
  return contentIndexesOf(cell)
    .map((index) => elements[index]?.box)
    .filter((box): box is Rect => box !== undefined);
}

/** セルを構成する要素インデックス（背景と内容。装飾は含めない） */
function contentIndexesOf(cell: InterpretedCell): number[] {
  return [
    ...(cell.backgroundIndex !== undefined ? [cell.backgroundIndex] : []),
    ...(cell.content.indexIndex !== undefined ? [cell.content.indexIndex] : []),
    cell.content.labelIndex,
    ...(cell.content.metricIndex !== undefined ? [cell.content.metricIndex] : []),
    ...(cell.content.descriptionIndexes ?? []),
  ];
}

function boundingBox(boxes: Rect[]): Rect {
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  const right = Math.max(...boxes.map((b) => b.x + b.w));
  const bottom = Math.max(...boxes.map((b) => b.y + b.h));
  return { x, y, w: right - x, h: bottom - y };
}

/** 解釈から Declarative Pattern の下書きを組み立てる */
export function buildCapturedPattern(
  elements: ExtractedElement[],
  interpretation: CaptureInterpretation,
  options: BuildCaptureOptions,
): BuildCaptureResult {
  const theme = resolveTheme(options.theme);
  const warnings = [...(options.warnings ?? [])];
  const slots: SlotDefinition[] = [];
  const layers: LayeredCapturedNode[] = [];
  const layer = (sourceOrder: number, child: AbsoluteChild): void => {
    layers.push({ sourceOrder, child });
  };

  const textOf = (index: number | undefined): ExtractedElement | undefined =>
    index === undefined ? undefined : elements[index];

  // --- 見出し・補足・役割ラベル・出典 --------------------------------------
  const headline = textOf(interpretation.headlineIndex);
  if (headline) {
    slots.push({ id: "headline", type: "text", source: "keyMessage" });
    layer(
      interpretation.headlineIndex!,
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
    // 合成した見出しは元 PPTX に対応する要素が無いので最前面へ置く
    layer(
      elements.length,
      absoluteChild(
        { x: 0.6, y: 0.4, w: options.slideSize.widthIn - 1.2, h: 0.6 },
        { type: "text", value: { ref: "headline" }, style: { role: "keyMessage", bold: true } },
      ),
    );
  }

  const eyebrow = textOf(interpretation.eyebrowIndex);
  if (eyebrow) {
    slots.push({ id: "eyebrow", type: "text", source: "slideRoleLabel", optional: true });
    layer(
      interpretation.eyebrowIndex!,
      absoluteChild(eyebrow.box, {
        type: "text",
        value: { ref: "eyebrow" },
        style: textStyleOf(theme, eyebrow, warnings, { color: "primary" }),
        when: { hasValue: "eyebrow" },
      }),
    );
  }

  const subline = textOf(interpretation.sublineIndex);
  if (subline) {
    slots.push({ id: "subline", type: "text", source: "supportingMessage", optional: true });
    layer(
      interpretation.sublineIndex!,
      absoluteChild(subline.box, {
        type: "text",
        value: { ref: "subline" },
        style: textStyleOf(theme, subline, warnings, { color: "mutedText" }),
        when: { hasValue: "subline" },
      }),
    );
  }

  const source = textOf(interpretation.sourceTextIndex);
  if (source) {
    slots.push({ id: "sourceText", type: "text", source: "sourceText", optional: true });
    layer(
      interpretation.sourceTextIndex!,
      absoluteChild(source.box, {
        type: "text",
        value: { ref: "sourceText" },
        style: textStyleOf(theme, source, warnings, { color: "mutedText" }),
        when: { hasValue: "sourceText" },
      }),
    );
  }

  // --- 固定デザイン（ロゴ・装飾・固定ラベル） --------------------------------
  for (const index of interpretation.staticIndexes) {
    const element = elements[index];
    if (!element) continue;

    if (element.kind === "image") {
      const assetId = options.imageAssetIds?.get(index);
      if (!assetId) {
        warnings.push(`画像を解決できなかったため除外しました: ${element.name ?? element.id}`);
        continue;
      }
      layer(
        index,
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

    // コネクタと線プリセットの図形はどちらも線。塗りベースの分岐より前に判定する
    // （線は塗りを持たないため、後段の solid 判定では拾えない）。
    if (isLineElement(element)) {
      layer(
        index,
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

    if (element.text && element.text.plain.trim().length > 0) {
      layer(
        index,
        absoluteChild(element.box, {
          type: "text",
          value: { literal: element.text.plain.replace(/\n/g, " ").slice(0, 400) },
          style: textStyleOf(theme, element, warnings, {}),
        }),
      );
      continue;
    }

    if (element.fill?.type === "solid") {
      layer(
        index,
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
  // Compiler は解釈された役割をそのまま構造へ写す。どのテキストが label なのか、
  // どれが metric なのかを *ここで判断することはない*。
  const cells = interpretation.cells;
  if (cells.length >= 2) {
    const cellBoxes = cells.map((cell) => boundingBox(cellBoxesOf(elements, cell)));
    const region = boundingBox(cellBoxes);
    const gaps = cellBoxes
      .slice(1)
      .map((box, i) => box.x - (cellBoxes[i]!.x + cellBoxes[i]!.w))
      .filter((g) => g > 0);
    const gap = gaps.length > 0 ? round(gaps.reduce((a, b) => a + b, 0) / gaps.length) : 0.2;

    const count = cells.length;
    slots.push({
      id: "items",
      type: "items",
      source: "contentOrEvidence",
      min: count,
      max: count,
      labelMaxLength: 40,
    });

    // 1 セル目の見た目を反復テンプレートの基準にする（全セルは同じ役割構成）
    const template = cells[0]!;
    const stackChildren: LayoutNode[] = [];

    const push = (
      element: ExtractedElement | undefined,
      node: (style: DeclarativeTextStyle) => LayoutNode,
      defaults: Partial<DeclarativeTextStyle>,
    ): void => {
      if (!element) return;
      stackChildren.push(node(textStyleOf(theme, element, warnings, defaults)));
    };

    // index: 元の "01" を literal で残さず、item の順番から作る（再利用しても正しく並ぶ）
    push(
      textOf(template.content.indexIndex),
      (style) => ({
        type: "text",
        value: { template: "{{item.paddedOrdinal}}" },
        style,
      }),
      { bold: true, color: "primary" },
    );

    push(
      textOf(template.content.labelIndex),
      (style) => ({ type: "text", value: { ref: "item.label" }, style }),
      { bold: true, color: "text" },
    );

    push(
      textOf(template.content.metricIndex),
      (style) => ({
        type: "text",
        value: { ref: "item.metric" },
        style,
        when: { hasValue: "item.metric" },
      }),
      { bold: true, color: "primary" },
    );

    // divider: セル内の区切り線は内容と一緒に動くよう、カードのフローへ入れる
    const divider = elements[template.decoration?.dividerIndexes?.[0] ?? -1];
    if (divider) {
      stackChildren.push({
        type: "shape",
        shape: "line",
        // 元 PPTX と同じ「高さ 0 の水平線」のまま流し込む（Phase 1 の契約）
        height: 0,
        style: {
          line: {
            color: colorTokenOf(theme, divider.line?.color, "border", warnings, "区切り線の色"),
            ...(divider.line?.widthPt ? { width: round(divider.line.widthPt) } : {}),
          },
        },
        when: { hasValue: "item.description" },
      });
    }
    if ((template.decoration?.dividerIndexes ?? []).length > 1) {
      warnings.push(
        "セル内に複数の区切り線が指定されたため、先頭の 1 本だけをテンプレートに使いました",
      );
    }

    // description: 同じ値を複数ノードへ出さない（重複描画の原因になる）
    const descriptionIndexes = template.content.descriptionIndexes ?? [];
    push(
      textOf(descriptionIndexes[0]),
      (style) => ({
        type: "text",
        value: { ref: "item.description" },
        style,
        height: "fill" as const,
        when: { hasValue: "item.description" },
      }),
      { color: "mutedText" },
    );
    if (descriptionIndexes.length > 1) {
      warnings.push(
        `セルに ${descriptionIndexes.length} 件の description が指定されましたが、` +
          "同じ内容を重複描画しないよう先頭の 1 件だけを使いました",
      );
    }

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

    // 反復セルは複数の PPTX 要素を 1 ノードへまとめるため、
    // 構成要素の最小インデックス（＝最も後ろにあった要素）を Z-order の代表値にする。
    // こうすると、カード背景より後ろにあった装飾（アクセントバー）はカードより前面に残る。
    const cellSourceOrder = Math.min(
      ...cells.map((cell) => Math.min(...contentIndexesOf(cell))),
    );

    layer(
      cellSourceOrder,
      absoluteChild(region, {
        type: "columns",
        gap,
        repeat: { slot: "items" },
        item: cellNode,
      }),
    );

    // accent: セルの縁を飾る図形。元の座標のまま、カードより前面に置く
    // （セル内フローには含めない。カードと同じ絶対位置に固定されるため崩れない）
    for (const cell of cells) {
      for (const index of cell.decoration?.accentIndexes ?? []) {
        const accent = elements[index];
        if (!accent) continue;
        if (accent.fill?.type === "solid") {
          layer(
            index,
            absoluteChild(accent.box, {
              type: "shape",
              shape: accent.geometry === "roundRect" ? "roundRect" : "rect",
              style: {
                fill: colorTokenOf(theme, accent.fill.color, "primary", warnings, "アクセントの色"),
              },
            }),
          );
          continue;
        }
        if (isLineElement(accent)) {
          layer(
            index,
            absoluteChild(accent.box, {
              type: "shape",
              shape: "line",
              style: {
                line: {
                  color: colorTokenOf(theme, accent.line?.color, "primary", warnings, "アクセントの色"),
                  ...(accent.line?.widthPt ? { width: round(accent.line.widthPt) } : {}),
                },
              },
            }),
          );
          continue;
        }
        warnings.push(`アクセントとして解釈された要素を描画できませんでした: index ${index}`);
      }
    }
  } else if (interpretation.cells.length === 1) {
    warnings.push("反復セルが 1 件しかないため、反復構造にしませんでした");
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
        // 解釈の出どころ（heuristic / human、将来は ai）を成果物に残す
        interpretationSource: interpretation.interpretationSource,
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
      // 元 PPTX の描画順（Z-order）を保ったまま並べる。
      // Array.prototype.sort は安定なので、同じ sourceOrder の要素は追加順のまま。
      body: {
        type: "absolute",
        origin: "canvas",
        children: [...layers]
          .sort((a, b) => a.sourceOrder - b.sourceOrder)
          .map((entry) => entry.child),
      },
    },
    ...(options.assets && options.assets.length > 0
      ? { assets: options.assets.map((a) => ({ id: a.id, path: a.path, kind: "image" as const })) }
      : {}),
  };

  return { document, warnings };
}
