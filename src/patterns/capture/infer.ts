import { z } from "zod";
import type { ExtractedElement } from "../../pptx/model.js";
import type { Rect } from "../../render/geometry.js";
import { isTemplatePlaceholder } from "./element-roles.js";

/**
 * Slot 推定 — 「固定デザイン」と「毎回差し替わるコンテンツ」を分ける工程。
 *
 * ここは AI を差し込める場所として設計してある（`SlotInferencer` 実装を差し替える）。
 * ただし **AI が無くても必ず動く**: 既定は幾何とフォントサイズだけを見る
 * 決定論的ヒューリスティクス。AI の出力も同じ `SlotInferenceResult` スキーマで
 * 検証され、範囲外のインデックスは拒否される。
 *
 * どちらの場合でも結果は「下書き」であり、人間のレビューを経て登録される。
 */

export interface SlotInferenceInput {
  slideNumber: number;
  slideSize: { widthIn: number; heightIn: number };
  elements: ExtractedElement[];
}

/** 反復要素 1 つ分（カード 1 枚に相当） */
export interface InferredCell {
  /** 背景図形の要素インデックス */
  backgroundIndex?: number;
  /** セル内のテキスト要素インデックス（上から順） */
  textIndexes: number[];
}

export interface SlotInferenceResult {
  /** KeyMessage に対応するテキスト要素 */
  headlineIndex?: number;
  /** SupportingMessage に対応するテキスト要素 */
  sublineIndex?: number;
  /** 役割ラベル（「課題」など）に対応するテキスト要素 */
  eyebrowIndex?: number;
  /** 出典行に対応するテキスト要素 */
  sourceIndex?: number;
  /** 反復要素（3 カードなど） */
  cells: InferredCell[];
  /** 固定デザインとして扱う要素 */
  staticIndexes: number[];
  notes: string[];
}

export const SlotInferenceResultSchema = z
  .object({
    headlineIndex: z.number().int().min(0).optional(),
    sublineIndex: z.number().int().min(0).optional(),
    eyebrowIndex: z.number().int().min(0).optional(),
    sourceIndex: z.number().int().min(0).optional(),
    cells: z
      .array(
        z
          .object({
            backgroundIndex: z.number().int().min(0).optional(),
            textIndexes: z.array(z.number().int().min(0)).max(12),
          })
          .strict(),
      )
      .max(12),
    staticIndexes: z.array(z.number().int().min(0)).max(200),
    notes: z.array(z.string().max(400)).max(50),
  })
  .strict();

export interface SlotInferencer {
  readonly id: string;
  infer(input: SlotInferenceInput): Promise<SlotInferenceResult>;
}

export class SlotInferenceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SlotInferenceError";
  }
}

/**
 * 推定結果の妥当性検証。
 * AI 由来でもヒューリスティクス由来でも、必ずここを通してから使う。
 */
export function validateInference(
  raw: unknown,
  elementCount: number,
): SlotInferenceResult {
  const parsed = SlotInferenceResultSchema.safeParse(raw);
  if (!parsed.success) {
    throw new SlotInferenceError(
      `Slot 推定結果の形式が不正です:\n${parsed.error.issues
        .map((i) => `  - ${i.path.join(".")}: ${i.message}`)
        .join("\n")}`,
    );
  }
  const result = parsed.data as SlotInferenceResult;

  const used = new Set<number>();
  const check = (index: number | undefined, where: string): void => {
    if (index === undefined) return;
    if (index >= elementCount) {
      throw new SlotInferenceError(
        `${where}: 存在しない要素インデックスです (${index} >= ${elementCount})`,
      );
    }
    if (used.has(index)) {
      throw new SlotInferenceError(`要素 ${index} が複数の役割に割り当てられています (${where})`);
    }
    used.add(index);
  };

  check(result.headlineIndex, "headlineIndex");
  check(result.sublineIndex, "sublineIndex");
  check(result.eyebrowIndex, "eyebrowIndex");
  check(result.sourceIndex, "sourceIndex");
  result.cells.forEach((cell, i) => {
    check(cell.backgroundIndex, `cells[${i}].backgroundIndex`);
    cell.textIndexes.forEach((index, j) => check(index, `cells[${i}].textIndexes[${j}]`));
  });
  for (const index of result.staticIndexes) check(index, "staticIndexes");

  // セルの構造は揃っている必要がある（1 つだけ項目数が違う = 反復ではない）
  const counts = new Set(result.cells.map((c) => c.textIndexes.length));
  if (result.cells.length > 0 && counts.size > 1) {
    result.notes.push(
      "セルごとにテキスト数が異なります。反復要素として扱うため、共通部分のみを使用します",
    );
  }
  return result;
}

const center = (box: Rect): { x: number; y: number } => ({
  x: box.x + box.w / 2,
  y: box.y + box.h / 2,
});

const containsPoint = (box: Rect, point: { x: number; y: number }): boolean =>
  point.x >= box.x - 0.02 &&
  point.x <= box.x + box.w + 0.02 &&
  point.y >= box.y - 0.02 &&
  point.y <= box.y + box.h + 0.02;

const near = (a: number, b: number, tolerance: number): boolean => Math.abs(a - b) <= tolerance;

function fontSizeOf(element: ExtractedElement): number {
  return element.text?.maxFontSize ?? 12;
}

function hasText(element: ExtractedElement): boolean {
  return !!element.text && element.text.plain.trim().length > 0;
}

/**
 * 決定論的な Slot 推定（AI 不要の既定実装）。
 *
 * 見ているのは「幾何」と「フォントサイズ」だけ:
 *  - 上部で最大フォントのテキスト → 見出し（keyMessage）
 *  - 同じ大きさの図形が横に 2 つ以上並んでいる → 反復セル
 *  - 下端の小さなテキスト → 出典
 *  - 残り（ロゴ・装飾・固定ラベル） → 固定デザイン
 */
export class HeuristicSlotInferencer implements SlotInferencer {
  readonly id = "heuristic";

  // eslint-disable-next-line @typescript-eslint/require-await
  async infer(input: SlotInferenceInput): Promise<SlotInferenceResult> {
    return this.inferSync(input);
  }

  inferSync(input: SlotInferenceInput): SlotInferenceResult {
    const { elements, slideSize } = input;
    const notes: string[] = [];
    const assigned = new Set<number>();

    // ページ番号・フッターのプレースホルダは Template の責務なので、
    // 見出しや出典の候補にしない（"2" というページ番号を出典と推定しないため）
    const textIndexes = elements
      .map((element, index) => ({ element, index }))
      .filter(({ element }) => hasText(element) && !isTemplatePlaceholder(element));

    // --- 見出し -------------------------------------------------------------
    const topBand = slideSize.heightIn * 0.45;
    const headlineCandidates = textIndexes.filter(({ element }) => element.box.y <= topBand);
    const headline = [...(headlineCandidates.length > 0 ? headlineCandidates : textIndexes)].sort(
      (a, b) =>
        fontSizeOf(b.element) - fontSizeOf(a.element) ||
        a.element.box.y - b.element.box.y ||
        a.element.box.x - b.element.box.x,
    )[0];
    if (headline) {
      assigned.add(headline.index);
      notes.push(
        `見出しを推定: "${truncate(headline.element.text!.plain)}" (${fontSizeOf(headline.element)}pt)`,
      );
    } else {
      notes.push("テキストが 1 つも無いため見出しを推定できませんでした");
    }

    const headlineSize = headline ? fontSizeOf(headline.element) : 0;

    // --- 役割ラベル（見出しの上にある小さな短いテキスト） ---------------------
    let eyebrowIndex: number | undefined;
    if (headline) {
      const candidate = textIndexes
        .filter(({ index, element }) => {
          if (assigned.has(index)) return false;
          return (
            element.box.y < headline.element.box.y &&
            fontSizeOf(element) < headlineSize * 0.75 &&
            element.text!.plain.trim().length <= 20
          );
        })
        .sort((a, b) => b.element.box.y - a.element.box.y)[0];
      if (candidate) {
        eyebrowIndex = candidate.index;
        assigned.add(candidate.index);
      }
    }

    // --- 補足メッセージ（見出しのすぐ下） -------------------------------------
    let sublineIndex: number | undefined;
    if (headline) {
      const headlineBottom = headline.element.box.y + headline.element.box.h;
      const candidate = textIndexes
        .filter(({ index, element }) => {
          if (assigned.has(index)) return false;
          return (
            element.box.y >= headlineBottom - 0.05 &&
            element.box.y <= headlineBottom + 0.9 &&
            fontSizeOf(element) < headlineSize &&
            element.box.w >= headline.element.box.w * 0.5
          );
        })
        .sort((a, b) => a.element.box.y - b.element.box.y)[0];
      if (candidate) {
        sublineIndex = candidate.index;
        assigned.add(candidate.index);
      }
    }

    // --- 出典行（下端の小さなテキスト） ---------------------------------------
    let sourceIndex: number | undefined;
    const bottomBand = slideSize.heightIn * 0.86;
    const sourceCandidate = textIndexes
      .filter(({ index, element }) => {
        if (assigned.has(index)) return false;
        return element.box.y >= bottomBand && fontSizeOf(element) <= Math.max(11, headlineSize * 0.5);
      })
      .sort((a, b) => b.element.box.y - a.element.box.y)[0];
    if (sourceCandidate) {
      sourceIndex = sourceCandidate.index;
      assigned.add(sourceCandidate.index);
    }

    // --- 反復セル -------------------------------------------------------------
    const cells = this.detectCells(elements, assigned, notes);
    for (const cell of cells) {
      if (cell.backgroundIndex !== undefined) assigned.add(cell.backgroundIndex);
      for (const index of cell.textIndexes) assigned.add(index);
    }

    const staticIndexes = elements
      .map((_, index) => index)
      .filter((index) => !assigned.has(index));

    if (cells.length === 0) {
      notes.push(
        "反復構造を検出できませんでした。固定レイアウトの Pattern として下書きします（必要なら items slot を手で追加してください）",
      );
    } else {
      notes.push(`${cells.length} 件の反復セルを検出しました`);
    }

    return {
      ...(headline ? { headlineIndex: headline.index } : {}),
      ...(sublineIndex !== undefined ? { sublineIndex } : {}),
      ...(eyebrowIndex !== undefined ? { eyebrowIndex } : {}),
      ...(sourceIndex !== undefined ? { sourceIndex } : {}),
      cells,
      staticIndexes,
      notes,
    };
  }

  /** 同じ大きさで横に並ぶ図形（またはテキスト群）を反復セルとみなす */
  private detectCells(
    elements: ExtractedElement[],
    assigned: Set<number>,
    notes: string[],
  ): InferredCell[] {
    const shapes = elements
      .map((element, index) => ({ element, index }))
      .filter(
        ({ element, index }) =>
          !assigned.has(index) &&
          (element.kind === "shape" || element.kind === "text") &&
          element.fill?.type === "solid",
      );

    // 幅・高さ・y がほぼ同じ図形のグループを探す
    const groups: Array<Array<{ element: ExtractedElement; index: number }>> = [];
    for (const candidate of shapes) {
      const group = groups.find((g) => {
        const first = g[0]!.element.box;
        return (
          near(first.w, candidate.element.box.w, Math.max(0.08, first.w * 0.12)) &&
          near(first.h, candidate.element.box.h, Math.max(0.08, first.h * 0.15)) &&
          near(first.y, candidate.element.box.y, 0.25)
        );
      });
      if (group) group.push(candidate);
      else groups.push([candidate]);
    }

    const repeated = groups
      .filter((g) => g.length >= 2)
      .sort((a, b) => b.length - a.length)[0];
    if (!repeated) return [];

    const sorted = [...repeated].sort((a, b) => a.element.box.x - b.element.box.x);
    if (sorted.length > 6) {
      notes.push(`反復セルが ${sorted.length} 件見つかりましたが、6 件までに制限します`);
      sorted.length = 6;
    }

    return sorted.map(({ element, index }) => {
      const inner = elements
        .map((el, i) => ({ el, i }))
        .filter(
          ({ el, i }) =>
            i !== index &&
            !assigned.has(i) &&
            hasText(el) &&
            containsPoint(element.box, center(el.box)),
        )
        .sort((a, b) => a.el.box.y - b.el.box.y)
        .map(({ i }) => i);

      // 図形自体がテキストを持つ場合（PowerPoint ではよくある）はセル本文として扱う
      const ownText = hasText(element);
      return {
        ...(ownText ? {} : { backgroundIndex: index }),
        textIndexes: ownText ? [index, ...inner] : inner,
      };
    });
  }
}

function truncate(text: string, max = 24): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > max ? `${flat.slice(0, max)}…` : flat;
}

export const heuristicSlotInferencer = new HeuristicSlotInferencer();
