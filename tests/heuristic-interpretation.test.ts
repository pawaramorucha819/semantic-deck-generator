import { describe, expect, it } from "vitest";
import {
  HEURISTIC_THRESHOLDS,
  looksLikeMetric,
  ordinalToken,
  relativeBox,
  sameFontTier,
  sameRelativePosition,
} from "../src/patterns/capture/heuristic-roles.js";
import type { SlotInferenceResult } from "../src/patterns/capture/infer.js";
import {
  interpretFromSlotInference,
  parseAndValidateCaptureInterpretation,
  type CaptureInterpretation,
} from "../src/patterns/capture/interpretation.js";
import type { ExtractedElement } from "../src/pptx/model.js";
import type { Rect } from "../src/render/geometry.js";

/**
 * Phase 3: 決定論的ヒューリスティクスによる役割推定。
 *
 * ここで固定したいのは Recall ではなく **Precision**。
 * 「明白な反復規則を認識できる」ことと同じくらい、
 * **「曖昧なものを無理に推定しない（abstain する）」** ことを検証する。
 */

const CARD = { y: 1.6, w: 2.9, h: 3.4, gap: 0.2, firstX: 0.6 };
const cardBox = (cell: number): Rect => ({
  x: CARD.firstX + cell * (CARD.w + CARD.gap),
  y: CARD.y,
  w: CARD.w,
  h: CARD.h,
});

function element(overrides: Partial<ExtractedElement> = {}): ExtractedElement {
  return { id: "e", kind: "shape", box: { x: 0, y: 0, w: 1, h: 1 }, ...overrides };
}

function textElement(id: string, plain: string, fontSize: number, box: Rect): ExtractedElement {
  return element({
    id,
    kind: "text",
    box,
    text: { paragraphs: [plain], runs: [], plain, maxFontSize: fontSize },
  });
}

interface TextRow {
  /** セル 0..n のテキスト */
  values: string[];
  fontSize: number;
}

/** カード 1 枚に足す装飾（セルごとに変えられる） */
type Decoration = (card: Rect, cell: number) => ExtractedElement | undefined;

/**
 * 3 カードのスライドと、それに対応する SlotInference の結果を作る。
 * 要素順は 背景 → 装飾 → テキスト行（上から）で、元 PPTX の Z-order を模す。
 */
function makeSlide(
  rows: TextRow[],
  decorations: Decoration[] = [],
): { elements: ExtractedElement[]; inference: SlotInferenceResult } {
  const cellCount = rows[0]!.values.length;
  const elements: ExtractedElement[] = [];
  const cells: SlotInferenceResult["cells"] = [];
  const staticIndexes: number[] = [];

  for (let c = 0; c < cellCount; c += 1) {
    const card = cardBox(c);
    elements.push(
      element({
        id: `bg-${c}`,
        geometry: "roundRect",
        box: card,
        fill: { type: "solid", color: "F2F6FA" },
      }),
    );
    const backgroundIndex = elements.length - 1;

    for (const decorate of decorations) {
      const decoration = decorate(card, c);
      if (!decoration) continue;
      elements.push(decoration);
      staticIndexes.push(elements.length - 1);
    }

    const textIndexes: number[] = [];
    rows.forEach((row, r) => {
      elements.push(
        textElement(`r${r}-c${c}`, row.values[c]!, row.fontSize, {
          x: card.x + 0.2,
          y: card.y + (0.08 + r * 0.22) * card.h,
          w: card.w - 0.4,
          h: 0.3,
        }),
      );
      textIndexes.push(elements.length - 1);
    });

    cells.push({ backgroundIndex, textIndexes });
  }

  return { elements, inference: { cells, staticIndexes, notes: [] } };
}

const leftAccent: Decoration = (card) =>
  element({
    id: "accent",
    geometry: "rect",
    box: { x: card.x, y: card.y, w: 0.06, h: card.h },
    fill: { type: "solid", color: "0B4F8A" },
  });

const horizontalDivider: Decoration = (card) =>
  element({
    id: "divider",
    geometry: "line",
    box: { x: card.x + 0.2, y: card.y + card.h * 0.55, w: card.w - 0.4, h: 0 },
    fill: { type: "none" },
    line: { color: "D5DCE4" },
  });

/** 標準的な 3 カード（連番 / 見出し / 数値 / 説明 + 縁の帯 + 区切り線） */
function standardCards(overrides: Partial<Record<number, TextRow>> = {}) {
  const rows: TextRow[] = [
    { values: ["01", "02", "03"], fontSize: 11 },
    { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
    { values: ["30%", "−25%", "2.4×"], fontSize: 26 },
    {
      values: [
        "型化とレビュー準備の時間を短縮する",
        "チェック観点を標準化して属人性を下げる",
        "過去の設計を再利用して立ち上がりを早める",
      ],
      fontSize: 10.5,
    },
  ];
  for (const [k, row] of Object.entries(overrides)) {
    if (row) rows[Number(k)] = row;
  }
  return makeSlide(rows, [leftAccent, horizontalDivider]);
}

function interpret(slide: {
  elements: ExtractedElement[];
  inference: SlotInferenceResult;
}): CaptureInterpretation {
  const interpretation = interpretFromSlotInference(slide.elements, slide.inference);
  // ヒューリスティクスの出力も human と同じ安全ゲートを必ず通る
  parseAndValidateCaptureInterpretation(interpretation, slide.elements);
  return interpretation;
}

const notesOf = (interpretation: CaptureInterpretation): string =>
  (interpretation.notes ?? []).join("\n");

describe("テキストの形（内容ではなく書式だけを見る）", () => {
  it("裸の連番トークンだけを ordinal として認める", () => {
    expect(ordinalToken("01")).toEqual({ value: 1, width: 2 });
    expect(ordinalToken(" 3 ")).toEqual({ value: 3, width: 1 });
    expect(ordinalToken("001")).toEqual({ value: 1, width: 3 });
  });

  it("単位・符号・小数・桁数の多いものは ordinal にしない", () => {
    for (const text of ["30%", "-1", "−2", "2.4", "¥100", "2026年", "2024", "第1"]) {
      expect(ordinalToken(text)).toBeUndefined();
    }
  });

  it("数値が主役の短いテキストを metric として認める", () => {
    for (const text of ["30%", "−25%", "2.4×", "12件", "3部門", "¥100万", "+15pt"]) {
      expect(looksLikeMetric(text)).toBe(true);
    }
  });

  it("文章は metric にしない", () => {
    for (const text of [
      "前年比30%改善した",
      "2026年8月21日時点で30%改善した",
      "開発スピード",
      "レビュー時間を3割短縮する",
      "",
    ]) {
      expect(looksLikeMetric(text)).toBe(false);
    }
  });
});

describe("セル横断の比較 helper", () => {
  it("相対座標はセル矩形に対する比率になる", () => {
    const relative = relativeBox({ x: 1.2, y: 2.0, w: 1.0, h: 0.5 }, { x: 1.0, y: 1.6, w: 2.0, h: 4.0 });
    expect(relative.x).toBeCloseTo(0.1, 6);
    expect(relative.y).toBeCloseTo(0.1, 6);
    expect(relative.w).toBeCloseTo(0.5, 6);
    expect(relative.h).toBeCloseTo(0.125, 6);
  });

  it("相対位置は tolerance の範囲で同じとみなす", () => {
    const a = { x: 0.1, y: 0.1, w: 0.5, h: 0.1 };
    const within = { x: 0.1 + HEURISTIC_THRESHOLDS.relativePositionTolerance / 2, y: 0.1, w: 0.5, h: 0.1 };
    const outside = { x: 0.1 + HEURISTIC_THRESHOLDS.relativePositionTolerance * 2, y: 0.1, w: 0.5, h: 0.1 };
    expect(sameRelativePosition(a, within)).toBe(true);
    expect(sameRelativePosition(a, outside)).toBe(false);
  });

  it("フォント階層は相対差で判定する", () => {
    const at = (size: number) => textElement("t", "x", size, { x: 0, y: 0, w: 1, h: 1 });
    expect(sameFontTier(at(17), at(17.5))).toBe(true);
    expect(sameFontTier(at(17), at(10.5))).toBe(false);
  });
});

describe("index 推定（連番の系列として確認する）", () => {
  it("01 / 02 / 03 を index と推定する", () => {
    const interpretation = interpret(standardCards());
    expect(interpretation.cells.map((c) => c.content.indexIndex)).toEqual([3, 10, 17]);
    expect(notesOf(interpretation)).toContain("index と推定");
  });

  it("1 / 2 / 3 と 001 / 002 / 003 も index と推定する", () => {
    for (const values of [
      ["1", "2", "3"],
      ["001", "002", "003"],
    ]) {
      const interpretation = interpret(standardCards({ 0: { values, fontSize: 11 } }));
      expect(interpretation.cells[0]!.content.indexIndex).toBe(3);
    }
  });

  it("10% / 20% / 30% は index にしない（単位が付いている）", () => {
    const interpretation = interpret(
      standardCards({ 0: { values: ["10%", "20%", "30%"], fontSize: 11 } }),
    );
    expect(interpretation.cells.every((c) => c.content.indexIndex === undefined)).toBe(true);
  });

  it("2024 / 2025 / 2026 は index にしない（1 から始まる連番ではない）", () => {
    const interpretation = interpret(
      standardCards({ 0: { values: ["2024", "2025", "2026"], fontSize: 11 } }),
    );
    expect(interpretation.cells.every((c) => c.content.indexIndex === undefined)).toBe(true);
  });

  it("1 / 3 / 5 は index にしない（連続していない）", () => {
    const interpretation = interpret(
      standardCards({ 0: { values: ["1", "3", "5"], fontSize: 11 } }),
    );
    expect(interpretation.cells.every((c) => c.content.indexIndex === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("1 から始まる連番ではない");
  });

  it("01 / 02 / 04 は index にしない（欠番がある）", () => {
    const interpretation = interpret(
      standardCards({ 0: { values: ["01", "02", "04"], fontSize: 11 } }),
    );
    expect(interpretation.cells.every((c) => c.content.indexIndex === undefined)).toBe(true);
  });

  it("01 / 02 / 3 は index にしない（桁の書式が揃わない）", () => {
    const interpretation = interpret(
      standardCards({ 0: { values: ["01", "02", "3"], fontSize: 11 } }),
    );
    expect(interpretation.cells.every((c) => c.content.indexIndex === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("桁の書式が揃わない");
  });

  it("index として採用した連番を metric に再利用しない", () => {
    const interpretation = interpret(standardCards());
    for (const cell of interpretation.cells) {
      expect(cell.content.metricIndex).not.toBe(cell.content.indexIndex);
    }
    expect(interpretation.cells.map((c) => c.content.metricIndex)).toEqual([5, 12, 19]);
  });
});

describe("metric 推定", () => {
  it("同じ相対位置の数値テキストを metric と推定する", () => {
    const interpretation = interpret(standardCards());
    expect(interpretation.cells.map((c) => c.content.metricIndex)).toEqual([5, 12, 19]);
    expect(notesOf(interpretation)).toContain("metric と推定");
  });

  it("数字を含むだけの文章は metric にしない", () => {
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
        {
          values: ["前年比30%改善した", "前年比25%削減した", "前年比2倍に増えた"],
          fontSize: 10.5,
        },
      ],
      [leftAccent, horizontalDivider],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.content.metricIndex === undefined)).toBe(true);
    // label / description としては使える
    expect(interpretation.cells[0]!.content.labelIndex).toBe(4);
    expect(interpretation.cells[0]!.content.descriptionIndexes).toEqual([5]);
  });

  it("数値の行が複数あり決め手が無ければ metric にしない", () => {
    const slide = makeSlide([
      { values: ["12件", "8件", "5件"], fontSize: 17 },
      { values: ["30%", "25%", "12%"], fontSize: 17 },
    ]);
    const interpretation = interpret(slide);
    expect(interpretation.cells.length).toBe(0);
    expect(notesOf(interpretation)).toContain("決め手が無いため metric と推定しませんでした");
  });
});

describe("label / description の区別（フォント階層が明確なときだけ）", () => {
  it("フォント階層と配置順が明確なら label / description を分ける", () => {
    const interpretation = interpret(standardCards());
    expect(interpretation.cells.map((c) => c.content.labelIndex)).toEqual([4, 11, 18]);
    expect(interpretation.cells.map((c) => c.content.descriptionIndexes)).toEqual([[6], [13], [20]]);
    expect(notesOf(interpretation)).toContain("フォント階層");
  });

  it("残るテキストが 1 行なら label だけにする", () => {
    const slide = makeSlide([
      { values: ["01", "02", "03"], fontSize: 11 },
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
    ]);
    const interpretation = interpret(slide);
    expect(interpretation.cells.map((c) => c.content.labelIndex)).toEqual([2, 5, 8]);
    expect(interpretation.cells.every((c) => c.content.descriptionIndexes === undefined)).toBe(
      true,
    );
  });

  it("同じフォントサイズなら「上だから label」で決めない（abstain）", () => {
    const slide = makeSlide([
      { values: ["01", "02", "03"], fontSize: 11 },
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 14 },
      { values: ["時間を短縮する", "属人性を下げる", "立ち上がりを早める"], fontSize: 14 },
    ]);
    const interpretation = interpret(slide);
    expect(interpretation.cells).toEqual([]);
    expect(notesOf(interpretation)).toContain("フォント階層が不明確");
  });

  it("フォント差が閾値未満なら abstain する", () => {
    const slide = makeSlide([
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 12 },
      { values: ["時間を短縮する", "属人性を下げる", "早く始められる"], fontSize: 11 },
    ]);
    const interpretation = interpret(slide);
    expect(interpretation.cells).toEqual([]);
    expect(notesOf(interpretation)).toContain("フォント階層が不明確");
  });
});

describe("divider 推定（繰り返される軸沿いの線だけ）", () => {
  it("各セルの同じ位置にある水平線を divider と推定する", () => {
    const interpretation = interpret(standardCards());
    expect(interpretation.cells.map((c) => c.decoration?.dividerIndex)).toEqual([2, 9, 16]);
    expect(notesOf(interpretation)).toContain("水平線を divider と推定");
  });

  it("垂直線でも推定できる", () => {
    const verticalDivider: Decoration = (card) =>
      element({
        id: "vdiv",
        geometry: "line",
        box: { x: card.x + card.w * 0.5, y: card.y + 0.2, w: 0, h: card.h - 0.4 },
        line: { color: "D5DCE4" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [verticalDivider],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.dividerIndex !== undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("垂直線を divider と推定");
  });

  it("斜めのコネクタは divider にしない", () => {
    const diagonal: Decoration = (card) =>
      element({
        id: "arrow",
        kind: "line",
        box: { x: card.x + 0.3, y: card.y + 1.0, w: 1.6, h: 1.2 },
        line: { color: "D5DCE4" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [diagonal],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.dividerIndex === undefined)).toBe(true);
    // 認識できなかった線は固定デザインとして残る
    expect(interpretation.staticIndexes).toContain(1);
  });

  it("セルごとに線の位置が違えば divider にしない", () => {
    const shifted: Decoration = (card, cell) =>
      element({
        id: `div-${cell}`,
        geometry: "line",
        box: { x: card.x + 0.2, y: card.y + card.h * (0.3 + cell * 0.25), w: card.w - 0.4, h: 0 },
        line: { color: "D5DCE4" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [shifted],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.dividerIndex === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("相対位置が揃わない");
  });

  it("1 セルにしか線が無ければ divider にしない", () => {
    const onlyFirst: Decoration = (card, cell) =>
      cell === 0 ? horizontalDivider(card, cell) : undefined;
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [onlyFirst],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.dividerIndex === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("区切り線の本数が揃わない");
  });

  it("セル間で線の色が違えば divider にしない", () => {
    const recolored: Decoration = (card, cell) =>
      element({
        id: `div-${cell}`,
        geometry: "line",
        box: { x: card.x + 0.2, y: card.y + card.h * 0.55, w: card.w - 0.4, h: 0 },
        line: { color: cell === 1 ? "FF0000" : "D5DCE4" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [recolored],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.dividerIndex === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("区切り線の色が揃わない");
  });
});

describe("accent 推定（縁の細い帯だけ）", () => {
  it("各カード左端の細い帯を accent と推定する", () => {
    const interpretation = interpret(standardCards());
    expect(interpretation.cells.map((c) => c.decoration?.accentIndexes)).toEqual([[1], [8], [15]]);
    expect(notesOf(interpretation)).toContain("accent と推定");
  });

  it("大きな色面は accent にしない", () => {
    const bigShape: Decoration = (card) =>
      element({
        id: "panel",
        geometry: "rect",
        box: { x: card.x, y: card.y, w: card.w, h: card.h * 0.4 },
        fill: { type: "solid", color: "0B4F8A" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [bigShape],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.accentIndexes === undefined)).toBe(true);
    expect(interpretation.staticIndexes).toContain(1);
  });

  it("縁から離れたアイコン的な図形は accent にしない", () => {
    const icon: Decoration = (card) =>
      element({
        id: "icon",
        geometry: "rect",
        box: { x: card.x + card.w * 0.4, y: card.y + card.h * 0.4, w: 0.25, h: 0.25 },
        fill: { type: "solid", color: "0B4F8A" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [icon],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.accentIndexes === undefined)).toBe(true);
  });

  it("セルによって accent の数が違えば採用しない", () => {
    const extra: Decoration = (card, cell) =>
      cell === 2
        ? element({
            id: "accent-top",
            geometry: "rect",
            box: { x: card.x, y: card.y, w: card.w, h: 0.05 },
            fill: { type: "solid", color: "0B4F8A" },
          })
        : undefined;
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [leftAccent, extra],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.accentIndexes === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("アクセントの数が揃わない");
  });

  it("セル間で帯の色が違えば accent にしない", () => {
    const recolored: Decoration = (card, cell) =>
      element({
        id: `accent-${cell}`,
        geometry: "rect",
        box: { x: card.x, y: card.y, w: 0.06, h: card.h },
        fill: { type: "solid", color: cell === 1 ? "FF0000" : "0B4F8A" },
      });
    const slide = makeSlide(
      [
        { values: ["01", "02", "03"], fontSize: 11 },
        { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
      ],
      [recolored],
    );
    const interpretation = interpret(slide);
    expect(interpretation.cells.every((c) => c.decoration?.accentIndexes === undefined)).toBe(true);
    expect(notesOf(interpretation)).toContain("塗り・線が揃わない");
  });
});

describe("abstain（曖昧なら反復構造を作らない）", () => {
  it("セルごとのテキスト数が揃わなければ推定を中止する", () => {
    const slide = makeSlide([
      { values: ["01", "02", "03"], fontSize: 11 },
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 17 },
    ]);
    // 1 セルだけテキストを 1 つ減らす
    slide.inference.cells[1]!.textIndexes = slide.inference.cells[1]!.textIndexes.slice(0, 1);

    const interpretation = interpret(slide);
    expect(interpretation.cells).toEqual([]);
    expect(notesOf(interpretation)).toContain("テキスト数が揃わない");
  });

  it("セル間で相対位置が揃わなければ推定を中止する", () => {
    const slide = standardCards();
    // cell 1 の見出しだけ大きく下げる
    slide.elements[10]!.box = { ...slide.elements[10]!.box, y: slide.elements[10]!.box.y + 1.2 };
    const interpretation = interpret(slide);
    expect(interpretation.cells).toEqual([]);
    expect(notesOf(interpretation)).toContain("相対位置が揃わない");
  });

  it("abstain しても要素を捨てず、固定デザインとして残す", () => {
    const slide = makeSlide([
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 14 },
      { values: ["時間を短縮する", "属人性を下げる", "早く始める"], fontSize: 14 },
    ]);
    const interpretation = interpret(slide);
    expect(interpretation.cells).toEqual([]);
    // 背景もテキストも固定デザインとして残っている
    for (let index = 0; index < slide.elements.length; index += 1) {
      expect(interpretation.staticIndexes).toContain(index);
    }
  });

  it("abstain した解釈も検証を通り、Compiler へ安全に渡せる", () => {
    const slide = makeSlide([
      { values: ["開発スピード", "品質の安定", "知識の再利用"], fontSize: 14 },
      { values: ["時間を短縮する", "属人性を下げる", "早く始める"], fontSize: 14 },
    ]);
    expect(() =>
      parseAndValidateCaptureInterpretation(
        interpretFromSlotInference(slide.elements, slide.inference),
        slide.elements,
      ),
    ).not.toThrow();
  });
});

describe("Interpretation としての整合", () => {
  it("役割として採用した装飾は staticIndexes に残さない", () => {
    const slide = standardCards();
    const interpretation = interpret(slide);
    const claimed = interpretation.cells.flatMap((cell) => [
      cell.backgroundIndex!,
      cell.content.indexIndex!,
      cell.content.labelIndex,
      cell.content.metricIndex!,
      ...(cell.content.descriptionIndexes ?? []),
      ...(cell.decoration?.accentIndexes ?? []),
      cell.decoration!.dividerIndex!,
    ]);
    expect(claimed).toHaveLength(21);
    for (const index of claimed) expect(interpretation.staticIndexes).not.toContain(index);
    expect(interpretation.staticIndexes).toEqual([]);
  });

  it("全セルの Visual Role 構成が揃う（Phase 2.1 の検証を通る）", () => {
    const interpretation = interpret(standardCards());
    const signature = (cell: (typeof interpretation.cells)[number]) => ({
      background: cell.backgroundIndex !== undefined,
      index: cell.content.indexIndex !== undefined,
      metric: cell.content.metricIndex !== undefined,
      description: (cell.content.descriptionIndexes ?? []).length,
      accent: (cell.decoration?.accentIndexes ?? []).length,
      divider: cell.decoration?.dividerIndex !== undefined,
    });
    const first = signature(interpretation.cells[0]!);
    for (const cell of interpretation.cells) expect(signature(cell)).toEqual(first);
  });

  it("同じ入力からは必ず同じ解釈になる（決定論的）", () => {
    const a = interpretFromSlotInference(standardCards().elements, standardCards().inference);
    const b = interpretFromSlotInference(standardCards().elements, standardCards().inference);
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });

  it("ページ番号プレースホルダは Pattern へ取り込まない（Template の責務）", () => {
    const slide = standardCards();
    slide.elements.push(
      element({
        id: "slide-number",
        kind: "placeholder",
        box: { x: 9.0, y: 5.2, w: 0.5, h: 0.2 },
        placeholder: { type: "sldNum" },
        text: { paragraphs: ["2"], runs: [], plain: "2" },
      }),
    );
    slide.inference.staticIndexes.push(slide.elements.length - 1);

    const interpretation = interpret(slide);
    expect(interpretation.staticIndexes).not.toContain(slide.elements.length - 1);
    expect(notesOf(interpretation)).toContain("Template の責務");
  });
});
