import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildCapturedPattern } from "../src/patterns/capture/build.js";
import {
  CAPTURE_INTERPRETATION_SCHEMA_VERSION,
  CaptureInterpretationError,
  assignmentsOf,
  cellRoleSignature,
  interpretFromSlotInference,
  parseAndValidateCaptureInterpretation,
  parseCaptureInterpretation,
  type CaptureInterpretation,
} from "../src/patterns/capture/interpretation.js";
import { compileDeclarativePattern } from "../src/patterns/declarative/compile.js";
import { lookupPath } from "../src/patterns/declarative/values.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import type { ExtractedElement } from "../src/pptx/model.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { defaultTheme } from "../src/theme/theme.js";
import {
  extractTexts,
  metricEvidence,
  pointsContent,
  projectRoot,
  readPptx,
  slide,
} from "./helpers.js";

/**
 * Phase 2 / 2.1: 解釈（Interpretation）と Pattern Compiler。
 *
 * ここで固定したいのは「意味推定の賢さ」ではなく、
 * **正しく解釈された入力なら正しい Pattern が出る** という変換の正確さと、
 * **壊れた解釈は Compiler へ到達しない** という境界。
 */

function element(overrides: Partial<ExtractedElement> = {}): ExtractedElement {
  return { id: "e", kind: "shape", box: { x: 0, y: 0, w: 1, h: 1 }, ...overrides };
}

function text(
  id: string,
  plain: string,
  fontSize: number,
  box: { x: number; y: number; w: number; h: number },
): ExtractedElement {
  return element({
    id,
    kind: "text",
    box,
    text: { paragraphs: [plain], runs: [], plain, maxFontSize: fontSize },
  });
}

/**
 * 3 カードのスライドを模した抽出結果。
 * 各カードは 背景 → アクセント → 連番 → 見出し → 数値 → 区切り線 → 説明 の順（Z-order）。
 *
 *   0            headline
 *   base+0 (1)   背景 roundRect（solid fill）
 *   base+1 (2)   アクセント rect（solid fill）
 *   base+2 (3)   連番 "01"
 *   base+3 (4)   見出し
 *   base+4 (5)   数値
 *   base+5 (6)   区切り線（h = 0）
 *   base+6 (7)   説明
 */
function cardSlideElements(): ExtractedElement[] {
  const elements: ExtractedElement[] = [
    text("headline", "3 つの観点で効果が出る", 24, { x: 0.6, y: 0.5, w: 8.8, h: 0.6 }),
  ];
  for (let i = 0; i < 3; i += 1) {
    const x = 0.6 + i * 3.1;
    elements.push(
      element({
        id: `bg-${i}`,
        geometry: "roundRect",
        box: { x, y: 1.6, w: 2.9, h: 3.4 },
        fill: { type: "solid", color: "F2F6FA" },
      }),
      element({
        id: `accent-${i}`,
        geometry: "rect",
        box: { x, y: 1.6, w: 0.06, h: 3.4 },
        fill: { type: "solid", color: "0B4F8A" },
      }),
      text(`index-${i}`, `0${i + 1}`, 11, { x: x + 0.2, y: 1.8, w: 0.5, h: 0.24 }),
      text(`label-${i}`, `見出し${i + 1}`, 17, { x: x + 0.2, y: 2.2, w: 2.5, h: 0.4 }),
      text(`metric-${i}`, `${(i + 1) * 10}%`, 26, { x: x + 0.2, y: 2.75, w: 2.5, h: 0.55 }),
      element({
        id: `divider-${i}`,
        geometry: "line",
        box: { x: x + 0.2, y: 3.45, w: 2.5, h: 0 },
        fill: { type: "none" },
        line: { color: "D5DCE4", widthPt: 1 },
      }),
      text(`desc-${i}`, `説明文${i + 1}`, 10.5, { x: x + 0.2, y: 3.6, w: 2.5, h: 1.2 }),
    );
  }
  return elements;
}

const ELEMENT_COUNT = 22;
const cellBase = (i: number): number => 1 + i * 7;

/** cardSlideElements() に対応する、人手で書いた想定の解釈 */
function humanInterpretation(
  overrides: Partial<CaptureInterpretation> = {},
): CaptureInterpretation {
  const cells = [0, 1, 2].map((i) => {
    const base = cellBase(i);
    return {
      backgroundIndex: base,
      content: {
        indexIndex: base + 2,
        labelIndex: base + 3,
        metricIndex: base + 4,
        descriptionIndexes: [base + 6],
      },
      decoration: { accentIndexes: [base + 1], dividerIndex: base + 5 },
    };
  });
  return {
    schemaVersion: CAPTURE_INTERPRETATION_SCHEMA_VERSION,
    interpretationSource: "human",
    headlineIndex: 0,
    cells,
    staticIndexes: [],
    ...overrides,
  };
}

function compile(interpretation: CaptureInterpretation, elements = cardSlideElements()) {
  return buildCapturedPattern(elements, interpretation, {
    id: "captured",
    theme: defaultTheme,
    slideSize: { widthIn: 10, heightIn: 5.625 },
    source: { file: "test.pptx", slideNumber: 1 },
  });
}

type LayoutChild = { box: { w: number; h: number }; node: Record<string, any> };

function bodyChildren(interpretation: CaptureInterpretation): LayoutChild[] {
  const body = compile(interpretation).document.layout.body as { children: LayoutChild[] };
  return body.children;
}

function cellStack(interpretation: CaptureInterpretation): Array<Record<string, any>> {
  const columns = bodyChildren(interpretation).find((c) => c.node.type === "columns")!;
  return columns.node.item.child.children as Array<Record<string, any>>;
}

/** 要素を 1 つだけ差し替えた抽出結果（役割 ↔ 要素種別の検証用） */
function withElement(index: number, replacement: ExtractedElement): ExtractedElement[] {
  const elements = cardSlideElements();
  elements[index] = replacement;
  return elements;
}

const imageElement = element({
  id: "img",
  kind: "image",
  box: { x: 1, y: 2, w: 1, h: 1 },
  image: { relationshipId: "rId9" },
});
const plainText = text("plain", "ただの文字", 12, { x: 1, y: 2, w: 1, h: 0.3 });
const filledRect = element({
  id: "rect",
  geometry: "rect",
  box: { x: 1, y: 2, w: 1, h: 1 },
  fill: { type: "solid", color: "112233" },
});
const horizontalLine = element({
  id: "hline",
  geometry: "line",
  box: { x: 1, y: 2, w: 2, h: 0 },
  fill: { type: "none" },
  line: { color: "D5DCE4" },
});
const verticalLine = element({
  id: "vline",
  geometry: "line",
  box: { x: 1, y: 2, w: 0, h: 2 },
  fill: { type: "none" },
  line: { color: "D5DCE4" },
});
const connectorLine = element({
  id: "connector",
  kind: "line",
  box: { x: 1, y: 2, w: 2, h: 0 },
  line: { color: "D5DCE4" },
});

describe("Interpretation スキーマ（untrusted input として検証する）", () => {
  it("正しい解釈を受け入れる", () => {
    const parsed = parseCaptureInterpretation(humanInterpretation(), ELEMENT_COUNT);
    expect(parsed.interpretationSource).toBe("human");
    expect(parsed.cells).toHaveLength(3);
  });

  it("要素数を超えるインデックスを拒否する", () => {
    expect(() => parseCaptureInterpretation(humanInterpretation(), 10)).toThrow(
      /存在しない要素インデックス/,
    );
  });

  it("同じ要素を複数の役割へ割り当てると拒否する", () => {
    const conflict = humanInterpretation();
    conflict.cells[0]!.content.metricIndex = conflict.cells[0]!.content.labelIndex;
    expect(() => parseCaptureInterpretation(conflict, ELEMENT_COUNT)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("content と decoration の衝突も検出する", () => {
    const conflict = humanInterpretation();
    conflict.cells[0]!.decoration = {
      accentIndexes: [conflict.cells[0]!.content.labelIndex],
      dividerIndex: 6,
    };
    expect(() => parseCaptureInterpretation(conflict, ELEMENT_COUNT)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("static と cell の衝突も検出する", () => {
    const conflict = humanInterpretation({ staticIndexes: [4] });
    expect(() => parseCaptureInterpretation(conflict, ELEMENT_COUNT)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("label が無いセルを拒否する", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation())) as {
      cells: Array<{ content: { labelIndex?: number } }>;
    };
    delete invalid.cells[0]!.content.labelIndex;
    expect(() => parseCaptureInterpretation(invalid, ELEMENT_COUNT)).toThrow(
      CaptureInterpretationError,
    );
  });

  it("未知のキーを拒否する（strict schema）", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation()));
    invalid.cells[0].content.tagline = 3;
    expect(() => parseCaptureInterpretation(invalid, ELEMENT_COUNT)).toThrow(
      CaptureInterpretationError,
    );
    expect(() =>
      parseCaptureInterpretation({ ...humanInterpretation(), script: "x" }, ELEMENT_COUNT),
    ).toThrow(CaptureInterpretationError);
  });

  it("負のインデックス・小数を拒否する", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation()));
    invalid.cells[0].content.labelIndex = -1;
    expect(() => parseCaptureInterpretation(invalid, ELEMENT_COUNT)).toThrow(
      CaptureInterpretationError,
    );
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() =>
      parseCaptureInterpretation(humanInterpretation({ schemaVersion: "9.9" }), ELEMENT_COUNT),
    ).toThrow(/schemaVersion/);
  });

  it("未知の interpretationSource（例: ai）は今は拒否する", () => {
    expect(() =>
      parseCaptureInterpretation(
        { ...humanInterpretation(), interpretationSource: "ai" },
        ELEMENT_COUNT,
      ),
    ).toThrow(CaptureInterpretationError);
  });

  it("要素 → 役割の逆引きが取れる（Capture Report 用）", () => {
    const assignments = assignmentsOf(humanInterpretation());
    expect(assignments.get(4)).toEqual({
      scope: "cell",
      cell: 0,
      kind: "content",
      role: "label",
    });
    expect(assignments.get(2)).toEqual({
      scope: "cell",
      cell: 0,
      kind: "structural",
      role: "accent",
    });
    expect(assignments.get(6)).toEqual({
      scope: "cell",
      cell: 0,
      kind: "structural",
      role: "divider",
    });
  });
});

describe("反復セルの構造一致（Content Role + Structural Role）", () => {
  const reject = (mutate: (interpretation: CaptureInterpretation) => void, expected: RegExp) => {
    const interpretation = humanInterpretation();
    mutate(interpretation);
    expect(() => parseCaptureInterpretation(interpretation, ELEMENT_COUNT)).toThrow(expected);
  };

  it("3 セルすべて同じ役割構成なら通る", () => {
    expect(() => parseCaptureInterpretation(humanInterpretation(), ELEMENT_COUNT)).not.toThrow();
    expect(cellRoleSignature(humanInterpretation().cells[0]!)).toEqual({
      background: true,
      index: true,
      label: true,
      metric: true,
      description: true,
      accentCount: 1,
      divider: true,
    });
  });

  it("background の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      delete i.cells[1]!.backgroundIndex;
    }, /background: expected=true, actual=false/);
  });

  it("index の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      delete i.cells[1]!.content.indexIndex;
    }, /index: expected=true, actual=false/);
  });

  it("metric の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      delete i.cells[1]!.content.metricIndex;
    }, /metric: expected=true, actual=false/);
  });

  it("description の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      delete i.cells[1]!.content.descriptionIndexes;
    }, /description: expected=true, actual=false/);
  });

  it("accent の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      i.cells[1]!.decoration = { dividerIndex: cellBase(1) + 5 };
    }, /accentCount: expected=1, actual=0/);
  });

  it("divider の有無がセル間で違うと拒否する", () => {
    reject((i) => {
      i.cells[1]!.decoration = { accentIndexes: [cellBase(1) + 1] };
    }, /divider: expected=true, actual=false/);
  });

  it("accent の個数がセル間で違うと拒否する（同じ Visual Structure ではない）", () => {
    const interpretation = humanInterpretation();
    // 区切り線をアクセント扱いにして、cell 0 だけアクセントを 2 つにする
    for (const [i, cell] of interpretation.cells.entries()) {
      cell.decoration = { accentIndexes: [cellBase(i) + 1] };
    }
    interpretation.cells[0]!.decoration = { accentIndexes: [cellBase(0) + 1, cellBase(0) + 5] };
    expect(() => parseCaptureInterpretation(interpretation, ELEMENT_COUNT)).toThrow(
      /accentCount: expected=2, actual=1/,
    );
  });

  it("どのセルが違うかがエラーに出る", () => {
    const interpretation = humanInterpretation();
    delete interpretation.cells[2]!.content.metricIndex;
    expect(() => parseCaptureInterpretation(interpretation, ELEMENT_COUNT)).toThrow(
      /cells\[2\] の Visual Role 構成が cells\[0\] と一致しません/,
    );
  });

  it("セルが 1 件だけの場合を拒否する", () => {
    const single = humanInterpretation({ cells: [humanInterpretation().cells[0]!] });
    expect(() => parseCaptureInterpretation(single, ELEMENT_COUNT)).toThrow(/2 件以上/);
  });
});

describe("役割 ↔ 要素種別の整合（描画できる要素だけを許す）", () => {
  const validate = (interpretation: CaptureInterpretation, elements: ExtractedElement[]) =>
    parseAndValidateCaptureInterpretation(interpretation, elements);

  const accept = (elements: ExtractedElement[]) => {
    expect(() => validate(humanInterpretation(), elements)).not.toThrow();
  };
  const reject = (elements: ExtractedElement[], expected: RegExp) => {
    expect(() => validate(humanInterpretation(), elements)).toThrow(expected);
  };

  describe("Content Role は text であること", () => {
    it("label → text は許可", () => accept(cardSlideElements()));

    it("label → rect は拒否", () => {
      reject(
        withElement(4, filledRect),
        /cells\[0\]\.content\.labelIndex: 要素 4 \(shape\(rect\)\) は label として利用できません/,
      );
    });

    it("metric → image は拒否", () => {
      reject(withElement(5, imageElement), /cells\[0\]\.content\.metricIndex: .*metric として/);
    });

    it("description → 空テキストは拒否", () => {
      reject(
        withElement(7, text("blank", "   ", 10, { x: 1, y: 4, w: 2, h: 0.5 })),
        /cells\[0\]\.content\.descriptionIndexes\[0\]: .*description として/,
      );
    });

    it("index → text は許可（内容が数字らしいかは判定しない）", () => {
      accept(withElement(3, text("idx", "壱", 11, { x: 1, y: 1.8, w: 0.5, h: 0.3 })));
    });

    it("意味の妥当性は判定しない（数値を label と解釈しても text なら通す）", () => {
      accept(withElement(4, text("num", "30%", 17, { x: 1, y: 2.2, w: 2, h: 0.4 })));
    });
  });

  describe("Slide-level Role は text であること", () => {
    it("headline → text は許可", () => accept(cardSlideElements()));

    it("headline → 図形のみは拒否", () => {
      reject(withElement(0, filledRect), /headlineIndex: 要素 0 .* headline として利用できません/);
    });

    it("sourceText → text は許可", () => {
      const elements = cardSlideElements();
      elements.push(text("src", "出典: 社内調査", 9, { x: 0.6, y: 5.1, w: 4, h: 0.2 }));
      expect(() =>
        validate(humanInterpretation({ sourceTextIndex: 22 }), elements),
      ).not.toThrow();
    });

    it("sourceText → 画像は拒否", () => {
      const elements = [...cardSlideElements(), imageElement];
      expect(() => validate(humanInterpretation({ sourceTextIndex: 22 }), elements)).toThrow(
        /sourceTextIndex: .* sourceText として/,
      );
    });
  });

  describe("divider は線として描ける要素であること", () => {
    it("水平線 (h = 0) は許可", () => accept(withElement(6, horizontalLine)));
    it("垂直線 (w = 0) は許可", () => accept(withElement(6, verticalLine)));
    it("コネクタ (kind = line) は許可", () => accept(withElement(6, connectorLine)));

    it("text は拒否", () => {
      reject(
        withElement(6, plainText),
        /cells\[0\]\.decoration\.dividerIndex: 要素 6 \(text "ただの文字"\) は divider として利用できません \(線 \(line\) 要素が必要です\)/,
      );
    });

    it("rect は拒否", () => {
      reject(withElement(6, filledRect), /dividerIndex: .* divider として/);
    });

    it("長さ 0 の線は拒否（描画できない）", () => {
      reject(
        withElement(6, element({ id: "dot", geometry: "line", box: { x: 1, y: 2, w: 0, h: 0 } })),
        /dividerIndex: .* divider として/,
      );
    });
  });

  describe("background は塗りを持つ図形であること", () => {
    it("solid fill のカード図形は許可", () => accept(cardSlideElements()));

    it("line は拒否", () => {
      reject(
        withElement(1, horizontalLine),
        /cells\[0\]\.backgroundIndex: .* background として利用できません \(塗り \(solid fill\) を持つ図形が必要です\)/,
      );
    });

    it("塗りの無いテキストは拒否", () => {
      reject(withElement(1, plainText), /backgroundIndex: .* background として/);
    });

    it("画像は拒否", () => {
      reject(withElement(1, imageElement), /backgroundIndex: .* background として/);
    });

    it("非対応の塗り（グラデーション等）は拒否", () => {
      reject(
        withElement(
          1,
          element({ id: "grad", box: { x: 1, y: 1, w: 2, h: 2 }, fill: { type: "unsupported", detail: "gradient" } }),
        ),
        /backgroundIndex: .* background として/,
      );
    });

    it("テキストを持つ塗り図形は許可（Compiler は card として描ける）", () => {
      accept(
        withElement(
          1,
          element({
            id: "bg-with-text",
            geometry: "roundRect",
            box: { x: 0.6, y: 1.6, w: 2.9, h: 3.4 },
            fill: { type: "solid", color: "F2F6FA" },
            text: { paragraphs: ["x"], runs: [], plain: "x" },
          }),
        ),
      );
    });
  });

  describe("accent は塗り図形か線であること", () => {
    it("solid fill の rect は許可", () => accept(cardSlideElements()));
    it("線は許可", () => accept(withElement(2, horizontalLine)));

    it("plain text は拒否", () => {
      reject(
        withElement(2, plainText),
        /cells\[0\]\.decoration\.accentIndexes\[0\]: .* accent として利用できません \(塗りを持つ図形または線が必要です\)/,
      );
    });

    it("画像は拒否", () => {
      reject(withElement(2, imageElement), /accentIndexes\[0\]: .* accent として/);
    });

    it("描画できない図形（面積 0）は拒否", () => {
      reject(
        withElement(
          2,
          element({ id: "z", geometry: "rect", box: { x: 1, y: 1, w: 0, h: 0 }, fill: { type: "solid", color: "0B4F8A" } }),
        ),
        /accentIndexes\[0\]: .* accent として/,
      );
    });
  });

  it("staticIndexes は広く許可する（元デザインとして残すため）", () => {
    const elements = [...cardSlideElements(), imageElement, horizontalLine, plainText];
    expect(() =>
      validate(humanInterpretation({ staticIndexes: [22, 23, 24] }), elements),
    ).not.toThrow();
  });

  it("Compiler へ渡す前に形式検証と要素検証の両方を通す", () => {
    // 形式（未知キー）
    expect(() =>
      parseAndValidateCaptureInterpretation(
        { ...humanInterpretation(), script: "x" },
        cardSlideElements(),
      ),
    ).toThrow(CaptureInterpretationError);
    // 構造（セル間の不一致）
    const uneven = humanInterpretation();
    delete uneven.cells[1]!.decoration;
    expect(() => parseAndValidateCaptureInterpretation(uneven, cardSlideElements())).toThrow(
      /Visual Role 構成/,
    );
    // 要素種別
    expect(() =>
      parseAndValidateCaptureInterpretation(humanInterpretation(), withElement(6, plainText)),
    ).toThrow(/divider として/);
  });
});

describe("divider は単数役割（Capture Interpretation 1.1）", () => {
  it("dividerIndex を受け入れる", () => {
    const parsed = parseCaptureInterpretation(humanInterpretation(), ELEMENT_COUNT);
    expect(parsed.cells[0]!.decoration?.dividerIndex).toBe(6);
  });

  it("最新スキーマで dividerIndexes は拒否される（strict schema）", () => {
    const legacy = JSON.parse(JSON.stringify(humanInterpretation())) as Record<string, any>;
    for (const cell of legacy.cells) {
      cell.decoration = { accentIndexes: cell.decoration.accentIndexes, dividerIndexes: [6] };
    }
    expect(() => parseCaptureInterpretation(legacy, ELEMENT_COUNT)).toThrow(
      CaptureInterpretationError,
    );
  });

  it("1.0 の dividerIndexes: [x] は dividerIndex へ移行する", () => {
    const legacy = JSON.parse(JSON.stringify(humanInterpretation())) as Record<string, any>;
    legacy.schemaVersion = "1.0";
    legacy.cells.forEach((cell: Record<string, any>, i: number) => {
      cell.decoration = { accentIndexes: [cellBase(i) + 1], dividerIndexes: [cellBase(i) + 5] };
    });

    const parsed = parseCaptureInterpretation(legacy, ELEMENT_COUNT);
    expect(parsed.schemaVersion).toBe(CAPTURE_INTERPRETATION_SCHEMA_VERSION);
    expect(parsed.cells.map((c) => c.decoration?.dividerIndex)).toEqual([6, 13, 20]);
    // 読み替えた痕跡を残す
    expect(parsed.notes?.join()).toContain("1.0 から 1.1 へ移行しました");
  });

  it("1.0 の空配列は「divider 無し」として移行する", () => {
    const legacy = JSON.parse(JSON.stringify(humanInterpretation())) as Record<string, any>;
    legacy.schemaVersion = "1.0";
    for (const cell of legacy.cells) {
      cell.decoration = { accentIndexes: cell.decoration.accentIndexes, dividerIndexes: [] };
    }
    const parsed = parseCaptureInterpretation(legacy, ELEMENT_COUNT);
    expect(parsed.cells.every((c) => c.decoration?.dividerIndex === undefined)).toBe(true);
  });

  it("1.0 で divider が複数指定されていたら、黙って先頭を採らず明示エラーにする", () => {
    const legacy = JSON.parse(JSON.stringify(humanInterpretation())) as Record<string, any>;
    legacy.schemaVersion = "1.0";
    legacy.cells[0].decoration = { accentIndexes: [2], dividerIndexes: [6, 7] };
    expect(() => parseCaptureInterpretation(legacy, ELEMENT_COUNT)).toThrow(
      /divider は単数です.*\[6, 7\].*dividerIndex に指定してください/s,
    );
  });

  it("Compiler に「複数から 1 件を黙って選ぶ」コードが残っていない", async () => {
    const source = await readFile(join(projectRoot, "src/patterns/capture/build.ts"), "utf8");
    expect(source).not.toContain("dividerIndexes");
  });
});

describe("Pattern Compiler（解釈された役割をそのまま写す）", () => {
  it("index / label / metric / divider / description の順で組み立てる", () => {
    const stack = cellStack(humanInterpretation());
    expect(stack.map((n) => n.type)).toEqual(["text", "text", "text", "shape", "text"]);
    expect(stack[0]!.value).toEqual({ template: "{{item.paddedOrdinal}}" });
    expect(stack[1]!.value).toEqual({ ref: "item.label" });
    expect(stack[2]!.value).toEqual({ ref: "item.metric" });
    expect(stack[3]!.shape).toBe("line");
    expect(stack[4]!.value).toEqual({ ref: "item.description" });
  });

  it("divider は高さ 0 の線のまま流し込む（Phase 1 の契約）", () => {
    const divider = cellStack(humanInterpretation())[3]!;
    expect(divider.height).toBe(0);
  });

  it("description は 1 ノードしか作らない（重複描画しない）", () => {
    const stack = cellStack(humanInterpretation());
    const descriptions = stack.filter(
      (n) =>
        n.type === "text" &&
        JSON.stringify(n.value) === JSON.stringify({ ref: "item.description" }),
    );
    expect(descriptions).toHaveLength(1);
  });

  it("description が複数指定されても 1 ノードに畳み、理由を警告に残す", () => {
    const interpretation = humanInterpretation();
    for (const [i, cell] of interpretation.cells.entries()) {
      const base = cellBase(i);
      // 見出しも説明も description として解釈された状況（現在のヒューリスティクスの結果に相当）
      cell.content = { labelIndex: base + 2, descriptionIndexes: [base + 6, base + 3] };
    }
    const result = compile(interpretation);
    const stack = (
      (result.document.layout.body as { children: LayoutChild[] }).children.find(
        (c) => c.node.type === "columns",
      )!.node as Record<string, any>
    ).item.child.children as Array<Record<string, any>>;

    const descriptions = stack.filter(
      (n) => JSON.stringify(n.value) === JSON.stringify({ ref: "item.description" }),
    );
    expect(descriptions).toHaveLength(1);
    expect(result.warnings.join()).toContain("重複描画しない");
  });

  it("Compiler は意味を推測しない（数字を label と解釈されたらそのまま label にする）", () => {
    const interpretation = humanInterpretation();
    for (const [i, cell] of interpretation.cells.entries()) {
      const base = cellBase(i);
      // 数値テキストを label、見出しテキストを metric として解釈させる
      cell.content.labelIndex = base + 4;
      cell.content.metricIndex = base + 3;
    }
    const stack = cellStack(interpretation);
    const labelNode = stack[1]!;
    // 解釈どおり: 26pt（数値テキスト）のスタイルが label ノードへ入る
    expect(labelNode.value).toEqual({ ref: "item.label" });
    expect(labelNode.style.role).toBe("keyMessage");
  });

  it("背景・アクセント・区切り線を構造として保持する", () => {
    const children = bodyChildren(humanInterpretation());
    const columnsAt = children.findIndex((c) => c.node.type === "columns");
    const accents = children.flatMap((c, i) =>
      c.node.type === "shape" && c.node.shape === "rect" ? [i] : [],
    );

    // 背景は card ノードとして反復テンプレートに入る
    const columns = children[columnsAt]!.node;
    expect(columns.item.type).toBe("card");
    expect(columns.item.fill).toBe("surface");

    // 区切り線はカード内のフローへ、アクセントはカードより前面の装飾として残る
    expect(cellStack(humanInterpretation()).some((n) => n.shape === "line")).toBe(true);
    expect(accents).toHaveLength(3);
    for (const index of accents) expect(index).toBeGreaterThan(columnsAt);
  });

  it("線として解釈されたアクセントも描ける", () => {
    const children = compile(humanInterpretation(), withElement(2, horizontalLine)).document.layout
      .body as { children: LayoutChild[] };
    const lines = children.children.filter(
      (c) => c.node.type === "shape" && c.node.shape === "line",
    );
    expect(lines).toHaveLength(1);
  });

  it("解釈の出どころを成果物に残す", () => {
    expect(compile(humanInterpretation()).document.metadata?.source?.interpretationSource).toBe(
      "human",
    );
  });

  it("同じ解釈からは同じ Pattern が出る（決定論的）", () => {
    expect(JSON.stringify(compile(humanInterpretation()).document)).toBe(
      JSON.stringify(compile(humanInterpretation()).document),
    );
  });

  it("役割を減らした解釈でも壊れない（index / metric / divider 無し）", () => {
    const minimal = humanInterpretation();
    for (const [i, cell] of minimal.cells.entries()) {
      const base = cellBase(i);
      cell.content = { labelIndex: base + 3, descriptionIndexes: [base + 6] };
      delete cell.decoration;
    }
    const stack = cellStack(minimal);
    expect(stack.map((n) => JSON.stringify(n.value ?? n.shape))).toEqual([
      JSON.stringify({ ref: "item.label" }),
      JSON.stringify({ ref: "item.description" }),
    ]);
  });
});

describe("連番（item の順番から作る値）", () => {
  it("item.paddedOrdinal / item.ordinal は 1 始まり", () => {
    const scope = { bindings: {}, item: { label: "x" }, index: 0 };
    expect(lookupPath(scope, "item.ordinal")).toBe("1");
    expect(lookupPath(scope, "item.paddedOrdinal")).toBe("01");
    expect(lookupPath({ ...scope, index: 9 }, "item.paddedOrdinal")).toBe("10");
  });

  it("既存の item.index / index の契約は変えていない（1 始まりのまま）", () => {
    const scope = { bindings: {}, item: { index: 1, label: "x" }, index: 0 };
    expect(lookupPath(scope, "item.index")).toBe(1);
    expect(lookupPath(scope, "index")).toBe(1);
  });

  it("item が同名フィールドを持つ場合は実データを優先する", () => {
    const scope = { bindings: {}, item: { ordinal: "壱" }, index: 0 };
    expect(lookupPath(scope, "item.ordinal")).toBe("壱");
  });

  it("repeat の外では解決されない", () => {
    expect(lookupPath({ bindings: {} }, "item.paddedOrdinal")).toBeUndefined();
  });

  it("3 セルの Pattern から 01 / 02 / 03 が描画される", async () => {
    const pattern = compileDeclarativePattern(
      JSON.parse(
        JSON.stringify({
          ...compile(humanInterpretation()).document,
          metadata: { scope: "company", status: "approved" },
        }),
      ),
    );
    const registry = new PatternRegistry([{ pattern, scope: "company" }, messageOnlyPattern]);
    const target = slide({
      keyMessage: "3 つの効果がある",
      content: pointsContent(
        { label: "設計知の再利用", description: "構造を保存する", evidenceRefs: ["e1"] },
        { label: "作成時間の短縮", description: "ゼロから考えない", evidenceRefs: ["e2"] },
        { label: "組織標準化", description: "共通資産にする", evidenceRefs: ["e3"] },
      ),
      evidence: [
        metricEvidence("e1", "件数", 12, "件"),
        metricEvidence("e2", "短縮率", 40, "%"),
        metricEvidence("e3", "部門数", 3, "部門"),
      ],
    });

    const plan = buildVisualPlan({ title: "ordinal", slides: [target] }, { registry });
    expect(plan.slides[0]!.patternId).toBe("captured");

    const { pptx, issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);

    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
    const texts = extractTexts(slideXmls[0]!);
    expect(texts).toContain("01");
    expect(texts).toContain("02");
    expect(texts).toContain("03");
    // 元 PPTX の "01" が literal として焼き付いていない（順番から作られている）
    expect(texts.filter((t) => t === "01")).toHaveLength(1);
    // description は 1 回だけ
    expect(texts.filter((t) => t === "構造を保存する")).toHaveLength(1);
  });
});

describe("ヒューリスティクスとの互換（Interpretation へ寄せても壊れない）", () => {
  it("推定結果を Interpretation へ変換できる", () => {
    const elements = cardSlideElements();
    const interpretation = interpretFromSlotInference(elements, {
      headlineIndex: 0,
      cells: [
        { backgroundIndex: 1, textIndexes: [3, 4, 5, 7] },
        { backgroundIndex: 8, textIndexes: [10, 11, 12, 14] },
        { backgroundIndex: 15, textIndexes: [17, 18, 19, 21] },
      ],
      staticIndexes: [2, 6, 9, 13, 16, 20],
      notes: [],
    });

    expect(interpretation.interpretationSource).toBe("heuristic");
    expect(interpretation.cells).toHaveLength(3);
    // Phase 2 では推定の賢さを変えていない: 先頭テキストが label、数値らしいものが metric
    expect(interpretation.cells[0]!.content.labelIndex).toBe(3);
    expect(interpretation.cells[0]!.content.metricIndex).toBe(5);
    expect(interpretation.cells[0]!.content.descriptionIndexes).toEqual([4, 7]);
    // accent / divider は判別しない（Phase 3 の課題）
    expect(interpretation.cells[0]!.decoration).toBeUndefined();
    // 変換結果は human と同じ入口をそのまま通る
    expect(() => parseAndValidateCaptureInterpretation(interpretation, elements)).not.toThrow();
  });

  it("推定が揃わない場合は反復セルにせず、要素を固定デザインとして残す", () => {
    const elements = cardSlideElements();
    const interpretation = interpretFromSlotInference(elements, {
      cells: [
        { backgroundIndex: 1, textIndexes: [3, 5] },
        { backgroundIndex: 8, textIndexes: [10] },
      ],
      staticIndexes: [],
      notes: [],
    });
    expect(interpretation.cells).toEqual([]);
    expect(interpretation.notes?.join()).toContain("役割の構成が揃わなかった");
    expect(interpretation.staticIndexes).toContain(3);
  });

  it("描画できない要素へ役割を割り当てず、固定デザインへ回す", () => {
    // headline に図形、background に線が推定されてしまった場合
    const elements = withElement(0, filledRect);
    elements[1] = horizontalLine;
    const interpretation = interpretFromSlotInference(elements, {
      headlineIndex: 0,
      cells: [
        { backgroundIndex: 1, textIndexes: [3, 4, 5, 7] },
        { backgroundIndex: 8, textIndexes: [10, 11, 12, 14] },
      ],
      staticIndexes: [],
      notes: [],
    });

    expect(interpretation.headlineIndex).toBeUndefined();
    expect(interpretation.staticIndexes).toContain(0);
    expect(interpretation.staticIndexes).toContain(1);
    expect(interpretation.notes?.join()).toContain("固定デザインとして残しました");
    // 要素を捨てず、かつ検証を通る解釈になっている
    expect(() => parseAndValidateCaptureInterpretation(interpretation, elements)).not.toThrow();
  });
});
