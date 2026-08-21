import { describe, expect, it } from "vitest";
import { buildCapturedPattern } from "../src/patterns/capture/build.js";
import {
  CAPTURE_INTERPRETATION_SCHEMA_VERSION,
  CaptureInterpretationError,
  assignmentsOf,
  interpretFromSlotInference,
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
import { extractTexts, metricEvidence, pointsContent, readPptx, slide } from "./helpers.js";

/**
 * Phase 2: 解釈（Interpretation）と Pattern Compiler。
 *
 * ここで固定したいのは「意味推定の賢さ」ではなく、
 * **正しく解釈された入力なら正しい Pattern が出る** という変換の正確さ。
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

/** cardSlideElements() に対応する、人手で書いた想定の解釈 */
function humanInterpretation(
  overrides: Partial<CaptureInterpretation> = {},
): CaptureInterpretation {
  const cells = [0, 1, 2].map((i) => {
    const base = 1 + i * 7;
    return {
      backgroundIndex: base,
      content: {
        indexIndex: base + 2,
        labelIndex: base + 3,
        metricIndex: base + 4,
        descriptionIndexes: [base + 6],
      },
      decoration: { accentIndexes: [base + 1], dividerIndexes: [base + 5] },
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

function compile(interpretation: CaptureInterpretation) {
  return buildCapturedPattern(cardSlideElements(), interpretation, {
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

describe("Interpretation スキーマ（untrusted input として検証する）", () => {
  const elementCount = 22;

  it("正しい解釈を受け入れる", () => {
    const parsed = parseCaptureInterpretation(humanInterpretation(), elementCount);
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
    expect(() => parseCaptureInterpretation(conflict, elementCount)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("content と decoration の衝突も検出する", () => {
    const conflict = humanInterpretation();
    conflict.cells[0]!.decoration = { accentIndexes: [conflict.cells[0]!.content.labelIndex] };
    expect(() => parseCaptureInterpretation(conflict, elementCount)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("static と cell の衝突も検出する", () => {
    const conflict = humanInterpretation({ staticIndexes: [4] });
    expect(() => parseCaptureInterpretation(conflict, elementCount)).toThrow(
      /複数の役割へ割り当てられています/,
    );
  });

  it("label が無いセルを拒否する", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation())) as {
      cells: Array<{ content: { labelIndex?: number } }>;
    };
    delete invalid.cells[0]!.content.labelIndex;
    expect(() => parseCaptureInterpretation(invalid, elementCount)).toThrow(
      CaptureInterpretationError,
    );
  });

  it("未知のキーを拒否する（strict schema）", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation()));
    invalid.cells[0].content.tagline = 3;
    expect(() => parseCaptureInterpretation(invalid, elementCount)).toThrow(
      CaptureInterpretationError,
    );
    expect(() =>
      parseCaptureInterpretation({ ...humanInterpretation(), script: "x" }, elementCount),
    ).toThrow(CaptureInterpretationError);
  });

  it("負のインデックス・小数を拒否する", () => {
    const invalid = JSON.parse(JSON.stringify(humanInterpretation()));
    invalid.cells[0].content.labelIndex = -1;
    expect(() => parseCaptureInterpretation(invalid, elementCount)).toThrow(
      CaptureInterpretationError,
    );
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() =>
      parseCaptureInterpretation(humanInterpretation({ schemaVersion: "9.9" }), elementCount),
    ).toThrow(/schemaVersion/);
  });

  it("未知の interpretationSource（例: ai）は今は拒否する", () => {
    expect(() =>
      parseCaptureInterpretation(
        { ...humanInterpretation(), interpretationSource: "ai" },
        elementCount,
      ),
    ).toThrow(CaptureInterpretationError);
  });

  it("セルが 1 件だけ、または役割構成が揃わない場合を拒否する", () => {
    const single = humanInterpretation({ cells: [humanInterpretation().cells[0]!] });
    expect(() => parseCaptureInterpretation(single, elementCount)).toThrow(/2 件以上/);

    const uneven = humanInterpretation();
    uneven.cells[1]!.content = { labelIndex: uneven.cells[1]!.content.labelIndex };
    expect(() => parseCaptureInterpretation(uneven, elementCount)).toThrow(/役割の構成が異なります/);
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

  it("description は 1 ノードしか作らない（重複描画しない）", () => {
    const stack = cellStack(humanInterpretation());
    const descriptions = stack.filter(
      (n) => n.type === "text" && JSON.stringify(n.value) === JSON.stringify({ ref: "item.description" }),
    );
    expect(descriptions).toHaveLength(1);
  });

  it("description が複数指定されても 1 ノードに畳み、理由を警告に残す", () => {
    const interpretation = humanInterpretation();
    for (const [i, cell] of interpretation.cells.entries()) {
      const base = 1 + i * 7;
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
      const base = 1 + i * 7;
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
      const base = 1 + i * 7;
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
    // 変換結果も同じ検証を通る
    expect(() => parseCaptureInterpretation(interpretation, elements.length)).not.toThrow();
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
});
