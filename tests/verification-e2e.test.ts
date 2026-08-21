import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { parsePresentation } from "../src/domain/schema.js";
import { capturePatternFromPptx, writePatternPackage } from "../src/patterns/capture/capture.js";
import { loadPatternPackage } from "../src/patterns/declarative/loader.js";
import { PatternRegistry, builtinPatterns } from "../src/patterns/registry.js";
import { generate } from "../src/pipeline.js";
import { readDeck } from "../src/pptx/extract.js";
import { PptxPackage } from "../src/pptx/package.js";
import {
  importTemplateFromPptx,
  writeTemplatePackage,
} from "../src/templates/importer/import-pptx.js";
import { loadTemplateFromDirectory } from "../src/templates/loader.js";
import { TemplateRegistry } from "../src/templates/registry.js";
import { projectRoot } from "./helpers.js";

/**
 * 実機で作った検証用 PPTX を使った E2E。
 *
 *   PPTX → Template Import + Pattern Capture → 人間レビュー → 同じ Template で再利用 → PPTX
 *
 * 目的は「PowerPoint から客観的に読み取れる構造（幾何・重なり順・種別）を
 * 変換途中で失わない」ことの回帰固定。意味推定（label / metric / description）の
 * 精度はここでは検証しない。
 */

const SOURCE_PPTX = join(projectRoot, "verification", "semantic-deck-v2-import-test.pptx");
const REUSE_PRESENTATION = join(projectRoot, "verification", "reuse-presentation.json");
const INTERPRETATION = join(projectRoot, "verification", "interpretation-slide2.json");

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-verification-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

async function importTemplate() {
  return importTemplateFromPptx(SOURCE_PPTX, {
    id: "verification-source",
    organization: "SDG V2 検証用サンプル",
    confidentialityLabel: "CONFIDENTIAL",
    sourceLabel: "semantic-deck-v2-import-test.pptx",
  });
}

async function capturePattern(interpretation?: unknown) {
  return capturePatternFromPptx(SOURCE_PPTX, {
    id: "test-three-column",
    slideNumber: 2,
    sourceLabel: "semantic-deck-v2-import-test.pptx",
    ...(interpretation !== undefined ? { interpretation } : {}),
  });
}

async function humanInterpretation(): Promise<unknown> {
  return JSON.parse(await readFile(INTERPRETATION, "utf8"));
}

describe("検証用 PPTX の抽出（前提）", () => {
  it("Content Layout に水平線があり、Slide 2 はカード → アクセント → 区切り線の順で並ぶ", async () => {
    const pkg = await PptxPackage.open(SOURCE_PPTX);
    const deck = await readDeck(pkg, { slideNumbers: [2], includeNotes: false });
    const elements = deck.slides[0]!.elements;

    // 元 PPTX の重なり順: 背景 → アクセント → テキスト … → 区切り線
    const backgrounds = elements
      .map((e, i) => ({ e, i }))
      .filter(({ e }) => e.geometry === "roundRect" && e.fill?.type === "solid");
    const accents = elements
      .map((e, i) => ({ e, i }))
      .filter(({ e }) => e.geometry === "rect" && e.fill?.type === "solid" && e.box.w < 0.2);
    const dividers = elements
      .map((e, i) => ({ e, i }))
      .filter(({ e }) => e.geometry === "line");

    expect(backgrounds).toHaveLength(3);
    expect(accents).toHaveLength(3);
    expect(dividers).toHaveLength(3);
    // 区切り線は高さ 0 の水平線として表現されている
    for (const { e } of dividers) {
      expect(e.box.h).toBe(0);
      expect(e.box.w).toBeGreaterThan(0);
    }
    // 背景より後ろ（= 前面）にアクセントと区切り線がある
    expect(accents[0]!.i).toBeGreaterThan(backgrounds[0]!.i);
    expect(dividers[0]!.i).toBeGreaterThan(backgrounds[0]!.i);
  });
});

describe("Template Import（Integration）", () => {
  it("Footer の水平線が Content Master に残る", async () => {
    const imported = await importTemplate();
    const content = imported.template.masters.find((m) => m.role === "content")!;
    const lines = (content.elements ?? []).filter(
      (e) => e.kind === "shape" && e.shape === "line",
    );

    expect(lines).toHaveLength(1);
    expect(lines[0]!.box.h).toBe(0);
    expect(lines[0]!.box.w).toBeGreaterThan(10);
    // フッターテキストのすぐ上にある
    expect(lines[0]!.box.y).toBeGreaterThan(6);
  });

  it("書き出した template.json を読み直しても線が残る", async () => {
    const root = await makeTempDir();
    const written = await writeTemplatePackage(root, await importTemplate());
    const loaded = await loadTemplateFromDirectory(written.dir);

    const content = loaded.template.masters.find((m) => m.role === "content")!;
    expect(
      (content.elements ?? []).filter((e) => e.kind === "shape" && e.shape === "line"),
    ).toHaveLength(1);

    const json = await readFile(written.templatePath, "utf8");
    expect(json).toContain('"shape": "line"');
  });
});

describe("Pattern Capture（Integration）", () => {
  it("区切り線が Pattern に残る", async () => {
    const captured = await capturePattern();
    const body = captured.document.layout.body as {
      type: string;
      children: Array<{ box: { w: number; h: number }; node: { type: string; shape?: string } }>;
    };

    expect(body.type).toBe("absolute");
    const dividers = body.children.filter(
      (c) => c.node.type === "shape" && c.node.shape === "line",
    );
    expect(dividers).toHaveLength(3);
    for (const divider of dividers) {
      expect(divider.box.h).toBe(0);
      expect(divider.box.w).toBeGreaterThan(0);
    }
  });

  it("元 PPTX の重なり順を保つ（カードより後ろにアクセント・区切り線）", async () => {
    const captured = await capturePattern();
    const body = captured.document.layout.body as {
      children: Array<{ box: { w: number }; node: { type: string; shape?: string } }>;
    };

    const columnsAt = body.children.findIndex((c) => c.node.type === "columns");
    expect(columnsAt).toBeGreaterThanOrEqual(0);

    const accents = body.children.flatMap((c, i) =>
      c.node.type === "shape" && c.node.shape === "rect" && c.box.w < 0.2 ? [i] : [],
    );
    const dividers = body.children.flatMap((c, i) =>
      c.node.type === "shape" && c.node.shape === "line" ? [i] : [],
    );
    expect(accents).toHaveLength(3);
    expect(dividers).toHaveLength(3);
    for (const index of [...accents, ...dividers]) {
      expect(index).toBeGreaterThan(columnsAt);
    }
  });

  it("Capture 結果は draft で、決定論的", async () => {
    const first = await capturePattern();
    const second = await capturePattern();
    expect(first.document.metadata?.status).toBe("draft");
    expect(JSON.stringify(second.document)).toBe(JSON.stringify(first.document));
  });
});

/**
 * PPTX → Template Import + Pattern Capture → 人間レビュー → 同じ Template で再利用。
 * `interpretation` を渡すと人手の解釈、省略するとヒューリスティクスを使う。
 */
async function reuse(interpretation?: unknown) {
  const root = await makeTempDir();

  const importedTemplate = await importTemplate();
  const templateDir = await writeTemplatePackage(join(root, "templates"), importedTemplate);
  const loadedTemplate = await loadTemplateFromDirectory(templateDir.dir);

  const captured = await capturePattern(interpretation);
  // 人間レビュー相当（draft → approved）
  const approved = {
    ...captured.document,
    metadata: { ...captured.document.metadata, status: "approved" as const },
  };
  const written = await writePatternPackage(join(root, "patterns"), {
    document: approved,
    assets: captured.assets,
  });
  const pkg = await loadPatternPackage(written.dir);

  const registry = new PatternRegistry(builtinPatterns).registerPackage(pkg);
  const templates = new TemplateRegistry().registerLoaded(loadedTemplate);
  const spec = parsePresentation(JSON.parse(await readFile(REUSE_PRESENTATION, "utf8")));

  return generate(spec, {
    registry,
    templateRegistry: templates,
    templateId: "verification-source",
    policy: {
      mode: "manual",
      overrides: {
        "slide-reuse": { patternId: "test-three-column", reason: "Capture した Pattern の再利用" },
      },
    },
  });
}

describe("Capture した Pattern を同じ Template で再利用（E2E）", () => {
  it("生成が完了し、レイアウトエラーが出ない", async () => {
    const result = await reuse();
    expect(result.plan.slides[0]!.patternId).toBe("test-three-column");
    expect(result.plan.templateId).toBe("verification-source");
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    expect(result.pptxBuffer.length).toBeGreaterThan(10_000);
  });

  it("カード背景・アクセント・区切り線がすべて描画され、アクセントと区切り線が前面に来る", async () => {
    const result = await reuse();
    const placements = result.reports.find((r) => r.slideId === "slide-reuse")!.placements;

    const cards = placements.flatMap((p, i) => (p.name.includes("card") ? [i] : []));
    const accents = placements.flatMap((p, i) =>
      p.kind === "shape" && !p.name.includes("card") && p.box.w < 0.2 ? [i] : [],
    );
    const dividers = placements.flatMap((p, i) =>
      p.kind === "shape" && p.box.h === 0 ? [i] : [],
    );

    expect(cards).toHaveLength(3);
    expect(accents).toHaveLength(3);
    expect(dividers).toHaveLength(3);

    // 描画順 = PPTX の Z-order。カードより後に描かれる = カードの上に見える
    for (const index of [...accents, ...dividers]) {
      expect(index).toBeGreaterThan(Math.max(...cards));
    }
  });

  it("生成した PPTX に Template の Footer 線とカードが含まれる", async () => {
    const result = await reuse();
    const { readPptx } = await import("./helpers.js");
    const pptx = await readPptx(result.pptxBuffer);

    // Template の線は Slide Layout（Master）側
    expect(pptx.slideLayoutXmls.join("")).toContain('prst="line"');
    // Pattern の線・カードはスライド本体側
    const slideXml = pptx.slideXmls[pptx.slideXmls.length - 1]!;
    expect(slideXml).toContain('prst="line"');
    expect(slideXml).toContain('prst="roundRect"');
  });
});

describe("Phase 2: 人手の解釈で取り込む（E2E）", () => {
  it("解釈 JSON の各要素が役割どおりに Pattern へ入る", async () => {
    const captured = await capturePattern(await humanInterpretation());
    expect(captured.document.metadata?.source?.interpretationSource).toBe("human");

    const body = captured.document.layout.body as {
      children: Array<{ box: { w: number }; node: Record<string, any> }>;
    };
    const columns = body.children.find((c) => c.node.type === "columns")!;
    const stack = columns.node.item.child.children as Array<Record<string, any>>;

    // index → 順番から作る / label / metric / divider / description の順
    expect(stack.map((n) => n.type)).toEqual(["text", "text", "text", "shape", "text"]);
    expect(stack[0]!.value).toEqual({ template: "{{item.paddedOrdinal}}" });
    expect(stack[1]!.value).toEqual({ ref: "item.label" });
    expect(stack[2]!.value).toEqual({ ref: "item.metric" });
    expect(stack[3]!.shape).toBe("line");
    expect(stack[4]!.value).toEqual({ ref: "item.description" });

    // 元 PPTX の "01" が literal として焼き付いていない
    expect(JSON.stringify(captured.document)).not.toContain('"literal":"01"');

    // 背景は card、アクセントはカードより前面の装飾として残る
    expect(columns.node.item.type).toBe("card");
    const columnsAt = body.children.indexOf(columns);
    const accents = body.children.flatMap((c, i) =>
      c.node.type === "shape" && c.node.shape === "rect" && c.box.w < 0.2 ? [i] : [],
    );
    expect(accents).toHaveLength(3);
    for (const index of accents) expect(index).toBeGreaterThan(columnsAt);
  });

  it("Capture Report に inventory と採用された解釈の両方が残る", async () => {
    const captured = await capturePattern(await humanInterpretation());

    expect(captured.report.inventory).toHaveLength(captured.report.elements.total);
    const label = captured.report.inventory.find((e) => e.index === 6)!;
    expect(label.role).toBe("cells[0].content.label");
    expect(label.text).toContain("開発スピード");
    expect(label.sourceOrder).toBe(6);

    expect(captured.report.interpretation.source).toBe("human");
    expect(captured.report.interpretation.cellRoles).toHaveLength(3);
    expect(captured.report.interpretation.cellRoles[0]).toMatchObject({
      background: 3,
      index: 5,
      label: 6,
      metric: 7,
      descriptions: [9],
      accents: [4],
      dividers: [8],
    });
  });

  it("別内容へ再利用すると 01 / 02 / 03 と label / metric が正しく入る", async () => {
    const result = await reuse(await humanInterpretation());
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);

    const { readPptx } = await import("./helpers.js");
    const { extractTexts } = await import("./helpers.js");
    const pptx = await readPptx(result.pptxBuffer);
    const texts = extractTexts(pptx.slideXmls[pptx.slideXmls.length - 1]!);

    // 連番は item の順番から作られる
    expect(texts).toContain("01");
    expect(texts).toContain("02");
    expect(texts).toContain("03");

    // label / metric が再利用先の内容になる
    for (const label of ["設計知の再利用", "作成時間の短縮", "組織標準化"]) {
      expect(texts.filter((t) => t === label)).toHaveLength(1);
    }
    for (const metric of ["12件", "40%", "3部門"]) {
      expect(texts).toContain(metric);
    }

    // description は 1 回だけ（重複描画しない）
    const description = "良いスライドの構造をPatternとして保存し、別案件でも再利用する";
    expect(texts.filter((t) => t === description)).toHaveLength(1);
  });

  it("再利用後もカード・アクセント・区切り線が残り、Z-order が保たれる（Phase 1 の維持）", async () => {
    const result = await reuse(await humanInterpretation());
    const placements = result.reports.find((r) => r.slideId === "slide-reuse")!.placements;

    const cards = placements.flatMap((p, i) => (p.name.includes("card") ? [i] : []));
    const accents = placements.flatMap((p, i) =>
      p.kind === "shape" && !p.name.includes("card") && p.box.w < 0.2 ? [i] : [],
    );
    const dividers = placements.flatMap((p, i) =>
      p.kind === "shape" && p.box.h === 0 ? [i] : [],
    );

    expect(cards).toHaveLength(3);
    expect(accents).toHaveLength(3);
    expect(dividers).toHaveLength(3);
    // アクセントはカードより後に描かれる = カードの上に見える
    for (const index of accents) expect(index).toBeGreaterThan(Math.max(...cards));
  });

  it("解釈なし（ヒューリスティクス）でも従来どおり取り込める（後方互換）", async () => {
    const captured = await capturePattern();
    expect(captured.document.metadata?.source?.interpretationSource).toBe("heuristic");
    expect(captured.report.interpretation.source).toBe("heuristic");
    expect(captured.report.interpretation.inferencer).toBe("heuristic");

    // 解釈なしでも生成まで到達する
    const result = await reuse();
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
  });

  it("壊れた解釈 JSON は Compiler へ到達しない", async () => {
    await expect(
      capturePattern({ schemaVersion: "1.0", interpretationSource: "human", cells: [], statics: [] }),
    ).rejects.toThrow(/Capture Interpretation/);

    const outOfRange = JSON.parse(JSON.stringify(await humanInterpretation())) as {
      cells: Array<{ content: { labelIndex: number } }>;
    };
    outOfRange.cells[0]!.content.labelIndex = 9999;
    await expect(capturePattern(outOfRange)).rejects.toThrow(/存在しない要素インデックス/);
  });
});
