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

async function capturePattern() {
  return capturePatternFromPptx(SOURCE_PPTX, {
    id: "test-three-column",
    slideNumber: 2,
    sourceLabel: "semantic-deck-v2-import-test.pptx",
  });
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

describe("Capture した Pattern を同じ Template で再利用（E2E）", () => {
  async function reuse() {
    const root = await makeTempDir();

    const importedTemplate = await importTemplate();
    const templateDir = await writeTemplatePackage(join(root, "templates"), importedTemplate);
    const loadedTemplate = await loadTemplateFromDirectory(templateDir.dir);

    const captured = await capturePattern();
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
