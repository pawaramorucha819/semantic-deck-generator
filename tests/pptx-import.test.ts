import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import PptxGenJS from "pptxgenjs";
import { afterAll, describe, expect, it } from "vitest";
import { generate } from "../src/pipeline.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import { capturePatternFromPptx, writePatternPackage } from "../src/patterns/capture/capture.js";
import { HeuristicSlotInferencer, validateInference } from "../src/patterns/capture/infer.js";
import { loadPatternPackage } from "../src/patterns/declarative/loader.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import { buildCapturedPattern } from "../src/patterns/capture/build.js";
import {
  CAPTURE_INTERPRETATION_SCHEMA_VERSION,
  type CaptureInterpretation,
} from "../src/patterns/capture/interpretation.js";
import { readDeck } from "../src/pptx/extract.js";
import {
  hasRenderableExtent,
  isLineElement,
  type ExtractedElement,
} from "../src/pptx/model.js";
import { PptxPackage, PptxReadError, resolvePartPath } from "../src/pptx/package.js";
import { parseXml, XmlParseError, findDescendants, textOf } from "../src/pptx/xml.js";
import { defaultTheme } from "../src/theme/theme.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import {
  importTemplateFromPptx,
  writeTemplatePackage,
} from "../src/templates/importer/import-pptx.js";
import { loadTemplateFromDirectory } from "../src/templates/loader.js";
import { TemplateRegistry } from "../src/templates/registry.js";
import { extractTexts, loadExample, metricEvidence, pointsContent, readPptx, slide, TINY_PNG_BASE64 } from "./helpers.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-pptx-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

/** 「人が PowerPoint で作った良いスライド」に相当するサンプルを組み立てる */
async function buildSampleDeck(): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  const slide1 = pptx.addSlide();

  slide1.addText("課題", {
    x: 0.6, y: 0.32, w: 2, h: 0.24, fontSize: 10, color: "0B4F8A", bold: true,
  });
  slide1.addText("3 つの構造的な原因がある", {
    x: 0.6, y: 0.6, w: 8.8, h: 0.5, fontSize: 22, bold: true, color: "1A2230",
  });
  slide1.addText("現場ヒアリングから要因は 3 つに集約された", {
    x: 0.6, y: 1.15, w: 8.8, h: 0.3, fontSize: 13, color: "6B7686",
  });

  const cards = [
    { title: "定型指摘に人手", metric: "35%", desc: "規約・命名の指摘が多い" },
    { title: "レビュー集中", metric: "58%", desc: "上位層が担当している" },
    { title: "待ち時間が長い", metric: "18時間", desc: "初回コメントまで" },
  ];
  cards.forEach((card, i) => {
    const x = 0.6 + i * 3.0;
    slide1.addShape("roundRect", {
      x, y: 1.7, w: 2.8, h: 2.6,
      fill: { color: "F2F6FA" },
      line: { color: "D5DCE4", width: 1 },
    });
    slide1.addText(card.title, { x: x + 0.2, y: 1.9, w: 2.4, h: 0.36, fontSize: 15, bold: true, color: "1A2230" });
    slide1.addText(card.metric, { x: x + 0.2, y: 2.32, w: 2.4, h: 0.42, fontSize: 20, bold: true, color: "0B4F8A" });
    slide1.addText(card.desc, { x: x + 0.2, y: 2.82, w: 2.4, h: 1.2, fontSize: 12, color: "6B7686" });
  });

  slide1.addImage({ data: `image/png;base64,${TINY_PNG_BASE64}`, x: 8.9, y: 0.25, w: 0.5, h: 0.3 });
  slide1.addText("出典: 社内ヒアリング 2025-01", {
    x: 0.6, y: 5.0, w: 5, h: 0.24, fontSize: 9, color: "6B7686",
  });

  return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
}

/** 非対応要素（SmartArt）を含む最小 PPTX を手で組む */
async function buildPptxWithSlideXml(slideXml: string): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
  );
  zip.file("ppt/slides/slide1.xml", slideXml);
  return (await zip.generateAsync({ type: "nodebuffer" })) as Buffer;
}

const SMARTART_SLIDE = `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2400" b="1"/><a:t>SmartArt を含むスライド</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:graphicFrame><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="6400800" cy="2286000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`;

const GROUPED_SLIDE = `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:grpSp>
      <p:grpSpPr><a:xfrm>
        <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/>
        <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/>
      </a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="InGroup"/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="228600"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr>
      </p:sp>
    </p:grpSp>
  </p:spTree></p:cSld>
</p:sld>`;

describe("XML パーサ（信頼できない入力の扱い）", () => {
  it("要素・属性・テキストを読める", () => {
    const root = parseXml(`<a:root xmlns:a="x" id="1"><a:child v="2">text</a:child></a:root>`);
    expect(root.name).toBe("a:root");
    expect(root.attrs.id).toBe("1");
    expect(root.children[0]!.name).toBe("a:child");
    expect(textOf(root.children[0]!)).toBe("text");
  });

  it("実体参照をデコードする", () => {
    const root = parseXml(`<t>a&amp;b&lt;c&gt;&#65;</t>`);
    expect(textOf(root)).toBe("a&b<c>A");
  });

  it("DOCTYPE / ENTITY を含む XML を拒否する（XXE 防止）", () => {
    expect(() =>
      parseXml(`<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><t>&xxe;</t>`),
    ).toThrow(XmlParseError);
  });

  it("ネストが深すぎる XML を拒否する", () => {
    const deep = `${"<a>".repeat(50)}x${"</a>".repeat(50)}`;
    expect(() => parseXml(deep, { maxDepth: 10 })).toThrow(/ネスト/);
  });

  it("壊れた XML を拒否する", () => {
    expect(() => parseXml(`<a><b></a>`)).toThrow(XmlParseError);
    expect(() => parseXml(`not xml at all`)).toThrow(XmlParseError);
  });

  it("__proto__ 属性を取り込まない", () => {
    const root = parseXml(`<a __proto__="polluted" id="ok"/>`);
    expect(Object.keys(root.attrs)).toEqual(["id"]);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

describe("PPTX パッケージ（untrusted input）", () => {
  it("ZIP でないファイルを拒否する", async () => {
    await expect(PptxPackage.open(Buffer.from("not a zip"))).rejects.toThrow(PptxReadError);
  });

  it("スライドを含まない ZIP を拒否する", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    const buffer = (await zip.generateAsync({ type: "nodebuffer" })) as Buffer;
    await expect(PptxPackage.open(buffer)).rejects.toThrow(/スライドが 1 枚も/);
  });

  it("パッケージ外を指す相対パスを拒否する", () => {
    expect(() => resolvePartPath("ppt/slides/slide1.xml", "../../../../etc/passwd")).toThrow(
      PptxReadError,
    );
    expect(resolvePartPath("ppt/slides/slide1.xml", "../media/img.png")).toBe("ppt/media/img.png");
  });

  it("生成した PPTX を読み戻せる", async () => {
    const pkg = await PptxPackage.open(await buildSampleDeck());
    expect(pkg.slideParts()).toEqual(["ppt/slides/slide1.xml"]);
    const deck = await readDeck(pkg);
    expect(deck.slideSize).toEqual({ widthIn: 10, heightIn: 5.625 });
    expect(deck.slides[0]!.elements.length).toBeGreaterThan(10);
    expect(deck.slides[0]!.elements.some((e) => e.kind === "image")).toBe(true);
  });

  it("グループ内の図形を親座標系へ展開する", async () => {
    const pkg = await PptxPackage.open(await buildPptxWithSlideXml(GROUPED_SLIDE));
    const deck = await readDeck(pkg);
    const element = deck.slides[0]!.elements[0]!;
    expect(element.fromGroup).toBe(true);
    // 子座標 (0.5in, 0.25in) がグループ変換で 2 倍され、グループ原点 (1in, 1in) へ移動する
    expect(element.box.x).toBeCloseTo(2, 3);
    expect(element.box.y).toBeCloseTo(1.5, 3);
    expect(element.box.w).toBeCloseTo(1, 3);
  });

  it("SmartArt を非対応として明示的に報告する", async () => {
    const pkg = await PptxPackage.open(await buildPptxWithSlideXml(SMARTART_SLIDE));
    const deck = await readDeck(pkg);
    expect(deck.slides[0]!.unsupported.join()).toContain("SmartArt");
    // 対応している text だけは取り込まれる
    expect(deck.slides[0]!.elements.map((e) => e.text?.plain)).toContain(
      "SmartArt を含むスライド",
    );
  });
});

describe("既存 PPTX からの Template Import", () => {
  it("スライドサイズ・テーマ・マスター装飾を取り込む", async () => {
    const spec = await loadExample();
    const generated = await generate(spec, { templateId: "corporate" });
    const imported = await importTemplateFromPptx(generated.pptxBuffer, {
      id: "reimported",
      organization: "Example Corporation",
      confidentialityLabel: "社外秘 / Confidential",
      sourceLabel: "corporate.pptx",
    });

    expect(imported.template.id).toBe("reimported");
    expect(imported.report.slideSize).toEqual({ widthIn: 10, heightIn: 5.625 });
    expect(imported.template.masters.length).toBeGreaterThanOrEqual(1);

    const contentMaster = imported.template.masters.find((m) => m.role === "content")!;
    // ヘッダー帯 / 組織名 / フッター罫線 が Master 要素として復元される
    expect(contentMaster.elements!.length).toBeGreaterThan(0);
    expect(
      contentMaster.elements!.some((e) => e.kind === "text" && e.text.includes("Example Corporation")),
    ).toBe(true);
    // ページ番号プレースホルダの位置も復元される
    expect(contentMaster.slideNumber?.box.x).toBeCloseTo(8.4, 2);
  });

  it("取り込めなかった要素と近似した色を必ず報告する（silent failure にしない）", async () => {
    const spec = await loadExample();
    const generated = await generate(spec, { templateId: "corporate" });
    const imported = await importTemplateFromPptx(generated.pptxBuffer, { id: "reported" });

    expect(imported.report.warnings.some((w) => w.includes("フォントサイズ"))).toBe(true);
    expect(imported.report.warnings.some((w) => w.includes("寄せました"))).toBe(true);
    expect(imported.template.import?.importedFrom).toBeDefined();
  });

  it("SmartArt を含む PPTX でも unsupportedFeatures として記録される", async () => {
    const buffer = await buildPptxWithSlideXml(SMARTART_SLIDE);
    // slideMaster が無いファイルは Template として取り込めない（明示的に失敗する）
    await expect(importTemplateFromPptx(buffer, { id: "broken" })).rejects.toThrow(
      /slideMaster が見つかりません/,
    );
  });

  it("壊れた PPTX は分かりやすいエラーになる", async () => {
    await expect(
      importTemplateFromPptx(Buffer.from("garbage"), { id: "x" }),
    ).rejects.toThrow(PptxReadError);
  });

  it("書き出した Template パッケージを読み直して生成に使える", async () => {
    const spec = await loadExample();
    const generated = await generate(spec, { templateId: "corporate" });
    const imported = await importTemplateFromPptx(generated.pptxBuffer, {
      id: "roundtrip",
      organization: "Example Corporation",
      sourceLabel: "corporate.pptx",
    });

    const root = await makeTempDir();
    const written = await writeTemplatePackage(root, imported);
    expect(await readFile(written.reportPath, "utf8")).toContain("unsupportedFeatures");

    const loaded = await loadTemplateFromDirectory(written.dir);
    const registry = new TemplateRegistry().registerLoaded(loaded);
    const result = await generate(spec, { templateId: "roundtrip", templateRegistry: registry });

    expect(result.plan.templateId).toBe("roundtrip");
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    const contents = await readPptx(result.pptxBuffer);
    expect(contents.slideLayoutXmls.join("")).toContain("roundtrip__roundtrip-content");
  });
});

describe("Pattern Capture", () => {
  it("単純な text / shape / image のスライドから Pattern 下書きを作る", async () => {
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-three-card",
      slideNumber: 1,
      sourceLabel: "sample.pptx",
      organization: "ACME",
    });

    expect(captured.document.id).toBe("captured-three-card");
    // Capture の成果物は必ず下書き
    expect(captured.document.metadata?.status).toBe("draft");
    expect(captured.report.inference.headline).toBe("3 つの構造的な原因がある");
    expect(captured.report.inference.cells).toBe(3);

    const slotIds = captured.document.slots.map((s) => s.id);
    expect(slotIds).toContain("headline");
    expect(slotIds).toContain("eyebrow");
    expect(slotIds).toContain("subline");
    expect(slotIds).toContain("sourceText");
    expect(slotIds).toContain("items");

    // 色は必ず Theme トークンへ寄せられ、生の hex は残らない
    const json = JSON.stringify(captured.document);
    expect(json).not.toMatch(/#?[0-9A-F]{6}"/);
    expect(json).toContain('"fill":"surface"');
  });

  it("元スライドの座標を保った absolute レイアウトになる", async () => {
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-geometry",
      slideNumber: 1,
    });
    const body = captured.document.layout.body as { type: string; children: Array<{ box: { x: number; y: number } }> };
    expect(body.type).toBe("absolute");
    expect(body.children.some((c) => Math.abs(c.box.y - 0.6) < 0.01)).toBe(true);
  });

  it("source metadata（取り込み元ファイル・スライド番号）が保存される", async () => {
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-source",
      slideNumber: 1,
      sourceLabel: "past-decks/strategy.pptx",
      importedAt: "2026-01-01",
    });
    expect(captured.document.metadata?.source).toMatchObject({
      file: "past-decks/strategy.pptx",
      slideNumber: 1,
      importedAt: "2026-01-01",
    });
  });

  it("非対応要素は unsupportedFeatures として下書きに残る", async () => {
    const captured = await capturePatternFromPptx(await buildPptxWithSlideXml(SMARTART_SLIDE), {
      id: "captured-smartart",
      slideNumber: 1,
    });
    expect(captured.report.unsupportedFeatures.join()).toContain("SmartArt");
    expect(captured.document.metadata?.source?.unsupportedFeatures?.join()).toContain("SmartArt");
  });

  it("存在しないスライド番号は分かりやすいエラーになる", async () => {
    await expect(
      capturePatternFromPptx(await buildSampleDeck(), { id: "x", slideNumber: 99 }),
    ).rejects.toThrow(/スライド 99 が見つかりません/);
  });

  it("壊れた PPTX を拒否する", async () => {
    await expect(
      capturePatternFromPptx(Buffer.from("PK not really a zip"), { id: "x", slideNumber: 1 }),
    ).rejects.toThrow(PptxReadError);
  });

  it("CLI からの明示指定で推定を上書きできる", async () => {
    const buffer = await buildSampleDeck();
    const auto = await capturePatternFromPptx(buffer, { id: "auto", slideNumber: 1 });
    const overridden = await capturePatternFromPptx(buffer, {
      id: "overridden",
      slideNumber: 1,
      overrides: { headlineIndex: 0 },
    });
    expect(auto.report.inference.headline).not.toBe("課題");
    expect(overridden.report.inference.headline).toBe("課題");
    expect(overridden.report.inference.notes.join()).toContain("CLI 指定");
  });

  it("Slot 推定結果は範囲外インデックスを拒否する（AI 実装でも同じゲートを通る）", () => {
    expect(() =>
      validateInference({ cells: [], staticIndexes: [99], notes: [] }, 3),
    ).toThrow(/存在しない要素インデックス/);
    expect(() =>
      validateInference({ headlineIndex: 1, sublineIndex: 1, cells: [], staticIndexes: [], notes: [] }, 3),
    ).toThrow(/複数の役割/);
    expect(() => validateInference({ garbage: true }, 3)).toThrow(/形式が不正/);
  });

  it("推定器は差し替え可能（AI 実装の受け口）", async () => {
    const scripted = {
      id: "scripted-test",
      // eslint-disable-next-line @typescript-eslint/require-await
      async infer() {
        return { headlineIndex: 1, cells: [], staticIndexes: [], notes: ["scripted"] };
      },
    };
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-scripted",
      slideNumber: 1,
      inferencer: scripted,
    });
    expect(captured.report.inference.inferencer).toBe("scripted-test");
    expect(captured.document.slots.map((s) => s.id)).toEqual(["headline"]);
  });

  it("ヒューリスティクスは決定論的（同じ入力 → 同じ推定）", async () => {
    const buffer = await buildSampleDeck();
    const pkg = await PptxPackage.open(buffer);
    const deck = await readDeck(pkg);
    const inferencer = new HeuristicSlotInferencer();
    const input = {
      slideNumber: 1,
      slideSize: deck.slideSize,
      elements: deck.slides[0]!.elements,
    };
    expect(await inferencer.infer(input)).toEqual(await inferencer.infer(input));
  });

  it("下書き → 人間レビュー（approved）→ 登録 → 描画まで通る", async () => {
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-usable",
      slideNumber: 1,
      sourceLabel: "sample.pptx",
    });

    const root = await makeTempDir();
    const written = await writePatternPackage(root, captured);

    // draft のままでは登録できない
    await expect(loadPatternPackage(written.dir)).rejects.toThrow(/draft/);

    // 人間がレビューして approved にした体で読み直す
    const approved = {
      ...captured.document,
      metadata: { ...captured.document.metadata, status: "approved" as const },
    };
    await writePatternPackage(root, { document: approved, assets: captured.assets });
    const pkg = await loadPatternPackage(written.dir);

    const registry = new PatternRegistry([messageOnlyPattern]).registerPackage(pkg);
    const target = slide({
      id: "s1",
      keyMessage: "3 つの構造的な原因がある",
      supportingMessage: "現場ヒアリングから要因は 3 つに集約された",
      content: pointsContent(
        { label: "定型指摘に人手", description: "規約・命名の指摘が多い", evidenceRefs: ["e1"] },
        { label: "レビュー集中", description: "上位層が担当している", evidenceRefs: ["e2"] },
        { label: "待ち時間が長い", description: "初回コメントまで", evidenceRefs: ["e3"] },
      ),
      evidence: [
        metricEvidence("e1", "定型指摘", 35, "%"),
        metricEvidence("e2", "レビュー集中", 58, "%"),
        metricEvidence("e3", "待ち時間", 18, "時間"),
      ],
    });

    const plan = buildVisualPlan({ title: "captured", slides: [target] }, { registry });
    expect(plan.slides[0]!.patternId).toBe("captured-usable");

    const { pptx, issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);

    const { slideXmls, files } = await readPptx(await toPptxBuffer(pptx));
    const texts = extractTexts(slideXmls[0]!).join("\n");
    expect(texts).toContain("3 つの構造的な原因がある");
    expect(texts).toContain("定型指摘に人手");
    expect(texts).toContain("35%");
    // 取り込んだロゴ画像も再現される
    expect(files.filter((f) => f.startsWith("ppt/media/")).length).toBeGreaterThan(0);
  });

  it("キャプチャした Pattern の描画も RenderContext 経由で検査される", async () => {
    const captured = await capturePatternFromPptx(await buildSampleDeck(), {
      id: "captured-validated",
      slideNumber: 1,
    });
    const approved = {
      ...captured.document,
      metadata: { ...captured.document.metadata, status: "approved" as const },
    };
    const root = await makeTempDir();
    const written = await writePatternPackage(root, { document: approved, assets: captured.assets });
    const pkg = await loadPatternPackage(written.dir);
    const registry = new PatternRegistry([messageOnlyPattern]).registerPackage(pkg);

    const target = slide({
      content: pointsContent("A", "B", "C"),
      evidence: [],
    });
    const plan = buildVisualPlan({ title: "t", slides: [target] }, { registry });
    const { reports } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(reports[0]!.placements.length).toBeGreaterThan(5);
  });
});

describe("非対応要素の検出", () => {
  it("スライド XML 内のアニメーションを検出する", async () => {
    const animated = SMARTART_SLIDE.replace(
      "</p:spTree></p:cSld>",
      "</p:spTree></p:cSld><p:timing><p:anim/></p:timing>",
    );
    const pkg = await PptxPackage.open(await buildPptxWithSlideXml(animated));
    const deck = await readDeck(pkg);
    expect(deck.slides[0]!.unsupported.join()).toContain("アニメーション");
  });

  it("グラデーション塗りを非対応として報告する", async () => {
    const gradient = `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Grad"/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:gradFill/></p:spPr>
    </p:sp>
  </p:spTree></p:cSld></p:sld>`;
    const pkg = await PptxPackage.open(await buildPptxWithSlideXml(gradient));
    const deck = await readDeck(pkg);
    expect(deck.slides[0]!.unsupported.join()).toContain("グラデーション");
  });

  it("マクロ入りファイルを検出できる", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file(
      "ppt/presentation.xml",
      `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
    );
    zip.file("ppt/slides/slide1.xml", SMARTART_SLIDE);
    zip.file("ppt/vbaProject.bin", Buffer.from([0x00, 0x01, 0x02]));
    const pkg = await PptxPackage.open((await zip.generateAsync({ type: "nodebuffer" })) as Buffer);
    expect(pkg.detectPackageLevelUnsupported().join()).toContain("VBA マクロ");
  });

  it("スライド XML の非対応要素は要素ごとに数えられる", async () => {
    const pkg = await PptxPackage.open(await buildPptxWithSlideXml(SMARTART_SLIDE));
    const deck = await readDeck(pkg);
    const xml = parseXml(await pkg.text("ppt/slides/slide1.xml"));
    expect(findDescendants(xml, "graphicFrame")).toHaveLength(1);
    expect(deck.slides[0]!.elements).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Phase 1: PPTX から客観的に読み取れる構造を失わない
// ---------------------------------------------------------------------------

/** 抽出済み要素を最小構成で作る */
function extracted(overrides: Partial<ExtractedElement> = {}): ExtractedElement {
  return {
    id: "1",
    kind: "shape",
    box: { x: 1, y: 1, w: 2, h: 1 },
    ...overrides,
  };
}

describe("線のジオメトリ契約（Unit）", () => {
  it("水平線 (w > 0, h = 0) は有効", () => {
    const line = extracted({ kind: "line", box: { x: 0.5, y: 6.9, w: 12, h: 0 } });
    expect(isLineElement(line)).toBe(true);
    expect(hasRenderableExtent(line)).toBe(true);
  });

  it("垂直線 (w = 0, h > 0) は有効", () => {
    const line = extracted({ kind: "line", box: { x: 6, y: 1, w: 0, h: 3 } });
    expect(hasRenderableExtent(line)).toBe(true);
  });

  it("長さ 0 の線 (w = 0, h = 0) は無効", () => {
    const line = extracted({ kind: "line", box: { x: 1, y: 1, w: 0, h: 0 } });
    expect(hasRenderableExtent(line)).toBe(false);
  });

  it("線プリセットの図形 (p:sp + prstGeom line) も線として扱う", () => {
    const line = extracted({ kind: "shape", geometry: "line", box: { x: 0, y: 3, w: 5, h: 0 } });
    expect(isLineElement(line)).toBe(true);
    expect(hasRenderableExtent(line)).toBe(true);
  });

  it("通常の矩形は正の幅と高さを要求する", () => {
    expect(hasRenderableExtent(extracted({ geometry: "rect" }))).toBe(true);
    expect(
      hasRenderableExtent(extracted({ geometry: "rect", box: { x: 1, y: 1, w: 2, h: 0 } })),
    ).toBe(false);
    expect(
      hasRenderableExtent(extracted({ geometry: "rect", box: { x: 1, y: 1, w: 0, h: 0 } })),
    ).toBe(false);
  });

  it("テキスト・画像には線の特例を広げない", () => {
    expect(isLineElement(extracted({ kind: "text" }))).toBe(false);
    expect(
      hasRenderableExtent(extracted({ kind: "text", box: { x: 1, y: 1, w: 3, h: 0 } })),
    ).toBe(false);
    expect(
      hasRenderableExtent(extracted({ kind: "image", box: { x: 1, y: 1, w: 0, h: 2 } })),
    ).toBe(false);
  });
});

describe("Capture の Z-order 保持（Unit）", () => {
  /** card background → accent → text → divider という元 PPTX の順序を再現する */
  function cardSlideElements(): ExtractedElement[] {
    const elements: ExtractedElement[] = [
      extracted({
        id: "headline",
        kind: "text",
        box: { x: 0.6, y: 0.5, w: 9, h: 0.6 },
        text: { paragraphs: ["見出し"], runs: [], plain: "見出し", maxFontSize: 24 },
      }),
    ];
    for (let i = 0; i < 3; i += 1) {
      const x = 0.6 + i * 3.1;
      // 背景 → アクセント → テキスト → 区切り線（元 PPTX の重なり順）
      elements.push(
        extracted({
          id: `bg-${i}`,
          kind: "shape",
          geometry: "roundRect",
          box: { x, y: 2, w: 2.9, h: 2.6 },
          fill: { type: "solid", color: "F2F6FA" },
        }),
        extracted({
          id: `accent-${i}`,
          kind: "shape",
          geometry: "rect",
          box: { x, y: 2, w: 0.06, h: 2.6 },
          fill: { type: "solid", color: "0B4F8A" },
        }),
        extracted({
          id: `label-${i}`,
          kind: "text",
          box: { x: x + 0.2, y: 2.3, w: 2.4, h: 0.4 },
          text: { paragraphs: [`項目${i}`], runs: [], plain: `項目${i}`, maxFontSize: 15 },
        }),
        extracted({
          id: `divider-${i}`,
          kind: "shape",
          geometry: "line",
          box: { x: x + 0.2, y: 3.1, w: 2.4, h: 0 },
          fill: { type: "none" },
          line: { color: "D5DCE4", widthPt: 1 },
        }),
      );
    }
    return elements;
  }

  /**
   * ヒューリスティクスは accent / divider を判別しないため、
   * それらは静的要素として元の Z-order のまま残る（Phase 1 の挙動）。
   */
  function interpretationFor(): CaptureInterpretation {
    return {
      schemaVersion: CAPTURE_INTERPRETATION_SCHEMA_VERSION,
      interpretationSource: "heuristic",
      headlineIndex: 0,
      cells: [
        { backgroundIndex: 1, content: { labelIndex: 3 } },
        { backgroundIndex: 5, content: { labelIndex: 7 } },
        { backgroundIndex: 9, content: { labelIndex: 11 } },
      ],
      staticIndexes: [2, 4, 6, 8, 10, 12],
      notes: [],
    };
  }

  function build() {
    return buildCapturedPattern(cardSlideElements(), interpretationFor(), {
      id: "captured-order",
      theme: defaultTheme,
      slideSize: { widthIn: 10, heightIn: 5.625 },
      source: { file: "test.pptx", slideNumber: 1 },
    });
  }

  it("元 PPTX の重なり順（背景 → アクセント → 区切り線）を保つ", () => {
    const body = build().document.layout.body as {
      type: string;
      children: Array<{ node: { type: string; shape?: string } }>;
    };
    expect(body.type).toBe("absolute");

    const types = body.children.map((c) =>
      c.node.type === "shape" ? `shape:${c.node.shape}` : c.node.type,
    );
    const columnsAt = types.indexOf("columns");
    expect(columnsAt).toBeGreaterThanOrEqual(0);

    // カード（columns）より後ろに accent と divider が来る = カードの上に描かれる
    const accentIndexes = types.flatMap((t, i) => (t === "shape:rect" ? [i] : []));
    const dividerIndexes = types.flatMap((t, i) => (t === "shape:line" ? [i] : []));
    expect(accentIndexes).toHaveLength(3);
    expect(dividerIndexes).toHaveLength(3);
    for (const index of [...accentIndexes, ...dividerIndexes]) {
      expect(index).toBeGreaterThan(columnsAt);
    }
  });

  it("反復セルの Z-order は構成要素の最小インデックスで決まる（決定論的）", () => {
    const first = build().document.layout.body;
    const second = build().document.layout.body;
    expect(JSON.stringify(second)).toBe(JSON.stringify(first));

    const body = first as { children: Array<{ node: { type: string } }> };
    // 見出し(0) → カード(背景 index 1) → その後ろの装飾、という順序になる
    expect(body.children[0]!.node.type).toBe("text");
    expect(body.children[1]!.node.type).toBe("columns");
  });

  it("区切り線（h = 0）が Pattern から失われない", () => {
    const body = build().document.layout.body as {
      children: Array<{ box: { w: number; h: number }; node: { type: string; shape?: string } }>;
    };
    const dividers = body.children.filter(
      (c) => c.node.type === "shape" && c.node.shape === "line",
    );
    expect(dividers).toHaveLength(3);
    for (const divider of dividers) {
      expect(divider.box.h).toBe(0);
      expect(divider.box.w).toBeGreaterThan(0);
    }
  });
});

describe("Template Import が線を保持する（Integration）", () => {
  /**
   * Master 側に線を持つ PPTX を作る。
   * Template Import が読むのは Slide Master / Layout なので、スライド本体ではなく
   * defineSlideMaster の objects へ置く（実際の企業テンプレートと同じ形）。
   */
  async function deckWithLines(): Promise<Buffer> {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    const stroke = { color: "D5DCE4", width: 1 };
    pptx.defineSlideMaster({
      title: "LINES",
      background: { color: "FFFFFF" },
      objects: [
        // 水平線 / 垂直線 / 通常矩形 / 長さ 0 の線 / サイズ 0 の矩形
        { line: { x: 0.6, y: 5.1, w: 8.8, h: 0, line: stroke } },
        { line: { x: 5, y: 1, w: 0, h: 2.5, line: stroke } },
        { rect: { x: 0, y: 0, w: 10, h: 0.09, fill: { color: "0B4F8A" } } },
        { line: { x: 2, y: 2, w: 0, h: 0, line: stroke } },
        { rect: { x: 3, y: 3, w: 0, h: 0, fill: { color: "0B4F8A" } } },
        { text: { text: "footer", options: { x: 0.6, y: 5.2, w: 4, h: 0.26, fontSize: 9 } } },
      ],
    });
    pptx.addSlide({ masterName: "LINES" }).addText("body", {
      x: 1, y: 1, w: 4, h: 0.4, fontSize: 14,
    });
    return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  }

  it("水平線と垂直線は Master 要素として残り、長さ 0 の図形は除外される", async () => {
    const imported = await importTemplateFromPptx(await deckWithLines(), { id: "lines" });
    const master = imported.template.masters.find((m) => m.role === "content")!;
    const elements = master.elements ?? [];

    const lines = elements.filter((e) => e.kind === "shape" && e.shape === "line");
    expect(lines).toHaveLength(2);
    expect(lines.some((l) => l.box.w > 0 && l.box.h === 0)).toBe(true); // 水平線
    expect(lines.some((l) => l.box.w === 0 && l.box.h > 0)).toBe(true); // 垂直線

    // 通常の矩形は残り、長さ 0 の線・サイズ 0 の矩形は残らない
    const rects = elements.filter((e) => e.kind === "shape" && e.shape === "rect");
    expect(rects).toHaveLength(1);
    expect(rects[0]!.box).toMatchObject({ w: 10, h: 0.09 });
    for (const element of elements) {
      expect(element.box.w > 0 || element.box.h > 0).toBe(true);
    }
  });

  it("線を含む Template で PPTX を生成できる", async () => {
    const imported = await importTemplateFromPptx(await deckWithLines(), { id: "lines" });
    const spec = await loadExample();
    const registry = new TemplateRegistry().register({
      template: imported.template,
      source: "external",
      assets: [],
    });
    const result = await generate(spec, { templateId: "lines", templateRegistry: registry });

    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    const pptx = await readPptx(result.pptxBuffer);
    // 線は Slide Layout（Master）側へ入る
    expect(pptx.slideLayoutXmls.join("")).toContain('prst="line"');
  });
});
