import { describe, expect, it } from "vitest";
import { buildVisualPlan, serializeVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { validatePlacements } from "../src/render/validate.js";
import type { Placement } from "../src/render/surface.js";
import { midnightTheme } from "../src/theme/theme.js";
import { countMatches, extractTexts, loadExample, readPptx } from "./helpers.js";

describe("PPTX 生成", () => {
  it("16:9 のスライドが表紙 + 本文ぶん生成される", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { slideXmls, presentationXml } = await readPptx(await toPptxBuffer(pptx));

    expect(slideXmls).toHaveLength(7); // 表紙 1 + 本文 6
    // 9144000 x 5143500 EMU = 10in x 5.625in = 16:9
    expect(presentationXml).toContain('cx="9144000"');
    expect(presentationXml).toContain('cy="5143500"');
  });

  it("KeyMessage が編集可能なテキストとして格納されている", async () => {
    const spec = await loadExample();
    const { pptx } = renderVisualPlan(buildVisualPlan(spec));
    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));

    spec.slides.forEach((s, i) => {
      const texts = extractTexts(slideXmls[i + 1]!).join("");
      expect(texts).toContain(s.keyMessage);
    });
  });

  it("画像化されておらず、テキストと図形で構成されている", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { slideXmls, files } = await readPptx(await toPptxBuffer(pptx));

    expect(files.filter((f) => f.startsWith("ppt/media/"))).toHaveLength(0);
    for (const xml of slideXmls) {
      expect(xml).not.toContain("<p:pic>");
      expect(countMatches(xml, /<p:sp>/g)).toBeGreaterThan(0);
      expect(countMatches(xml, /<a:t>/g)).toBeGreaterThan(0);
    }
  });

  it("出典・補足テキストもスライドに含まれる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));

    const all = slideXmls.flatMap(extractTexts).join("\n");
    expect(all).toContain("出典: A社 在庫管理システム導入事例, 2024-11");
    expect(all).toContain("STEP 1");
    expect(all).toContain("4,800万円");
    expect(all).toContain("-30%");
  });

  it("表紙を省略できる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan, { includeTitleSlide: false });
    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
    expect(slideXmls).toHaveLength(6);
  });

  it("Theme を差し替えても Pattern を変更せずに配色が変わる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const light = renderVisualPlan(plan);
    const dark = renderVisualPlan(plan, { theme: midnightTheme });

    const lightXml = (await readPptx(await toPptxBuffer(light.pptx))).slideXmls[1]!;
    const darkXml = (await readPptx(await toPptxBuffer(dark.pptx))).slideXmls[1]!;

    expect(lightXml).toContain("0B4F8A"); // default primary
    expect(darkXml).toContain("5AA9F0"); // midnight primary
    expect(darkXml).not.toContain("0B4F8A");
    // 構造（要素数）は同じ
    expect(countMatches(darkXml, /<p:sp>/g)).toBe(countMatches(lightXml, /<p:sp>/g));
    expect(extractTexts(darkXml)).toEqual(extractTexts(lightXml));
  });

  it("同じ Visual Plan からは同じスライド XML が生成される（決定論的）", async () => {
    const plan = buildVisualPlan(await loadExample());
    const first = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    const second = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    expect(second.slideXmls).toEqual(first.slideXmls);
  });

  /**
   * 決定論性の正確な保証範囲を明文化しておく。
   * - スライド XML / グラフの中身: 完全に一致する
   * - グラフのパート番号 (chart1.xml → chart2.xml): PptxGenJS がプロセス内で
   *   通し番号を振るため、同一プロセスで 2 回描くとずれる（中身は同一）
   * - docProps/core.xml: 生成時刻が埋め込まれるため一致しない
   */
  it("同一プロセスで 2 回描いてもグラフの内容は同一（パート番号のみ変わる）", async () => {
    const plan = buildVisualPlan(await loadExample());
    const first = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    const second = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));

    const chartOf = (files: string[]) => files.find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))!;
    const nameA = chartOf(first.files);
    const nameB = chartOf(second.files);

    const xmlA = await first.zip.file(nameA)!.async("string");
    const xmlB = await second.zip.file(nameB)!.async("string");
    expect(xmlB).toBe(xmlA);
  });

  it("Renderer は Visual Plan の JSON だけで描画できる（SlideSpec に依存しない）", async () => {
    const plan = buildVisualPlan(await loadExample());
    const roundTripped = JSON.parse(serializeVisualPlan(plan));

    const fromObject = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    const fromJson = await readPptx(await toPptxBuffer(renderVisualPlan(roundTripped).pptx));
    expect(fromJson.slideXmls).toEqual(fromObject.slideXmls);
  });
});

describe("レイアウト検査", () => {
  it("サンプル資料はエラー・警告ゼロで描画できる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { issues } = renderVisualPlan(plan);
    expect(issues).toEqual([]);
  });

  it("全要素が安全領域に収まっている", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { reports } = renderVisualPlan(plan);
    for (const report of reports) {
      expect(report.placements.length).toBeGreaterThan(0);
      expect(report.issues.filter((i) => i.severity === "error")).toEqual([]);
    }
  });

  it("content レイヤの重なりを検出する", () => {
    const placements: Placement[] = [
      { kind: "text", layer: "content", box: { x: 1, y: 1, w: 3, h: 1 }, name: "a" },
      { kind: "text", layer: "content", box: { x: 2, y: 1.2, w: 3, h: 1 }, name: "b" },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues.map((i) => i.code)).toContain("overlap");
  });

  it("decoration レイヤ同士の重なりは許容する（カード背景など）", () => {
    const placements: Placement[] = [
      { kind: "shape", layer: "decoration", box: { x: 1, y: 1, w: 3, h: 1 }, name: "card" },
      { kind: "text", layer: "content", box: { x: 1.2, y: 1.1, w: 2, h: 0.5 }, name: "label" },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues).toEqual([]);
  });

  it("スライド外への飛び出しを検出する", () => {
    const placements: Placement[] = [
      { kind: "shape", layer: "content", box: { x: 9.5, y: 1, w: 2, h: 1 }, name: "overflow" },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues[0]!.code).toBe("off-canvas");
    expect(issues[0]!.severity).toBe("error");
  });

  it("テキストがボックスに収まらない場合を検出する", () => {
    const placements: Placement[] = [
      {
        kind: "text",
        layer: "content",
        box: { x: 1, y: 1, w: 2, h: 0.3 },
        name: "long",
        text: "非常に長い日本語のテキストが小さなボックスに入れられている状況を検出できるかどうかを確認する",
        fontSize: 14,
        lineHeight: 1.3,
      },
    ];
    const issues = validatePlacements(placements, {
      slideId: "s",
      slideNumber: 1,
      patternId: "p",
      canvas: { x: 0, y: 0, w: 10, h: 5.625 },
      frame: { x: 0.5, y: 0.5, w: 9, h: 4.6 },
    });
    expect(issues.map((i) => i.code)).toContain("text-overflow");
  });
});
