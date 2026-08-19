import { describe, expect, it } from "vitest";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { PatternRegistry } from "../src/patterns/registry.js";
import {
  buildVisualPlan,
  parseVisualPlan,
  serializeVisualPlan,
  VisualPlanValidationError,
} from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { validatePlacements } from "../src/render/validate.js";
import type { Placement } from "../src/render/surface.js";
import { midnightTheme } from "../src/theme/theme.js";
import {
  countMatches,
  extractTexts,
  loadExample,
  metricEvidence,
  pointsContent,
  readPptx,
  slide,
} from "./helpers.js";

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

  it("plan.themeId だけで Theme が解決される（オプション無しの再生で色が変わらない）", async () => {
    const plan = buildVisualPlan(await loadExample(), { themeId: "midnight" });
    // options.theme を渡さずに描画しても plan.themeId (midnight) が使われる
    const { pptx } = renderVisualPlan(plan);
    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
    expect(slideXmls[1]).toContain("5AA9F0"); // midnight primary
    expect(slideXmls[1]).not.toContain("0B4F8A"); // default primary は現れない
  });

  it("完全な出典（URL・取得日）がスピーカーノートに格納される", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { zip, files } = await readPptx(await toPptxBuffer(pptx));

    const notesFiles = files.filter((f) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f));
    expect(notesFiles.length).toBeGreaterThan(0);

    const allNotes = (
      await Promise.all(notesFiles.map((f) => zip.file(f)!.async("string")))
    ).join("\n");
    // フッターでは失われる URL / 取得日 / 参照箇所がノートには残る
    expect(allNotes).toContain("https://example.com/case/a-company");
    expect(allNotes).toContain("取得日: 2025-02-10");
    expect(allNotes).toContain("p.12 表2-4");
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
    // 保存 → 再読込 → parseVisualPlan による検証 → 描画、という再生経路
    const roundTripped = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));

    const fromObject = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
    const fromJson = await readPptx(await toPptxBuffer(renderVisualPlan(roundTripped).pptx));
    expect(fromJson.slideXmls).toEqual(fromObject.slideXmls);
  });

  /**
   * renderVisualPlan は描画の境界であり、parseVisualPlan と同じ検証を必ず通す。
   * 「parseVisualPlan は拒否するが renderVisualPlan は受け入れる」抜け道を塞ぐ。
   */
  describe("renderVisualPlan の入口検証", () => {
    async function tamperedPlan(
      mutate: (plan: Record<string, any>) => void,
    ): Promise<Record<string, any>> {
      const plan = JSON.parse(serializeVisualPlan(buildVisualPlan(await loadExample())));
      mutate(plan);
      return plan;
    }

    it("patternVersion が一致しない Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        p.slides[0].patternVersion = "0.0.1";
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(/バージョンが不一致/);
    });

    it("bindings が契約違反の Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        p.slides[0].bindings = { garbage: true };
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(/契約に違反/);
    });

    it("schemaVersion が一致しない Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        p.schemaVersion = "0.9";
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(VisualPlanValidationError);
      expect(() => renderVisualPlan(plan as any)).toThrow(/schemaVersion/);
    });

    it("未実装の variantId を持つ Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        p.slides[0].variantId = "compact";
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(VisualPlanValidationError);
      expect(() => renderVisualPlan(plan as any)).toThrow(/variant/);
    });

    it("未登録の patternId を持つ Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        p.slides[0].patternId = "nonexistent-pattern";
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(/未登録の patternId/);
    });

    it("構造が壊れた Plan の描画は拒否される", async () => {
      const plan = await tamperedPlan((p) => {
        delete p.slides[0].slideId;
      });
      expect(() => renderVisualPlan(plan as any)).toThrow(VisualPlanValidationError);
    });

    it("parseVisualPlan 済みの正常な Plan は問題なく描画できる（検証は冪等）", async () => {
      const plan = buildVisualPlan(await loadExample());
      const validated = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)));

      const direct = await readPptx(await toPptxBuffer(renderVisualPlan(plan).pptx));
      const viaParse = await readPptx(await toPptxBuffer(renderVisualPlan(validated).pptx));

      expect(viaParse.slideXmls).toEqual(direct.slideXmls);
      expect(renderVisualPlan(validated).issues).toEqual([]);
    });

    it("カスタム Registry を渡すとその Registry を基準に検証・描画される", async () => {
      const registry = new PatternRegistry([messageOnlyPattern]);
      const spec = {
        title: "カスタム Registry",
        slides: [
          slide({
            id: "slide-custom",
            keyMessage: "message-only だけのレジストリでも描画できる",
            evidence: [],
          }),
        ],
      };
      const plan = buildVisualPlan(spec, { registry });
      expect(plan.slides[0]!.patternId).toBe("message-only");

      const { pptx, issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
      expect(issues).toEqual([]);
      const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
      expect(extractTexts(slideXmls[0]!).join("")).toContain(
        "message-only だけのレジストリでも描画できる",
      );
    });

    it("既定 Registry に無い Pattern の Plan は、そのカスタム Registry でのみ描画できる", async () => {
      // 既定 Registry で作った Plan を message-only だけの Registry へ渡すと拒否される
      const plan = buildVisualPlan(await loadExample());
      const narrowRegistry = new PatternRegistry([messageOnlyPattern]);
      expect(() => renderVisualPlan(plan, { registry: narrowRegistry })).toThrow(
        /未登録の patternId/,
      );
    });
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

  it("three-reasons の数値付きカードでも数値行がはみ出さない", () => {
    // 数値は heading より一段大きく描かれるため、行の高さを固定値にすると必ずはみ出す
    const s = slide({
      keyMessage: "3 つの構造的な原因がある",
      content: pointsContent(
        { label: "定型的な指摘に人手を使っている", description: "規約・命名の指摘の割合", evidenceRefs: ["e1"] },
        { label: "レビューが特定の担当者に集中している", description: "上位層が担当する割合", evidenceRefs: ["e2"] },
        { label: "レビュー開始までの待ち時間が長い", description: "初回コメントまでの中央値", evidenceRefs: ["e3"] },
      ),
      evidence: [
        metricEvidence("e1", "定型指摘の割合", 35, "%"),
        metricEvidence("e2", "上位層が占める比率", 58, "%"),
        metricEvidence("e3", "初回コメントまでの中央値", 18, "時間"),
      ],
    });

    const plan = buildVisualPlan({ title: "metric cards", slides: [s] });
    expect(plan.slides[0]!.patternId).toBe("three-reasons");

    const { issues } = renderVisualPlan(plan, { includeTitleSlide: false });
    expect(issues).toEqual([]);
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
