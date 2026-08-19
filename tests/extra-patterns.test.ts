import { describe, expect, it } from "vitest";
import type { BigNumberBindings } from "../src/patterns/big-number/index.js";
import { bigNumberPattern } from "../src/patterns/big-number/index.js";
import type { ChartWithMessageBindings } from "../src/patterns/chart-with-message/index.js";
import { chartWithMessagePattern } from "../src/patterns/chart-with-message/index.js";
import { checkConstraints } from "../src/patterns/pattern.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { extractTexts, loadExample, metricEvidence, readPptx, slide, textEvidence } from "./helpers.js";

const comparisonSlide = slide({
  id: "s-chart",
  keyMessage: "コストは3年連続で増加している",
  slideRole: "problem",
  logicalRelation: "claim_evidence",
  evidence: [
    {
      id: "e1",
      type: "comparison",
      statement: "コストは3,800万円から4,800万円へ増加した",
      data: {
        unit: "万円",
        items: [
          { label: "2022年度", value: 3800 },
          { label: "2023年度", value: 4250 },
          { label: "2024年度", value: 4800 },
        ],
      },
    },
  ],
});

describe("BigNumber", () => {
  const metricSlide = slide({
    keyMessage: "2年以内に投資回収できる",
    evidence: [metricEvidence("e1", "投資回収期間", 1.7, "年")],
  });

  it("metric が 1 件のときに適合する", () => {
    const result = bigNumberPattern.match(metricSlide);
    expect(result.compatible).toBe(true);
    expect(result.score).toBeGreaterThan(0.9);
  });

  it("metric が 2 件以上あるときは適合しない（主役を決められない）", () => {
    const result = bigNumberPattern.match(
      slide({ evidence: [metricEvidence("e1", "A", 1), metricEvidence("e2", "B", 2)] }),
    );
    expect(result.compatible).toBe(false);
  });

  it("metric が無いスライドは constraints で落ちる", () => {
    const violations = checkConstraints(
      slide({ evidence: [textEvidence("e1", "定性的な根拠")] }),
      bigNumberPattern.constraints,
    );
    expect(violations.map((v) => v.constraint)).toContain("requiresAnyEvidenceType");
  });

  it("buildBindings が単位付きの整形済み値を返す", () => {
    const bindings = bigNumberPattern.buildBindings(metricSlide, "standard") as BigNumberBindings;
    expect(bindings.value).toBe("1.7年");
    expect(bindings.caption).toBe("投資回収期間");
    expect(bindings.note).toContain("投資回収期間");
  });

  it("千区切りはロケールに依存しない", () => {
    const bindings = bigNumberPattern.buildBindings(
      slide({ evidence: [metricEvidence("e1", "削減額", 1440, "万円")] }),
      "standard",
    ) as BigNumberBindings;
    expect(bindings.value).toBe("1,440万円");
  });
});

describe("ChartWithMessage", () => {
  it("comparison evidence から系列を抽出できる", () => {
    const result = chartWithMessagePattern.match(comparisonSlide);
    expect(result.compatible).toBe(true);
    expect(result.score).toBeGreaterThan(0.9);
  });

  it("数値系列が無ければ適合しない", () => {
    const result = chartWithMessagePattern.match(
      slide({ evidence: [textEvidence("e1", "定性的な根拠")] }),
    );
    expect(result.compatible).toBe(false);
  });

  it("単位が揃った metric 3 件以上も系列として扱える", () => {
    const result = chartWithMessagePattern.match(
      slide({
        evidence: [
          metricEvidence("e1", "拠点A", 30, "%"),
          metricEvidence("e2", "拠点B", 45, "%"),
          metricEvidence("e3", "拠点C", 22, "%"),
        ],
      }),
    );
    expect(result.compatible).toBe(true);
  });

  it("単位が揃っていない metric 群は系列にしない", () => {
    const result = chartWithMessagePattern.match(
      slide({
        evidence: [
          metricEvidence("e1", "初期投資", 2400, "万円"),
          metricEvidence("e2", "削減額", 1440, "万円"),
          metricEvidence("e3", "回収期間", 1.7, "年"),
        ],
      }),
    );
    expect(result.compatible).toBe(false);
  });

  it("buildBindings がカテゴリ・系列・変化量を作る", () => {
    const bindings = chartWithMessagePattern.buildBindings(
      comparisonSlide,
      "standard",
    ) as ChartWithMessageBindings;

    expect(bindings.chart.categories).toEqual(["2022年度", "2023年度", "2024年度"]);
    expect(bindings.chart.series).toHaveLength(1);
    expect(bindings.chart.series[0]!.values).toEqual([3800, 4250, 4800]);
    // 系列名は凡例に収まるよう切り詰められる
    expect(bindings.chart.series[0]!.name).toBe("コストは3,800万円から4,800万円へ増加…");
    expect(bindings.chart.unit).toBe("万円");
    expect(bindings.changeText).toBe("3,800万円 → 4,800万円 (+26.3%)");
  });
});

describe("ネイティブグラフの出力", () => {
  it("画像ではなく編集可能なグラフとして埋め込まれる", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { files, zip } = await readPptx(await toPptxBuffer(pptx));

    const chartFiles = files.filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f));
    expect(chartFiles).toHaveLength(1);

    const chartXml = await zip.file(chartFiles[0]!)!.async("string");
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("2022年度");
    expect(chartXml).toContain("<c:v>3800</c:v>");
    // 系列色は Theme の primary
    expect(chartXml).toContain("0B4F8A");
    // 元データが埋め込まれている = PowerPoint 上でデータ編集できる
    expect(files.some((f) => f.startsWith("ppt/embeddings/"))).toBe(true);
  });

  it("グラフスライドにも読み取り文がテキストとして載る", async () => {
    const plan = buildVisualPlan(await loadExample());
    const { pptx } = renderVisualPlan(plan);
    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
    const texts = extractTexts(slideXmls[1]!).join("\n");
    expect(texts).toContain("3,800万円 → 4,800万円 (+26.3%)");
  });
});
