import { readFile, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePresentation } from "../src/domain/schema.js";
import { buildVisualPlan, parseVisualPlan } from "../src/plan/visual-plan.js";
import { generate, generateFromFile } from "../src/pipeline.js";
import { renderVisualPlan } from "../src/render/renderer.js";
import { defaultRegistry, PatternRegistry } from "../src/patterns/registry.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { partitionWithUnknownsPattern, PartitionBindingsSchema } from "../src/patterns/partition-with-unknowns/index.js";
import { checkSelectable } from "../src/selector/selector.js";
import { buildArgumentReviewPacket, validateArgumentReview, ARGUMENT_CHECKS } from "../src/review/argument.js";
import { main } from "../src/cli/main.js";
import { extractTexts, readPptx } from "./helpers.js";

const inputFile = "examples/arguments/nec-candidate-quality-2026-10-16.json";
const load = async () => parsePresentation(JSON.parse(await readFile(inputFile, "utf8")));
const reviewed = (input: unknown) => ({ schemaVersion: "1.0", inputHash: buildArgumentReviewPacket(input).inputHash,
  actor: { kind: "assistant", id: "test-fixture" }, decision: "accept",
  findings: ARGUMENT_CHECKS.map(check => ({ check, assessment: "Test fixture only, not a real semantic review." })) });

describe("partition argument in normal generation", () => {
  it("selects the dedicated pattern deterministically and preserves original semantic design", async () => {
    const spec = await load(), first = buildVisualPlan(spec), second = buildVisualPlan(spec);
    expect(first).toEqual(second);
    const selected = first.slides[0]!;
    expect(selected.patternId).toBe("partition-with-unknowns");
    expect(selected.fallbackUsed).toBe(false);
    const b = PartitionBindingsSchema.parse(selected.bindings);
    expect(b.evidence).toEqual(spec.slides[0]!.evidence[0]);
    expect(b.argument).toEqual(spec.slides[0]!.argument);
    expect(b.sources).toEqual(spec.slides[0]!.sources);
    expect(b.calculation.remainder).toBe(4491);
    expect(b.calculation.remainderPercent).toBe(89.2);
    expect(b.evidence.data.quality.selected).not.toHaveProperty("value");
  });

  it("rejects unsuitable overrides and fails closed if the dedicated pattern is unavailable", async () => {
    const spec = await load();
    for (const p of defaultRegistry.all().filter(p => !p.fallback && p.id !== partitionWithUnknownsPattern.id)) {
      expect(checkSelectable(p, spec.slides[0]!).ok).toBe(false);
    }
    expect(() => buildVisualPlan(spec, { registry: new PatternRegistry([messageOnlyPattern]) })).toThrow("refusing lossy fallback");
    expect(() => buildVisualPlan(spec, { minimumScore: 1.1 })).toThrow("refusing lossy fallback");
    const broken = { ...partitionWithUnknownsPattern, buildBindings: () => { throw new Error("bug"); } };
    expect(() => buildVisualPlan(spec, { registry: new PatternRegistry([broken,messageOnlyPattern]) })).toThrow("refusing lossy fallback");
  });

  it.each(["selectedFraction", "remainder", "remainderPercent"])("rejects tampered %s during replay", async key => {
    const plan = buildVisualPlan(await load());
    (plan.slides[0]!.bindings as any).calculation[key] = .5;
    expect(() => parseVisualPlan(plan)).toThrow();
    expect(() => renderVisualPlan(plan)).toThrow();
  });

  it("rejects altered trace, missing source, unsafe URL and invented quality values", async () => {
    const plan = buildVisualPlan(await load());
    for (const mutate of [
      (b: any) => { b.calculation.trace.inputs.selected = 99; },
      (b: any) => { b.sources = []; },
      (b: any) => { b.sources[0].url = "javascript:alert(1)"; },
      (b: any) => { b.evidence.data.quality.selected.value = 0; },
      (b: any) => { b.argument.evidenceId = "missing"; },
    ]) {
      const copy = structuredClone(plan); mutate(copy.slides[0]!.bindings);
      expect(() => renderVisualPlan(copy)).toThrow();
    }
  });

  it("uses unrounded quantity geometry and native editable objects, without quality area or arrowheads", async () => {
    const result = await generate(await load(), { themeId:"portable", includeTitleSlide:false });
    expect(result.issues).toEqual([]);
    const placements = result.reports[0]!.placements;
    const selected = placements.find(p => p.name === "selected-quantity-band")!.box;
    const remainder = placements.find(p => p.name === "remainder-quantity-band")!.box;
    expect(selected.w / (selected.w + remainder.w)).toBeCloseTo(544 / 5035, 4);
    expect(selected.x + selected.w).toBeCloseTo(remainder.x, 4);
    expect(placements.filter(p => p.name.startsWith("unknown-quality")).every(p => p.kind === "text")).toBe(true);
    const pptx = await readPptx(result.pptxBuffer), xml = pptx.slideXmls[0]!;
    const text = extractTexts(xml).join("\n");
    for (const expected of ["5,035", "平均 544", "4,491", "12事象", "10.8%", "89.2%", "計算", "未確認", "発表本文・注記", "各検索の再現率", "候補平均544"]) expect(text).toContain(expected);
    expect(xml).not.toContain("<p:pic>");
    expect(xml).not.toMatch(/(?:headEnd|tailEnd)[^>]*type="triangle"/);
    expect(xml).toContain("hlinkClick");
    const rels = await pptx.zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    expect(rels).toContain("https://jpn.nec.com/press/202610/20261014_01.html");
  });

  it("supports different values/units and endpoint quantities without fabricated area", async () => {
    for (const selected of [0, 25, 200]) {
      const spec = await load(), e = spec.slides[0]!.evidence[0]!;
      if (e.type !== "partition") throw new Error("fixture");
      e.data.whole = 200; e.data.selected = selected; e.data.unit = "件"; e.data.aggregation = {kind:"single"};
      const result = await generate(spec, {themeId:"portable",includeTitleSlide:false});
      expect(result.issues).toEqual([]);
      const band = result.reports[0]!.placements.filter(p => p.name.endsWith("quantity-band"));
      expect(band.length).toBe(selected===0 || selected===200 ? 1 : 2);
      if (selected===0 || selected===200) expect(result.reports[0]!.placements.filter(p=>p.name.startsWith("unknown-quality")).map(p=>p.text).join()).toContain("対象がないため内訳なし");
    }
  });

  it("rejects overflow without truncation or a generated output file", async () => {
    const spec = await load(); spec.slides[0]!.argument!.inference = "長い説明".repeat(45);
    const dir = await mkdtemp(join(tmpdir(),"sdg-overflow-"));
    try {
      const file = join(dir,"input.json"), outDir=join(dir,"output"); await writeFile(file,JSON.stringify(spec));
      await expect(generateFromFile(file,{themeId:"portable",includeTitleSlide:false,outDir})).rejects.toThrow("No content was omitted");
      await expect(readdir(outDir)).rejects.toMatchObject({code:"ENOENT"});
    } finally { await rm(dir,{recursive:true,force:true}); }
  });
});

describe("revision-bound argument review", () => {
  it("creates unapproved packets and invalidates review for claims, data, source or interpretation changes", async () => {
    const spec = await load(), review = reviewed(spec), packet = buildArgumentReviewPacket(spec);
    expect(packet.slides[0]!.supportingMessage).toBe(spec.slides[0]!.supportingMessage);
    expect(packet.status).toBe("unreviewed"); expect(packet.reviewTemplate.decision).toBe("request-changes");
    expect(validateArgumentReview(spec,review).actor.kind).toBe("assistant");
    for (const mutate of [
      (s: any) => { s.slides[0].keyMessage = "Changed claim"; },
      (s: any) => { s.slides[0].evidence[0].data.selected = 100; },
      (s: any) => { s.slides[0].sources[0].accessedAt = "2026-10-17"; },
      (s: any) => { s.slides[0].argument.annotations.quantity.text = "Changed interpretation"; },
    ]) { const copy=structuredClone(spec);mutate(copy);expect(()=>validateArgumentReview(copy,review)).toThrow("stale"); }
    expect(()=>validateArgumentReview(spec,packet.reviewTemplate)).toThrow();
    expect(()=>validateArgumentReview(spec,{...review,decision:"request-changes"})).toThrow();
    expect(()=>validateArgumentReview(spec,{...review,findings:review.findings.map(()=>review.findings[0])})).toThrow();
  });

  it("CLI rejects stale supplied review before output and retains actual reviewer in report", async () => {
    const spec = await load(), dir=await mkdtemp(join(tmpdir(),"sdg-review-")), reviewFile=join(dir,"review.json");
    try {
      const review=reviewed(spec); review.inputHash="0".repeat(64);await writeFile(reviewFile,JSON.stringify(review));
      const args=["generate",inputFile,"--out-dir",join(dir,"out"),"--argument-review",reviewFile,"--theme","portable","--no-title-slide","--report","--strict","--quiet"];
      await expect(main(args)).rejects.toThrow("stale"); await expect(readdir(join(dir,"out"))).rejects.toMatchObject({code:"ENOENT"});
      await writeFile(reviewFile,JSON.stringify(reviewed(spec))); expect(await main(args)).toBe(0);
      const report=JSON.parse(await readFile(join(dir,"out","nec-candidate-quality-2026-10-16.structure-review.json"),"utf8"));
      expect(report.argumentReview.status).toBe("review-recorded"); expect(report.argumentReview.review.actor.kind).toBe("assistant");
      expect(await main(["argument-review",inputFile,"--out",join(dir,"packet.json")])).toBe(0);
      expect(JSON.parse(await readFile(join(dir,"packet.json"),"utf8")).reviewTemplate.decision).toBe("request-changes");
    } finally { await rm(dir,{recursive:true,force:true}); }
  });
});
