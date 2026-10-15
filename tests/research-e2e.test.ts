import { readFile, mkdtemp, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { describe, it, expect } from "vitest";
import { generateFromFile } from "../src/pipeline.js";
import { parseVisualPlan } from "../src/plan/visual-plan.js";
import { readPptx, extractTexts } from "./helpers.js";

describe("verified research end-to-end", () => {
  it("creates the real four-slide deck with provenance, caveats and structural diagnostics", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "sdg-research-"));
    try {
      const result = await generateFromFile("examples/research/nec-disaster-vlm-2026-10-14.json", { outDir, themeId: "portable", includeTitleSlide: false, writeReport: true });
      expect(result.issues).toEqual([]);
      expect(result.plan.slides.map(s => s.patternId)).toEqual(["three-reasons", "convergence", "process", "three-reasons"]);
      expect(result.plan.slides.some(s => s.fallbackUsed)).toBe(false);
      const savedPlan = parseVisualPlan(JSON.parse(await readFile(result.planPath, "utf8")));
      expect(savedPlan).toEqual(result.plan);
      const review = JSON.parse(await readFile(result.structureReviewPath!, "utf8"));
      expect(review.slides).toHaveLength(4);
      expect(review.slides.every((s: { semanticMECE: string }) => s.semanticMECE === "unverified")).toBe(true);
      const pptx = await readPptx(await readFile(result.pptxPath));
      expect(pptx.slideXmls).toHaveLength(4);
      const text = pptx.slideXmls.flatMap(extractTexts).join("\n");
      for (const fact of ["2027年度中", "平均544枚", "12事象", "1枚5秒", "約7時間", "約45分", "解釈・提案", "再現率の数値がない"]) expect(text).toContain(fact);
      const notes = (await Promise.all(pptx.files.filter(f => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f)).map(f => pptx.zip.file(f)!.async("string")))).join("\n");
      expect(notes).toContain("https://jpn.nec.com/press/202610/20261014_01.html");
      expect(notes).toContain("2026-10-15");
      expect(notes).toContain("ROI");
      expect(JSON.stringify(result.plan.slides[1]!.bindings)).toContain("nec-02-learning-inputs-outcome");
    } finally { await rm(outDir, { recursive: true, force: true }); }
  });
});
