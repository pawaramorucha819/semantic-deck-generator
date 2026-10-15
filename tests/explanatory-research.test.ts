import { readFile, mkdtemp, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { describe, it, expect } from "vitest";
import { generateFromFile } from "../src/pipeline.js";
import { readPptx, extractTexts } from "./helpers.js";

describe("standalone NEC explanatory deck", () => {
  it("keeps explanations and reading-critical conditions in visible slides, not just notes", async () => {
    const file = "examples/research/nec-explained-2026-10-15.json";
    const input = JSON.parse(await readFile(file, "utf8"));
    const outDir = await mkdtemp(join(tmpdir(), "sdg-explained-"));
    try {
      const result = await generateFromFile(file, { outDir, themeId: "portable", includeTitleSlide: false, writeReport: true });
      expect(result.issues).toEqual([]);
      expect(result.plan.slides.some(s => s.fallbackUsed)).toBe(false);
      const pptx = await readPptx(await readFile(result.pptxPath));
      expect(pptx.slideXmls).toHaveLength(input.slides.length);
      for (let i = 0; i < input.slides.length; i++) {
        const text = extractTexts(pptx.slideXmls[i]!).join("\n");
        const s = input.slides[i];
        expect(text).toContain(s.explanation.readerQuestion);
        for (const section of s.explanation.sections) {
          expect(text).toContain(section.heading);
          for (const paragraph of section.body.split("\n")) expect(text).toContain(paragraph);
        }
      }
      const visible = pptx.slideXmls.flatMap(extractTexts).join("\n");
      for (const requirement of ["視覚言語モデル", "忘却", "領域情報", "5,035", "12種類", "平均", "1枚5秒", "ROI", "再現率", "実用化", "執筆者", "https://jpn.nec.com/press/202610/20261014_01.html"]) expect(visible).toContain(requirement);
      // Text preservation is not a semantic comprehension certificate.
      expect(input.slides.every((s: {decomposition: unknown}) => Boolean(s.decomposition))).toBe(true);
    } finally { await rm(outDir, { recursive: true, force: true }); }
  });
});
