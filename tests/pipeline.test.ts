import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { PresentationValidationError } from "../src/domain/schema.js";
import { generate, generateFromFile } from "../src/pipeline.js";
import { examplePath, extractTexts, loadExample, readPptx } from "./helpers.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-test-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

describe("End-to-End パイプライン", () => {
  it("presentation.json から .pptx と .visual-plan.json を出力する", async () => {
    const outDir = await makeTempDir();
    const result = await generateFromFile(examplePath, { outDir });

    expect(result.pptxPath).toBe(join(outDir, "inventory-system.pptx"));
    expect(result.planPath).toBe(join(outDir, "inventory-system.visual-plan.json"));

    const pptxBytes = await readFile(result.pptxPath);
    expect(pptxBytes.length).toBeGreaterThan(10_000);
    expect(pptxBytes.subarray(0, 2).toString("latin1")).toBe("PK"); // zip

    const plan = JSON.parse(await readFile(result.planPath, "utf8"));
    expect(plan.slides).toHaveLength(6);
    expect(plan.themeId).toBe("default");
    expect(result.issues).toEqual([]);
  });

  it("出力した .pptx を再読込して内部構造を検証できる", async () => {
    const outDir = await makeTempDir();
    const result = await generateFromFile(examplePath, { outDir });
    const { slideXmls } = await readPptx(await readFile(result.pptxPath));

    expect(slideXmls).toHaveLength(7);
    const titleTexts = extractTexts(slideXmls[0]!);
    expect(titleTexts).toContain("在庫管理システム導入提案");
    expect(titleTexts).toContain("経営企画部");

    const lastTexts = extractTexts(slideXmls[6]!).join("");
    expect(lastTexts).toContain("したがって在庫管理システムを導入するべきである");
    expect(lastTexts).toContain("6 / 6");
  });

  it("テーマを指定して生成できる", async () => {
    const outDir = await makeTempDir();
    const result = await generateFromFile(examplePath, { outDir, themeId: "midnight" });
    expect(result.plan.themeId).toBe("midnight");
    const { slideXmls } = await readPptx(result.pptxBuffer);
    expect(slideXmls[1]).toContain("10151F"); // midnight background
  });

  it("未知のテーマ ID を拒否する", async () => {
    await expect(generateFromFile(examplePath, { themeId: "neon" })).rejects.toThrow(
      /未知のテーマ/,
    );
  });

  it("不正な presentation.json は生成前に拒否される", async () => {
    const dir = await makeTempDir();
    const badPath = join(dir, "bad.json");
    await writeFile(
      badPath,
      JSON.stringify({ title: "x", slides: [{ id: "s", keyMessage: "m" }] }),
      "utf8",
    );
    await expect(generateFromFile(badPath, { outDir: dir })).rejects.toThrow(
      PresentationValidationError,
    );
  });

  it("JSON として壊れている場合は分かりやすいエラーになる", async () => {
    const dir = await makeTempDir();
    const badPath = join(dir, "broken.json");
    await writeFile(badPath, "{ not json", "utf8");
    await expect(generateFromFile(badPath, { outDir: dir })).rejects.toThrow(
      /JSON として読み込めませんでした/,
    );
  });

  it("メモリ上でも生成できる（ファイル I/O 不要）", async () => {
    const result = await generate(await loadExample());
    expect(result.pptxBuffer.length).toBeGreaterThan(10_000);
    expect(result.plan.slides).toHaveLength(6);
    expect(result.reports).toHaveLength(7);
  });
});
