import { readFile, readdir } from "node:fs/promises";
import { join, relative, sep } from "node:path";
import { describe, expect, it } from "vitest";
import { projectRoot } from "./helpers.js";

/**
 * アーキテクチャ境界のテスト。
 *
 * CLAUDE.md に書いた不変条件のうち「import してはいけない」系を機械的に固定する。
 * ドキュメントだけだと必ず崩れるので、ここで壊れたら気付けるようにしておく。
 */

const SRC = join(projectRoot, "src");

async function listTsFiles(dir: string): Promise<string[]> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...(await listTsFiles(path)));
    else if (entry.name.endsWith(".ts")) files.push(path);
  }
  return files.sort();
}

async function filesUnder(...segments: string[]): Promise<string[]> {
  return listTsFiles(join(SRC, ...segments));
}

async function importsOf(file: string): Promise<string[]> {
  const source = await readFile(file, "utf8");
  return [...source.matchAll(/(?:^|\n)\s*(?:import|export)[^;]*?from\s+"([^"]+)"/g)].map(
    (match) => match[1]!,
  );
}

const rel = (file: string): string => relative(projectRoot, file).split(sep).join("/");

describe("アーキテクチャ境界", () => {
  it("Renderer は AI を知らない（Visual Plan 確定後は決定論的）", async () => {
    for (const file of await filesUnder("render")) {
      for (const specifier of await importsOf(file)) {
        expect(specifier, `${rel(file)} が AI を import しています`).not.toMatch(/\/ai\//);
      }
    }
  });

  it("domain / patterns / templates / knowledge / theme も AI を import しない", async () => {
    const dirs = ["domain", "patterns", "templates", "knowledge", "theme"];
    for (const dir of dirs) {
      for (const file of await filesUnder(dir)) {
        for (const specifier of await importsOf(file)) {
          expect(specifier, `${rel(file)} が AI を import しています`).not.toMatch(/\/ai\//);
        }
      }
    }
  });

  it("AI は Selection Policy / Config / 公開 API からのみ使われる", async () => {
    const allowed = new Set([
      "src/selector/selection-policy.ts",
      "src/config/context.ts",
      "src/index.ts",
      "src/api.ts",
    ]);
    const files = await listTsFiles(SRC);
    for (const file of files) {
      const path = rel(file);
      if (path.startsWith("src/ai/") || allowed.has(path)) continue;
      for (const specifier of await importsOf(file)) {
        expect(specifier, `${path} が AI を import しています`).not.toMatch(/\/ai\//);
      }
    }
  });

  it("Pattern は PptxGenJS を直接 import しない（RenderContext 経由）", async () => {
    for (const file of await filesUnder("patterns")) {
      const source = await readFile(file, "utf8");
      expect(source, `${rel(file)} が PptxGenJS を直接使っています`).not.toContain(
        'from "pptxgenjs"',
      );
    }
  });

  it("PptxGenJS を直接触るのは描画バックエンドと Master 生成だけ", async () => {
    const allowed = new Set(["src/render/surface.ts", "src/render/renderer.ts", "src/templates/master.ts"]);
    for (const file of await listTsFiles(SRC)) {
      const source = await readFile(file, "utf8");
      if (!source.includes('from "pptxgenjs"')) continue;
      expect(allowed.has(rel(file)), `${rel(file)} が PptxGenJS を import しています`).toBe(true);
    }
  });

  it("domain は他のレイヤに依存しない（意味の純粋性）", async () => {
    for (const file of await filesUnder("domain")) {
      for (const specifier of await importsOf(file)) {
        const isLocal = specifier.startsWith("./");
        const isExternal = !specifier.startsWith(".");
        expect(
          isLocal || isExternal,
          `${rel(file)} が domain の外を import しています: ${specifier}`,
        ).toBe(true);
      }
    }
  });

  it("Renderer は Selector / Selection Policy を import しない", async () => {
    for (const file of await filesUnder("render")) {
      for (const specifier of await importsOf(file)) {
        expect(specifier, `${rel(file)} が Selector を import しています`).not.toMatch(
          /\/selector\//,
        );
      }
    }
  });

  it("Knowledge は Pattern 実装に依存しない（差し替え可能に保つ）", async () => {
    for (const file of await filesUnder("knowledge")) {
      for (const specifier of await importsOf(file)) {
        expect(specifier, `${rel(file)} が Pattern を import しています`).not.toMatch(
          /\/patterns\//,
        );
      }
    }
  });

  it("Declarative Pattern のスキーマに eval 系の語彙が無い", async () => {
    const source = await readFile(join(SRC, "patterns", "declarative", "schema.ts"), "utf8");
    for (const word of ["eval(", "new Function", '"expr"', '"script"']) {
      expect(source).not.toContain(word);
    }
  });

  it("Declarative Pattern の描画は RenderContext 経由（PptxGenJS を含まない）", async () => {
    const source = await readFile(join(SRC, "patterns", "declarative", "renderer.ts"), "utf8");
    expect(source).not.toContain("pptxgenjs");
    expect(source).toContain("RenderContext");
  });
});
