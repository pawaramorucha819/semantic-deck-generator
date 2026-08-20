import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { generate } from "../src/pipeline.js";
import { buildVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { corporateTemplate } from "../src/templates/builtin/corporate.js";
import { defaultTemplate } from "../src/templates/builtin/default.js";
import {
  loadTemplateFromDirectory,
  resolveInsideDirectory,
  TemplateLoadError,
} from "../src/templates/loader.js";
import { TemplateRegistry, defaultTemplateRegistry } from "../src/templates/registry.js";
import { parseTemplate, TemplateValidationError } from "../src/templates/schema.js";
import {
  contentFrameOf,
  expandTemplateText,
  masterForRole,
  TEMPLATE_SCHEMA_VERSION,
} from "../src/templates/template.js";
import { midnightTheme, resolveTheme } from "../src/theme/theme.js";
import { extractTexts, loadExample, readPptx, tinyPngBytes } from "./helpers.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-template-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

function minimalTemplate(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    schemaVersion: TEMPLATE_SCHEMA_VERSION,
    id: "acme",
    version: "1.0.0",
    name: "ACME",
    theme: {
      id: "acme",
      name: "ACME Theme",
      fonts: { heading: "Yu Gothic", body: "Yu Gothic" },
      fontSizes: { keyMessage: 22, heading: 15, body: 13, caption: 10 },
      spacing: { xs: 0.06, sm: 0.12, md: 0.22, lg: 0.36, xl: 0.56 },
      colors: {
        background: "FFFFFF",
        text: "111111",
        mutedText: "666666",
        primary: "003366",
        secondary: "336699",
        accent: "CC6600",
      },
    },
    masters: [{ id: "main", role: "content", elements: [] }],
    defaultMasterId: "main",
    ...overrides,
  };
}

describe("Template スキーマ検証", () => {
  it("最小構成の Template を受け入れる", () => {
    const template = parseTemplate(minimalTemplate());
    expect(template.id).toBe("acme");
    expect(template.masters).toHaveLength(1);
  });

  it("未知のカラートークンを拒否する", () => {
    const bad = minimalTemplate({
      masters: [
        {
          id: "main",
          role: "content",
          elements: [
            { kind: "shape", shape: "rect", box: { x: 0, y: 0, w: 1, h: 1 }, style: { fill: "brandBlue" } },
          ],
        },
      ],
    });
    expect(() => parseTemplate(bad)).toThrow(TemplateValidationError);
  });

  it("存在しない master を defaultMasterId が指している場合を拒否する", () => {
    expect(() => parseTemplate(minimalTemplate({ defaultMasterId: "nope" }))).toThrow(
      /defaultMasterId/,
    );
  });

  it("未登録のアセットを参照する image 要素を拒否する", () => {
    const bad = minimalTemplate({
      masters: [
        {
          id: "main",
          role: "content",
          elements: [{ kind: "image", assetId: "logo", box: { x: 0, y: 0, w: 1, h: 1 } }],
        },
      ],
    });
    expect(() => parseTemplate(bad)).toThrow(/未登録のアセット/);
  });

  it("'..' を含むアセットパスを拒否する（パストラバーサル対策）", () => {
    const bad = minimalTemplate({
      assets: [{ id: "logo", kind: "image", path: "../../secret.png" }],
    });
    expect(() => parseTemplate(bad)).toThrow(TemplateValidationError);
  });

  it("絶対パスのアセットを拒否する", () => {
    const bad = minimalTemplate({
      assets: [{ id: "logo", kind: "image", path: "C:/windows/logo.png" }],
    });
    expect(() => parseTemplate(bad)).toThrow(TemplateValidationError);
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() => parseTemplate(minimalTemplate({ schemaVersion: "9.9" }))).toThrow(
      /schemaVersion/,
    );
  });

  it("未知のキー（描画以外の混入）を拒否する", () => {
    expect(() => parseTemplate(minimalTemplate({ patternId: "three-reasons" }))).toThrow(
      TemplateValidationError,
    );
  });

  it("resolveInsideDirectory はディレクトリ外を拒否する", () => {
    expect(() => resolveInsideDirectory("/base/dir", "../escape.png")).toThrow(TemplateLoadError);
    expect(resolveInsideDirectory("/base/dir", "assets/logo.png")).toContain("logo.png");
  });
});

describe("Template Registry", () => {
  it("ビルトイン Template を引ける", () => {
    expect(defaultTemplateRegistry.ids()).toEqual(["default", "corporate"]);
    expect(defaultTemplateRegistry.template("corporate").name).toBe("Corporate Standard");
  });

  it("未登録 ID は分かりやすいエラーになる", () => {
    expect(() => defaultTemplateRegistry.get("nope")).toThrow(/未登録の templateId/);
  });

  it("外部ディレクトリの Template を登録できる（アセット込み）", async () => {
    const root = await makeTempDir();
    const dir = join(root, "acme");
    await mkdir(join(dir, "assets"), { recursive: true });
    await writeFile(join(dir, "assets", "logo.png"), tinyPngBytes());
    await writeFile(
      join(dir, "template.json"),
      JSON.stringify(
        minimalTemplate({
          assets: [{ id: "logo", kind: "image", path: "assets/logo.png" }],
          masters: [
            {
              id: "main",
              role: "content",
              elements: [{ kind: "image", assetId: "logo", box: { x: 8.8, y: 0.2, w: 0.8, h: 0.3 } }],
            },
          ],
        }),
        null,
        2,
      ),
      "utf8",
    );

    const registry = new TemplateRegistry();
    const loaded = await registry.registerDirectory(root);
    expect(loaded).toHaveLength(1);
    expect(registry.get("acme").source).toBe("external");
    expect(registry.get("acme").assets[0]!.data.startsWith("image/png;base64,")).toBe(true);
  });

  it("theme.json を別ファイルに分けた Template を読める", async () => {
    const dir = await makeTempDir();
    const base = minimalTemplate();
    const { theme, ...withoutTheme } = base as Record<string, unknown> & { theme: unknown };
    await writeFile(join(dir, "template.json"), JSON.stringify(withoutTheme), "utf8");
    await writeFile(join(dir, "theme.json"), JSON.stringify(theme), "utf8");

    const loaded = await loadTemplateFromDirectory(dir);
    expect(loaded.template.theme.id).toBe("acme");
  });

  it("存在しないアセットファイルはロード時に失敗する（silent failure にしない）", async () => {
    const dir = await makeTempDir();
    await writeFile(
      join(dir, "template.json"),
      JSON.stringify(
        minimalTemplate({
          assets: [{ id: "logo", kind: "image", path: "assets/missing.png" }],
        }),
      ),
      "utf8",
    );
    await expect(loadTemplateFromDirectory(dir)).rejects.toThrow(/アセットが見つかりません/);
  });

  it("template.json が無いディレクトリは分かりやすいエラーになる", async () => {
    const dir = await makeTempDir();
    await expect(loadTemplateFromDirectory(dir)).rejects.toThrow(/template.json が見つかりません/);
  });
});

describe("Master 選択と本文領域", () => {
  it("役割ごとに Master が選ばれる", () => {
    expect(masterForRole(corporateTemplate, "title").id).toBe("corporate-title");
    expect(masterForRole(corporateTemplate, "section").id).toBe("corporate-section");
    expect(masterForRole(corporateTemplate, "content").id).toBe("corporate-content");
  });

  it("役割の Master が無ければ default master へフォールバックする", () => {
    const single = parseTemplate(minimalTemplate());
    expect(masterForRole(single, "title").id).toBe("main");
    expect(masterForRole(single, "section").id).toBe("main");
  });

  it("contentMargins から本文領域が決まる（Theme の margins より優先）", () => {
    const theme = resolveTheme(corporateTemplate.theme);
    const master = masterForRole(corporateTemplate, "content");
    const frame = contentFrameOf(master, theme, { widthIn: 10, heightIn: 5.625 });
    expect(frame.y).toBeCloseTo(0.62, 5);
    expect(frame.h).toBeCloseTo(5.625 - 0.62 - 0.66, 5);
  });

  it("Master のテキスト変数は列挙された語だけを展開する", () => {
    expect(
      expandTemplateText("{{organization}} / {{confidentiality}}", {
        deckTitle: "t",
        organization: "ACME",
        confidentiality: "社外秘",
      }),
    ).toBe("ACME / 社外秘");
    expect(() => expandTemplateText("{{apiKey}}", { deckTitle: "t" })).toThrow(/未知の変数/);
  });
});

describe("Template 適用後の PPTX 構造", () => {
  it("Master が Slide Layout として書き出され、ロゴ・フッターがスライド本体に混ざらない", async () => {
    const spec = await loadExample();
    const result = await generate(spec, { templateId: "corporate" });
    const pptx = await readPptx(result.pptxBuffer);

    const layouts = pptx.slideLayoutXmls.join("\n");
    expect(layouts).toContain("corporate__corporate-title");
    expect(layouts).toContain("corporate__corporate-content");
    // 組織名・機密表示は Master 側にあり、
    expect(layouts).toContain("Example Corporation");
    expect(layouts).toContain("社外秘 / Confidential");
    // スライド本体には現れない（= Pattern が会社固有の要素を描いていない）
    for (const xml of pptx.slideXmls) {
      expect(extractTexts(xml).join("")).not.toContain("社外秘 / Confidential");
    }
  });

  it("ページ番号は PowerPoint のページ番号フィールドとして Master に入る", async () => {
    const spec = await loadExample();
    const result = await generate(spec, { templateId: "corporate" });
    const pptx = await readPptx(result.pptxBuffer);
    expect(pptx.slideMasterXmls.join("")).toContain('type="slidenum"');
    // 本文スライドは sldNum プレースホルダを持ち、表紙は持たない
    expect(pptx.slideXmls[0]).not.toContain('type="sldNum"');
    expect(pptx.slideXmls[1]).toContain('type="sldNum"');
  });

  it("表紙・本文で異なる Master（Layout）が参照される", async () => {
    const spec = await loadExample();
    const result = await generate(spec, { templateId: "corporate" });
    const pptx = await readPptx(result.pptxBuffer);
    const [titleRef, ...bodyRefs] = pptx.slideLayoutRefs;
    expect(titleRef).toBeDefined();
    expect(new Set(bodyRefs).size).toBe(1);
    expect(bodyRefs[0]).not.toBe(titleRef);
  });

  it("section に指定したスライドだけ section master を使う", async () => {
    const spec = await loadExample();
    const result = await generate(spec, {
      templateId: "corporate",
      sectionSlideIds: ["slide-004"],
    });
    expect(result.plan.slides.find((s) => s.slideId === "slide-004")!.masterRole).toBe("section");

    const pptx = await readPptx(result.pptxBuffer);
    // 表紙 + 6 枚。slide-004 は 4 枚目なので index 4（表紙ぶんずれる）
    const refs = pptx.slideLayoutRefs;
    expect(refs[4]).not.toBe(refs[1]);
  });

  it("Template の Theme が使われ、--theme 相当の上書きもできる", async () => {
    const spec = await loadExample();
    const branded = await generate(spec, { templateId: "corporate" });
    expect(branded.plan.themeId).toBe("corporate");
    expect((await readPptx(branded.pptxBuffer)).slideXmls[1]).toContain("10365F");

    const overridden = await generate(spec, { templateId: "corporate", themeId: "midnight" });
    expect(overridden.plan.themeId).toBe("midnight");
    const xml = (await readPptx(overridden.pptxBuffer)).slideXmls[1]!;
    expect(xml).toContain("5AA9F0"); // midnight primary
    expect(xml).not.toContain("10365F");
  });

  it("Template 使用時もレイアウト検査はエラーゼロ", async () => {
    const spec = await loadExample();
    const result = await generate(spec, { templateId: "corporate" });
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
  });

  it("plan.templateId から Template が復元される（Plan だけで再生できる）", async () => {
    const spec = await loadExample();
    const plan = buildVisualPlan(spec, { themeId: "corporate", templateId: "corporate" });
    const { pptx, templateId } = renderVisualPlan(plan);
    expect(templateId).toBe("corporate");
    const contents = await readPptx(await toPptxBuffer(pptx));
    expect(contents.slideLayoutXmls.join("")).toContain("corporate__corporate-content");
  });

  it("未登録の templateId を持つ Plan は黙って Template 無しにならず例外になる", async () => {
    const spec = await loadExample();
    const plan = buildVisualPlan(spec, { templateId: "unknown-template" });
    expect(() => renderVisualPlan(plan)).toThrow(/未登録の templateId/);
  });

  it("外部 Template のロゴがマスターへ埋め込まれる", async () => {
    const dir = await makeTempDir();
    await mkdir(join(dir, "assets"), { recursive: true });
    await writeFile(join(dir, "assets", "logo.png"), tinyPngBytes());
    await writeFile(
      join(dir, "template.json"),
      JSON.stringify(
        minimalTemplate({
          assets: [{ id: "logo", kind: "image", path: "assets/logo.png" }],
          masters: [
            {
              id: "main",
              role: "content",
              contentMargins: { top: 0.6, right: 0.6, bottom: 0.5, left: 0.6 },
              elements: [{ kind: "image", assetId: "logo", box: { x: 8.9, y: 0.16, w: 0.5, h: 0.3 } }],
            },
          ],
        }),
      ),
      "utf8",
    );

    const loaded = await loadTemplateFromDirectory(dir);
    const spec = await loadExample();
    const result = await generate(spec, {
      template: loaded.template,
      templateAssets: loaded.assets,
    });
    const pptx = await readPptx(result.pptxBuffer);
    // 画像は media として同梱され、slideLayout から参照される
    expect(pptx.files.filter((f) => f.startsWith("ppt/media/")).length).toBeGreaterThan(0);
  });
});

describe("Template 非使用時の後方互換性", () => {
  it("Template を指定しなければ Slide Master は追加されず、V1 と同じページ番号表記になる", async () => {
    const spec = await loadExample();
    const result = await generate(spec);
    const pptx = await readPptx(result.pptxBuffer);

    expect(result.plan.templateId).toBeUndefined();
    // pptxgenjs の既定 layout 1 枚だけ（Template の master は増えていない）
    expect(pptx.slideLayoutXmls).toHaveLength(1);
    expect(extractTexts(pptx.slideXmls[6]!).join("")).toContain("6 / 6");
  });

  it("Template 非使用時のスライド XML は Template 導入前と同じ（テーマ経路も不変）", async () => {
    const spec = await loadExample();
    const a = await generate(spec);
    const b = await generate(spec, { themeId: "default" });
    expect((await readPptx(b.pptxBuffer)).slideXmls).toEqual(
      (await readPptx(a.pptxBuffer)).slideXmls,
    );
  });

  it("ビルトインの default Template は V1 と同じ余白を持つ", () => {
    const theme = resolveTheme(defaultTemplate.theme);
    const frame = contentFrameOf(masterForRole(defaultTemplate, "content"), theme, {
      widthIn: 10,
      heightIn: 5.625,
    });
    expect(frame).toEqual({ x: 0.6, y: 0.42, w: 8.8, h: 4.805 });
  });

  it("midnight テーマ + Template 無しの既存挙動が維持される", async () => {
    const spec = await loadExample();
    const result = await generate(spec, { themeId: "midnight" });
    expect(result.plan.themeId).toBe("midnight");
    expect((await readPptx(result.pptxBuffer)).slideXmls[1]).toContain(
      midnightTheme.colors.primary,
    );
  });
});
