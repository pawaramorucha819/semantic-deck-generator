import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { main } from "../src/cli/main.js";
import { buildGenerationContext, loadGenerationConfig } from "../src/config/context.js";
import { GenerationConfigError, parseGenerationConfig } from "../src/config/schema.js";
import { parsePresentation } from "../src/domain/schema.js";
import { generate } from "../src/pipeline.js";
import { parseVisualPlan } from "../src/plan/visual-plan.js";
import { extractTexts, projectRoot, readPptx } from "./helpers.js";

const ENTERPRISE_DIR = join(projectRoot, "examples", "enterprise");
const CONFIG_PATH = join(ENTERPRISE_DIR, "generation.config.json");
const PRESENTATION_PATH = join(ENTERPRISE_DIR, "presentation.json");

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-e2e-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

async function loadContext() {
  const loaded = await loadGenerationConfig(CONFIG_PATH);
  return buildGenerationContext(loaded);
}

async function loadSpec() {
  return parsePresentation(JSON.parse(await readFile(PRESENTATION_PATH, "utf8")));
}

describe("Generation Config", () => {
  it("サンプルの config を検証できる", async () => {
    const loaded = await loadGenerationConfig(CONFIG_PATH);
    expect(loaded.config.template?.id).toBe("acme");
    expect(loaded.config.ai?.mode).toBe("auto");
  });

  it("未知のキーを拒否する", () => {
    expect(() =>
      parseGenerationConfig({ schemaVersion: "1.0", patternId: "three-reasons" }),
    ).toThrow(GenerationConfigError);
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() => parseGenerationConfig({ schemaVersion: "9.9" })).toThrow(/schemaVersion/);
  });

  it("assist / auto で Provider が無い場合は明確なエラーになる（silent fallback しない）", () => {
    expect(() =>
      parseGenerationConfig({ schemaVersion: "1.0", ai: { mode: "assist" } }),
    ).toThrow(/AI Provider の設定が必要です/);
    expect(() =>
      parseGenerationConfig({
        schemaVersion: "1.0",
        ai: { mode: "auto", provider: { type: "none" } },
      }),
    ).toThrow(/AI Provider の設定が必要です/);
  });

  it("ai.mode=off なら Provider 無しで通る", () => {
    expect(() => parseGenerationConfig({ schemaVersion: "1.0", ai: { mode: "off" } })).not.toThrow();
  });

  it("config から Template / Pattern / Knowledge / AI が組み立てられる", async () => {
    const context = await loadContext();
    expect(context.templateId).toBe("acme");
    expect(context.templateRegistry.has("acme")).toBe(true);
    expect(context.patternRegistry.has("acme-three-column")).toBe(true);
    expect(context.patternRegistry.scopeOf("acme-three-column")).toBe("company");
    expect(context.knowledgeStore).toBeDefined();
    expect(context.aiProvider?.id).toContain("scripted");
    expect(context.policy?.mode).toBe("auto");
  });

  it("CLI の --ai-mode off で AI 経路を切れる", async () => {
    const loaded = await loadGenerationConfig(CONFIG_PATH);
    const context = await buildGenerationContext(loaded, { aiMode: "off" });
    expect(context.policy?.mode).toBe("manual");
    expect(context.policy?.advisor).toBeUndefined();
  });

  it("存在しない template.id は分かりやすいエラーになる", async () => {
    const loaded = await loadGenerationConfig(CONFIG_PATH);
    await expect(buildGenerationContext(loaded, { templateId: "nope" })).rejects.toThrow(
      /template.id 'nope' が見つかりません/,
    );
  });
});

describe("Enterprise End-to-End", () => {
  it("presentation + template + declarative pattern + knowledge + mock AI → PPTX", async () => {
    const context = await loadContext();
    const spec = await loadSpec();
    const result = await generate(spec, context.generateOptions);

    // --- Visual Plan --------------------------------------------------------
    expect(result.plan.templateId).toBe("acme");
    expect(result.plan.themeId).toBe("acme");
    expect(result.plan.slides).toHaveLength(6);

    // 企業 Declarative Pattern が実際に使われている
    const patternIds = result.plan.slides.map((s) => s.patternId);
    expect(patternIds).toContain("acme-three-column");

    // --- Decision Trace -----------------------------------------------------
    for (const planSlide of result.plan.slides) {
      const trace = planSlide.decisionTrace!;
      expect(trace.decision.mode).toBe("auto");
      expect(trace.deterministic.candidatePatternIds.length).toBeGreaterThan(0);
      expect(trace.knowledge?.referenceSlides.length).toBeGreaterThan(0);
      expect(trace.ai?.provider).toContain("scripted");
    }

    const accepted = result.plan.slides.filter((s) => s.decisionTrace?.ai?.status === "accepted");
    const rejected = result.plan.slides.filter((s) => s.decisionTrace?.ai?.status === "rejected");
    expect(accepted.length).toBeGreaterThan(0);
    // 意味的に不適合なスライドでは AI 推薦が棄却される（ハードゲートの実演）
    expect(rejected.length).toBeGreaterThan(0);
    for (const slide of rejected) {
      expect(slide.patternId).not.toBe("acme-three-column");
      expect(slide.decisionTrace!.decision.selectedBy).toBe("deterministic");
    }
    for (const slide of accepted) {
      expect(slide.decisionTrace!.decision.selectedBy).toBe("ai");
    }

    // --- Advisor Report -----------------------------------------------------
    expect(result.advisorReport?.summary.slides).toBe(6);
    expect(result.advisorReport?.summary.knowledgeReferenced).toBe(6);
    expect(result.advisorReport?.slides[0]!.references[0]!.slideId).toMatch(/#\d+$/);

    // --- PPTX ---------------------------------------------------------------
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    const pptx = await readPptx(result.pptxBuffer);
    expect(pptx.slideXmls).toHaveLength(7); // 表紙 + 6

    // 企業テンプレートの装飾は Master 側にある
    const layouts = pptx.slideLayoutXmls.join("\n");
    expect(layouts).toContain("acme__acme-content");
    expect(layouts).toContain("ACME 株式会社");
    expect(layouts).toContain("社外秘 / ACME Confidential");
    // ロゴ画像が同梱されている
    expect(pptx.files.filter((f) => f.startsWith("ppt/media/")).length).toBeGreaterThan(0);

    // 本文は Pattern が描いている
    const texts = pptx.slideXmls.flatMap(extractTexts).join("\n");
    expect(texts).toContain("コスト増の要因は3つの手作業に集中している");
    expect(texts).toContain("領収書の手入力");
    expect(texts).toContain("-45%");
  });

  it("生成した Visual Plan は保存・再生できる（AI 抜きで決定論的に再現）", async () => {
    const context = await loadContext();
    const spec = await loadSpec();
    const result = await generate(spec, context.generateOptions);

    const json = JSON.parse(JSON.stringify(result.plan));
    const reparsed = parseVisualPlan(json, { registry: context.patternRegistry });
    expect(reparsed.slides.map((s) => s.patternId)).toEqual(
      result.plan.slides.map((s) => s.patternId),
    );

    // 再生時に AI Provider は不要
    const { renderVisualPlan, toPptxBuffer } = await import("../src/render/renderer.js");
    const replay = renderVisualPlan(reparsed, {
      registry: context.patternRegistry,
      templateRegistry: context.templateRegistry,
    });
    const first = await readPptx(result.pptxBuffer);
    const second = await readPptx(await toPptxBuffer(replay.pptx));
    expect(second.slideXmls).toEqual(first.slideXmls);
  });

  it("同じ入力からは同じ Visual Plan が得られる（決定論的）", async () => {
    const spec = await loadSpec();
    const a = await generate(spec, (await loadContext()).generateOptions);
    const b = await generate(spec, (await loadContext()).generateOptions);
    expect(JSON.stringify(b.plan)).toBe(JSON.stringify(a.plan));
  });

  it("AI を切っても生成できる（AI は optional）", async () => {
    const loaded = await loadGenerationConfig(CONFIG_PATH);
    const context = await buildGenerationContext(loaded, { aiMode: "off" });
    const spec = await loadSpec();
    const result = await generate(spec, context.generateOptions);

    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    for (const planSlide of result.plan.slides) {
      expect(planSlide.decisionTrace?.ai).toBeUndefined();
      expect(planSlide.decisionTrace?.decision.selectedBy).toBe("deterministic");
    }
    // Template と企業 Pattern の登録は AI 無しでも有効
    expect(result.plan.templateId).toBe("acme");
    // ただし決定論的スコアでは builtin の three-reasons が勝つ。
    // 「企業標準を使う」という判断は Knowledge + AI（または人間の指定）が担う。
    expect(result.plan.slides.map((s) => s.patternId)).toContain("three-reasons");
    expect(result.plan.slides.map((s) => s.patternId)).not.toContain("acme-three-column");
  });

  it("CLI から一式を実行できる", async () => {
    const outDir = await makeTempDir();
    const code = await main([
      "generate",
      PRESENTATION_PATH,
      "--config",
      CONFIG_PATH,
      "--out-dir",
      outDir,
      "--advisor-report",
      "--quiet",
    ]);
    expect(code).toBe(0);

    const plan = JSON.parse(await readFile(join(outDir, "presentation.visual-plan.json"), "utf8"));
    expect(plan.templateId).toBe("acme");
    const advisor = JSON.parse(
      await readFile(join(outDir, "presentation.advisor-report.json"), "utf8"),
    );
    expect(advisor.mode).toBe("auto");
    expect(advisor.slides).toHaveLength(6);

    const pptx = await readFile(join(outDir, "presentation.pptx"));
    expect(pptx.subarray(0, 2).toString("latin1")).toBe("PK");
  });

  it("CLI は未知のオプションを拒否する", async () => {
    await expect(main(["generate", PRESENTATION_PATH, "--nonexistent"])).rejects.toThrow(
      /未知のオプション/,
    );
  });

  it("CLI の knowledge search が動く", async () => {
    const logs: string[] = [];
    const original = console.log;
    console.log = (...args: unknown[]) => logs.push(args.join(" "));
    try {
      const code = await main([
        "knowledge",
        "search",
        "--store",
        join(ENTERPRISE_DIR, "knowledge"),
        "--role",
        "analysis",
        "--json",
      ]);
      expect(code).toBe(0);
    } finally {
      console.log = original;
    }
    const results = JSON.parse(logs.join("\n"));
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].entry.semantic.slideRole).toBe("analysis");
  });

  it("V1 のサンプルは V2 でも同じ Pattern 選択になる（後方互換）", async () => {
    const spec = parsePresentation(
      JSON.parse(await readFile(join(projectRoot, "examples", "inventory-system.json"), "utf8")),
    );
    const result = await generate(spec);
    expect(result.plan.slides.map((s) => s.patternId)).toEqual([
      "chart-with-message",
      "three-reasons",
      "process",
      "before-after",
      "big-number",
      "message-only",
    ]);
    expect(result.plan.templateId).toBeUndefined();
    expect(result.plan.slides[0]!.decisionTrace).toBeUndefined();
  });
});
