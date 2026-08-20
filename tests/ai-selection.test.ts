import { describe, expect, it } from "vitest";
import { z } from "zod";
import {
  AnthropicAIProvider,
  createAnthropicProvider,
  DEFAULT_ANTHROPIC_MODEL,
} from "../src/ai/anthropic.js";
import {
  PatternAdvisor,
  PatternRecommendationSchema,
  toAdvisableCandidates,
} from "../src/ai/pattern-advisor.js";
import {
  AIResponseValidationError,
  AIUnavailableError,
  extractJsonObject,
  nullAIProvider,
  ScriptedAIProvider,
} from "../src/ai/provider.js";
import { StructureAdvisor } from "../src/ai/structure-advisor.js";
import { PresentationValidationError } from "../src/domain/schema.js";
import { InMemoryKnowledgeStore } from "../src/knowledge/store/store.js";
import { DeterministicRetriever } from "../src/knowledge/retrieval/retrieval.js";
import { KNOWLEDGE_SCHEMA_VERSION, type DeckKnowledgeEntry } from "../src/knowledge/schema.js";
import { beforeAfterPattern } from "../src/patterns/before-after/index.js";
import { compileDeclarativePattern } from "../src/patterns/declarative/compile.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import type { AnyVisualPattern } from "../src/patterns/pattern.js";
import { PatternRegistry, defaultRegistry } from "../src/patterns/registry.js";
import { buildAdvisorReport } from "../src/plan/advisor-report.js";
import { buildVisualPlanWithPolicy } from "../src/plan/visual-plan.js";
import { generate } from "../src/pipeline.js";
import {
  decideSlidePattern,
  SelectionPolicyError,
} from "../src/selector/selection-policy.js";
import {
  checkSelectable,
  evaluatePattern,
  meetsSelectionGate,
  selectPattern,
} from "../src/selector/selector.js";
import { loadExample, metricEvidence, pointsContent, slide, textEvidence } from "./helpers.js";

function threeItemSlide() {
  return slide({
    id: "slide-ai",
    keyMessage: "3 つの構造的な原因がある",
    slideRole: "analysis",
    logicalRelation: "claim_evidence",
    content: pointsContent(
      { label: "定型指摘に人手", description: "規約の指摘が多い", evidenceRefs: ["e1"] },
      { label: "レビュー集中", description: "上位層に偏っている", evidenceRefs: ["e2"] },
      { label: "待ち時間が長い", description: "初回コメントまで", evidenceRefs: ["e3"] },
    ),
    evidence: [
      metricEvidence("e1", "定型指摘", 35, "%"),
      metricEvidence("e2", "レビュー集中", 58, "%"),
      metricEvidence("e3", "待ち時間", 18, "時間"),
    ],
  });
}

function knowledgeEntry(overrides: Partial<DeckKnowledgeEntry> = {}): DeckKnowledgeEntry {
  return {
    schemaVersion: KNOWLEDGE_SCHEMA_VERSION,
    deckId: "fy2025-strategy",
    slideId: "fy2025-strategy#18",
    metadata: { title: "FY2025 戦略", purpose: "投資判断", audience: "経営会議" },
    semantic: {
      keyMessage: "3 つの原因が構造的に存在する",
      slideRole: "analysis",
      logicalRelation: "claim_evidence",
      keywords: ["原因", "構造"],
    },
    visual: { patternId: "three-reasons" },
    content: { text: "3 つの構造的な原因を並べて示したスライド" },
    source: { file: "fy2025-strategy.pptx", slideNumber: 18 },
    ...overrides,
  };
}

/** 決定論的候補のうち、指定 index の Pattern を推薦する Provider を作る */
function recommender(patternId: string, extra: Record<string, unknown> = {}) {
  return new ScriptedAIProvider({
    responses: {
      "pattern-selection": (input) => ({
        recommendedPatternId: patternId,
        confidence: 0.87,
        reason: "過去の類似スライドで同じ Pattern が使われているため",
        referenceSlides: (input as { references: Array<{ slideId: string }> }).references.map(
          (r) => r.slideId,
        ),
        ...extra,
      }),
    },
  });
}

describe("AI Provider 抽象", () => {
  it("Provider 未設定なら明示的に失敗する（黙って AI を使ったふりをしない）", () => {
    expect(nullAIProvider.available).toBe(false);
    expect(() => nullAIProvider.generateStructured({} as never)).toThrow(AIUnavailableError);
  });

  it("スクリプト Provider の出力もスキーマ検証を通る", async () => {
    const provider = new ScriptedAIProvider({
      responses: { test: { value: 42 } },
    });
    const result = await provider.generateStructured({
      purpose: "test",
      instructions: "",
      input: {},
      schema: z.object({ value: z.number() }).strict(),
    });
    expect(result).toEqual({ value: 42 });
  });

  it("スキーマに合わない AI 出力は拒否される", async () => {
    const provider = new ScriptedAIProvider({ responses: { test: { value: "not a number" } } });
    await expect(
      provider.generateStructured({
        purpose: "test",
        instructions: "",
        input: {},
        schema: z.object({ value: z.number() }).strict(),
      }),
    ).rejects.toThrow(AIResponseValidationError);
  });

  it("コードフェンス付きの応答から JSON を取り出せる", () => {
    expect(extractJsonObject('説明\n```json\n{"a":1}\n```')).toEqual({ a: 1 });
    expect(() => extractJsonObject("JSON ではない")).toThrow(AIResponseValidationError);
  });

  it("Anthropic Adapter は API キーが無ければ available=false", () => {
    const provider = new AnthropicAIProvider({ apiKey: "" });
    const previous = process.env.ANTHROPIC_API_KEY;
    delete process.env.ANTHROPIC_API_KEY;
    try {
      expect(provider.available).toBe(false);
    } finally {
      if (previous !== undefined) process.env.ANTHROPIC_API_KEY = previous;
    }
  });

  it("Anthropic Adapter は注入したクライアント経由で構造化出力を検証する（外部 API は呼ばない）", async () => {
    const requests: Array<Record<string, unknown>> = [];
    const provider = createAnthropicProvider({
      client: {
        messages: {
          // eslint-disable-next-line @typescript-eslint/require-await
          async create(params: Record<string, unknown>) {
            requests.push(params);
            return { content: [{ type: "text", text: '{"value": 7}' }], stop_reason: "end_turn" };
          },
        },
      },
    });

    const result = await provider.generateStructured({
      purpose: "test",
      instructions: "JSON を返してください",
      input: { a: 1 },
      schema: z.object({ value: z.number() }).strict(),
      jsonSchema: { type: "object" },
    });

    expect(result).toEqual({ value: 7 });
    expect(requests[0]!.model).toBe(DEFAULT_ANTHROPIC_MODEL);
    expect((requests[0]!.output_config as Record<string, unknown>).format).toEqual({
      type: "json_schema",
      schema: { type: "object" },
    });
  });

  it("Anthropic Adapter の応答もスキーマ検証を通らなければ拒否される", async () => {
    const provider = createAnthropicProvider({
      client: {
        // eslint-disable-next-line @typescript-eslint/require-await
        messages: { async create() { return { content: [{ type: "text", text: '{"value":"x"}' }] }; } },
      },
    });
    await expect(
      provider.generateStructured({
        purpose: "test",
        instructions: "",
        input: {},
        schema: z.object({ value: z.number() }).strict(),
      }),
    ).rejects.toThrow(AIResponseValidationError);
  });
});

describe("Pattern Advisor（候補集合の中からしか選べない）", () => {
  it("推薦は Zod で検証される", () => {
    expect(() =>
      PatternRecommendationSchema.parse({
        recommendedPatternId: "three-reasons",
        confidence: 1.4,
        reason: "範囲外の confidence",
      }),
    ).toThrow();
  });

  it("AI には決定論的に適合した候補だけが渡される", async () => {
    const provider = recommender("three-reasons");
    const advisor = new PatternAdvisor(provider);
    const target = threeItemSlide();
    const evaluations = selectPattern(target).candidates;

    await advisor.advise({
      slide: target,
      candidates: toAdvisableCandidates(evaluations, (id) => {
        const pattern = defaultRegistry.get(id);
        return { description: pattern.description, variants: pattern.variants };
      }),
      references: [],
    });

    const sent = provider.calls[0]!.input as { candidates: Array<{ patternId: string }> };
    const incompatible = evaluations.filter((e) => !e.compatible).map((e) => e.patternId);
    expect(sent.candidates.length).toBeGreaterThan(0);
    for (const id of incompatible) {
      expect(sent.candidates.map((c) => c.patternId)).not.toContain(id);
    }
  });

  it("候補集合に無い Pattern を推薦したら棄却する", async () => {
    const advisor = new PatternAdvisor(recommender("nonexistent-pattern"));
    const result = await advisor.advise({
      slide: threeItemSlide(),
      candidates: [
        {
          patternId: "three-reasons",
          variantId: "standard",
          variants: ["standard"],
          description: "d",
          deterministicScore: 0.9,
          reason: "r",
        },
      ],
      references: [],
    });
    expect(result.status).toBe("rejected");
    expect(result.rejectionReason).toContain("候補集合に無い patternId");
  });

  it("未実装の variant を推薦したら棄却する", async () => {
    const advisor = new PatternAdvisor(
      recommender("three-reasons", { recommendedVariantId: "compact" }),
    );
    const result = await advisor.advise({
      slide: threeItemSlide(),
      candidates: [
        {
          patternId: "three-reasons",
          variantId: "standard",
          variants: ["standard"],
          description: "d",
          deterministicScore: 0.9,
          reason: "r",
        },
      ],
      references: [],
    });
    expect(result.status).toBe("rejected");
    expect(result.rejectionReason).toContain("未実装の variant");
  });

  it("Provider が使えないときは unavailable を返す（例外にしない）", async () => {
    const advisor = new PatternAdvisor(nullAIProvider);
    const result = await advisor.advise({
      slide: threeItemSlide(),
      candidates: [],
      references: [],
    });
    expect(result.status).toBe("unavailable");
  });

  it("Provider がエラーになっても生成は止まらない（error として記録される）", async () => {
    const failing = {
      id: "failing",
      available: true,
      generateStructured: () => Promise.reject(new Error("network down")),
    };
    const advisor = new PatternAdvisor(failing);
    const result = await advisor.advise({
      slide: threeItemSlide(),
      candidates: [
        {
          patternId: "three-reasons",
          variantId: "standard",
          variants: ["standard"],
          description: "d",
          deterministicScore: 0.9,
          reason: "r",
        },
      ],
      references: [],
    });
    expect(result.status).toBe("error");
    expect(result.rejectionReason).toContain("network down");
  });
});

describe("Selection Policy: manual / assist / auto", () => {
  const store = new InMemoryKnowledgeStore([knowledgeEntry()]);
  const retriever = new DeterministicRetriever(() => store.all());

  it("manual: AI を呼ばず、決定論的勝者を使う", async () => {
    const provider = recommender("message-only");
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "manual",
      advisor: new PatternAdvisor(provider),
      retriever,
    });

    expect(provider.calls).toHaveLength(0);
    expect(outcome.decisionTrace.decision.selectedBy).toBe("deterministic");
    expect(outcome.patternId).toBe("three-reasons");
    // Knowledge は参考情報として集まる
    expect(outcome.references.length).toBeGreaterThan(0);
    expect(outcome.decisionTrace.knowledge?.referenceSlides).toContain("fy2025-strategy#18");
  });

  it("assist: AI 推薦を記録するが採用しない（非対話 CLI の安全な挙動）", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "assist",
      advisor: new PatternAdvisor(recommender("message-only")),
      retriever,
    });

    expect(outcome.decisionTrace.ai?.status).toBe("advisory");
    expect(outcome.decisionTrace.ai?.recommendedPatternId).toBe("message-only");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("deterministic");
    expect(outcome.patternId).toBe("three-reasons");
  });

  it("auto: 検証を通った AI 推薦を採用する", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      advisor: new PatternAdvisor(recommender("message-only")),
      retriever,
    });

    expect(outcome.decisionTrace.ai?.status).toBe("accepted");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("ai");
    expect(outcome.patternId).toBe("message-only");
    // 決定論的な推薦も記録として残る（何を上書きしたか分かる）
    expect(outcome.decisionTrace.deterministic.recommendedPatternId).toBe("three-reasons");
  });

  it("auto でも constraints を突破できない（AI が不適合な Pattern を推薦しても棄却）", async () => {
    // logicalRelation=comparison のスライドで process を推薦させる
    const target = slide({
      id: "slide-incompatible",
      logicalRelation: "comparison",
      evidence: [textEvidence("e1", "A"), textEvidence("e2", "B")],
    });
    const outcome = await decideSlidePattern(target, {
      mode: "auto",
      advisor: new PatternAdvisor(recommender("process")),
    });

    expect(outcome.decisionTrace.ai?.status).toBe("rejected");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("deterministic");
    expect(outcome.patternId).not.toBe("process");
  });

  it("AI が使えない場合も決定論的に選択される（生成は止まらない）", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      advisor: new PatternAdvisor(nullAIProvider),
    });
    expect(outcome.decisionTrace.ai?.status).toBe("unavailable");
    expect(outcome.patternId).toBe("three-reasons");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("deterministic");
  });

  it("human override は AI より優先される", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      advisor: new PatternAdvisor(recommender("message-only")),
      overrides: { "slide-ai": { patternId: "three-reasons", reason: "レビューで指定" } },
    });
    expect(outcome.decisionTrace.decision.selectedBy).toBe("human");
    expect(outcome.patternId).toBe("three-reasons");
    expect(outcome.decisionTrace.ai?.status).toBe("advisory");
  });

  it("不適合な human override は黙って無視せず例外にする", async () => {
    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        overrides: { "slide-ai": { patternId: "process" } },
      }),
    ).rejects.toThrow(SelectionPolicyError);

    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        overrides: { "slide-ai": { patternId: "does-not-exist" } },
      }),
    ).rejects.toThrow(/登録されていません/);

    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        overrides: { "slide-ai": { patternId: "three-reasons", variantId: "compact" } },
      }),
    ).rejects.toThrow(/variant/);
  });

  it("company Pattern も候補に入り、AI が選べる", async () => {
    const company = compileDeclarativePattern({
      schemaVersion: "1.0",
      kind: "declarative",
      id: "company-cards",
      version: "1.0.0",
      description: "企業標準の 3 カード",
      metadata: { scope: "company", status: "approved" },
      compatibility: { logicalRelations: ["claim_evidence"] },
      slots: [
        { id: "headline", type: "text", source: "keyMessage" },
        { id: "items", type: "items", source: "contentOrEvidence", min: 3, max: 3 },
      ],
      match: { baseScore: 0.5, requireSlots: ["items"] },
      layout: {
        header: { headline: { ref: "headline" } },
        body: {
          type: "columns",
          repeat: { slot: "items" },
          item: { type: "text", value: { ref: "item.label" } },
        },
      },
    });
    const registry = new PatternRegistry([
      { pattern: company, scope: "company" },
      messageOnlyPattern,
    ]);

    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry,
      advisor: new PatternAdvisor(recommender("company-cards")),
    });
    expect(outcome.patternId).toBe("company-cards");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("ai");
  });
});

describe("Decision Trace と Advisor Report", () => {
  const store = new InMemoryKnowledgeStore([knowledgeEntry()]);
  const retriever = new DeterministicRetriever(() => store.all());

  it("Visual Plan に decisionTrace が保存される", async () => {
    const spec = await loadExample();
    const { plan, outcomes } = await buildVisualPlanWithPolicy(spec, {
      policy: {
        mode: "assist",
        advisor: new PatternAdvisor(recommender("message-only")),
        retriever,
      },
    });

    expect(outcomes).toHaveLength(6);
    for (const planSlide of plan.slides) {
      expect(planSlide.decisionTrace).toBeDefined();
      expect(planSlide.decisionTrace!.decision.mode).toBe("assist");
      expect(planSlide.decisionTrace!.deterministic.candidatePatternIds.length).toBeGreaterThan(0);
    }
  });

  it("Advisor Report に探索の詳細が残る（Visual Plan とは責務が違う）", async () => {
    const spec = await loadExample();
    const { outcomes } = await buildVisualPlanWithPolicy(spec, {
      policy: {
        mode: "auto",
        advisor: new PatternAdvisor(recommender("message-only")),
        retriever,
      },
    });
    const report = buildAdvisorReport(outcomes, { mode: "auto", provider: "scripted" });

    expect(report.slides).toHaveLength(6);
    expect(report.summary.aiRecommended).toBe(6);
    expect(report.summary.aiAccepted).toBeGreaterThan(0);
    // 候補の全評価（落ちた理由込み）は Advisor Report 側にだけある
    expect(report.slides[0]!.deterministic.candidates.length).toBeGreaterThan(1);
    expect(report.slides[0]!.references.length).toBeGreaterThanOrEqual(0);
  });

  it("decisionTrace を含む Visual Plan も検証を通り描画できる", async () => {
    const spec = await loadExample();
    const result = await generate(spec, {
      policy: { mode: "assist", advisor: new PatternAdvisor(recommender("message-only")), retriever },
    });
    expect(result.advisorReport?.mode).toBe("assist");
    expect(result.issues.filter((i) => i.severity === "error")).toEqual([]);
    expect(result.plan.slides[0]!.decisionTrace).toBeDefined();
  });

  it("不正な decisionTrace を持つ Visual Plan は拒否される", async () => {
    const spec = await loadExample();
    const { plan } = await buildVisualPlanWithPolicy(spec, { policy: { mode: "manual" } });
    const tampered = JSON.parse(JSON.stringify(plan));
    tampered.slides[0].decisionTrace.decision.selectedBy = "robot";
    const { parseVisualPlan } = await import("../src/plan/visual-plan.js");
    expect(() => parseVisualPlan(tampered)).toThrow(/構造が不正/);
  });

  it("AI を使わないモードでは V1 と同じ Visual Plan になる（決定論性の維持）", async () => {
    const spec = await loadExample();
    const { buildVisualPlan } = await import("../src/plan/visual-plan.js");
    const deterministic = buildVisualPlan(spec);
    const { plan } = await buildVisualPlanWithPolicy(spec, { policy: { mode: "manual" } });

    expect(plan.slides.map((s) => s.patternId)).toEqual(
      deterministic.slides.map((s) => s.patternId),
    );
    expect(plan.slides.map((s) => s.bindings)).toEqual(
      deterministic.slides.map((s) => s.bindings),
    );
  });
});

describe("Structure Advisor", () => {
  it("AI 出力は必ず presentation.json スキーマを通る", async () => {
    const provider = new ScriptedAIProvider({
      responses: {
        "structure-planning": {
          title: "AI が構造化した資料",
          slides: [
            {
              id: "slide-001",
              keyMessage: "在庫管理コストの増加が課題である",
              slideRole: "problem",
              logicalRelation: "claim_evidence",
              evidence: [
                { id: "e1", type: "text", statement: "コストが3年間で12%増加している" },
              ],
            },
          ],
        },
      },
    });
    const result = await new StructureAdvisor(provider).advise({ request: "在庫管理の課題" });
    expect(result.spec.slides).toHaveLength(1);
    expect(result.spec.slides[0]!.slideRole).toBe("problem");
  });

  it("AI が描画情報を混ぜた場合はスキーマ検証で落ちる", async () => {
    const provider = new ScriptedAIProvider({
      responses: {
        "structure-planning": {
          title: "描画情報を混ぜた資料",
          slides: [
            {
              id: "slide-001",
              keyMessage: "これは通らない",
              slideRole: "problem",
              logicalRelation: "claim_evidence",
              evidence: [],
              patternId: "three-reasons",
              color: "#FF0000",
            },
          ],
        },
      },
    });
    await expect(new StructureAdvisor(provider).advise({ request: "x" })).rejects.toThrow(
      PresentationValidationError,
    );
  });

  it("Provider が無ければ明示的に失敗する", async () => {
    await expect(
      new StructureAdvisor(nullAIProvider).advise({ request: "x" }),
    ).rejects.toThrow(AIUnavailableError);
  });

  it("過去スライドを参照情報として渡せる", async () => {
    const provider = new ScriptedAIProvider({
      responses: {
        "structure-planning": (input) => {
          const references = (input as { references: unknown[] }).references;
          return {
            title: `参照 ${references.length} 件`,
            slides: [
              {
                id: "s1",
                keyMessage: "参照を踏まえた主張",
                slideRole: "analysis",
                logicalRelation: "claim_evidence",
                evidence: [],
              },
            ],
          };
        },
      },
    });
    const store = new InMemoryKnowledgeStore([knowledgeEntry()]);
    const references = await store.search({ slideRole: "analysis" });
    const result = await new StructureAdvisor(provider).advise({ request: "x", references });
    expect(result.spec.title).toBe("参照 1 件");
    expect(result.referenceSlides).toEqual(["fy2025-strategy#18"]);
  });
});

// ---------------------------------------------------------------------------
// 安全ゲート: AI も人間も同じ境界を通る
// ---------------------------------------------------------------------------

/** 常に一定の match スコアを返すテスト用 Pattern */
function fixedScorePattern(
  id: string,
  matchScore: number,
  overrides: Partial<AnyVisualPattern> = {},
): AnyVisualPattern {
  return {
    id,
    version: "1.0.0",
    description: `テスト用 (match=${matchScore})`,
    variants: ["standard"],
    compatibility: { logicalRelations: ["claim_evidence"] },
    bindingsSchema: z.object({ headline: z.string() }).strict(),
    match: () => ({ compatible: true, score: matchScore, reason: `固定スコア ${matchScore}` }),
    buildBindings: (slide) => ({ headline: slide.keyMessage }),
    render: () => undefined,
    ...overrides,
  };
}

/** 候補集合の検証を無視して低スコア Pattern を押し込んでくる Advisor */
function rogueAdvisor(patternId: string): PatternAdvisor {
  return {
    providerId: "rogue",
    available: true,
    advise: () =>
      Promise.resolve({
        provider: "rogue",
        status: "advisory" as const,
        recommendation: {
          recommendedPatternId: patternId,
          confidence: 1,
          reason: "候補集合を無視して推薦する",
        },
      }),
  } as unknown as PatternAdvisor;
}

describe("Human Override も安全ゲートを突破できない", () => {
  /** logicalRelation は合うが match() が不適合を返すスライド（before-after の典型） */
  function matchFailsSlide() {
    return slide({
      id: "slide-match-fails",
      logicalRelation: "before_after",
      evidence: [textEvidence("e1", "定性的な変化"), textEvidence("e2", "別の変化")],
    });
  }

  it("match() が compatible: false を返す Pattern の指定は拒否される", async () => {
    const target = matchFailsSlide();
    // 前提: compatibility と constraints は通るが match で落ちる
    const evaluation = evaluatePattern(beforeAfterPattern, target);
    expect(evaluation.rejectedAt).toBe("match");

    await expect(
      decideSlidePattern(target, {
        mode: "manual",
        overrides: { "slide-match-fails": { patternId: "before-after" } },
      }),
    ).rejects.toThrow(SelectionPolicyError);
    await expect(
      decideSlidePattern(target, {
        mode: "manual",
        overrides: { "slide-match-fails": { patternId: "before-after" } },
      }),
    ).rejects.toThrow(/match 段階/);
  });

  it("compatibility 違反は引き続き拒否される", async () => {
    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        overrides: { "slide-ai": { patternId: "process" } },
      }),
    ).rejects.toThrow(/compatibility 段階/);
  });

  it("constraints 違反は引き続き拒否される", async () => {
    const target = slide({ id: "slide-no-evidence", logicalRelation: "before_after", evidence: [] });
    await expect(
      decideSlidePattern(target, {
        mode: "manual",
        overrides: { "slide-no-evidence": { patternId: "before-after" } },
      }),
    ).rejects.toThrow(/constraints 段階/);
  });

  it("未実装 variant は引き続き拒否される", async () => {
    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        overrides: { "slide-ai": { patternId: "three-reasons", variantId: "compact" } },
      }),
    ).rejects.toThrow(/variant 'compact' は実装されていません/);
  });

  it("minimumScore を下回る Pattern の指定も拒否される", async () => {
    const registry = new PatternRegistry([
      fixedScorePattern("high-score", 1),
      fixedScorePattern("low-score", 0.2),
      messageOnlyPattern,
    ]);
    await expect(
      decideSlidePattern(threeItemSlide(), {
        mode: "manual",
        registry,
        minimumScore: 0.8,
        overrides: { "slide-ai": { patternId: "low-score" } },
      }),
    ).rejects.toThrow(/minimumScore/);
  });

  it("正常な human override は AI より優先される（既存仕様は維持）", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      advisor: new PatternAdvisor(recommender("message-only")),
      overrides: { "slide-ai": { patternId: "three-reasons", reason: "レビューで指定" } },
    });
    expect(outcome.decisionTrace.decision.selectedBy).toBe("human");
    expect(outcome.patternId).toBe("three-reasons");
    expect(outcome.variantId).toBe("standard");
    expect(outcome.decisionTrace.ai?.status).toBe("advisory");
  });

  it("deterministic / AI / human の適合性判定の境界が揃っている", async () => {
    const target = matchFailsSlide();

    // 1) deterministic: 候補としては残るが勝者にはならない
    const deterministic = selectPattern(target);
    expect(deterministic.patternId).not.toBe("before-after");
    expect(
      deterministic.candidates.find((c) => c.patternId === "before-after")?.compatible,
    ).toBe(false);

    // 2) AI: 候補集合に入らないため推薦しても棄却される
    const provider = recommender("before-after");
    const aiOutcome = await decideSlidePattern(target, {
      mode: "auto",
      advisor: new PatternAdvisor(provider),
    });
    const sent = provider.calls[0]!.input as { candidates: Array<{ patternId: string }> };
    expect(sent.candidates.map((c) => c.patternId)).not.toContain("before-after");
    expect(aiOutcome.decisionTrace.ai?.status).toBe("rejected");
    expect(aiOutcome.patternId).not.toBe("before-after");

    // 3) human: 例外になる
    await expect(
      decideSlidePattern(target, {
        mode: "manual",
        overrides: { [target.id]: { patternId: "before-after" } },
      }),
    ).rejects.toThrow(SelectionPolicyError);
  });

  it("checkSelectable は 3 経路が共有する 1 つのゲート", () => {
    const ok = checkSelectable(messageOnlyPattern, threeItemSlide());
    expect(ok.ok).toBe(true);

    const failed = checkSelectable(beforeAfterPattern, matchFailsSlide());
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.reason).toContain("match 段階");

    const belowScore = checkSelectable(messageOnlyPattern, threeItemSlide(), {
      minimumScore: 0.99,
    });
    expect(belowScore.ok).toBe(false);
    if (!belowScore.ok) expect(belowScore.reason).toContain("minimumScore");

    const badVariant = checkSelectable(messageOnlyPattern, threeItemSlide(), {
      variantId: "compact",
    });
    expect(badVariant.ok).toBe(false);
  });
});

describe("minimumScore は AI 候補にも適用される", () => {
  const MINIMUM = 0.8;

  function scoredRegistry() {
    return new PatternRegistry([
      fixedScorePattern("high-score", 1), // 最終スコア 1.00
      fixedScorePattern("low-score", 0.2), // 最終スコア 0.44
      messageOnlyPattern,
    ]);
  }

  it("前提: 2 つの Pattern は compatible だがスコアが違う", () => {
    const registry = scoredRegistry();
    const { candidates } = selectPattern(threeItemSlide(), { registry });
    const high = candidates.find((c) => c.patternId === "high-score")!;
    const low = candidates.find((c) => c.patternId === "low-score")!;
    expect(high.compatible && low.compatible).toBe(true);
    expect(high.score).toBeGreaterThan(MINIMUM);
    expect(low.score).toBeLessThan(MINIMUM);
    expect(meetsSelectionGate(low, MINIMUM)).toBe(false);
  });

  it("minimumScore 未満の Pattern は AI へ渡らない", async () => {
    const provider = recommender("high-score");
    await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry: scoredRegistry(),
      minimumScore: MINIMUM,
      advisor: new PatternAdvisor(provider),
    });

    const sent = provider.calls[0]!.input as { candidates: Array<{ patternId: string }> };
    const ids = sent.candidates.map((c) => c.patternId);
    expect(ids).toContain("high-score");
    expect(ids).not.toContain("low-score");
    expect(ids).not.toContain("message-only");
  });

  it("AI が minimumScore 未満の Pattern を返しても棄却される", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry: scoredRegistry(),
      minimumScore: MINIMUM,
      advisor: new PatternAdvisor(recommender("low-score")),
    });
    expect(outcome.decisionTrace.ai?.status).toBe("rejected");
    expect(outcome.patternId).toBe("high-score");
    expect(outcome.decisionTrace.decision.selectedBy).toBe("deterministic");
  });

  it("Advisor の絞り込みを迂回しても Selection Policy が再検証して棄却する", async () => {
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry: scoredRegistry(),
      minimumScore: MINIMUM,
      advisor: rogueAdvisor("low-score"),
    });
    expect(outcome.decisionTrace.ai?.status).toBe("rejected");
    expect(outcome.decisionTrace.ai?.rejectionReason).toContain("minimumScore");
    expect(outcome.patternId).toBe("high-score");
  });

  it("AI は minimumScore 未満の Pattern で fallback 判定を迂回できない", async () => {
    // minimumScore を満たす Pattern が 1 つも無い状況
    const registry = new PatternRegistry([fixedScorePattern("low-score", 0.2), messageOnlyPattern]);
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry,
      minimumScore: MINIMUM,
      advisor: rogueAdvisor("low-score"),
    });

    // 決定論的 Selector は従来どおり fallback する
    expect(outcome.decisionTrace.deterministic.fallbackUsed).toBe(true);
    expect(outcome.patternId).toBe("message-only");
    // AI が低スコア Pattern を押し込むことはできない
    expect(outcome.decisionTrace.ai?.status).toBe("rejected");
  });

  it("minimumScore 未指定なら従来どおり低スコア Pattern も候補に入る", async () => {
    const provider = recommender("low-score");
    const outcome = await decideSlidePattern(threeItemSlide(), {
      mode: "auto",
      registry: scoredRegistry(),
      advisor: new PatternAdvisor(provider),
    });

    const sent = provider.calls[0]!.input as { candidates: Array<{ patternId: string }> };
    expect(sent.candidates.map((c) => c.patternId)).toContain("low-score");
    expect(outcome.decisionTrace.ai?.status).toBe("accepted");
    expect(outcome.patternId).toBe("low-score");
  });

  it("toAdvisableCandidates 自体も minimumScore で絞る", () => {
    const registry = scoredRegistry();
    const { candidates } = selectPattern(threeItemSlide(), { registry });
    const describe_ = (id: string) => {
      const pattern = registry.get(id);
      return { description: pattern.description, variants: pattern.variants };
    };

    expect(toAdvisableCandidates(candidates, describe_).map((c) => c.patternId)).toContain(
      "low-score",
    );
    expect(
      toAdvisableCandidates(candidates, describe_, { minimumScore: MINIMUM }).map(
        (c) => c.patternId,
      ),
    ).toEqual(["high-score"]);
  });

  it("minimumScore を使っても deterministic な選択結果は変わらない", () => {
    const registry = scoredRegistry();
    const withMinimum = selectPattern(threeItemSlide(), { registry, minimumScore: MINIMUM });
    const without = selectPattern(threeItemSlide(), { registry });
    expect(withMinimum.patternId).toBe(without.patternId);
    expect(withMinimum.fallbackUsed).toBe(false);
  });
});
