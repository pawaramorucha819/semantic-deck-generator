import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { compileDeclarativePattern } from "../src/patterns/declarative/compile.js";
import { loadPatternPackage, loadPatternPackages } from "../src/patterns/declarative/loader.js";
import {
  DeclarativePatternValidationError,
  parseDeclarativePattern,
} from "../src/patterns/declarative/schema.js";
import { messageOnlyPattern } from "../src/patterns/message-only/index.js";
import { PatternRegistry, SCOPE_PRIORITY } from "../src/patterns/registry.js";
import { buildVisualPlan, parseVisualPlan, serializeVisualPlan } from "../src/plan/visual-plan.js";
import { renderVisualPlan, toPptxBuffer } from "../src/render/renderer.js";
import { selectPattern } from "../src/selector/selector.js";
import { midnightTheme } from "../src/theme/theme.js";
import {
  extractTexts,
  metricEvidence,
  pointsContent,
  readPptx,
  slide,
  textEvidence,
  tinyPngBytes,
} from "./helpers.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "sdg-decl-"));
  tempDirs.push(dir);
  return dir;
}

afterAll(async () => {
  await Promise.all(tempDirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

/** テスト用の企業 Pattern（3 列カード） */
function threeColumnDoc(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    schemaVersion: "1.0",
    kind: "declarative",
    id: "company-three-column",
    version: "1.0.0",
    description: "3 つの要点を横並びのカードで示す企業標準 Pattern",
    metadata: { scope: "company", status: "approved", organization: "ACME" },
    compatibility: {
      logicalRelations: ["claim_evidence", "problem_solution"],
      slideRoles: ["analysis", "recommendation"],
    },
    constraints: { maxKeyMessageLength: 80 },
    slots: [
      { id: "eyebrow", type: "text", source: "slideRoleLabel", optional: true },
      { id: "headline", type: "text", source: "keyMessage" },
      { id: "subline", type: "text", source: "supportingMessage", optional: true },
      { id: "sourceText", type: "text", source: "sourceText", optional: true },
      {
        id: "items",
        type: "items",
        source: "contentOrEvidence",
        min: 3,
        max: 3,
        labelMaxLength: 30,
      },
    ],
    match: {
      baseScore: 0.6,
      requireSlots: ["items"],
      rules: [
        {
          when: { itemCount: { eq: 3 }, slot: "items" },
          score: 0.3,
          reason: "要点が 3 件で 3 分割レイアウトに最適",
        },
      ],
    },
    layout: {
      header: {
        eyebrow: { ref: "eyebrow" },
        headline: { ref: "headline" },
        subline: { ref: "subline" },
      },
      footer: { text: { ref: "sourceText" } },
      body: {
        type: "columns",
        gap: "md",
        repeat: { slot: "items" },
        item: {
          type: "card",
          fill: "surface",
          border: "border",
          accent: "primary",
          child: {
            type: "stack",
            gap: "sm",
            children: [
              {
                type: "text",
                value: { template: "0{{item.index}}" },
                style: { role: "caption", color: "primary", bold: true, font: "numeric" },
              },
              {
                type: "text",
                value: { ref: "item.label" },
                style: { role: "heading", bold: true, lineHeight: 1.28 },
              },
              {
                type: "text",
                value: { ref: "item.metric" },
                when: { hasValue: "item.metric" },
                style: {
                  role: "heading",
                  fontScale: 1.3,
                  font: "numeric",
                  color: "primary",
                  bold: true,
                },
              },
              {
                type: "text",
                value: { ref: "item.description" },
                when: { hasValue: "item.description" },
                style: { role: "body", color: "mutedText", lineHeight: 1.38, shrink: true },
                height: "fill",
              },
            ],
          },
        },
      },
    },
    ...overrides,
  };
}

function threeItemSlide() {
  return slide({
    id: "slide-decl",
    keyMessage: "3 つの構造的な原因がある",
    supportingMessage: "現場ヒアリングから、遅延の要因は 3 つに集約された",
    slideRole: "analysis",
    logicalRelation: "claim_evidence",
    content: pointsContent(
      { label: "定型的な指摘に人手を使っている", description: "規約・命名の指摘が多い", evidenceRefs: ["e1"] },
      { label: "レビューが特定の担当者に集中", description: "上位層が担当する割合が高い", evidenceRefs: ["e2"] },
      { label: "レビュー開始までの待ち時間が長い", description: "初回コメントまでが長い", evidenceRefs: ["e3"] },
    ),
    evidence: [
      metricEvidence("e1", "定型指摘の割合", 35, "%"),
      metricEvidence("e2", "上位層が占める比率", 58, "%"),
      metricEvidence("e3", "初回コメントまでの中央値", 18, "時間"),
    ],
  });
}

describe("Declarative Pattern スキーマ検証", () => {
  it("正しい宣言をコンパイルできる", () => {
    const pattern = compileDeclarativePattern(threeColumnDoc());
    expect(pattern.id).toBe("company-three-column");
    expect(pattern.version).toBe("1.0.0");
    expect(pattern.variants).toEqual(["standard"]);
    expect(pattern.constraints).toEqual({ maxKeyMessageLength: 80 });
  });

  it("未知のカラートークンを拒否する", () => {
    const doc = threeColumnDoc();
    (doc as any).layout.body.item.fill = "brandBlue";
    expect(() => parseDeclarativePattern(doc)).toThrow(DeclarativePatternValidationError);
  });

  it("未定義の slot を参照するレイアウトを拒否する", () => {
    const doc = threeColumnDoc();
    (doc as any).layout.body.item.child.children[1].value = { ref: "notASlot" };
    expect(() => compileDeclarativePattern(doc)).toThrow(/未定義の slot/);
  });

  it("repeat の外で item.* を参照するレイアウトを拒否する", () => {
    const doc = threeColumnDoc({
      layout: {
        header: { headline: { ref: "headline" } },
        body: { type: "text", value: { ref: "item.label" } },
      },
    });
    expect(() => compileDeclarativePattern(doc)).toThrow(/repeat の中でしか参照できません/);
  });

  it("repeat と children の併用を拒否する", () => {
    const doc = threeColumnDoc({
      layout: {
        header: { headline: { ref: "headline" } },
        body: {
          type: "columns",
          repeat: { slot: "items" },
          item: { type: "text", value: { ref: "item.label" } },
          children: [{ type: "text", value: { ref: "headline" } }],
        },
      },
    });
    expect(() => compileDeclarativePattern(doc)).toThrow(/併用できません/);
  });

  it("未知の要素タイプを拒否する", () => {
    const doc = threeColumnDoc({
      layout: {
        header: { headline: { ref: "headline" } },
        body: { type: "video", src: "http://example.com/a.mp4" },
      },
    });
    expect(() => parseDeclarativePattern(doc)).toThrow(DeclarativePatternValidationError);
  });

  it("スクリプト・式のような未知キーを拒否する（任意コード実行の経路を作らない）", () => {
    const doc = threeColumnDoc({ script: "process.exit(1)" });
    expect(() => parseDeclarativePattern(doc)).toThrow(DeclarativePatternValidationError);

    const withExpr = threeColumnDoc();
    (withExpr as any).layout.body.item.child.children[0].value = { expr: "1+1" };
    expect(() => parseDeclarativePattern(withExpr)).toThrow(DeclarativePatternValidationError);
  });

  it("プロトタイプ汚染を狙うパスを拒否する", () => {
    const doc = threeColumnDoc();
    (doc as any).layout.body.item.child.children[1].value = { ref: "__proto__.polluted" };
    expect(() => parseDeclarativePattern(doc)).toThrow(DeclarativePatternValidationError);
  });

  it("存在しない slot を requireSlots が参照している場合を拒否する", () => {
    const doc = threeColumnDoc();
    (doc as any).match.requireSlots = ["nope"];
    expect(() => parseDeclarativePattern(doc)).toThrow(/存在しない slot/);
  });

  it("draft はレビュー前として登録を拒否する", () => {
    const doc = threeColumnDoc({
      metadata: { scope: "company", status: "draft" },
    });
    expect(() => compileDeclarativePattern(doc)).toThrow(/draft/);
    expect(() => compileDeclarativePattern(doc, { allowDraft: true })).not.toThrow();
  });

  it("未対応の schemaVersion を拒否する", () => {
    expect(() => parseDeclarativePattern(threeColumnDoc({ schemaVersion: "2.0" }))).toThrow(
      /schemaVersion/,
    );
  });
});

describe("Declarative Pattern の bindings", () => {
  const pattern = compileDeclarativePattern(threeColumnDoc());

  it("content.points から items を抽出する", () => {
    const bindings = pattern.buildBindings(threeItemSlide(), "standard") as any;
    expect(bindings.headline).toBe("3 つの構造的な原因がある");
    expect(bindings.eyebrow).toBe("分析");
    expect(bindings.items).toHaveLength(3);
    expect(bindings.items[0]).toMatchObject({ index: 1, metric: "35%" });
  });

  it("content が無ければ evidence から items を抽出する", () => {
    const bindings = pattern.buildBindings(
      slide({
        evidence: [
          textEvidence("e1", "理由A"),
          textEvidence("e2", "理由B"),
          textEvidence("e3", "理由C"),
        ],
      }),
      "standard",
    ) as any;
    expect(bindings.items.map((i: any) => i.label)).toEqual(["理由A", "理由B", "理由C"]);
  });

  it("生成された bindingsSchema が不正な bindings を拒否する", () => {
    expect(() => pattern.bindingsSchema.parse({ headline: "x" })).toThrow();
    expect(() => pattern.bindingsSchema.parse({ garbage: true })).toThrow();
    const valid = pattern.buildBindings(threeItemSlide(), "standard");
    expect(() => pattern.bindingsSchema.parse(valid)).not.toThrow();
  });

  it("slot を抽出できない場合 buildBindings は失敗する（fallback へ退避できる）", () => {
    expect(() => pattern.buildBindings(slide({ evidence: [] }), "standard")).toThrow(
      /slot を抽出できません/,
    );
  });

  it("match は宣言されたルールでスコアリングする", () => {
    const result = pattern.match(threeItemSlide());
    expect(result.compatible).toBe(true);
    expect(result.score).toBeCloseTo(0.9, 5);
    expect(result.reason).toContain("3 分割");
  });

  it("素材が足りないスライドには compatible: false を返す", () => {
    const result = pattern.match(slide({ evidence: [textEvidence("e1", "1件だけ")] }));
    expect(result.compatible).toBe(false);
    expect(result.score).toBe(0);
  });
});

describe("Declarative Pattern の描画（RenderContext 経由）", () => {
  function planFor(doc = threeColumnDoc()) {
    const pattern = compileDeclarativePattern(doc);
    const registry = new PatternRegistry([
      { pattern, scope: "company" },
      messageOnlyPattern,
    ]);
    const plan = buildVisualPlan(
      { title: "declarative", slides: [threeItemSlide()] },
      { registry },
    );
    return { registry, plan, pattern };
  }

  it("Declarative Pattern が選ばれ、テキストが PPTX に出力される", async () => {
    const { registry, plan } = planFor();
    expect(plan.slides[0]!.patternId).toBe("company-three-column");

    const { pptx, issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);

    const { slideXmls } = await readPptx(await toPptxBuffer(pptx));
    const texts = extractTexts(slideXmls[0]!).join("\n");
    expect(texts).toContain("3 つの構造的な原因がある");
    expect(texts).toContain("定型的な指摘に人手を使っている");
    expect(texts).toContain("35%");
    expect(texts).toContain("01");
  });

  it("配置がすべて Render Report に記録される（レイアウト検査の対象になる）", () => {
    const { registry, plan } = planFor();
    const { reports } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    const placements = reports[0]!.placements;
    expect(placements.length).toBeGreaterThan(8);
    expect(placements.some((p) => p.kind === "text")).toBe(true);
    expect(placements.some((p) => p.kind === "shape")).toBe(true);
  });

  it("Theme を差し替えると色だけが変わる（Pattern は Theme を知らない）", async () => {
    const { registry, plan } = planFor();
    const light = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    const dark = renderVisualPlan(plan, {
      registry,
      includeTitleSlide: false,
      theme: midnightTheme,
    });
    const lightXml = (await readPptx(await toPptxBuffer(light.pptx))).slideXmls[0]!;
    const darkXml = (await readPptx(await toPptxBuffer(dark.pptx))).slideXmls[0]!;
    expect(lightXml).toContain("0B4F8A");
    expect(darkXml).toContain("5AA9F0");
    expect(extractTexts(darkXml)).toEqual(extractTexts(lightXml));
  });

  it("キャンバス外へ出る absolute 配置はレイアウト検査が検出する", () => {
    const doc = threeColumnDoc({
      layout: {
        header: { headline: { ref: "headline" } },
        body: {
          type: "absolute",
          children: [
            { box: { x: 9.4, y: 1, w: 2, h: 0.4 }, node: { type: "text", value: { ref: "headline" } } },
          ],
        },
      },
    });
    const { registry, plan } = planFor(doc);
    const { issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.map((i) => i.code)).toContain("off-canvas");
  });

  it("重なる absolute 配置はレイアウト検査が検出する", () => {
    const doc = threeColumnDoc({
      layout: {
        body: {
          type: "absolute",
          children: [
            { box: { x: 1, y: 1, w: 3, h: 1 }, node: { type: "text", value: { ref: "headline" } } },
            { box: { x: 2, y: 1.2, w: 3, h: 1 }, node: { type: "text", value: { ref: "headline" } } },
          ],
        },
      },
    });
    const { registry, plan } = planFor(doc);
    const { issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.map((i) => i.code)).toContain("overlap");
  });

  it("カードを横に並べたとき行の高さが揃う（列ごとにずれない）", () => {
    const { registry, plan } = planFor();
    const { reports } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    const placements = reports[0]!.placements.filter((p) => p.kind === "text");

    // 各カードの 2 番目のテキスト（item.label）が同じ y に並ぶ
    const labels = placements.filter((p) => p.text?.startsWith("定型") || p.text?.startsWith("レビュー") || p.text?.startsWith("待ち"));
    expect(labels.length).toBe(3);
    const ys = labels.map((p) => Number(p.box.y.toFixed(3)));
    expect(new Set(ys).size).toBe(1);

    // 数値（metric）も同じ y に並ぶ
    const metrics = placements.filter((p) => /^\d+(%|時間)$/.test(p.text ?? ""));
    expect(metrics.length).toBe(3);
    expect(new Set(metrics.map((p) => Number(p.box.y.toFixed(3)))).size).toBe(1);
  });

  it("align: middle は内容の高さに縮めて縦中央に置く", () => {
    const stretched = threeColumnDoc();
    const centered = threeColumnDoc();
    (centered as any).layout.body.align = "middle";

    const heightOf = (doc: Record<string, unknown>): number => {
      const { registry, plan } = planFor(doc);
      const { reports } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
      const cards = reports[0]!.placements.filter((p) => p.name.includes("card"));
      return cards[0]!.box.h;
    };

    const stretchedHeight = heightOf(stretched);
    const centeredHeight = heightOf(centered);
    expect(centeredHeight).toBeLessThan(stretchedHeight);
  });

  it("保存した Visual Plan を再生できる（bindings が再検証される）", async () => {
    const { registry, plan } = planFor();
    const roundTripped = parseVisualPlan(JSON.parse(serializeVisualPlan(plan)), { registry });
    const a = await readPptx(
      await toPptxBuffer(renderVisualPlan(plan, { registry, includeTitleSlide: false }).pptx),
    );
    const b = await readPptx(
      await toPptxBuffer(
        renderVisualPlan(roundTripped, { registry, includeTitleSlide: false }).pptx,
      ),
    );
    expect(b.slideXmls).toEqual(a.slideXmls);
  });

  it("改竄された bindings は Declarative Pattern でも拒否される", () => {
    const { registry, plan } = planFor();
    const tampered = JSON.parse(serializeVisualPlan(plan));
    tampered.slides[0].bindings.items = "not-an-array";
    expect(() => parseVisualPlan(tampered, { registry })).toThrow(/契約に違反/);
  });
});

describe("Pattern パッケージのロードと Registry 階層", () => {
  async function writePackage(
    root: string,
    name: string,
    doc: Record<string, unknown>,
    extra?: { preview?: boolean; asset?: boolean },
  ): Promise<string> {
    const dir = join(root, name);
    await mkdir(dir, { recursive: true });
    if (extra?.preview) await writeFile(join(dir, "preview.png"), tinyPngBytes());
    if (extra?.asset) {
      await mkdir(join(dir, "assets"), { recursive: true });
      await writeFile(join(dir, "assets", "icon.png"), tinyPngBytes());
    }
    await writeFile(join(dir, "pattern.json"), JSON.stringify(doc, null, 2), "utf8");
    return dir;
  }

  it("pattern.json / preview.png / assets を読み込める", async () => {
    const root = await makeTempDir();
    const dir = await writePackage(
      root,
      "three-column",
      threeColumnDoc({
        metadata: {
          scope: "company",
          status: "approved",
          previewImage: "preview.png",
        },
        assets: [{ id: "icon", kind: "image", path: "assets/icon.png" }],
      }),
      { preview: true, asset: true },
    );

    const pkg = await loadPatternPackage(dir);
    expect(pkg.scope).toBe("company");
    expect(pkg.previewImagePath).toContain("preview.png");
    expect(pkg.assets[0]!.id).toBe("pattern:company-three-column:icon");
  });

  it("ルート配下の複数パッケージをまとめて読める", async () => {
    const root = await makeTempDir();
    await writePackage(root, "a", threeColumnDoc({ id: "pattern-a" }));
    await writePackage(root, "b", threeColumnDoc({ id: "pattern-b" }));
    const packages = await loadPatternPackages(root);
    expect(packages.map((p) => p.document.id).sort()).toEqual(["pattern-a", "pattern-b"]);
  });

  it("draft はディレクトリ読み込みでも既定で拒否される", async () => {
    const root = await makeTempDir();
    await writePackage(
      root,
      "draft",
      threeColumnDoc({ metadata: { scope: "company", status: "draft" } }),
    );
    await expect(loadPatternPackages(root)).rejects.toThrow(/draft/);
    await expect(loadPatternPackages(root, { allowDraft: true })).resolves.toHaveLength(1);
  });

  it("参照先の無いプレビュー画像はロード時に失敗する", async () => {
    const root = await makeTempDir();
    const dir = await writePackage(
      root,
      "no-preview",
      threeColumnDoc({
        metadata: { scope: "company", status: "approved", previewImage: "preview.png" },
      }),
    );
    await expect(loadPatternPackage(dir)).rejects.toThrow(/プレビュー画像が見つかりません/);
  });

  it("パッケージ外を参照するアセットパスを拒否する", async () => {
    const root = await makeTempDir();
    const doc = threeColumnDoc({
      assets: [{ id: "icon", kind: "image", path: "../../etc/passwd.png" }],
    });
    const dir = await writePackage(root, "escape", doc);
    await expect(loadPatternPackage(dir)).rejects.toThrow();
  });

  it("scope から優先度が決まり、同点時に project > company > builtin となる", () => {
    const company = compileDeclarativePattern(threeColumnDoc({ id: "company-p" }));
    const project = compileDeclarativePattern(threeColumnDoc({ id: "project-p" }));
    const registry = new PatternRegistry([
      { pattern: company, scope: "company" },
      { pattern: project, scope: "project" },
      messageOnlyPattern,
    ]);

    expect(registry.priorityOf("company-p")).toBe(SCOPE_PRIORITY.company);
    expect(registry.priorityOf("project-p")).toBe(SCOPE_PRIORITY.project);

    // 同一宣言＝同スコアなので、scope 優先度の高い project が勝つ
    const selection = selectPattern(threeItemSlide(), { registry });
    expect(selection.patternId).toBe("project-p");
  });

  it("scope 優先度は意味的ハードゲートを突破しない", () => {
    // project scope だが logicalRelation が合わない Pattern
    const project = compileDeclarativePattern(
      threeColumnDoc({
        id: "project-only-timeline",
        compatibility: { logicalRelations: ["timeline"] },
      }),
    );
    const registry = new PatternRegistry([{ pattern: project, scope: "project" }, messageOnlyPattern]);
    const selection = selectPattern(threeItemSlide(), { registry });
    expect(selection.patternId).toBe("message-only");
    expect(
      selection.candidates.find((c) => c.patternId === "project-only-timeline")?.rejectedAt,
    ).toBe("compatibility");
  });

  it("Registry の一覧に code / declarative の区別が出る", () => {
    const registry = new PatternRegistry([
      { pattern: compileDeclarativePattern(threeColumnDoc()), scope: "company", source: "/tmp/x" },
      messageOnlyPattern,
    ]);
    const list = registry.list();
    expect(list[0]).toMatchObject({
      id: "company-three-column",
      kind: "declarative",
      scope: "company",
      source: "/tmp/x",
    });
    expect(list[1]).toMatchObject({ id: "message-only", kind: "code", scope: "builtin" });
  });

  it("Pattern パッケージの画像アセットが PPTX へ埋め込まれる", async () => {
    const root = await makeTempDir();
    const dir = await writePackage(
      root,
      "with-icon",
      threeColumnDoc({
        id: "with-icon",
        assets: [{ id: "icon", kind: "image", path: "assets/icon.png" }],
        layout: {
          header: { headline: { ref: "headline" } },
          body: {
            type: "stack",
            gap: "sm",
            children: [
              { type: "image", assetId: "icon", height: 0.4, width: 0.4 },
              { type: "text", value: { ref: "headline" }, height: "fill" },
            ],
          },
        },
      }),
      { asset: true },
    );

    const pkg = await loadPatternPackage(dir);
    const registry = new PatternRegistry([messageOnlyPattern]).registerPackage(pkg);
    const plan = buildVisualPlan({ title: "icon", slides: [threeItemSlide()] }, { registry });
    expect(plan.slides[0]!.patternId).toBe("with-icon");

    const { pptx, issues } = renderVisualPlan(plan, { registry, includeTitleSlide: false });
    expect(issues.filter((i) => i.severity === "error")).toEqual([]);
    const { files } = await readPptx(await toPptxBuffer(pptx));
    expect(files.filter((f) => f.startsWith("ppt/media/")).length).toBeGreaterThan(0);
  });
});
