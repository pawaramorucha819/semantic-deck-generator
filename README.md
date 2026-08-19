# Semantic Deck Generator (V1)

**論理が一目で伝わる PowerPoint を、意味構造から決定論的に生成する仕組み。**

このシステムは「AI に PowerPoint を直接デザインさせる」ものでは**ありません**。
意味（何を伝えるか）・表現（どう伝えるか）・描画（どう描くか）を明確に分離し、
最後の描画だけを完全に決定論的なコードに任せます。

```text
presentation.json          意味・論理        WHAT / WHY / LOGIC
        ↓
Visual Pattern Selector    パターン選択
        ↓
Visual Plan                表現の確定        HOW TO COMMUNICATE
        ↓
PowerPoint Renderer        描画              HOW TO DRAW
        ↓
      .pptx
```

---

## 1. 目的

資料作成において本当に難しいのは、図形を並べることではなく

- 何を主張するのか
- なぜそう言えるのか
- その主張を **どう見せれば一目で伝わるのか**

を決めることです。

本システムはこの 3 つを別々のレイヤーに分離し、
「意味さえ正しく書ければ、見せ方と描画は自動的に決まる」状態を目指します。

## 2. 設計思想 — なぜ意味と表現を分離するのか

LLM に PPTX を直接書かせると、次の問題がまとめて発生します。

| 問題 | 内容 |
| --- | --- |
| 再現性が無い | 同じ入力から毎回違うレイアウトが出る |
| 品質が安定しない | 座標がずれる、文字がはみ出る、要素が重なる |
| 検証できない | 「なぜこの図なのか」を後から説明できない |
| 改善が蓄積しない | 良いレイアウトを次回に再利用できない |
| テーマ変更に弱い | 色を変えるだけで全スライドを作り直すことになる |

そこで責務を次のように切ります。

| レイヤー | 担当 | 知っていること | 知らないこと |
| --- | --- | --- | --- |
| `presentation.json` | 意味・論理 | KeyMessage / Evidence / 論理関係 | 座標・色・図の種類 |
| Visual Pattern | 表現 | 意味構造 → 視覚表現の変換 | 具体的な色・フォント名 |
| Theme | 見た目 | 色・書体・余白・サイズ | レイアウト構造 |
| Renderer | 描画 | PptxGenJS の描き方 | 元の意味・論理 |

**判断に迷ったときの基準:「意味 → 表現 → 描画」の責務が混ざっていないか。**

具体的には、次の 2 つがスキーマ・型レベルで強制されています。

1. `presentation.json` に `fontSize` / `color` / `x` / `patternId` を書くと**検証エラー**になる（Zod `.strict()`）
2. Pattern の `render()` は `SlideSpec` を受け取れない（Bindings しか渡らない）

---

## 3. presentation.json（中間表現 / IR）

意味と論理だけを保持します。

```json
{
  "title": "在庫管理システム導入提案",
  "slides": [
    {
      "id": "slide-004",
      "keyMessage": "A社では導入後に在庫管理コストを30%削減している",
      "supportingMessage": "同業種・同規模のA社では、導入から12か月で年4,800万円から3,360万円へ低下した",
      "slideRole": "evidence",
      "logicalRelation": "before_after",
      "evidence": [
        {
          "id": "evidence-401",
          "type": "caseStudy",
          "statement": "A社では在庫管理システム導入後に在庫管理コストが30%削減された",
          "data": {
            "organization": "A社",
            "before": 4800,
            "after": 3360,
            "unit": "万円",
            "beforeLabel": "導入前",
            "afterLabel": "導入12か月後",
            "changeLabel": "-30%"
          },
          "sourceId": "source-401"
        }
      ],
      "sources": [
        {
          "id": "source-401",
          "title": "A社 在庫管理システム導入事例",
          "url": "https://example.com/case/a-company",
          "date": "2024-11"
        }
      ]
    }
  ]
}
```

### 書いてよいもの / 書いてはいけないもの

| 書いてよい | 書いてはいけない |
| --- | --- |
| keyMessage / supportingMessage | フォントサイズ・色 |
| slideRole（資料内での役割） | 座標・サイズ |
| logicalRelation（スライド内の論理関係） | 棒グラフにするかどうか |
| evidence（根拠とその数値） | 左右レイアウト・矢印の位置 |
| sources（出典） | patternId の直接指定 |

### 型定義

```ts
interface SlideSpec {
  id: string;
  keyMessage: string;
  supportingMessage?: string;
  slideRole: SlideRole;
  logicalRelation: LogicalRelation;
  evidence: Evidence[];
  sources?: Source[];
}

type SlideRole =
  | "context" | "problem" | "analysis" | "insight" | "evidence"
  | "comparison" | "solution" | "recommendation" | "decision" | "summary";

type LogicalRelation =
  | "claim_evidence" | "comparison" | "before_after" | "cause_effect"
  | "process" | "timeline" | "hierarchy" | "composition"
  | "correlation" | "problem_solution" | "pros_cons";

type EvidenceType = "metric" | "caseStudy" | "text" | "comparison";
```

Evidence は `type` による discriminated union です。種別を追加するときは
`src/domain/types.ts` に 1 つ interface を足し、`src/domain/schema.ts` の
union に 1 バリアント足すだけで済みます。

---

## 4. Visual Pattern

Visual Pattern は PowerPoint テンプレートではなく、
**意味構造を視覚表現へ変換する再利用可能な Component** です。

```ts
interface VisualPattern<TBindings> {
  id: string;
  version: string;
  description: string;
  variants: readonly string[];
  fallback?: boolean;

  compatibility: {
    logicalRelations: LogicalRelation[];   // ゲート（外れたら不適合）
    slideRoles?: SlideRole[];              // 減点要素
    evidenceTypes?: EvidenceType[];        // 減点要素
  };

  constraints?: PatternConstraints;        // 定量的な使用条件

  match(slide: SlideSpec): MatchResult;                    // 適合度判定
  buildBindings(slide: SlideSpec, variantId: string): TBindings;  // 意味 → 表現
  render(bindings: TBindings, context: RenderContext): void;      // 描画
}
```

### SlideSpec を直接 render しない

```text
✗  SlideSpec ─────────────────────────────→ render()

✓  SlideSpec ─ buildBindings() ─ Bindings ─→ render()
```

`render()` のシグネチャに `SlideSpec` は現れません。
これにより「描画時に意味へ戻って判断する」ことが構造的に不可能になります。
Bindings は React / Remotion における props に近い位置付けです。

```ts
interface BeforeAfterBindings {
  eyebrow?: string;
  headline: string;
  subline?: string;
  before: { label: string; value: string | number };
  after: { label: string; value: string | number };
  change?: { label?: string; value: string | number };
  unit?: string;
  sourceText?: string;
}
```

### 実装済み Pattern（V1）

| Pattern | 使いどころ | 適用条件（要点） |
| --- | --- | --- |
| `message-only` | KeyMessage を主役に見せる。**fallback** | 常に適用可能 |
| `before-after` | 導入前後・状態変化の対比 | before/after として読める 2 値が抽出できる |
| `three-reasons` | 主張を複数の並列な根拠で支える | evidence 2〜4 件 |
| `process` | 工程・順序 | `process` / `timeline` かつ evidence 2〜5 件 |
| `big-number` | 1 つの数字で成立するスライド | metric evidence がちょうど 1 件 |
| `chart-with-message` | 同一軸に並ぶ数値をネイティブグラフで | 数値系列が抽出できる |

---

## 5. Visual Pattern Selector

V1 では **LLM を使いません**。決定論的なスコアリングだけで選びます。

```text
compatibility  →  constraints  →  match()  →  スコアリング  →  最高スコア
   ゲート          定量条件       意味フィット度
```

```ts
score = 0.7 * match().score + 0.3 * compatibilityScore
```

- `compatibility.logicalRelations` に含まれない → その場で不適合
- `constraints` 違反 → 不適合（理由付きで記録）
- `match()` が `compatible: false` → 不適合（例: before/after の値が取れない）
- 同点時は「fallback でない方」→「レジストリ登録順」で決まるため**常に同じ結果**

出力例:

```json
{
  "patternId": "before-after",
  "variantId": "standard",
  "score": 1.0,
  "reason": "logicalRelation が before_after / caseStudy に before/after が定義されている / 変化量を提示できる"
}
```

適合する Pattern が 1 つも無い場合は、無理に図解せず `message-only` へ fallback します。

---

## 6. Visual Plan

Pattern 決定後、`presentation.json` とは別に Visual Plan を生成します。
ここが「表現が確定した状態」であり、Renderer が読むのはこの JSON だけです。

```json
{
  "schemaVersion": "1.0",
  "title": "在庫管理システム導入提案",
  "themeId": "default",
  "slides": [
    {
      "slideId": "slide-004",
      "slideNumber": 4,
      "patternId": "before-after",
      "patternVersion": "1.0.0",
      "variantId": "standard",
      "score": 1,
      "reason": "logicalRelation が before_after / caseStudy に before/after が定義されている / 変化量を提示できる",
      "fallbackUsed": false,
      "bindings": {
        "eyebrow": "根拠",
        "headline": "A社では導入後に在庫管理コストを30%削減している",
        "subline": "同業種・同規模のA社では、導入から12か月で在庫管理コストが年4,800万円から3,360万円へ低下した",
        "before": { "label": "導入前", "value": 4800 },
        "after": { "label": "導入12か月後", "value": 3360 },
        "change": { "label": "変化", "value": "-30%" },
        "unit": "万円",
        "sourceText": "出典: A社 在庫管理システム導入事例, 2024-11"
      },
      "candidates": [ /* 全 Pattern の評価結果（スコア降順・デバッグ用） */ ]
    }
  ]
}
```

- `candidates` には**落選した Pattern とその理由**も残るため、選択の妥当性を後から検証できます
- Visual Plan にも座標・色・フォントは含まれません（レイアウトは Renderer の責務）
- `--no-candidates` で候補一覧を省略できます

---

## 7. Renderer

- PptxGenJS / 16:9（10in × 5.625in）
- テキスト・図形・グラフはすべて **PowerPoint 上で編集可能**（画像化しない）
- Pattern は PptxGenJS を直接触らず、`RenderContext` 経由でのみ描画する

`RenderContext` は色を **Theme のトークン名**でしか受け付けないため、
Pattern 実装に `#0B4F8A` のようなリテラルが紛れ込みません。

```ts
ctx.text(box, "導入前", { role: "heading", color: "mutedText", align: "center" });
ctx.shape("roundRect", box, { fill: "surface", line: { color: "border" } });
ctx.chart(box, { kind: "bar", categories, series, colors: ["primary"] });
```

### レイアウト検査

`RenderContext` が配置した矩形をすべて記録しているため、生成と同時に検査できます。

| 検査 | 深刻度 | 内容 |
| --- | --- | --- |
| `off-canvas` | error | 要素がスライド外へ出た |
| `out-of-frame` | warning | 要素が安全領域（余白の内側）を超えた |
| `overlap` | error | content レイヤ同士が重なった |
| `text-overflow` | warning | テキストがボックスに収まらない可能性 |

カード背景など「重なるのが正しい」要素は `decoration` レイヤに置かれ、重なり検査の対象外です。

### 決定論性の保証範囲

同じ `presentation.json` からは常に同じ Visual Plan・同じスライド構造が生成されます。
ただし .pptx ファイルの**バイト列**は完全一致しません。理由は次の 2 点だけです。

| 差分 | 原因 |
| --- | --- |
| `docProps/core.xml` | PptxGenJS が生成時刻を埋め込む |
| `ppt/charts/chartN.xml` の連番 | PptxGenJS がプロセス内でグラフに通し番号を振る（中身は同一） |

`ppt/slides/slideN.xml` とグラフの中身はバイト単位で一致し、テストで検証しています。

---

## 8. Theme

Pattern は**構造**を、Theme は**見た目**を担当します。

```ts
interface PresentationTheme {
  id: string;
  name: string;
  fonts:     { heading: string; body: string; numeric?: string };
  fontSizes: { keyMessage: number; heading: number; body: number; caption: number; display?: number };  // pt
  spacing:   { xs: number; sm: number; md: number; lg: number; xl: number };                            // inch
  colors: {
    background: string; text: string; mutedText: string;
    primary: string; secondary: string; accent: string;
    surface?: string; border?: string; positive?: string; negative?: string; onPrimary?: string;
  };
  margins?: { top: number; right: number; bottom: number; left: number };
  shape?:   { cornerRadius: number; borderWidth: number };
}
```

`?` 付きの項目は `resolveTheme()` が自動で補完します（例: `surface` は `primary` の淡色から導出）。

---

## 9. CLI

```bash
npm install

# サンプルを生成
npm run generate -- examples/inventory-system.json
# または
npm run example
```

生成物:

```text
examples/generated/inventory-system.pptx
examples/generated/inventory-system.visual-plan.json
```

### オプション

| オプション | 内容 |
| --- | --- |
| `--out-dir <dir>` | 出力先ディレクトリ（既定: 入力と同階層の `generated/`） |
| `--theme <id>` | テーマ（`default` / `midnight`） |
| `--no-title-slide` | 表紙を付けない |
| `--no-candidates` | Visual Plan から候補一覧を省略 |
| `--strict` | レイアウト警告が 1 件でもあれば終了コード 1 |
| `--quiet` | ログを最小限に |

実行例:

```text
Pattern 選択結果
   1. slide-001  →  chart-with-message/standard  score=0.96
      理由: 3 系列で推移・比較を読み取れる / comparison evidence がそのまま系列になる
   2. slide-002  →  three-reasons/standard  score=1.00
   3. slide-003  →  process/standard  score=0.99
   4. slide-004  →  before-after/standard  score=1.00
   5. slide-005  →  big-number/standard  score=0.97
   6. slide-006  →  message-only/standard  score=0.75

レイアウト検査: 問題なし
```

### ライブラリとして使う

```ts
import { generate, generateFromFile, parsePresentation } from "./src/index.js";

const { plan, pptxBuffer, issues } = await generate(parsePresentation(json));
```

---

## 10. Pattern の追加方法

1. `src/patterns/<pattern-id>/index.ts` を作る
2. `VisualPattern<TBindings>` を実装する
   - `compatibility`: どの `logicalRelation` に適用できるか（ゲート）
   - `constraints`: evidence 件数・文字数などの定量条件
   - `match()`: 意味構造とのフィット度（0〜1）と**理由の文章**
   - `buildBindings()`: `SlideSpec` → Pattern 専用の Bindings（ここだけが意味を読む）
   - `render()`: Bindings → `RenderContext` への描画のみ
3. `src/patterns/registry.ts` の `builtinPatterns` に 1 行追加する

Selector / Renderer / CLI の変更は不要です。

```ts
export const myPattern: VisualPattern<MyBindings> = {
  id: "my-pattern",
  version: "1.0.0",
  description: "...",
  variants: ["standard"],
  compatibility: { logicalRelations: ["cause_effect"] },
  constraints: { minEvidence: 2, maxEvidence: 3 },

  match(slide) {
    const facts = extractSomething(slide);
    if (!facts) return { compatible: false, score: 0, reason: "必要な構造が無い" };
    return { compatible: true, score: 0.9, reason: "cause_effect 構造を持つため" };
  },

  buildBindings(slide) {
    return { headline: slide.keyMessage, /* ... */ };
  },

  render(b, ctx) {
    const body = drawFooter(ctx, drawHeader(ctx, b), b.sourceText);
    ctx.text(body, b.headline, { role: "keyMessage" });
  },
};
```

### 実装時のチェックリスト

- [ ] `render()` の中で意味（`slideRole` / `logicalRelation` / `evidence`）を参照していないか
- [ ] 色を `#RRGGBB` で直書きしていないか（Theme トークンを使う）
- [ ] `match()` の `reason` が人間に読める説明になっているか
- [ ] 適用できないケースで `compatible: false` を返しているか（無理に描かない）

## 11. Variant の追加方法

Variant は「同じ意味構造に対する見せ方の差分」です。

1. `variants` 配列に ID を足す（例: `["standard", "compact"]`）
2. `match()` が状況に応じて `variantId` を返す（例: evidence が多ければ `compact`）
3. `render()` で `ctx.variantId` を見て分岐する

```ts
render(b, ctx) {
  const dense = ctx.variantId === "compact";
  const cardHeight = dense ? 1.6 : 2.5;
  // ...
}
```

V1 では全 Pattern が `standard` のみを実装しています。
将来ユーザーの微調整結果を Pattern に還元するとき、この Variant が受け皿になります。

## 12. Theme の差し替え方法

```ts
import { generate, type PresentationTheme } from "./src/index.js";

const corporateTheme: PresentationTheme = {
  id: "corporate",
  name: "Corporate",
  fonts: { heading: "Meiryo", body: "Meiryo" },
  fontSizes: { keyMessage: 22, heading: 15, body: 12, caption: 9 },
  spacing: { xs: 0.06, sm: 0.12, md: 0.22, lg: 0.36, xl: 0.56 },
  colors: {
    background: "FFFFFF", text: "1A2230", mutedText: "6B7686",
    primary: "8A1538", secondary: "C46A85", accent: "D8A24A",
  },
};

await generate(spec, { theme: corporateTheme });
```

Pattern のコードは一切変更しません。
組み込みテーマ（`default` / `midnight`）は `src/theme/theme.ts` にあり、
CLI からは `--theme midnight` で切り替えられます。

---

## 13. サンプル

`examples/inventory-system.json` — **在庫管理システム導入提案**（6 枚）

| # | KeyMessage | slideRole | logicalRelation | 選ばれた Pattern |
| --- | --- | --- | --- | --- |
| 1 | 在庫管理コストの増加が経営課題となっている | problem | claim_evidence | `chart-with-message` |
| 2 | 現行業務では手作業の在庫管理がボトルネックになっている | analysis | claim_evidence | `three-reasons` |
| 3 | 在庫管理システムにより主要業務を自動化できる | solution | process | `process` |
| 4 | A社では導入後に在庫管理コストを30%削減している | evidence | before_after | `before-after` |
| 5 | 導入コストを考慮しても2年以内に投資回収が期待できる | analysis | claim_evidence | `big-number` |
| 6 | したがって在庫管理システムを導入するべきである | recommendation | claim_evidence | `message-only` |

同じ `claim_evidence` でも、evidence の**構造**が違えば別の Pattern が選ばれます。
これが「意味から表現が決まる」ということです。

---

## 14. テスト

```bash
npm test          # vitest run
npm run typecheck # tsc --noEmit
```

| テストファイル | 内容 |
| --- | --- |
| `tests/schema.test.ts` | Schema validation / 不正入力の拒否 / 描画語彙の混入拒否 |
| `tests/patterns.test.ts` | compatibility / constraints / match / buildBindings |
| `tests/extra-patterns.test.ts` | big-number / chart-with-message とネイティブグラフ出力 |
| `tests/selector.test.ts` | スコアリング / 落選理由 / fallback / 決定論性 |
| `tests/visual-plan.test.ts` | Visual Plan 生成 / JSON ラウンドトリップ / 座標が混ざらないこと |
| `tests/render.test.ts` | PPTX 内部構造（16:9・テキスト・図形）/ Theme 差し替え / レイアウト検査 |
| `tests/pipeline.test.ts` | End-to-End（ファイル入出力込み） |
| `tests/geometry.test.ts` | レイアウト計算・テキスト計測・配色ユーティリティ |

PPTX は「ファイルが存在する」だけでなく、zip を展開して
`ppt/slides/slideN.xml` / `ppt/charts/chart1.xml` の内容まで検証しています。

### 目視確認（Windows + PowerPoint）

自動テストではレイアウトの「見た目」までは分からないため、
生成物を PNG に書き出して確認する補助スクリプトを用意しています。

```powershell
powershell -File scripts/preview-pptx.ps1 -Pptx examples/generated/inventory-system.pptx -OutDir .preview
```

---

## 15. ディレクトリ構成

```text
src/
  domain/       意味・論理（型 + Zod スキーマ）      ← 座標・色を知らない
  theme/        見た目（色・書体・余白）
  patterns/     意味 → 表現（Visual Pattern）
    shared/     Evidence 抽出・共通レイアウト部品
  selector/     Pattern 選択（決定論的スコアリング）
  plan/         Visual Plan の生成と直列化
  render/       表現 → 描画（PptxGenJS / レイアウト検査）
  cli/          CLI エントリポイント
  pipeline.ts   End-to-End パイプライン
examples/
  inventory-system.json
tests/
```

## 16. V1 のスコープ外

GUI / Web アプリ / LLM API 連携 / PowerPoint 編集 UI / 過去資料 RAG /
ユーザー修正の自動学習 / AI による Visual Review / 高度な AI Pattern Selector /
完全自由レイアウト生成。

V1 は **End-to-End が一本通っていること** を最優先しています。
