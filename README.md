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
このアーキテクチャは React / Remotion のモデルに近く、**Pattern が「意味変換 + レイアウトを持つ Component」**、
**Visual Plan が「Component 呼び出し + 検証済み props」**、**Renderer が「描画バックエンド」**です。

| レイヤー | 担当 | 知っていること | 知らないこと |
| --- | --- | --- | --- |
| `presentation.json` | 意味・論理 | KeyMessage / Content / Evidence / 論理関係 | 座標・色・図の種類 |
| Visual Pattern | 意味変換 + レイアウト Component | Bindings への変換と、Bindings からの配置計算 | 具体的な色・フォント名（トークンのみ） |
| Visual Plan | Pattern 呼び出しの記録 | どの Pattern をどの props（Bindings）で呼ぶか + version / theme 契約 | 最終的な座標 |
| Theme | 見た目 | 色・書体・余白・サイズ | レイアウト構造 |
| Renderer | 描画バックエンド | RenderContext の提供・PPTX 変換・配置検査 | 元の意味・論理 |

つまり**最終座標は Visual Plan ではなく Pattern の `render()` が決めます**。
配置結果（矩形・検査結果）が必要な場合は Visual Plan ではなく **Render Report**（後述）を使ってください。

**判断に迷ったときの基準:「意味 → 表現 → 描画」の責務が混ざっていないか。**

具体的には、次の 3 つがスキーマ・型レベルで強制されています。

1. `presentation.json` に `fontSize` / `color` / `x` / `patternId` を書くと**検証エラー**になる（Zod `.strict()`）
2. Pattern の `render()` は `SlideSpec` を受け取れない（Bindings しか渡らない）
3. Bindings は各 Pattern の `bindingsSchema`（Zod）で**生成時と再生時の両方**で実行時検証される

---

## 3. presentation.json（中間表現 / IR）

意味と論理だけを保持します。

### content と evidence の区別（重要）

- **`content`** … スライドで「見せる対象」となる情報構造（工程・並列項目・アクション）。主張の中身そのもの。
- **`evidence`** … 主張を**裏付ける根拠**（数値・事例・観察された事実）。

工程やアクションを evidence に詰めてはいけません。`Evidence = 根拠` という原則が壊れると、
根拠の種類と情報構造の種類が同じ union に混ざり、Pattern の都合で Schema が際限なく膨張します。

```json
{
  "content": {
    "kind": "sequence",
    "steps": [
      { "label": "入荷登録", "description": "入荷時にスキャンし在庫数を即時更新する" },
      { "label": "在庫照合", "description": "受発注システムと自動照合する" }
    ]
  },
  "evidence": []
}
```

`content.kind` は V1.1 時点で `sequence`（順序あり）と `points`（並列）の 2 種。
各項目は `evidenceRefs` で `evidence[].id` を参照でき（未定義参照は検証エラー）、
「見せる項目」と「それを裏付ける根拠」を分離したまま接続できます。

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
          "sourceIds": ["source-401"]
        }
      ],
      "sources": [
        {
          "id": "source-401",
          "title": "A社 在庫管理システム導入事例",
          "url": "https://example.com/case/a-company",
          "date": "2024-11",
          "accessedAt": "2025-02-10"
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
| content（見せる情報構造: 工程・並列項目） | 左右レイアウト・矢印の位置 |
| evidence（根拠とその数値） | patternId の直接指定 |
| sources（出典・URL・参照箇所・取得日） | |

### 型定義

```ts
interface SlideSpec {
  id: string;
  keyMessage: string;
  supportingMessage?: string;
  slideRole: SlideRole;
  logicalRelation: LogicalRelation;
  content?: SemanticContent;   // 見せる対象（sequence | points）
  evidence: Evidence[];        // 根拠（sourceIds で sources を参照）
  sources?: Source[];          // url / locator / accessedAt を保持できる
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
**意味構造を視覚表現へ変換し、自身のレイアウトも担う再利用可能な Component** です
（Remotion の Component に相当。座標計算は Pattern の `render()` の責務です）。

```ts
interface VisualPattern<TBindings> {
  id: string;
  version: string;               // bindings の互換性が変わったら上げる
  description: string;
  variants: readonly string[];
  fallback?: boolean;

  compatibility: {
    logicalRelations: LogicalRelation[];   // ゲート（外れたら不適合）
    slideRoles?: SlideRole[];              // 減点要素
    evidenceTypes?: EvidenceType[];        // 減点要素
  };

  constraints?: PatternConstraints;        // 定量的な使用条件（ハードゲート）

  bindingsSchema: z.ZodType<TBindings>;    // Bindings の実行時契約（生成時・再生時に検証）

  match(slide: SlideSpec): MatchResult;                    // 適合度判定
  buildBindings(slide: SlideSpec, variantId: string): TBindings;  // 意味 → 表現
  render(bindings: TBindings, context: RenderContext): void;      // レイアウト + 描画
}
```

### match() のスコア契約

score は Pattern 間で比較されるため、共通ルーブリックに従います（範囲外は Selector が実行時に拒否）。

| score | 意味 |
| --- | --- |
| 0.90 - 1.00 | 意味構造がこの Pattern の理想形 |
| 0.70 - 0.90 | 適用可能で有力 |
| 0.40 - 0.70 | 可能だが他により適切な表現がありうる |
| 0.00 - 0.40 | ほぼ fallback 向け |

### fallback の契約

fallback（`message-only`）は「常に描画できる緊急表現」であり、通常の Pattern とは別の契約を持ちます。
**PatternRegistry が登録時に強制します。**

- `constraints` を持ってはならない（自分を失格させる条件を持てない）
- 全 `logicalRelation` に対応していなければならない
- render() はどんな入力でも収める（`message-only` は見出しサイズを段階的に下げて対応）

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

### 実装済み Pattern（V1.1）

| Pattern | 使いどころ | 適用条件（要点） |
| --- | --- | --- |
| `message-only` | KeyMessage を主役に見せる。**fallback** | 常に適用可能 |
| `before-after` | 導入前後・状態変化の対比 | before/after として読める 2 値が evidence から抽出できる |
| `three-reasons` | 主張を複数の並列項目で支える | `content.points` または evidence が 2〜4 件 |
| `process` | 工程・順序 | `process` / `timeline` かつ `content.sequence` が 2〜5 工程 |
| `big-number` | 1 つの数字で成立するスライド | metric evidence がちょうど 1 件 |
| `chart-with-message` | 同一軸に並ぶ数値をネイティブグラフで | 数値系列が evidence から抽出できる |

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

- `compatibility.logicalRelations` に含まれない → その場で不適合（ハードゲート）
- `constraints` 違反 → 不適合（理由付きで記録、ハードゲート）
- `match()` が `compatible: false` → 不適合（例: before/after の値が取れない）
- `match()` のスコアが 0〜1 の範囲外 → **実行時エラー**（Pattern 作者が範囲外の値で選択を支配できない）
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
Visual Plan は「**どの Pattern（とバージョン）を、どの検証済み props（Bindings）で、どの Theme の下で呼ぶか**」を
確定させた記録であり、Renderer が読むのはこの JSON だけです。
最終座標は含まれません（座標は Pattern の `render()` が Bindings から決定します）。

```json
{
  "schemaVersion": "1.1",
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
      "provenance": {
        "sources": [
          {
            "id": "source-401",
            "title": "A社 在庫管理システム導入事例",
            "url": "https://example.com/case/a-company",
            "date": "2024-11",
            "accessedAt": "2025-02-10"
          }
        ],
        "evidence": [
          {
            "id": "evidence-401",
            "type": "caseStudy",
            "statement": "A社では在庫管理システム導入後に在庫管理コストが30%削減された",
            "sourceIds": ["source-401"]
          }
        ]
      },
      "candidates": [ /* 全 Pattern の評価結果（スコア降順・デバッグ用） */ ]
    }
  ]
}
```

- `candidates` には**落選した Pattern とその理由**も残るため、選択の妥当性を後から検証できます
- `provenance` は出典・根拠の**構造化された記録**です。フッターの表示用出典と違い URL・参照箇所・取得日を失いません（→ スピーカーノートに完全な形で出力）
- Visual Plan には座標・色・フォントは含まれません
- `--no-candidates` で候補一覧を省略できます

### 保存・再生の契約（parseVisualPlan）

保存した Visual Plan は `parseVisualPlan()` を通してから再生します。次を検証し、
通らない Plan が「黙って違う出力」になることはありません（必ず例外になります）。

1. 構造と `schemaVersion` の一致
2. `patternId` が現在のレジストリに存在すること
3. `patternVersion` が現在の Pattern 実装と一致すること（不一致は再生成を要求）
4. `variantId` が実装済みであること
5. `bindings` が各 Pattern の `bindingsSchema` に適合すること

Theme は `plan.themeId` から自動解決されます（`--theme midnight` で生成した Plan は
オプション無しで再生しても midnight 配色になります）。ビルトイン以外のカスタム Theme で
生成した Plan は、同じ Theme オブジェクトを再度渡す必要があります。

---

## 7. Renderer（描画バックエンド）

- PptxGenJS / 16:9（10in × 5.625in）
- テキスト・図形・グラフはすべて **PowerPoint 上で編集可能**（画像化しない）
- Pattern は PptxGenJS を直接触らず、`RenderContext` 経由でのみ描画する
- 描画前に patternVersion / bindings を検証する（契約違反の Plan は例外）
- `provenance` をスピーカーノートへ完全な形（URL・参照箇所・取得日込み）で書き出す

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
| `text-shrink-below-min` | warning | 自動縮小で実効フォントが下限（既定 9pt）を下回る = 「生成できたが読めない」 |

カード背景など「重なるのが正しい」要素は `decoration` レイヤに置かれ、重なり検査の対象外です。
テキスト計測は PowerPoint（Yu Gothic）の実測に合わせた保守的な近似です
（過小評価してあふれるより、過大評価して余白が出る方を選ぶ）。

### Render Report

配置結果（全要素の矩形・検査結果）は **Render Report** として機械可読に出力できます。
最終レイアウトの検証や将来の AI Visual Review には、Visual Plan ではなくこちらを渡してください。

```bash
npx tsx src/cli/generate.ts examples/inventory-system.json --report
# → examples/generated/inventory-system.render-report.json
```

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
| `--report` | Render Report（配置結果 + 検査結果）も書き出す |
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
   6. slide-006  →  message-only/standard  score=0.80

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
   - `compatibility`: どの `logicalRelation` に適用できるか（ハードゲート）
   - `constraints`: evidence 件数・文字数などの定量条件（ハードゲート）
   - `bindingsSchema`: Bindings の Zod スキーマ（実行時契約。生成時・再生時に検証される）
   - `match()`: 意味構造とのフィット度（**共通ルーブリック準拠の 0〜1**）と**理由の文章**
   - `buildBindings()`: `SlideSpec` → Pattern 専用の Bindings（ここだけが意味を読む）
   - `render()`: Bindings → レイアウト計算 + `RenderContext` への描画
3. `src/patterns/registry.ts` の `builtinPatterns` に 1 行追加する
4. bindings の互換性を壊す変更をしたら `version` を上げる（古い Visual Plan は再生時に拒否される）

Selector / Renderer / CLI の変更は不要です。

```ts
export const MyBindingsSchema = z.object({
  headline: z.string().min(1),
  /* ... */
}).strict();

export const myPattern: VisualPattern<MyBindings> = {
  id: "my-pattern",
  version: "1.0.0",
  description: "...",
  variants: ["standard"],
  compatibility: { logicalRelations: ["cause_effect"] },
  constraints: { minEvidence: 2, maxEvidence: 3 },
  bindingsSchema: MyBindingsSchema,

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
- [ ] `match()` のスコアが共通ルーブリック（§4）のアンカーに沿っているか
- [ ] `match()` の `reason` が人間に読める説明になっているか
- [ ] 適用できないケースで `compatible: false` を返しているか（無理に描かない）
- [ ] 「見せる対象」は `content` から、「根拠」は `evidence` から読んでいるか
- [ ] 競合しうる既存 Pattern との pairwise 選択テストを追加したか

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
  fontSizes: { keyMessage: 22, heading: 15, body: 13, caption: 10 },
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

| # | KeyMessage | slideRole | logicalRelation | 意味の素材 | 選ばれた Pattern |
| --- | --- | --- | --- | --- | --- |
| 1 | 在庫管理コストの増加が経営課題となっている | problem | claim_evidence | comparison evidence | `chart-with-message` |
| 2 | 現行業務では手作業の在庫管理がボトルネックになっている | analysis | claim_evidence | text evidence ×3 | `three-reasons` |
| 3 | 在庫管理システムにより主要業務を自動化できる | solution | process | content.sequence ×4 | `process` |
| 4 | A社では導入後に在庫管理コストを30%削減している | evidence | before_after | caseStudy evidence | `before-after` |
| 5 | 導入コストを考慮しても2年以内に投資回収が期待できる | analysis | claim_evidence | metric evidence ×1 | `big-number` |
| 6 | したがって在庫管理システムを導入するべきである | recommendation | claim_evidence | content.points ×1 | `message-only` |

同じ `claim_evidence` でも、意味の素材（content / evidence）の**構造**が違えば別の Pattern が選ばれます。
これが「意味から表現が決まる」ということです。
なお #3 の工程は根拠ではないため evidence ではなく `content.sequence` に、
#2 の 3 項目はヒアリングで観察された事実＝根拠なので evidence に書かれています。

---

## 14. テスト

```bash
npm test          # vitest run
npm run typecheck # tsc --noEmit
```

| テストファイル | 内容 |
| --- | --- |
| `tests/schema.test.ts` | Schema validation / 不正入力の拒否 / 描画語彙の混入拒否 / content 参照整合性 |
| `tests/patterns.test.ts` | compatibility / constraints / match / buildBindings / fallback 登録契約 |
| `tests/extra-patterns.test.ts` | big-number / chart-with-message とネイティブグラフ出力 |
| `tests/selector.test.ts` | スコアリング / 落選理由 / fallback / 決定論性 / pairwise 選択 |
| `tests/visual-plan.test.ts` | Visual Plan 生成 / provenance / parseVisualPlan（改竄・version 不一致の拒否） |
| `tests/contracts.test.ts` | スコア範囲の実行時検証 / fallback 契約 / 長文ストレス / 実効フォントサイズ検査 |
| `tests/render.test.ts` | PPTX 内部構造 / Theme 再生（plan.themeId）/ スピーカーノート / レイアウト検査 |
| `tests/pipeline.test.ts` | End-to-End（ファイル入出力込み） |
| `tests/geometry.test.ts` | レイアウト計算・テキスト計測（較正済み）・配色ユーティリティ |

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
