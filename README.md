# Semantic Deck Generator (V2)

**論理が一目で伝わる PowerPoint を、意味構造から決定論的に生成する仕組み。**
そして **企業内に蓄積された「資料の作り方」を、再利用可能な知識資産にするための基盤。**

このシステムは「AI に PowerPoint を直接デザインさせる」ものでは**ありません**。
意味（何を伝えるか）・表現（どう伝えるか）・描画（どう描くか）を明確に分離し、
最後の描画だけを完全に決定論的なコードに任せます。
AI は「安全と分かっている候補の中から推薦する」役割だけを担います。

```text
                         ┌────────────────────────┐
                         │ Presentation Template  │
                         │ Master / Theme / Asset │
                         └───────────┬────────────┘
                                     │
                                     ▼
┌──────────────┐          ┌──────────────────────┐
│ 過去の PPTX  │─────────▶│ Deck Knowledge Base  │
└──────┬───────┘          └──────────┬───────────┘
       │ Pattern Capture             │ Retrieval
       ▼                             ▼
┌───────────────────┐      ┌──────────────────────┐
│ Company Patterns  │      │ AI Deck Advisor      │
│ (Declarative)     │◀─────│ Pattern / Structure  │
└────────┬──────────┘      └──────────┬───────────┘
         │                            │
         └──────────────┬─────────────┘
                        ▼
                Pattern Candidates
                        │
                        ▼
           Deterministic Validation   ← compatibility / constraints / match
                        │
                        ▼
            Selection Policy (manual / assist / auto)
                        │
                        ▼
                   Visual Plan        ← 決定の記録（Decision Trace 付き）
                        │
                        ▼
            Deterministic Renderer    ← ここに AI は一切いない
                        │
                        ▼
                      PPTX
```

## ボディ分解と六つの情報関係

`npm run example:relationships` で六関係のPPTX、`npm run example:research` で出典付きの実資料を生成できます。
分解の対象範囲・軸・未確認事項、任意の有限集合による重複/漏れの検査を追加しました。
意味上のMECEは自動認定しません。詳しくは [実行・レビュー手順](docs/semantic-body-workflow.md) と [学びの反映状況](docs/learning-coverage.json) を参照してください。

### 原文からレビュー可能な分解候補へ

`structure prepare/import/replay/review/compile` で、固定した本文・各ノードの引用・分解理由・明示レビューを一組として検査できます。
`npm run example:decomposition` はNECの調査メモと保存済みのアシスタント分解案を、APIキーなしで再生しPPTXにします。
自然文の意味判断は人/アシスタントの担当で、CLIが自動理解したとは扱いません。新しい入口は単一階層の説明ボディ（`evidence: []`）に限定します。
詳細とレビューの境界は [原文→分解→レビューの手順](docs/body-decomposition.md) を参照してください。

---

## 目次

1. [目的](#1-目的)
2. [設計思想](#2-設計思想--なぜ意味と表現を分離するのか)
3. [presentation.json（意味・論理）](#3-presentationjson意味論理)
4. [generation.config.json（生成方法）](#4-generationconfigjson生成方法)
5. [Visual Pattern（Code / Declarative）](#5-visual-patterncode--declarative)
6. [Pattern Capture（PowerPoint から Pattern を作る）](#6-pattern-capturepowerpoint-から-pattern-を作る)
7. [Presentation Template（企業テンプレート）](#7-presentation-template企業テンプレート)
8. [Selector と Selection Policy](#8-selector-と-selection-policy)
9. [Deck Knowledge Base](#9-deck-knowledge-base)
10. [AI Layer](#10-ai-layer)
11. [Visual Plan / Decision Trace / Advisor Report](#11-visual-plan--decision-trace--advisor-report)
12. [Renderer](#12-renderer描画バックエンド)
13. [Theme](#13-theme)
14. [CLI](#14-cli)
15. [Pattern の追加方法](#15-pattern-の追加方法)
16. [PPTX Import の対応範囲](#16-pptx-import-の対応範囲)
17. [セキュリティモデル](#17-セキュリティモデル)
18. [バージョニング](#18-バージョニング)
19. [テスト](#19-テスト)
20. [ディレクトリ構成](#20-ディレクトリ構成)
21. [V2 のスコープと今後](#21-v2-のスコープと今後)

---

## 1. 目的

資料作成において本当に難しいのは、図形を並べることではなく

- 何を主張するのか
- なぜそう言えるのか
- その主張を **どう見せれば一目で伝わるのか**

を決めることです。V1 はこの 3 つをレイヤーに分離し、「意味さえ正しく書ければ、
見せ方と描画は自動的に決まる」状態を作りました。

V2 はここに **企業の文脈** を足します。

| V2 で足したもの | 何のために |
| --- | --- |
| Presentation Template | 企業のブランド（ロゴ・配色・マスター）を Pattern から分離したまま適用する |
| Declarative Pattern | TypeScript を書かずに企業標準の見せ方を追加する |
| Pattern Capture | 人が PowerPoint で作った良いスライドを Pattern 化する |
| Deck Knowledge Base | 過去資料を「1 スライド 1 エントリ」で蓄積・検索する |
| AI Layer | 過去資料を踏まえて Pattern を推薦する（決定はさせない） |
| Selection Policy | manual / assist / auto と人間の override |
| Decision Trace | 「なぜその Pattern になったのか」を後から説明できるようにする |

目指しているのは次の循環です。

```text
人間が良い資料を作る
      ↓
Template / VisualPattern / Knowledge として蓄積される
      ↓
AI が過去の知識を検索・推薦する
      ↓
人間または Selection Policy が判断する
      ↓
決定された Visual Plan を deterministic Renderer が正確に再現する
      ↓
また良い資料作成ノウハウが蓄積される
```

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
アーキテクチャは React / Remotion のモデルに近く、**Pattern が「意味変換 + レイアウトを持つ Component」**、
**Visual Plan が「Component 呼び出し + 検証済み props」**、**Renderer が「描画バックエンド」**です。

| レイヤー | 担当 | 知っていること | 知らないこと |
| --- | --- | --- | --- |
| `presentation.json` | 意味・論理 | KeyMessage / Content / Evidence / 論理関係 | 座標・色・図の種類・Pattern |
| `generation.config.json` | 生成方法 | Template / Pattern ソース / Knowledge / AI | 資料の中身 |
| Theme | デザイントークン | 色・書体・サイズ・余白 | レイアウト構造 |
| Template | 企業の紙面 | Slide Master・ロゴ・フッター・スライドサイズ | 情報をどう見せるか |
| Visual Pattern | 意味 → 表現 | Bindings への変換と配置計算 | 具体的な色・フォント名（トークンのみ） |
| Knowledge | 過去の事実 | 過去スライドの意味・使われた Pattern | これから作る資料 |
| AI | 推薦 | 候補集合と過去事例 | 座標・最終決定権 |
| Selector | 検証と決定 | ハードゲートと選択方式 | 描画 |
| Visual Plan | 決定の記録 | どの Pattern をどの props で呼ぶか | 最終的な座標 |
| Renderer | 描画 | RenderContext の提供・PPTX 変換・配置検査 | 元の意味・AI の存在 |

**判断に迷ったときの基準:「意味 → 表現 → 描画」の責務が混ざっていないか。**

次のことがスキーマ・型レベルで強制されています。

1. `presentation.json` に `fontSize` / `color` / `x` / `patternId` を書くと**検証エラー**（Zod `.strict()`）
2. Pattern の `render()` は `SlideSpec` を受け取れない（Bindings しか渡らない）
3. Bindings は各 Pattern の `bindingsSchema` で**生成時と再生時の両方**で検証される
4. Declarative Pattern から任意の JavaScript は実行できない（データしか書けない）
5. AI は決定論的に生成された候補集合の外を選べない（候補外 ID は棄却される）
6. Renderer は AI Provider を import していない（Visual Plan 確定後は完全に決定論的）

---

## 3. presentation.json（意味・論理）

意味と論理だけを保持します。**V2 でもここは一切変わりません。**

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

`content.kind` は `sequence`（順序あり）、`points`（並列）、`distribution`（一対多）、`convergence`（多対一）、`cycle`（閉じた反復）、`context`（共通の環境）の6種。
各項目は `evidenceRefs` で `evidence[].id` を参照でき（未定義参照は検証エラー）、
「見せる項目」と「それを裏付ける根拠」を分離したまま接続できます。

### 書いてよいもの / 書いてはいけないもの

| 書いてよい | 書いてはいけない |
| --- | --- |
| keyMessage / supportingMessage | フォントサイズ・色・座標 |
| slideRole（資料内での役割） | 棒グラフにするかどうか |
| logicalRelation（スライド内の論理関係） | patternId / variantId |
| content（見せる情報構造） | templateId / masterId |
| evidence（根拠とその数値） | AI recommendation / retrieval 結果 |
| sources（出典・URL・参照箇所・取得日） | referenceDeckId |

**生成方法に関する指定はすべて `generation.config.json` 側**に置きます。

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

---

## 4. generation.config.json（生成方法）

「どう生成するか」は presentation.json ではなくこちらに書きます。

```json
{
  "schemaVersion": "1.0",

  "template": { "id": "acme", "dirs": ["./templates"] },
  "theme": { "id": "midnight" },

  "patterns": {
    "builtin": true,
    "sources": [{ "dir": "./patterns/company", "scope": "company" }],
    "allowDraft": false
  },

  "knowledge": {
    "stores": [{ "type": "json", "dir": "./knowledge" }],
    "context": { "purpose": "投資判断", "audience": "経営会議" },
    "limit": 3
  },

  "ai": {
    "mode": "assist",
    "provider": { "type": "anthropic", "model": "claude-opus-5", "effort": "medium" },
    "patternSelection": true,
    "structurePlanning": false
  },

  "selection": {
    "overrides": { "slide-006": { "patternId": "three-reasons", "reason": "レビューで指定" } },
    "sectionSlides": ["slide-003"],
    "minimumScore": 0
  },

  "output": {
    "dir": "./generated",
    "titleSlide": true,
    "renderReport": true,
    "advisorReport": true
  },

  "recordKnowledge": { "deckId": "fy2025-expense", "purpose": "投資判断" }
}
```

| セクション | 意味 |
| --- | --- |
| `template` | 企業テンプレートの ID と読み込み元 |
| `theme` | Template の theme を上書きする（ビルトインテーマ ID） |
| `patterns` | ビルトインの有無、外部 Pattern パッケージの読み込み元と scope |
| `knowledge` | Knowledge ストアと、検索時の共通文脈 |
| `ai` | `off` / `manual` / `assist` / `auto` と Provider |
| `selection` | 人間による Pattern 指定、章扉スライド、最低スコア |
| `output` | 出力先と、書き出す成果物 |
| `recordKnowledge` | 生成結果を Knowledge へ自動登録する（知識が循環する） |

すべて Zod で実行時検証されます。`ai.mode` が `assist` / `auto` なのに Provider が無い場合は
**黙って AI 無しにはならず、明確なエラー**になります。

`provider.type` は `none` / `anthropic` / `scripted` の 3 種。`scripted` は
purpose ごとの応答を書いた JSON を読む決定論的 Provider で、API キー無しのデモや CI に使います
（応答も本物と同じスキーマ検証を通ります）。

---

## 5. Visual Pattern（Code / Declarative）

Visual Pattern は PowerPoint テンプレートではなく、
**意味構造を視覚表現へ変換し、自身のレイアウトも担う再利用可能な Component** です。

```text
VisualPattern
├─ Code Pattern        … TypeScript で実装（複雑なレイアウト計算ができる）
└─ Declarative Pattern … JSON で宣言（TypeScript 不要・企業が持ち込める）
```

**両者はまったく同じ契約を満たします。** 実装手段が違うだけで、
Selector から見ても Renderer から見ても区別はありません。

```ts
interface VisualPattern<TBindings> {
  id: string;
  version: string;               // bindings の互換性が変わったら上げる
  description: string;
  variants: readonly string[];
  fallback?: boolean;

  compatibility: { logicalRelations: LogicalRelation[]; slideRoles?: SlideRole[]; evidenceTypes?: EvidenceType[] };
  constraints?: PatternConstraints;
  bindingsSchema: z.ZodType<TBindings>;   // 実行時の契約

  match(slide: SlideSpec): MatchResult;                    // 意味フィット度
  buildBindings(slide: SlideSpec, variantId: string): TBindings;  // 意味 → 表現
  render(bindings: TBindings, ctx: RenderContext): void;   // 表現 → 描画
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

fallback（`message-only`）は「常に描画できる緊急表現」であり、別の契約を持ちます。
**PatternRegistry が登録時に強制します。**

- `constraints` を持ってはならない（自分を失格させる条件を持てない）
- 全 `logicalRelation` に対応していなければならない
- render() はどんな入力でも収める（見出しサイズを段階的に下げて対応）

### SlideSpec を直接 render しない

```text
✗  SlideSpec ─────────────────────────────→ render()
✓  SlideSpec ─ buildBindings() ─ Bindings ─→ render()
```

`render()` のシグネチャに `SlideSpec` は現れません。
これにより「描画時に意味へ戻って判断する」ことが構造的に不可能になります。

### ビルトイン Pattern（Code）

| Pattern | 使いどころ | 適用条件（要点） |
| --- | --- | --- |
| `message-only` | KeyMessage を主役に見せる。**fallback** | 常に適用可能 |
| `before-after` | 導入前後・状態変化の対比 | before/after として読める 2 値が抽出できる |
| `three-reasons` | 主張を複数の並列項目で支える | `content.points` または evidence が 2〜4 件 |
| `process` | 工程・順序 | `process` / `timeline` かつ `content.sequence` が 2〜5 工程 |
| `big-number` | 1 つの数字で成立するスライド | metric evidence がちょうど 1 件 |
| `chart-with-message` | 同一軸に並ぶ数値をネイティブグラフで | 数値系列が抽出できる |

### Declarative Pattern

JSON だけで Pattern を定義できます。**任意のコードは書けません。**

```jsonc
{
  "schemaVersion": "1.0",
  "kind": "declarative",
  "id": "acme-three-column",
  "version": "1.0.0",
  "description": "ACME 標準の 3 列カード",
  "metadata": { "scope": "company", "status": "approved", "previewImage": "preview.png" },

  "compatibility": { "logicalRelations": ["claim_evidence"], "slideRoles": ["analysis"] },
  "constraints": { "maxKeyMessageLength": 70 },

  // 意味 → bindings の抽出規則（source は列挙された語彙のみ）
  "slots": [
    { "id": "headline", "type": "text", "source": "keyMessage" },
    { "id": "sourceText", "type": "text", "source": "sourceText", "optional": true },
    { "id": "items", "type": "items", "source": "contentOrEvidence", "min": 3, "max": 3 }
  ],

  // match() のスコアリング（条件はホワイトリスト語彙のみ）
  "match": {
    "baseScore": 0.62,
    "requireSlots": ["items"],
    "rules": [{ "when": { "itemCount": { "eq": 3 } }, "score": 0.3, "reason": "3 件で 3 列に最適" }]
  },

  "layout": {
    "header": { "headline": { "ref": "headline" } },
    "footer": { "text": { "ref": "sourceText" } },
    "body": {
      "type": "columns",
      "gap": "md",
      "repeat": { "slot": "items" },
      "item": {
        "type": "card",
        "fill": "surface",
        "child": {
          "type": "stack",
          "gap": "sm",
          "children": [
            { "type": "text", "value": { "template": "0{{item.index}}" }, "style": { "role": "caption" } },
            { "type": "text", "value": { "ref": "item.label" }, "style": { "role": "heading", "bold": true } },
            { "type": "text", "value": { "ref": "item.description" }, "height": "fill",
              "when": { "hasValue": "item.description" } }
          ]
        }
      }
    }
  }
}
```

対応している要素:

| 要素 | 説明 |
| --- | --- |
| `stack` | 縦積み（gap / padding / 子の高さ auto・fill・数値） |
| `columns` | 横並び。`repeat` で items slot を繰り返せる。`align` で縦寄せ（既定 stretch） |
| `card` | 角丸カード（builtin Pattern と同じ `drawCard`） |
| `text` | slot 参照・literal・`{{item.label}}` テンプレート |
| `shape` | rect / roundRect / ellipse / line / 矢印など |
| `image` | パッケージ同梱アセット（`assetId` 経由のみ） |
| `spacer` | 余白 |
| `absolute` | 取り込んだスライドの座標をそのまま使う（Pattern Capture 用） |

`header` / `footer` は builtin Pattern と同じ共有部品を使うため、
どの Pattern でも見出しが同じ位置に来ます。

`columns` を `repeat` で使うと、**行の高さがカラム間で自動的に揃います**
（ラベルの折り返し具合で数値や説明の位置がずれない）。
`"align": "middle"` を付けると、カードを内容の高さに縮めて縦中央へ配置します。

**Declarative Pattern も RenderContext を経由します。** したがって
はみ出し・重なり・実効フォントサイズ検査、Theme トークン制約、Render Report への記録が
コード Pattern とまったく同じように機能します。

ロード時に拒否されるもの:

- スキーマに無いキー（`script` / `expr` など）
- 未知のカラートークン・テキストロール
- 未定義の slot 参照（typo）
- `repeat` の外での `item.*` 参照
- `__proto__` 等を含む値パス
- `metadata.status: "draft"`（人間レビュー前）

### Pattern Registry の階層

```text
project (40) > user (30) > team (20) > company (10) > builtin (0)
```

scope の優先度が効くのは **スコア同点時のタイブレークだけ** です。
compatibility / constraints / match のハードゲートは scope より常に優先されるため、
「project scope だから意味的に不適合な Pattern が選ばれる」ことはありません。

---

## 6. Pattern Capture（PowerPoint から Pattern を作る）

人が PowerPoint で作った良いスライドを Pattern 化します。

```text
PPTX → Extract → Interpret → Compile → Draft → 人間レビュー → Approved
        客観構造   意味の解釈   構造化    pattern.json        登録可能
```

| 段階 | 責務 | やらないこと |
| --- | --- | --- |
| **Extract** | 図形・テキスト・座標・色を「見たまま」取り出す | 意味を推測しない |
| **Interpret** | 各要素が何なのかを決める（`CaptureInterpretation`） | 座標を動かさない |
| **Compile** | 解釈を Declarative Pattern の構造へ落とす | **意味を判断しない** |
| **Review** | 人が解釈と結果を確認し `approved` にする | — |

Extract は事実、Interpret は意味、Compile は構造です。
`buildCapturedPattern()`（Compile）は「この文字列は数値だから metric だろう」のような
**意味の判断をしません**。役割は必ず Interpretation から受け取ります。
だから「解釈が正しければ、その解釈どおりの Pattern が必ず出てくる」ことが保証されます。

```bash
npm run sdg -- pattern import sample.pptx --slide 3 --id company-three-column
```

出力:

```text
patterns/company/company-three-column/
  pattern.json          status=draft（このままでは登録できない）
  capture-report.json   要素の内訳・推定の根拠・非対応要素・レビュー項目
  assets/image-1.png    取り込んだ画像
```

### Interpretation（要素の意味を決める）

`CaptureInterpretation` は「どの要素が何なのか」だけを持つ小さなモデルです。
座標も色も持ちません（それは Extract の結果にあります）。

役割は 2 種類あり、**混ぜません**。

| 種類 | 役割 | 意味 |
| --- | --- | --- |
| **Content Role** | `index` | 連番（01 / 02 / 03）。**順番から作り直される** |
| | `label` | 項目の見出し（必須） |
| | `metric` | 主役の数値 |
| | `description` | 補足説明 |
| **Structural Role** | `background` | カードの下地 |
| | `accent` | 強調バー・カラーチップ |
| | `divider` | 区切り線（**1 セルに 1 本**） |

Content Role は**中身が入れ替わる**もの、Structural Role は**見た目を作る**ものです。
accent / divider は Pattern の中だけの表現であり、`presentation.json` の意味モデルには入りません。
「装飾を足したい」ために意味スキーマを拡張してはいけません。

```jsonc
{
  "schemaVersion": "1.1",
  "interpretationSource": "human",
  "headlineIndex": 1,
  "sublineIndex": 2,
  "cells": [
    {
      "backgroundIndex": 3,
      "content": { "indexIndex": 5, "labelIndex": 6, "metricIndex": 7, "descriptionIndexes": [9] },
      "decoration": { "accentIndexes": [4], "dividerIndex": 8 }
    }
    // 他のセルも同じ役割構成であること
  ],
  "staticIndexes": []
}
```

Interpretation は **untrusted input** です。検証は 2 段階あり、
**両方を通した解釈だけが Compiler へ届きます**（入口は
`parseAndValidateCaptureInterpretation()` ひとつ）。

**1. 形式・構造の検証**

- `.strict()`（未知キーは拒否）+ `schemaVersion` の一致
- すべてのインデックスが `0 <= i < 要素数` の整数
- **同じ要素へ 2 つの役割を割り当てられない**（`cells[0].content.label` と
  `cells[1].decoration.accent` のような衝突を検出して、どの経路が競合したかを示す）
- `labelIndex` は必須。セル数 / description 数 / 装飾数に上限
- **全セルの Visual Role 構成が一致すること**。Content Role だけでなく
  background / accent（**個数まで**）/ divider も比較する
- 値参照・条件・式は一切書けない。**eval も関数文字列も無い**

**2. 役割 ↔ 要素種別の検証**

「その要素をその役割として *描けるか*」だけを見ます。判定は Compiler と同じ
helper（`src/patterns/capture/element-roles.ts`）を使うので、
「検証は通ったのに Compiler が描けない」状態は作れません。

| 役割 | 必要な要素 |
| --- | --- |
| headline / subline / eyebrow / sourceText / index / label / metric / description | テキストを持つ要素 |
| `background` | 塗り (solid fill) を持つ図形 |
| `accent` | 塗りを持つ図形、または線 |
| `divider` | 線（コネクタ / 線プリセット / h = 0 / w = 0） |
| `staticIndexes` | 制限なし（元デザインとして残すため） |

ここで判定するのは客観的な描画可能性だけで、**意味の妥当性は見ません**。
`"30%"` を label と解釈するのは（不自然でも）text である以上通します。
それを判断するのは human / heuristic / 将来の AI の責務です。

```text
cells[1].decoration.dividerIndex: 要素 13 (text "開発スピードの向上") は
  divider として利用できません (線 (line) 要素が必要です)
```

**スキーマのバージョン**

現行は `1.1` です。`1.0` の `dividerIndexes: number[]` は読み込み時に
`dividerIndex` へ移行され、痕跡が `notes` に残ります。
ただし `dividerIndexes: [8, 9]` のように複数あった場合は、
黙って先頭を採らずエラーにします（Compiler は 1 本しか描けないため、
どれを使うかは人が決め直す）。

### 連番（index）の扱い

`index` に指定した要素のテキスト（"01"）は**焼き付けません**。
再利用時に項目の順番から作り直されるよう `{{item.paddedOrdinal}}` へ変換します。

| 参照 | 値 | 由来 |
| --- | --- | --- |
| `item.index` | `1`, `2`, `3` | **V1 からの意味のまま**（1 始まり・変更なし） |
| `item.ordinal` | `"1"`, `"2"`, `"3"` | 並び順から導出 |
| `item.paddedOrdinal` | `"01"`, `"02"`, `"03"` | 並び順から導出（ゼロ埋め 2 桁） |

`ordinal` / `paddedOrdinal` は**導出フィールド**で、実データに同名フィールドがあればそちらが勝ちます。
汎用テンプレート言語ではありません（列挙済みの安全パスだけで、式評価は入りません）。

### 解釈の作り方（heuristic / human）

| source | 使い方 | 状態 |
| --- | --- | --- |
| `heuristic` | 既定。明白な反復規則だけを高 Precision で推定し、曖昧なら abstain する | 実装済み |
| `human` | `--interpretation` で解釈 JSON を渡す | 実装済み |
| `ai` | LLM に解釈させる | **未実装**（拡張点だけ用意） |

既定のヒューリスティクスは、まずスライド全体の役割を幾何とフォントサイズだけで推定します。

| 推定 | 根拠 |
| --- | --- |
| 見出し（keyMessage） | 上部で最大フォントのテキスト |
| 役割ラベル（slideRoleLabel） | 見出しの上にある小さく短いテキスト |
| 補足（supportingMessage） | 見出しのすぐ下の、やや小さいテキスト |
| 出典（sourceText） | 下端の小さなテキスト |
| 反復項目（items） | 同じ大きさの図形が横に 2 つ以上並んでいる領域 |
| 固定デザイン | 上記以外（ロゴ・装飾・固定ラベル） |

さらに反復カードの中身については、**セルを横断した強い規則**だけを根拠に役割を決めます
（`src/patterns/capture/heuristic-roles.ts`）。単独の要素だけを見て意味を決めることはしません。

| 役割 | 採用する根拠（すべて全セルで揃っていること） |
| --- | --- |
| `index` | 数字だけの短いテキストで、セル順に並べた値が **1, 2, 3 …** になる。桁の書式・相対位置・フォントも一致 |
| `metric` | 数値が主役の短いテキスト（`30%` `−25%` `2.4×`）。候補が複数なら最大フォントの行だけ |
| `label` / `description` | label が description より**明確に大きいフォント**で、かつ**上に**ある |
| `divider` | 各セルに 1 本ずつある水平／垂直の線。相対位置・長さ・色も一致 |
| `accent` | 各セルの縁に接した細長い帯。相対位置・大きさ・塗りも一致 |
| `background` | SlotInferencer が見つけたカード背景（`canRenderAsBackground()` を満たすもの） |

**根拠が弱ければ推定しません（abstain）。** 曖昧なまま「多分 label」と割り当てるより、
役割を付けずに固定デザインとして残す方を選びます。判断の根拠も abstain の理由も notes に残ります。

```text
- 01 / 02 / 03 が同じ相対位置に並ぶ連番だったため index と推定しました
- 各セルの縁にある細い装飾を accent と推定しました
- label と description のフォント階層が不明確なため（14pt と 14pt）、反復セルの役割推定を中止しました
```

これは**意味理解ではありません**。認識できるのは「人が見てもまず迷わない」
強い反復規則を持つケースだけで、`10% / 20% / 30%` や `2024 / 2025 / 2026` を連番と誤認したり、
斜めのコネクタを区切り線と誤認したりしないことを優先しています。
一般的なスライドの解釈は将来の AI 解釈（**未実装**）の役割です。

ヒューリスティクスの結果は `interpretFromSlotInference()` で Interpretation へ変換されるので、
**Compiler から見えるのは常に Interpretation だけ**です。推定器を差し替えても Compiler は変わりません。

推定が違っていた場合や、ヒューリスティクスが abstain した場合は、
`capture-report.json` の **inventory** を見て人が解釈 JSON を書きます。

```bash
# 1) まず取り込んで inventory と推定の根拠（notes）を見る
npm run sdg -- pattern import sample.pptx --slide 3 --id x

# 2) 解釈 JSON を書いて、解釈どおりに取り込み直す
npm run sdg -- pattern import sample.pptx --slide 3 --id x \
  --interpretation interpretation.json
```


inventory は人がレビューするための要素一覧です。

```jsonc
{
  "inventory": [
    {
      "index": 6, "id": "card-label-1", "kind": "text", "text": "開発スピードの向上",
      "box": { "x": 1.02, "y": 2.94, "w": 3.2, "h": 0.4 },
      "sourceOrder": 6, "role": "cells[0].content.label"
    }
  ]
}
```

Slot だけを直接指定して上書きすることもできます（従来どおり）。

```bash
npm run sdg -- pattern import sample.pptx --slide 3 --id x \
  --headline-index 1 --source-index 12 --static 0,4
```

推定器は `SlotInferencer` インターフェースなので **AI 実装へ差し替えられます**。
AI 由来でも同じスキーマで検証され、範囲外のインデックスは拒否されます。

### 必ず人間のレビューを通る

Capture の成果物は `metadata.status: "draft"` 固定です。
PatternRegistry は draft を受け付けないので、**人がレビューして `approved` に変えるまで
生成には使われません**。「AI が勝手に Pattern を増やす」ことはできません。

取り込んだ色は必ず Theme トークンへ寄せられ（近似した場合は警告）、
フォントサイズは Theme のロール + 倍率に変換されます。生の hex が Pattern に残ることはありません。

どの解釈で作られたかは `metadata.source.interpretationSource` に残るので、
レビュー時に「人が決めたのか、推定なのか」が分かります。

---

## 7. Presentation Template（企業テンプレート）

### Theme との違い

```text
Theme    … 色・書体・サイズ・spacing などの *デザイントークン*
Template … Theme + Slide Master（ロゴ・ヘッダー・フッター・機密表示・ページ番号）
           + ブランドアセット + スライドサイズ
Pattern  … 情報をどう見せるか
```

**Pattern の中に会社ロゴやフッターを書いてはいけません。** それは Template の責務です。
逆に Template は「情報をどう見せるか」を知りません。

### 構造

```ts
interface PresentationTemplate {
  schemaVersion: string;
  id: string;
  version: string;
  name: string;
  theme: PresentationTheme;
  slideSize?: { widthIn: number; heightIn: number };
  masters: SlideMasterDefinition[];
  defaultMasterId: string;
  titleMasterId?: string;
  sectionMasterId?: string;
  assets?: TemplateAsset[];          // ロゴ等（ディレクトリ内の相対パスのみ）
  metadata?: { organization?; department?; description?; confidentialityLabel? };
  import?: { importedFrom?; unsupportedFeatures?; warnings? };
}
```

Master は `title` / `section` / `content` の 3 ロール。各 Master は
背景色・固定要素（rect / line / text / image）・本文領域の余白・ページ番号位置を持ちます。

固定テキストでは `{{deckTitle}}` `{{deckSubtitle}}` `{{author}}` `{{organization}}`
`{{department}}` `{{confidentiality}}` だけが展開されます（未知の変数は例外）。
日付など可変値は決定論性のため展開できません。

### PowerPoint の Slide Master として書き出される

Template の Master は PptxGenJS の Slide Master として登録されるため、
ロゴ・フッター・機密表示は **スライド本体ではなくマスター側**に入ります。
PowerPoint 上で一括編集でき、ページ番号は本物のページ番号フィールドになります。

```bash
npm run sdg -- generate deck.json --template acme --template-dir ./templates
```

### 既存 PPTX からの Import

```bash
npm run sdg -- template import company-template.pptx --id corporate \
  --organization "ACME 株式会社" --confidential "社外秘"
```

```text
templates/corporate/
  template.json
  import-report.json    ← 取り込めなかった要素が必ず記録される
  assets/*.png
```

取り込む範囲は [16. PPTX Import の対応範囲](#16-pptx-import-の対応範囲) を参照。
**完全互換 Import は目指していません。** 取り込めなかったものは silent failure にせず、
`unsupportedFeatures` / `warnings` に必ず残します。

### Template を使わない場合

`--template` を指定しなければ **V1 とまったく同じ紙面**（余白 + `n / total` のページ番号）で
生成されます。Template 非使用時の出力が変わらないことはテストで固定されています。

---

## 8. Selector と Selection Policy

### 決定論的 Selector（ここに LLM は入らない）

```text
compatibility  →  constraints  →  match()  →  スコアリング  →  最高スコア
   ゲート          定量条件       意味フィット度
```

```ts
score = 0.7 * match().score + 0.3 * compatibilityScore
```

- `compatibility.logicalRelations` に含まれない → その場で不適合（ハードゲート）
- `constraints` 違反 → 不適合（理由付きで記録、ハードゲート）
- `match()` が `compatible: false` → 不適合
- `match()` のスコアが 0〜1 の範囲外 → **実行時エラー**
- 同点時は「scope 優先度」→「fallback でない方」→「登録順」で決まるため**常に同じ結果**

適合する Pattern が 1 つも無い場合は、無理に図解せず `message-only` へ fallback します。

### Selection Policy

| モード | AI | 採用されるもの |
| --- | --- | --- |
| `off` | 呼ばない | 決定論的勝者（Decision Trace も付かない = V1 と同一） |
| `manual` | 呼ばない | 決定論的勝者。Knowledge は参考情報として Trace に記録 |
| `assist` | 呼ぶ | **決定論的勝者**（AI 推薦は記録のみ。非対話 CLI の安全な既定） |
| `auto` | 呼ぶ | 検証を通った AI 推薦 |

**人間の override はどのモードでも最優先**です。ただし override も
compatibility / constraints / variant 実装の検証を通ります。不適合な指定は
黙って無視せず `SelectionPolicyError` になります。

```text
                    ┌── 決定論的 Selector ──┐
SlideSpec ─────────▶│ compatibility          │──▶ 候補集合（安全と分かっているもの）
                    │ constraints            │          │
                    │ match()                │          ▼
                    └────────────────────────┘   ┌─ AI Advisor ─┐
                                                 │ 候補 + 過去例 │──▶ 推薦
Knowledge ──────────────── 過去スライド ────────▶└──────────────┘     │
                                                                      ▼
                                              ┌── 再検証（候補内か / variant / 適合）──┐
                                              └──────────────┬─────────────────────────┘
                                                             ▼
                                              human override > AI > deterministic
```

---

## 9. Deck Knowledge Base

過去資料を「1 スライド 1 エントリ」で蓄積します。

```ts
interface DeckKnowledgeEntry {
  schemaVersion: string;
  deckId: string;
  slideId: string;                       // `<deckId>#<slideNumber>`
  metadata: { title?; organization?; department?; purpose?; audience?; createdAt?; tags? };
  semantic?: { keyMessage?; slideRole?; logicalRelation?; keywords? };
  visual?: { patternId?; variantId?; templateId?; density?; elementCounts? };
  content?: { text: string; bulletCount?: number };
  source: { file: string; slideNumber: number; importedAt?; origin?: "pptx" | "generation" };
}
```

### 取り込み

```bash
# 過去 PPTX から（意味は分からないので構造とテキストのみ）
npm run sdg -- knowledge ingest ./past-decks --store ./knowledge \
  --organization "ACME 株式会社" --purpose 投資判断 --audience 経営会議

# 自分で生成したデッキから（意味も Pattern も分かる）
#   generation.config.json の recordKnowledge を設定すると generate 時に自動登録される
```

`slideRole` / `logicalRelation` は PPTX からは決まらないため未設定になり、
その旨が warning として報告されます（`annotations` で補えます）。
一方 **自分の生成結果から取り込む経路** では意味も Pattern も揃うため、
「作った資料がそのまま知識資産になる」循環が閉じます。

### 保存形式

```text
knowledge/
  index.json           デッキ一覧
  decks/<deckId>.json  1 デッキ分のエントリ
```

Git で差分が読める形式です。**外部 Vector DB は必須ではありません。**
`DeckKnowledgeStore` インターフェースを実装すれば PostgreSQL / pgvector /
Elasticsearch などへ差し替えられ、Selector や AI 層は影響を受けません。

### 検索（AI 不要）

```bash
npm run sdg -- knowledge search --store ./knowledge \
  --role recommendation --relation claim_evidence --text "投資回収"
```

メタデータ一致（slideRole / logicalRelation / purpose / audience / patternId / organization）と
テキスト重なり（日本語は文字バイグラム + IDF 風の重み）による**決定論的スコアリング**です。
Embedding API が無くても動き、同じクエリからは常に同じ順序が返ります。
将来 Embedding 検索へ差し替えるときは `KnowledgeRetriever` を実装します。

---

## 10. AI Layer

### AI Provider

```ts
interface AIProvider {
  readonly id: string;
  readonly available: boolean;
  generateStructured<T>(request: AIRequest<T>): Promise<T>;
}
```

- ベンダー SDK は `src/ai/` の Adapter の中だけに閉じ込められています。
  Domain / Pattern / Renderer から AI Provider を import してはいけません。
- 出力は必ず Zod で検証されます。自由テキストをそのまま信用しません。
- 同梱 Provider: `nullAIProvider`（常に明示エラー）/ `ScriptedAIProvider`（決定論的・テスト用）/
  `AnthropicAIProvider`（公式 SDK を動的 import）。
- `@anthropic-ai/sdk` は **optional peerDependency** です。コア生成には不要で、
  未インストール・API キー未設定なら明確なエラーになります。

### Pattern Advisor

```text
SlideSpec → 決定論的候補生成 → (候補 + 過去スライド) → AI → 推薦 → 検証 → Selection Policy
```

AI に渡すのは **compatibility / constraints / match をすべて通過した候補だけ**です。
次の場合は推薦を棄却します（生成は止まらず、決定論的勝者が使われます）。

- 候補集合に無い `patternId` を返した
- 未実装の `variantId` を返した
- 再検証で不適合と判定された
- Provider が利用できない / 呼び出しに失敗した

推薦の形（Zod で検証）:

```json
{
  "recommendedPatternId": "acme-three-column",
  "recommendedVariantId": "standard",
  "confidence": 0.87,
  "reason": "経営会議向けの資料では 3 列カードの利用実績が多いため",
  "referenceSlides": ["fy2024-cost-reduction#4", "fy2024-cost-reduction#12"]
}
```

### Structure Advisor（上流の拡張ポイント）

```text
ユーザーの依頼 → 過去資料の検索 → AI → presentation.json 下書き → parsePresentation()
```

出力は **必ず既存の `parsePresentation()` を通ります**。AI 専用の Semantic Model を作って
Renderer へ流すことはしません。AI が `patternId` や `color` を混ぜた場合、
presentation.json のスキーマが `.strict()` なので検証で落ちます。

### AI なしでも動く

これは V2 の必須要件です。

```text
presentation.json → deterministic selector → Visual Plan → Renderer
```

API キーが無くても build / test / 通常生成はすべて動作します。
`--ai-mode off` は「AI を使わない」であって「AI を使ったふりをする」ではありません。

---

## 11. Visual Plan / Decision Trace / Advisor Report

### Visual Plan

「**どの Pattern（とバージョン）を、どの検証済み props（Bindings）で、
どの Theme / Template の下で呼ぶか**」を確定させた記録です。
Renderer が読むのはこの JSON だけで、最終座標は含まれません。

```json
{
  "schemaVersion": "2.0",
  "title": "経費精算業務の自動化提案",
  "themeId": "acme",
  "templateId": "acme",
  "slides": [
    {
      "slideId": "slide-002",
      "slideNumber": 2,
      "patternId": "acme-three-column",
      "patternVersion": "1.0.0",
      "variantId": "standard",
      "score": 0.944,
      "reason": "...",
      "fallbackUsed": false,
      "bindings": { "headline": "...", "items": [ ... ] },
      "provenance": { "sources": [ ... ], "evidence": [ ... ] },
      "decisionTrace": {
        "deterministic": {
          "recommendedPatternId": "three-reasons",
          "recommendedVariantId": "standard",
          "score": 1.0,
          "fallbackUsed": false,
          "candidatePatternIds": ["three-reasons", "acme-three-column", "message-only"]
        },
        "knowledge": { "referenceSlides": ["fy2024-cost-reduction#4"] },
        "ai": {
          "provider": "scripted:./ai-responses.json",
          "status": "accepted",
          "recommendedPatternId": "acme-three-column",
          "confidence": 0.87,
          "reason": "..."
        },
        "decision": {
          "mode": "auto",
          "selectedPatternId": "acme-three-column",
          "selectedVariantId": "standard",
          "selectedBy": "ai",
          "reason": "..."
        }
      }
    }
  ]
}
```

### 保存・再生の契約（parseVisualPlan）

保存済み Visual Plan を読み込むとき、次がすべて検証されます。

1. 構造
2. `schemaVersion` が対応バージョン（`1.1` / `2.0`）か
3. `patternId` が現在のレジストリに存在するか
4. `patternVersion` が現在の実装と一致するか
5. `variantId` が実装済みか
6. `bindings` が Pattern の `bindingsSchema` に適合するか

V1.1 の Plan は「V2 で追加されたフィールドがすべて任意」であることを根拠に受け入れ、
`upgradedFrom: "1.1"` を残します。**黙って読み替えることはありません。**
検証を通らない Plan は描画されず、必ず例外になります。

### Advisor Report

Visual Plan とは責務が違います。

| 成果物 | 役割 |
| --- | --- |
| Visual Plan | 再生に必要な**決定の記録**。小さく保つ |
| Advisor Report | 決定に至るまでの**探索の詳細**（全候補のスコアと落ちた理由、参照した過去スライドの一致理由、AI の応答） |

```bash
npm run sdg -- generate deck.json --config generation.config.json --advisor-report
# → deck.advisor-report.json
```

---

## 12. Renderer（描画バックエンド）

Pattern は PptxGenJS を直接呼びません。`RenderContext` だけを使います。

```ts
interface RenderContext {
  readonly theme: ResolvedTheme;
  readonly canvas: Rect;    // スライド全面
  readonly frame: Rect;     // 本文領域（Template の Master が決める）
  readonly variantId: string;

  text(box, content, style?): void;
  shape(kind, box, style?): void;
  chart(box, spec): void;
  image(box, spec): void;   // 登録済み assetId のみ（外部 URL 不可）
  measureText(content, opts): number;
  fontSizeOf(role): number;
  tint(token, amount): ColorToken;
}
```

理由は 3 つです。

1. 描画バックエンドを差し替えられる
2. 配置された矩形をすべて記録できる → はみ出し / 重なり検査が可能
3. 色やフォントを *トークン* でしか指定できない → Theme 分離が構造的に守られる

### レイアウト検査

| 検査 | 深刻度 | 内容 |
| --- | --- | --- |
| `off-canvas` | error | 要素がスライド外へ出た |
| `out-of-frame` | warning | 要素が安全領域を超えた |
| `overlap` | error | content レイヤ同士が重なった |
| `text-overflow` | warning | テキストがボックスに収まらない可能性 |
| `text-shrink-below-min` | warning | 自動縮小で実効フォントが下限（既定 9pt）を下回る |

カード背景など「重なるのが正しい」要素は `decoration` レイヤに置かれ、重なり検査の対象外です。

### 決定論性の保証範囲

同じ Visual Plan からは常に同じスライド構造が生成されます。
**Renderer は AI Provider を import していません。** Visual Plan が確定した時点で
AI の関与は終わり、以降は完全に決定論的です。

ただし .pptx ファイルの**バイト列**は完全一致しません。理由は次の 2 点だけです。

| 差分 | 原因 |
| --- | --- |
| `docProps/core.xml` | PptxGenJS が生成時刻を埋め込む |
| `ppt/charts/chartN.xml` の連番 | PptxGenJS がプロセス内で通し番号を振る（中身は同一） |

---

## 13. Theme

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

`?` 付きの項目は `resolveTheme()` が自動で補完します。
ビルトインは `default` / `midnight`、Template 同梱の例として `corporate` / `acme` があります。

Theme の解決順序:

1. `--theme <id>`（明示指定）
2. Template の theme
3. ビルトイン `default`

---

## 14. CLI

```bash
npm run sdg -- <command> [options]
```

| コマンド | 内容 |
| --- | --- |
| `generate <presentation.json>` | PowerPoint を生成する |
| `template import <file.pptx>` | 既存 PPTX から企業テンプレートを取り込む |
| `template list` | テンプレート一覧 |
| `pattern import <file.pptx>` | PPTX の 1 スライドから Pattern 下書きを作る |
| `pattern list` | Pattern 一覧（code / declarative / scope / preview を表示） |
| `knowledge ingest <path...>` | 過去資料を Knowledge へ蓄積する |
| `knowledge search` | 過去スライドを検索する |
| `knowledge stats` | 蓄積状況を表示する |

```bash
# 通常生成（V1 からの入口もそのまま動く）
npm run generate -- examples/inventory-system.json
npm run sdg -- generate examples/inventory-system.json

# 企業テンプレートを指定
npm run sdg -- generate deck.json --template acme --template-dir ./templates

# Generation Config でまとめて指定
npm run sdg -- generate deck.json --config generation.config.json

# AI モードを切り替える
npm run sdg -- generate deck.json --config generation.config.json --ai-mode off
npm run sdg -- generate deck.json --config generation.config.json --ai-mode assist

# 既存 PPTX を企業テンプレートとして取り込む
npm run sdg -- template import company-template.pptx --id corporate

# 良いスライドを Pattern 化する（出力は draft）
npm run sdg -- pattern import sample.pptx --slide 3 --id company-three-column

# 要素の意味を人が指定して Pattern 化する
npm run sdg -- pattern import sample.pptx --slide 3 --id company-three-column \
  --interpretation interpretation.json

# 過去資料を蓄積して検索する
npm run sdg -- knowledge ingest ./past-decks --store ./knowledge
npm run sdg -- knowledge search --store ./knowledge --relation comparison --role recommendation

# 企業向けサンプル一式（API キー不要）
npm run example:enterprise
```

`generate` の主なオプション:

```text
--config <file>       generation.config.json
--out-dir <dir>       出力先
--theme <id>          テーマ
--template <id>       企業テンプレート
--template-dir <dir>  外部テンプレートディレクトリ（複数可）
--pattern-dir <dir>   外部 Pattern ディレクトリ（複数可）
--knowledge-dir <dir> Knowledge ストア（複数可）
--ai-mode <mode>      off | manual | assist | auto
--section <slideId>   章扉（section master）にする（複数可）
--report              Render Report を書き出す
--advisor-report      Advisor Report を書き出す
--strict              レイアウト警告が 1 件でもあれば終了コード 1
```

### ライブラリとして使う

```ts
import {
  generate, generateFromFile,
  loadGenerationConfig, buildGenerationContext,
  loadTemplateFromDirectory, loadPatternPackage,
  ingestDeckFromPptx, ingestDeckFromGeneration, JsonKnowledgeStore,
  PatternAdvisor, ScriptedAIProvider, createAnthropicProvider,
  buildVisualPlan, buildVisualPlanWithPolicy, renderVisualPlan,
} from "./src/index.js";

const result = await generate(spec, { templateId: "acme", templateRegistry });
```

---

## 15. Pattern の追加方法

### Declarative Pattern（推奨: TypeScript 不要）

1. `patterns/company/<id>/pattern.json` を書く（[5 章](#declarative-pattern)の例を参照）
2. `metadata.status` を `approved` にする
3. `--pattern-dir ./patterns/company` か `generation.config.json` の `patterns.sources` で読み込む
4. 生成して**実際に PPTX を開いて目視確認する**

または PowerPoint で作ったスライドから `pattern import` で下書きを作り、レビューして登録します。

### Code Pattern（複雑なレイアウト計算が要る場合）

1. `src/patterns/<pattern-id>/index.ts` に `VisualPattern<TBindings>` を実装
   （`bindingsSchema` 必須。「見せる対象」は `content` から、「根拠」は `evidence` から読む）
2. `src/patterns/registry.ts` の `builtinPatterns` に 1 行追加
3. テストを `tests/` に追加（compatibility / constraints / match / buildBindings +
   競合しうる既存 Pattern との pairwise 選択テスト）
4. サンプルを再生成し、**実際に PPTX を開いて目視確認する**

Selector / Renderer / CLI の変更は不要です。

### チェックリスト

- [ ] `match()` は共通ルーブリックに沿ったスコアを返すか（0〜1）
- [ ] `buildBindings()` が `bindingsSchema` に適合する値を返すか
- [ ] `render()` が Theme トークンだけを使っているか（hex のハードコード禁止）
- [ ] 会社ロゴ・フッターを Pattern に書いていないか（それは Template の責務）
- [ ] レイアウト検査でエラー・警告が出ないか

---

## 16. PPTX Import の対応範囲

**PowerPoint 仕様の完全な再実装は目指していません。** 対応範囲を明示的に絞り、
対応できないものは必ず報告します（silent failure 禁止）。

### 対応している要素

| 要素 | Template Import | Pattern Capture |
| --- | --- | --- |
| スライドサイズ | ✅ | ✅（座標の基準） |
| テーマ配色 / フォント | ✅ | — |
| 単色塗りの図形（rect / line） | ✅ | ✅（roundRect も可） |
| テキスト（段落・フォントサイズ・色・揃え） | ✅ | ✅ |
| 画像（png / jpg / gif） | ✅ | ✅ |
| 単純なグループ（座標変換して展開） | ✅ | ✅ |
| プレースホルダ（本文領域・ページ番号位置） | ✅ | — |
| マスター / レイアウトの固定装飾 | ✅ | — |

### 対応していない要素（必ず報告される）

- SmartArt（diagram）
- グラフ（chart）
- 表（table）
- OLE 埋め込みオブジェクト
- アニメーション・画面切り替え
- 3D 効果・影・光彩などの特殊効果
- グラデーション塗り・画像塗り・パターン塗り
- カスタム図形（custGeom。矩形として近似し報告）
- 外部リンク画像（**取得しません**）
- VBA マクロ（読み込まず、実行もしません）
- 動画・音声

これらは `import-report.json` / `capture-report.json` の `unsupportedFeatures` に記録されます。

---

## 17. セキュリティモデル

企業利用を想定し、次を**構造的に**禁止しています。

| 脅威 | 対策 |
| --- | --- |
| Declarative Pattern からの任意コード実行 | JSON にコード表現が無い。値参照は列挙された安全パスのみ、条件はホワイトリスト語彙のみ |
| Template JSON からの任意コード実行 | 同上。Master 要素は rect / line / text / image のデータだけ |
| プロトタイプ汚染 | `__proto__` / `constructor` / `prototype` を含むパス・XML 属性を拒否 |
| XXE（外部実体参照） | 自前 XML パーサが DOCTYPE / ENTITY を含む文書を拒否 |
| ZIP パストラバーサル | 絶対パス・`..` を含むエントリを拒否。アセットはパッケージ内に限定 |
| ZIP 爆弾 | エントリ数・パートサイズ・展開後合計サイズに上限 |
| PPTX マクロ | 読み込まず、存在を非対応として報告するだけ |
| 外部 URL の無条件 fetch | 画像は登録済み `assetId` 経由のみ。外部リンク画像は取得せず報告 |
| Knowledge 由来の文字列実行 | Knowledge は検索とプロンプト材料にしか使われない。シェル実行の経路が無い |
| AI 出力の無検証利用 | すべて Zod 検証 + 候補集合・variant・compatibility の再検証を通る |

`.pptx` は常に **untrusted input** として扱います。

---

## 18. バージョニング

保存された成果物を将来読み込むため、それぞれにスキーマバージョンがあります。

| 対象 | バージョン | 不一致のときの挙動 |
| --- | --- | --- |
| Visual Plan | `schemaVersion` 2.0（1.1 も受け入れ、`upgradedFrom` を記録） | 対応外は例外 |
| Pattern | `version`（bindings 互換が変わったら上げる） | Plan と実装が不一致なら例外 |
| Template | `schemaVersion` 1.0 | 対応外は例外 |
| Declarative Pattern | `schemaVersion` 1.0 | 対応外は例外 |
| Knowledge | `schemaVersion` 1.0 | 対応外は例外 |
| Generation Config | `schemaVersion` 1.0 | 対応外は例外 |
| Render Report / Advisor Report | `schemaVersion` 1.0 | — |

「スキーマが変わったが黙って読み込んで別結果になる」ことはありません。

---

## 19. テスト

```bash
npm run typecheck
npm test
```

| テストファイル | 内容 |
| --- | --- |
| `tests/schema.test.ts` | presentation.json の検証（描画情報の混入拒否を含む） |
| `tests/geometry.test.ts` | 幾何・テキスト計測 |
| `tests/patterns.test.ts` / `extra-patterns.test.ts` | 各 Pattern の match / buildBindings |
| `tests/selector.test.ts` | 決定論的選択・pairwise 競合 |
| `tests/contracts.test.ts` | スコア契約・fallback 契約・可読性検査 |
| `tests/visual-plan.test.ts` | Visual Plan の生成・保存・再生契約 |
| `tests/render.test.ts` | PPTX 構造・Theme 差し替え・決定論性・入口検証 |
| `tests/templates.test.ts` | Template スキーマ / Registry / Master 選択 / PPTX XML 検査 / 後方互換 |
| `tests/declarative-pattern.test.ts` | 宣言スキーマ / bindings / 描画 / 不正要素拒否 / Registry 階層 |
| `tests/pptx-import.test.ts` | XML パーサ / ZIP 安全性 / Template Import / Pattern Capture |
| `tests/knowledge.test.ts` | ingest / 永続化 / 決定論的検索 |
| `tests/ai-selection.test.ts` | AI Provider / Advisor / manual・assist・auto / Decision Trace |
| `tests/enterprise-e2e.test.ts` | Generation Config と企業向け E2E |
| `tests/pipeline.test.ts` | End-to-End |

外部 API はテストから一切呼びません（AI は `ScriptedAIProvider` を使用）。

### 目視確認（Windows + PowerPoint）

```powershell
powershell -File scripts/preview-pptx.ps1 -Pptx examples/generated/inventory-system.pptx -OutDir .preview
```

---

## 20. ディレクトリ構成

```text
src/
  domain/       意味・論理（型 + Zod スキーマ）
  theme/        デザイントークン
  templates/    企業テンプレート（型 / スキーマ / Registry / Master 生成 / PPTX Import）
  patterns/     意味 → 表現
    <builtin>/  コード Pattern
    declarative/ 宣言的 Pattern（スキーマ / 値解決 / bindings / 描画 / ローダ）
    capture/    PowerPoint からの Pattern 化（推定 / 組み立て）
    registry.ts 階層化された Registry
  pptx/         PPTX 読み取り（自前 XML パーサ / ZIP / 要素抽出）
  knowledge/    過去資料の蓄積（スキーマ / ストア / 検索 / 取り込み）
  ai/           AI Provider / Pattern Advisor / Structure Advisor
  selector/     決定論的選択 + Selection Policy
  plan/         Visual Plan / Decision Trace / Advisor Report
  render/       表現 → 描画（RenderContext + レイアウト検査）
  config/       generation.config.json
  cli/          CLI（サブコマンド）
  pipeline.ts   End-to-End

examples/
  inventory-system.json      V1 サンプル
  ai-code-review.json        V1.1 サンプル
  enterprise/                V2 企業サンプル一式
```

---

## 21. V2 のスコープと今後

### V2 で実装したもの

- Presentation Template（Slide Master / ロゴ / 機密表示 / スライドサイズ）
- 既存 PPTX からの Template Import（サブセット + 明示的な非対応報告）
- Declarative Pattern（JSON から安全にロード・描画）
- Pattern Capture（PowerPoint → Pattern 下書き → 人間レビュー → 登録）
- Deck Knowledge Base（ローカル JSON + 決定論的検索）
- AI Provider 抽象と Pattern Advisor / Structure Advisor
- Selection Policy（manual / assist / auto）と Human Override
- Decision Trace と Advisor Report
- Generation Config とサブコマンド CLI

### V2 のスコープ外（意図的に実装していない）

- PowerPoint 仕様の完全な Import（SmartArt / グラフ / 表 / アニメーション）
- GUI / Web UI（Pattern の preview 情報は将来の GUI 向けに保持済み）
- 外部 Vector DB / Embedding 検索（差し替え可能な抽象は用意済み）
- AI による自由座標生成（**永久にスコープ外**）
- Pattern の自動増殖（Capture は必ず人間レビューを通る）

### 今後の候補

- Embedding Retriever の実装（`KnowledgeRetriever` を差し替えるだけ）
- Structure Advisor の CLI 化（`presentation.json` の下書き生成）
- Variant の拡充（compact / detailed など）
- Render Report を使った AI Visual Review
- Pattern preview 画像の自動生成

### 補足説明なしで読む解説資料

`explanation` に読者の問いと説明セクションを持たせ、図と同じ紙面に表示できます。`npm run example:explanatory` は背景・用語・仕組み・評価条件・限界・導入判断を含む8ページのNEC解説版を生成します。詳細と品質基準は [説明型資料の手順](docs/standalone-explanatory-decks.md) を参照してください。

### 主張を支える採否の量と未知の品質内訳

`npm run example:partition` は、NECの平均候補544/5,035枚を比例帯にし、両側の正解数は「未確認」として示す1枚を生成します。数値の差分・比率は計算として追跡し、主張・分解軸・図の役割・必要な解説を明示します。未知の再現率は計算しません。記録済みassistantレビューは入力変更で失効します。従来の根拠付き入力経路を使い、evidence=[]の原文分解経路の制限は維持します。

入力、レビュー、再現コマンドと限界: [docs/partition-argument.md](docs/partition-argument.md)。
