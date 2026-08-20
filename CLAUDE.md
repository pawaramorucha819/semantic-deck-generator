# CLAUDE.md

このリポジトリで作業するときの指示です。

## 必須ルール: 変更したら必ず commit & push

**このリポジトリに何か手を加えたら、その作業単位ごとに必ず Git へ commit し、リモートへ push してください。**

- ユーザーからの明示的な依頼を待たずに実行する（このファイルが恒久的な承認です）
- 作業を「完了しました」と報告する前に、commit と push が済んでいること
- 1 つの意味のある変更 = 1 コミット。無関係な変更を 1 コミットに混ぜない
- コミット前に `npm run typecheck` と `npm test` が通ることを確認する
- コミットメッセージは日本語または英語で、**何を・なぜ**変えたかが分かる粒度で書く

```bash
npm run typecheck && npm test
git add -A
git commit -m "<変更内容>"
git push
```

例外: 失敗している状態を意図的に残す場合や、ユーザーから明示的に「まだコミットしないで」と
指示された場合のみ commit をスキップしてよい。その場合は理由を必ず伝えること。

## このプロジェクトについて

`presentation.json`（意味・論理）から決定論的に PowerPoint を生成する Semantic Deck Generator。
V2 では、企業のテンプレート・自作 Pattern・過去資料の知識を扱えるように拡張されています。

```text
presentation.json ─┐
                   ├─▶ Selector ─▶ Selection Policy ─▶ Visual Plan ─▶ Renderer ─▶ .pptx
generation.config ─┘   決定論的       manual/assist/auto   表現の確定     決定論的描画
                        ハードゲート    (+ Knowledge / AI)
```

設計の詳細は `README.md` を参照してください。

---

## 設計上、絶対に壊してはいけない不変条件

判断に迷ったら「**意味 → 表現 → 描画**」の責務が混ざっていないかを基準にしてください。
アーキテクチャの実態は Remotion に近い: **Pattern = 意味変換 + レイアウト Component**、
**Visual Plan = Pattern 呼び出し + 検証済み props (+ version / theme / template 契約)**、
**Renderer = 描画バックエンド**。最終座標は Pattern の `render()` が決め、
配置結果が必要なら Visual Plan ではなく Render Report を使います。

### コア（V1 から不変）

1. **`presentation.json` に描画情報・生成設定を入れない**
   座標・色・フォントサイズ・グラフ種別・`patternId` / `templateId` / `variantId` は書けません。
   生成方法の指定は `generation.config.json` 側です。
   Zod スキーマが全て `.strict()` なので、混入すると検証エラーになります。
2. **evidence を「根拠」以外に使わない**
   工程・アクション・並列項目など「見せる対象」は `content`（sequence / points）に書きます。
   `Evidence = 主張を裏付ける事実` という原則を Pattern の都合で崩さないでください。
3. **`render()` に `SlideSpec` を渡さない**
   必ず `SlideSpec → buildBindings() → Bindings → render()` の順を通します。
4. **Bindings は `bindingsSchema` で実行時検証される**
   コード Pattern も Declarative Pattern も同じ契約です。生成時（buildVisualPlan）と
   再生時（renderVisualPlan / parseVisualPlan）の両方で検証されます。
   bindings の互換性を壊したら Pattern の `version` を上げてください。
5. **Pattern の中で色やフォントをハードコードしない**
   `RenderContext` は Theme のトークン名（`"primary"` など）しか受け付けません。
   相対的な強弱が必要なら `ctx.tint(token, amount)` を使います。
6. **Pattern から PptxGenJS を直接呼ばない**
   描画は必ず `RenderContext`（`text` / `shape` / `chart` / `image`）経由で行います。
   レイアウト検査（はみ出し・重なり・実効フォントサイズ）がこれに依存しています。
7. **fallback（message-only）の契約を守る**
   fallback は constraints を持てず、全 logicalRelation 対応が必須です
   （PatternRegistry が登録時に強制）。どんな入力でも必ず描画できる緊急表現です。

### AI（V1 の「Selector に LLM を持ち込まない」の置き換え）

8. **AI は決定論的 Selector を置き換えない**
   - compatibility / constraints による候補生成と検証は **deterministic layer の責務**
   - AI はその安全な候補集合と Knowledge Retrieval を使って**推薦**できる
   - AI recommendation も最終的に Selection Policy と既存の Pattern 検証を通らなければならない
   - **Renderer は AI を知らない**（`src/render/` から `src/ai/` を import してはいけない）

   具体的に守ること:
   - `PatternAdvisor` に渡す候補は `compatible: true` のものだけ
   - 候補集合に無い patternId、未実装 variant は棄却する（`status: "rejected"`）
   - AI が使えない・失敗した場合も生成は止めない（決定論的勝者へ戻る）
   - AI 出力は必ず Zod で検証する。検証を飛ばす実装は禁止
   - `--ai-mode off` は「AI を使わない」であって「使ったふりをする」ではない。
     Provider が必要なモードで未設定なら **明確なエラー**にする（silent fallback 禁止）

9. **AI Provider をベンダーロックインさせない**
   具体 SDK は `src/ai/` の Adapter の中だけ。`domain` / `patterns` / `render` /
   `templates` / `knowledge` から AI Provider を import してはいけません。
   `@anthropic-ai/sdk` は optional peerDependency で、動的 import されます。

10. **AI 無しで完全に動くこと**
    API キーが無くても build / test / 通常生成がすべて成功する状態を維持してください。
    「LLM API が落ちたので PPT を作れない」システムにしないこと。

### Template

11. **Template と Theme と Pattern の責務を混ぜない**
    - Theme = デザイントークン（色・書体・サイズ・spacing）
    - Template = Theme + Slide Master + ブランドアセット + スライドサイズ
    - Pattern = 情報をどう見せるか

    **Pattern の中に会社ロゴ・企業フッター・機密表示を書かない。** それは Template の責務です。
    逆に Template は「情報をどう見せるか」を知りません。

12. **Template 非使用時の出力を変えない**
    `--template` を指定しない生成は V1 と同じ紙面（余白 + `n / total`）のままにしてください。
    テストで固定されています。

13. **Template のテキスト変数は決定論的なものだけ**
    展開できるのは `{{deckTitle}}` `{{deckSubtitle}}` `{{author}}` `{{organization}}`
    `{{department}}` `{{confidentiality}}` のみ。日付や乱数を足さないこと。

### Declarative Pattern

14. **JSON から任意コードを実行できるようにしない**
    値参照は列挙された安全パス（slot / `item.*` / `index`）のみ、条件はホワイトリスト語彙のみ。
    `eval` / 関数文字列 / 式評価を導入しないでください。
    `__proto__` などのプロパティ名は拒否したままにしてください。

15. **Declarative Pattern も必ず RenderContext を通す**
    `JSON → PptxGenJS 直接呼び出し` にしないこと。
    レイアウト検査・Theme トークン制約・Render Report がすべて効かなくなります。

16. **Capture の成果物は必ず draft**
    `metadata.status: "draft"` で出力し、PatternRegistry は draft を拒否します。
    「AI が勝手に Pattern を増やす」経路を作らないでください
    （Capture → Draft → **人間レビュー** → Register）。

17. **取り込んだ色・サイズは Theme トークン / ロールへ寄せる**
    生の hex を Declarative Pattern に残さないこと。近似したときは必ず警告に残します。

### Knowledge

18. **Knowledge は生成の入力を汚さない**
    過去資料の情報を `presentation.json` へ書き戻さないこと。
    Knowledge は「参照・推薦の材料」であって、意味の一部ではありません。

19. **Retrieval は AI 無しで動くこと**
    既定の検索はメタデータ一致 + トークン重なりによる決定論的スコアリングです。
    Embedding API を必須にしないでください（`KnowledgeRetriever` を実装して差し替える）。

20. **Store は差し替え可能に保つ**
    `DeckKnowledgeStore` インターフェース経由でのみ使い、
    Selector や AI 層が特定のストア実装に依存しないようにしてください。

### Decision Trace / 説明可能性

21. **決定の記録と探索の詳細を二重管理しない**
    - Visual Plan の `decisionTrace` = 再生に必要な **決定の記録**（小さく保つ）
    - Advisor Report = 決定に至る **探索の詳細**（全候補・参照・AI 応答）

22. **「なぜそうなったか」を必ず残す**
    deterministic candidates / retrieved references / AI recommendation /
    final selection / selectedBy / reason が追跡できる状態を保ってください。

### バージョニングと後方互換

23. **保存済み成果物を黙って読み替えない**
    Visual Plan / Pattern / Template / Declarative Pattern / Knowledge /
    Generation Config にはそれぞれ schemaVersion があります。
    非対応バージョンは例外にし、受け入れる場合は `upgradedFrom` のように痕跡を残してください。

24. **V1 の利用方法を壊さない**
    - `npm run generate -- examples/inventory-system.json` が動くこと
    - AI / Template / Knowledge を使わない場合、同じ入力から同じ Pattern 選択・
      同じ Visual Plan・同じ描画結果になること

### セキュリティ

25. **`.pptx` は untrusted input として扱う**
    ZIP パストラバーサル・ZIP 爆弾・XXE（DOCTYPE / ENTITY）・マクロを
    構造的に防いでいます。緩めないでください。
    外部 URL の fetch も禁止です（画像は登録済み `assetId` 経由のみ）。

---

## よく使うコマンド

```bash
npm run generate -- examples/inventory-system.json   # V1 サンプル（従来どおりの入口）
npm run example                                      # 同上のショートカット
npm run example:enterprise                           # V2 企業サンプル一式（API キー不要）

npm run sdg -- generate deck.json --config generation.config.json
npm run sdg -- template import company.pptx --id corporate
npm run sdg -- pattern import sample.pptx --slide 3 --id company-cards
npm run sdg -- knowledge ingest ./past-decks --store ./knowledge
npm run sdg -- knowledge search --store ./knowledge --role recommendation

npm test                                             # vitest run
npm run typecheck                                    # tsc --noEmit

# 生成物を目視確認（Windows + PowerPoint が必要）
powershell -File scripts/preview-pptx.ps1 -Pptx examples/generated/inventory-system.pptx -OutDir .preview
```

## ディレクトリ構成

```text
src/domain/     意味・論理（型 + Zod スキーマ）
src/theme/      デザイントークン
src/templates/  企業テンプレート（Master / Registry / PPTX Import）
src/patterns/   意味 → 表現（Code Pattern / Declarative Pattern / Capture / Registry）
src/pptx/       PPTX 読み取り（自前 XML パーサ / ZIP / 要素抽出）
src/knowledge/  過去資料の蓄積（スキーマ / ストア / 検索 / 取り込み）
src/ai/         AI Provider / Pattern Advisor / Structure Advisor
src/selector/   決定論的選択 + Selection Policy
src/plan/       Visual Plan / Decision Trace / Advisor Report
src/render/     表現 → 描画（RenderContext + レイアウト検査）
src/config/     generation.config.json
src/cli/        CLI（サブコマンド）
```

## Pattern を追加するとき

### Declarative Pattern（TypeScript 不要。企業固有の見せ方はこちらが基本）

1. `patterns/<scope>/<id>/pattern.json` を書く（`README.md` 5 章に例）
2. `metadata.status` を `approved` にする
3. `--pattern-dir` か `generation.config.json` の `patterns.sources` で読み込む
4. 生成して**実際に PPTX を開いて目視確認する**

### Code Pattern

1. `src/patterns/<pattern-id>/index.ts` に `VisualPattern<TBindings>` を実装
   （`bindingsSchema` 必須。「見せる対象」は `content` から、「根拠」は `evidence` から読む）
2. `src/patterns/registry.ts` の `builtinPatterns` に 1 行追加
3. テストを `tests/` に追加（compatibility / constraints / match / buildBindings +
   競合しうる既存 Pattern との pairwise 選択テスト）
4. サンプルを再生成し、**実際に PPTX を開いて目視確認する**

Selector / Renderer / CLI の変更は不要です。
