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

```
presentation.json ──▶ Visual Pattern Selector ──▶ Visual Plan ──▶ Renderer ──▶ .pptx
   意味・論理            決定論的スコアリング        表現の確定       決定論的描画
```

設計の詳細は `README.md` を参照してください。

## 設計上、絶対に壊してはいけない不変条件

判断に迷ったら「**意味 → 表現 → 描画**」の責務が混ざっていないかを基準にしてください。
アーキテクチャの実態は Remotion に近い: **Pattern = 意味変換 + レイアウト Component**、
**Visual Plan = Pattern 呼び出し + 検証済み props (+ version / theme 契約)**、
**Renderer = 描画バックエンド**。最終座標は Pattern の `render()` が決め、
配置結果が必要なら Visual Plan ではなく Render Report を使います。

1. **`presentation.json` に描画情報を入れない**
   座標・色・フォントサイズ・グラフ種別・`patternId` などは書けません。
   Zod スキーマが全て `.strict()` なので、混入すると検証エラーになります。
2. **evidence を「根拠」以外に使わない**
   工程・アクション・並列項目など「見せる対象」は `content`（sequence / points）に書きます。
   `Evidence = 主張を裏付ける事実` という原則を Pattern の都合で崩さないでください。
3. **`render()` に `SlideSpec` を渡さない**
   必ず `SlideSpec → buildBindings() → Bindings → render()` の順を通します。
   `render()` のシグネチャに `SlideSpec` を足さないでください。
4. **Bindings は `bindingsSchema` で実行時検証される**
   全 Pattern が Zod スキーマを持ち、生成時（buildVisualPlan）と再生時
   （renderVisualPlan / parseVisualPlan）の両方で検証されます。
   bindings の互換性を壊したら Pattern の `version` を上げてください
   （version 不一致の Visual Plan は黙って再生されず、例外になります）。
5. **Pattern の中で色やフォントをハードコードしない**
   `RenderContext` は Theme のトークン名（`"primary"` など）しか受け付けません。
   相対的な強弱が必要なら `ctx.tint(token, amount)` を使います。
6. **Pattern から PptxGenJS を直接呼ばない**
   描画は必ず `RenderContext`（`text` / `shape` / `chart`）経由で行います。
   配置矩形の記録＝レイアウト検査（はみ出し・重なり・実効フォントサイズ）がこれに依存しています。
7. **Selector に LLM を持ち込まない（V1 スコープ）**
   選択は決定論的スコアリングのみ。`match()` のスコアは共通ルーブリック
   （src/patterns/pattern.ts 参照）準拠の 0〜1 で、範囲外は実行時エラーになります。
8. **fallback（message-only）の契約を守る**
   fallback は constraints を持てず、全 logicalRelation 対応が必須です
   （PatternRegistry が登録時に強制）。どんな入力でも必ず描画できる緊急表現です。

## よく使うコマンド

```bash
npm run generate -- examples/inventory-system.json   # サンプル生成
npm run example                                      # 同上のショートカット
npm test                                             # vitest run
npm run typecheck                                    # tsc --noEmit

# 生成物を目視確認（Windows + PowerPoint が必要）
powershell -File scripts/preview-pptx.ps1 -Pptx examples/generated/inventory-system.pptx -OutDir .preview
```

## ディレクトリ構成

```
src/domain/     意味・論理（型 + Zod スキーマ）
src/theme/      見た目（色・書体・余白）
src/patterns/   意味 → 表現（Visual Pattern）
src/selector/   Pattern 選択（決定論的スコアリング）
src/plan/       Visual Plan の生成と直列化
src/render/     表現 → 描画（PptxGenJS + レイアウト検査）
src/cli/        CLI エントリポイント
```

## Pattern を追加するとき

1. `src/patterns/<pattern-id>/index.ts` に `VisualPattern<TBindings>` を実装
   （`bindingsSchema` 必須。「見せる対象」は `content` から、「根拠」は `evidence` から読む）
2. `src/patterns/registry.ts` の `builtinPatterns` に 1 行追加
3. テストを `tests/` に追加（compatibility / constraints / match / buildBindings +
   競合しうる既存 Pattern との pairwise 選択テスト）
4. サンプルを再生成し、**実際に PPTX を開いて目視確認する**

Selector / Renderer / CLI の変更は不要です。
