# Enterprise Example — 企業利用の一式サンプル

V2 の要素をひととおり使ったサンプルです。**API キーは不要**で、そのまま実行できます。

```bash
npm run example:enterprise
```

生成物は `examples/enterprise/generated/` に出ます（`.gitignore` 対象）。

- `presentation.pptx` — 企業テンプレート適用済みのスライド
- `presentation.visual-plan.json` — どの Pattern をどの props で呼んだか + Decision Trace
- `presentation.render-report.json` — 配置結果とレイアウト検査
- `presentation.advisor-report.json` — 候補一覧・参照した過去スライド・AI 推薦の詳細

## 中身

```text
examples/enterprise/
  presentation.json          意味・論理だけ（Pattern も色も書かれていない）
  generation.config.json     生成方法（Template / Pattern / Knowledge / AI）
  ai-responses.json          scripted AI Provider の応答（オフラインで AI 経路を再現）
  templates/acme/            企業テンプレート（Slide Master・ロゴ・機密表示）
    template.json
    assets/logo.png          scripts/make-example-logo.mjs で生成した小さな PNG
  patterns/company/          企業標準の Declarative Pattern（TypeScript なし）
    acme-three-column/pattern.json
  knowledge/                 過去資料の知識（Git 管理できる JSON）
    index.json
    decks/*.json
```

## 実行すると何が起きるか

1. **Corporate Template 適用** — ヘッダー帯・ロゴ・組織名・機密表示・ページ番号が
   PowerPoint の Slide Master 側に入る。Pattern 側には企業固有の要素が一切入らない。
2. **外部 Declarative Pattern の利用** — `acme-three-column` が Registry に company scope で
   登録され、要点 3 件のスライドで採用される。
3. **Knowledge Retrieval** — 各スライドについて、過去デッキから似たスライドを 3 件引く
   （slideRole / logicalRelation / purpose / audience / 本文の重なりで決定論的に検索）。
4. **AI 推薦（scripted）** — `ai.mode: "auto"` なので、検証を通った推薦は自動採用される。
5. **Decision Trace** — 「誰がなぜ選んだか」が Visual Plan と Advisor Report に残る。

## 出力例（Pattern 選択結果）

```text
1. slide-001  →  chart-with-message  [deterministic]
   AI  : rejected → acme-three-column (候補集合に無い patternId を推薦しました)
2. slide-002  →  acme-three-column   [ai]
   AI  : accepted → acme-three-column
3. slide-003  →  process             [deterministic]
   AI  : rejected → acme-three-column (候補集合に無い patternId を推薦しました)
...
```

**これがハードゲートと AI 協業の実演です。**

- サンプルの AI はどのスライドでも `acme-three-column` を推薦しますが、
  意味的に適合しないスライド（推移グラフ・工程図・前後比較）では
  *候補集合に入っていない* ため棄却され、決定論的な選択が使われます。
  AI は「安全と分かっている候補の中から選ぶ」ことしかできません。
- 一方 slide-002 / slide-006 では、決定論的スコアだけならビルトインの
  `three-reasons` が勝ちます（`--ai-mode off` で確認できます）。
  過去資料に「経営会議向けでは ACME 標準の 3 列カードを使ってきた」という
  実績があるため、AI がそれを踏まえて企業標準を推薦し、採用されています。

つまり **「意味的に妥当か」は決定論的レイヤが、「社内でどう見せてきたか」は
Knowledge + AI が担当する**、という役割分担になっています。

## 本物の AI に切り替える

`generation.config.json` の provider を差し替えるだけです。

```json
"ai": {
  "mode": "assist",
  "provider": { "type": "anthropic", "model": "claude-opus-5", "effort": "medium" }
}
```

`ANTHROPIC_API_KEY` を設定し、`npm install @anthropic-ai/sdk` を実行してください。
キーが無い状態で `assist` / `auto` を指定すると、**黙って AI 無しにはならず**
明確なエラーになります。AI を使わない場合は `--ai-mode off` を付けてください。

## 人間が最終決定する（override）

`generation.config.json` に書きます。人間の指定は AI より常に優先されます。

```json
"selection": {
  "overrides": {
    "slide-006": { "patternId": "three-reasons", "reason": "レビューで指定" }
  }
}
```

指定した Pattern が意味的に不適合な場合は、黙って無視せずエラーになります。
