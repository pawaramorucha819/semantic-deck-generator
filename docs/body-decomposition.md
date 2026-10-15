# 原文 → 分解候補 → 明示レビュー → 既存の図解

## できることと境界

単一の本文から作った分解候補を、原文との対応と判断理由を残してレビューし、既存の `presentation.json` に変換する。
この段階は一つの本文・一つの候補・一つのスライド・単一階層に限定する。
複数案は別の候補ファイルとして比較できる。入れ子ツリー、複合図、自動ページ分割は扱わない。
この入口は説明対象のボディ構造に限定し、`evidence: []` を必須とする。
根拠を伴う旧来の数値Patternが本文ノードを置き換える経路を避けるため、根拠付きスライドは明示的に拒否する。
元本文・出典メタデータ・引用範囲は保持する。根拠付きの主張スライドを扱う既存の生成入口は変更しない。

候補の意味理解は人またはアシスタントが行う。CLI は正規表現で意味を理解したふりをせず、原文保存・検査・レビュー記録・変換を担当する。
すべての CLI コマンドはオフライン。外部API呼び出し、APIキー、追加のAPI利用料金は不要。
保存した候補を再生できることと、自由文から同じ候補を自動推論できることは別である。

## 手順

```sh
# 1. 本文そのものを保存する。JSONなら source に公開元などを追加できる。
npm run sdg -- structure prepare body.txt --title "説明したいこと" --out work/request.json
# work/request.raw.json は正規化していない本文、request.json は候補作成用の指示。

# 2. 指示と原文から人/アシスタントが candidate.json を作る。ここは意味の判断。
npm run sdg -- structure import work/request.raw.json --draft candidate.json --out work/draft.json

# 3. 原文・要素・引用・関係・未対応箇所を読む。承認は自動作成されない。
npm run sdg -- structure review work/request.raw.json --draft work/draft.json --out work/packet.json
# packet.review-template.json は request-changes、空の記入欄を持つ未承認テンプレート。
# 実際にレビューして、別名 review.json に実際の actor / 判断 / 理由を記録する。

# 4. レビュー対象の版が一致した場合だけ、既存入力と lineage を出力する。
npm run sdg -- structure compile work/request.raw.json --draft work/draft.json --review work/review.json --out work/presentation.json
npm run sdg -- generate work/presentation.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir work/generated
npm run preview -- work/generated/presentation.pptx work/preview
```

コマンドは異なる内容の既存出力を上書きしない。同じ保存済み入力の再生は可能。
修正時は新しい出力先を指定し、変更後の候補を改めてレビューする。

## 原文と候補の契約

`RawBody` は `schemaVersion/id/title/body/source?`。本文は最大20,000 UTF-16コード単位。
改行・空白・Unicodeを正規化しない。URLはメタデータとして保持するだけで取得しない。
候補は `status: draft`、明示した候補ID、原文全体のSHA-256、実際の作成者区分を持つ。

- `slide`：既存の意味モデル。`content` と `decomposition` は必須、描画指定は禁止。
- `nodeSources`：中心・結果・環境を含む全表示ノードの明示ID、原文範囲、`source/interpretation`、理由。
- `messageSources`：要旨を確認する原文範囲。
- `relation`：関係を選んだ理由と原文範囲。
- `excludedSpans`：今回の説明から外す原文範囲と理由。

範囲は `{start,end,quote}`。JavaScriptのUTF-16単位、開始を含み末尾を含まない。
`body.slice(start,end) === quote` が必要。絵文字のサロゲートペア分断を拒否する。
`resolveExactQuote` は同じ引用が複数ある場合、明示した位置がなければ拒否する。
引用の一致は「そこに書いてある文字」を証明するだけで、要約の正確さや事実の真偽は証明しない。
「可能性がある」を「必ず起こる」に変えた要約は、引用が合っていても意味レビューで修正する。

対象範囲の根は図中のノードとは別。範囲・タイトルを勝手に中心、成果、共通基盤へ変換しない。
列挙2〜4、フロー2〜5、拡散/合流/循環2〜4、背景1〜4の専用図の容量を守る。
中心・成果・基盤と各要素は全て明示IDと引用を必要とする。IDは順番由来にせず、並べ替えても維持する。
出典メタデータは固定した `raw.source` と一致させ、候補が別のURLを付け足して出典に見せることを拒否する。

## レビューとMECE

レビュー用の表示は、元本文・候補・各ノードの引用・関係グラフ・未対応の本文を並べて残す。
本文のカバー範囲には `nodeSources` と明示的な除外だけを使う。
要旨や関係に本文全体を引用しても、ノードに対応づけていない箇所は隠れない。
原文範囲の重複と意味上の分類の重複は別物。一文が複数ノードを説明することもある。

レビュー記録は原文と候補のハッシュ、`approve/request-changes`、実際のレビュー担当者、判断理由を持つ。
`actor.kind` は `human/assistant/fixture`。アシスタント確認をユーザー承認と記録しない。
記録は自己申告であり、本人認証や電子署名ではない。
全警告について `accepted-limitation` または `out-of-scope` と具体的理由を記録する。
曖昧さを残して採用する場合、その理由と限界を残す。未確認事項を事実に書き換えない。
候補のラベル、順番、理由、除外、出典などを変えれば古いレビューは無効になる。

承認後も `semanticMECE: unverified` とする。有限集合の検査は既存の構造レビューで別に行う。
コンパイル結果にはレビュー用フィールドを混ぜず、完全な記録は `.lineage.json` に残す。
コンパイルは可視化の前処理にすぎないため、続けて strict 生成と全ページ目視確認が必要。
長いメッセージ等で専用Patternが選ばれない場合もある。生成後の `.structure-review.json` の `fallbackUsed` を確認し、原文との対応と図の関係を再確認する。

## 既存AIとの接続

`src/ai/body-decomposition-advisor.ts` は既存の `AIProvider.generateStructured` を利用する任意のライブラリ入口。
プロバイダーがスキーマ検証を省いた実装でも、受信後に同じ検証を再実行する。
通常CLIの `replay` は `ScriptedAIProvider` に保存済み応答を渡すだけで、ネットワークを呼ばない。
外部AIを使う入口はCLIに追加していない。既存の `ai.structurePlanning` も今回有効化しない。
ライブAPIのモデル互換性、費用、秘密情報の送信、認証設定は別の検証・許可が必要。
既存 `StructureAdvisor` のAPIと通常の生成・Selector・Rendererは変更しない。

## 次の改善

実際の利用で、要約の断定化、根拠の取り違え、分解軸の不一致、未対応箇所、関係選択の修正量を記録する。
原文・候補・レビュー・生成物を一組の回帰例にする。意味上の合否は機械テストの通過と分ける。
次段階で入れ子の論理ツリーや候補比較UIを検討する。モデルなしで任意の文章を理解できるとは約束しない。
先に、数値根拠を加えても本文ノードを失わないPattern選択契約と、根拠自体の原文対応を追加する必要がある。
