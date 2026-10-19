# 数値根拠を加えた本文の保持（2026-10-19）

PR #3 head `0d08c74730a86cb96a34d6b73bfb160d97447903` で再現。既存の六関係入力から並列2項目と時系列3工程を取り出し、仮の数値根拠を追加すると、big-number / chart-with-message が選ばれて本文が落ちた。入力は `examples/relationships/body-with-numeric-evidence.json`。値42・1/2/3は回帰用の人工値で、研究結果ではない。

数値専用の big-number / chart-with-message / before-after は全ての明示contentを受け取らない。以前のガードは拡散・背景・合流・循環だけを保護し、points/sequenceを除外していた。スコアを上げて競争させず、既存matchとbindings抽出の不適合判定を広げる。共通checkSelectableを使う手動指定・AI推薦も同じ判定を通る。bindings形式と既存保存Planの再生契約は変更しない。

数値だけを主役にする場合は、別のスライドでcontentを持たない根拠中心の入力を使う。本文付きページに必要な数値説明は既存description/explanationへ著者が明示する。根拠のすべてが自動で可視本文に出る機能ではない。raw-body compileのevidence=[]制限も解除しない。

## 検証

- Node 24.19.0 / npm 11.9.0。依存修正PR #4は別件のまま維持。
- 変更前 npm run check: 708件。新規回帰は変更前3件失敗/1件成功。変更後check: 712件、31ファイル、typecheck/build成功。
- 新規テストはpoints/sequence×数値3パターンの選択ゲートと直接bindings拒否、本文ラベル・説明のPPTX保持、根拠statement来歴、根拠のみの選択継続を確認。
- 既存structureテストが既知バグのbig-number選択を期待していたためthree-reasonsへ更新。分解入口の根拠拒否テストは保持。
- AI mode off / portable / strictで前後生成: NEC研究4頁、NEC解説8頁、六関係6頁、partition1頁、再現2頁。全て配置警告0。
- LibreOfficeDev 26.8.0.0.alpha0 / Poppler26.05.0 / Noto Sans CJK JP・Noto Sans。前後42ページのPNG生成、既存19頁の全ピクセル一致。変更2頁は並列2項目と受付→確認→承認の保持を目視。変更後全21頁も目視、切れ・重なりを認めなかった。
- 独立アシスタントが差分・回帰テストを確認。ユーザーや親レビュー担当の承認ではない。

## 範囲と未確認

これは上記の組込み数値専用3パターンへの限定修正。任意の外部Patternや全fallbackの無損失保証ではない。独立レビューで、sequence + claim_evidence + metric/textの2根拠という入力は既存three-reasonsの根拠列挙経路でsequenceを失い得ることが判明。数値図とは別の既存経路として残す。関係とcontentの整合や汎用無損失契約は今後の課題。

PowerPointネイティブ描画・親による実使用レビューは未実施。意味上のMECEは未検証。新たなニュース調査、外部AI/APIキー、有料API、mainマージは実施していない。

再現コマンド:

```sh
npm ci --cache /tmp/sdg-npm-cache
npm run check
npm run sdg -- generate examples/relationships/body-with-numeric-evidence.json --no-title-slide --ai-mode off --report --strict --theme portable
npm run preview -- examples/relationships/generated/body-with-numeric-evidence.pptx .preview/body-with-numeric-evidence
```
