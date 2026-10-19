# 工程を根拠カードへ置換しない（2026-10-19）

main `a818355e486741248e81b6aa0999321189af2d2b` で、sequence + claim_evidence + metric/text各1件を生成すると、three-reasons が根拠だけを描画して受付→確認→承認の工程を失った。strict配置検査は成功するため、配置だけでは発見できない内容欠落だった。

## 限定した変更

- three-reasons の素材抽出は、contentがある場合pointsだけを許す。sequenceを根拠の列挙へ置換しない。match／checkSelectable／直接buildBindingsの各経路で拒否する。
- sequence + claim_evidence は Selector の入口で生成を停止する。スライドIDを示し、工程ならprocess、時系列ならtimelineを著者が指定するか、根拠提示と工程を別ページに分けるよう案内する。通常生成とmanual/assist/autoは同じ入口を通る。根拠件数に依存させず、この組合せは根拠なしでも停止する。
- 不整合をfallbackに隠さず、著者の入力・関係・根拠を変更しない。CLIでは終了コード1、既存出力を上書きしない。過去に生成されたファイルの削除もしないため、終了コードを確認せず古い出力を新しい成功結果と扱わないこと。
- スキーマは元入力を読めるまま維持する。保存済みVisual Planの再生契約・bindings形式・依存は変更しない。既に工程が欠落したPlanは自動復元されないので、元入力を著者が見直して再生成する。
- 他の全content/関係の適合性や外部Patternの無損失性を保証する変更ではない。

## 再現と正常例

`examples/relationships/sequence-claim-evidence.invalid.json` はスキーマ上読めるが生成時に拒否すべき人工例。42件は実測ではない。個人情報や外部APIは含まない。

```sh
# 期待: 終了コード1、スライドID・process/timeline・ページ分割の案内。新規出力なし。
npm run sdg -- generate examples/relationships/sequence-claim-evidence.invalid.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir .preview/sequence-rejected

# 著者がprocess/timelineを明示した2ページ。元の工程・説明・根拠は同じ。
npm run sdg -- generate examples/relationships/sequence-relations.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir .preview/sequence-generated
npm run preview -- .preview/sequence-generated/sequence-relations.pptx .preview/sequence-rendered
```

正常例では全工程ラベル・説明・順序が紙面に残り、根拠statementはノートに残る。全ての数値根拠を可視本文へ自動配置する機能ではない。読者に必要な根拠は著者がdescription/explanationや別ページに明示する。

## 実施検証

- Node24.19.0、npm11.9.0、LibreOfficeDev26.8.0.0.alpha0、Poppler26.05.0、Noto Sans CJK JP/Noto Sans。
- 新規5回帰: 変更前3失敗/2成功、変更後5成功。通常/手動/AIモードの拒否、入力不変、出力なし、正常関係で全工程描画・全根拠statementのノート保持を検査。
- npm run check: 717テスト/32ファイル・型検査・build成功。
- AI mode off・portable・strictで正常例2頁、前回数値回帰2頁、NEC研究4頁、NEC解説8頁、六関係6頁、partition1頁を生成。全23頁で配置警告0、PDF/PNGを生成して全ページを目視。新規正常例の工程・説明に切れや重なりを認めなかった。
- 既存21頁は前回検証画像と全ピクセル一致。修正前の欠落例1頁も生成して目視し、工程が根拠2カードへ置換されることを確認。
- 独立アシスタントがコードと専用5回帰を検査。テスト内で部分文字列「確認」がタイトルへ誤一致する点を指摘し、XMLテキスト要素の完全一致による順序検査へ修正後、5件成功を再確認。実装の重大指摘なし。ユーザー承認ではない。

親による今回PRのレビュー・PowerPointネイティブ表示は未確認。mainマージ待ち。PR #4の依存変更は含まない。

## 次の実使用入力（親からの報告、当方では未再現）

親がIEA調査を4頁化。合流図の各入力に2文を入れると必要高1.05in/枠0.66inでstrictが拒否。入力ラベルを短くし、同じ順の根拠説明を次頁へ分離して解消、全4頁を親が目視済みとの報告。34地域/約90%は調査範囲であり効果値ではないと明記、事実と助手示唆も可視で分離したとのこと。

確定Library JSONは未受領。資料・原典条件を当方が検証したとは扱わず、受領後にexamples/research規約とレビュー票に沿って保存・再現する。今回のエラー回避のために文字を自動削除したり関係を勝手に変更したりしない。
