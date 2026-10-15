# 生の調査メモから分解候補を検査する例

- `nec-raw.json`：NEC公式発表に基づく短い調査メモ。本文は逐語引用ではない。
- `nec-candidate.json`：アシスタントが本文を読んで作った分解候補。プログラムの自動推論結果ではない。
- `nec-assistant-review.json`：候補と本文を確認したアシスタントのレビュー。ユーザー承認ではない。

```sh
npm run example:decomposition
npm run preview -- examples/decomposition/generated/presentation.pptx .preview/decomposition
```

コマンドは保存済みの原文・候補・レビューをオフラインで検査して再生する。新しい承認を作らず、原文や候補を変更するとレビューのハッシュが一致せず停止する。
`generated/sample-review-packet.json` で元本文、各要素への引用範囲、関係の理由、未対応範囲を確認できる。
`generated/presentation.lineage.json` はその判断履歴を保持する。

本文にある「新たに学習する被災事象のデータ」と「更新前モデルが生成した多様な説明文」を、継続学習に入る二つの情報として整理した。
対象範囲・タイトルを合流先に転用せず、本文にある継続学習を別ノードとして引用している。
接続語や述語には機械的な未対応範囲が残るため、レビュアーが関係の表現として確認し、理由を記録する。
学習の詳細や精度は未確認のまま。意味上のMECEや、製品提供開始、実測効果の証明ではない。

新しい本文を扱う手順と制約は [body-decomposition.md](../../docs/body-decomposition.md) を参照。
PDF/PNGの確認にはLibreOffice、PopplerとNotoフォントが必要。PPTXのOS・フォント依存は残る。
