# 主張を支える「採否の量と未知の品質内訳」

この限定機能は、全体から選ばれた量が分かっても、選ばれた／外れた対象の品質内訳が分からない場合を扱います。NECの「候補削減だけでは見逃しの少なさを判断できない」という1枚で検証しています。自動的な文章理解や一般的なMECE証明は行いません。

## 意味を先に指定する

従来の根拠付き `presentation.json` 経路に `evidence.type: "partition"` と `argument.kind: "partition-with-unknowns"` を追加しました。`structure compile` の evidence=[] 制限は維持します。そこへ数値を文字列として隠して迂回する使い方は対象外です。

- `whole` と `selected` は報告された値。全体は正数、部分は0以上かつ全体以下。
- `population` と `unit` は共有する一つの対象・単位。異なる母集団や単位の値を混ぜるフィールドはありません。宣言の正しさは出典レビューが必要です。
- `aggregation.kind` は single または mean。mean は `eventCount`、`eventUnit`、`wholeInvariant: true` を必須にします。全体は固定の母数、部分は複数事象の平均でよく、両者に同じ集計ラベルを強制しません。母数が変動する場合や中央値、重み付き集計は未対応です。
- 補集合と比率は `computePartition` の固定算術のみで導出。入力値・式・集計条件をbindingsに記録します。表示比率は小数1桁、帯幅は丸め前の比率です。任意式の評価はありません。
- `quality.selected` と `quality.remainder` は両方 `status: "unknown"`。数値、0、比率、再現率を与える入力は拒否します。未知を塗り面積に変換しません。
- `argument` は根拠ID、対象範囲、分類軸、推論、図が担う理解、3種類の注釈（quantity/quality/inference）の本文と必要理由を必須にします。推論・注釈は著者による解釈です。任意の文章が根拠から論理的に導けるかは機械判定しません。
- v1は一つのpartition evidenceに限定。既存contentやexplanation、追加evidenceとの混在は明示エラーです。別の関係図の本文を数値パターンが置き換えることを避けます。

専用patternの意味能力をSelectorで検査し、対応pattern不在やbindings失敗時は明示エラーにします。数値・未知内訳を落としたmessage-onlyへの退避はしません。従来入力のfallbackは変更しません。保存済みbindingsの計算値と来歴も再生時に再計算して検証します。

## 再現・レビュー・生成

```sh
npm ci
npm run check
npm run example:partition
npm run preview -- examples/arguments/generated/nec-candidate-quality-2026-10-16.pptx .preview/partition
```

新しい入力を作る場合:

```sh
npm run sdg -- argument-review deck.json --out review-packet.json
# packet内のreviewTemplateを別ファイルreview.jsonに取り出し、実際のレビュー主体と各findingを記入
npm run sdg -- generate deck.json --argument-review review.json --no-title-slide --ai-mode off --theme portable --report --strict
```

レビュー票は主張、補助文、数値、意味設計、出典を含む入力全体の正規化JSONのSHA256に結び付きます。一つでも変更したら旧レビューは拒否します。未承認テンプレートはdecision=request-changesで、生成側は承認を推定しません。actorは自己申告であり認証・電子署名ではありません。サンプルの記録はassistantによる実レビューで、ユーザー承認ではありません。

`--argument-review` を省略した通常生成は検討用の未レビュー出力を許可します。`--report` のstructure-reviewはunreviewedを明示します。記録済みの意味レビューを利用する生成では必ずこのオプションを指定してください。example:partitionはチェックイン済みのレビューを必ず読み、古い場合は出力前に止まります。意味レビューのハッシュはレンダラーのバージョンや見た目を保証しないため、コード変更後は改めて全ページを目視します。

## NECの数値と限界

一次出典: https://jpn.nec.com/press/202610/20261014_01.html

2026-10-16に発表本文・注記を再確認。固定母数5,035枚、12事象、平均候補544枚は報告値です。4,491枚は5,035−544の差分で、報告された平均値に基づく計算。候補割合約10.8%、候補削減約89.2%です。特定の検索における画像の所属一覧は取得していません。

正解数・再現率の数値未確認は、確認した発表本文・注記に限定します。図版、関連論文や他の資料に評価がないとは主張しません。NECは定性的に見逃し削減も説明しています。本図の主張は「候補件数の削減だけからは判断できない」に限定し、品質が悪い・実際に見逃したとは結論しません。

約7時間→約45分は1枚5秒という仮定による目視時間換算です。この1枚では主要な論拠ではないため数値を載せません。候補削減率を実工数削減、精度、ROIに読み替えません。導入時には、事象別の候補件数と見逃しを同じ評価画像で確かめる必要があります。

## 図と描画の検証

- 一つの帯だけが採否の数量を符号化します。下段の内訳はラベル付きの非因果参照線と文字で対応し、未知の正解割合を描きません。
- 出典リンクはRenderContext経由のネイティブ文字リンク。HTTP(S)のみ許可し、自動取得しません。PPTXとLibreOffice PDFの実リンクを検査しています。
- 部分が0または全体に一致する場合、空側は「対象がないため内訳なし」と表示します。未知品質を数値0へ代入したり、再現率を計算したりはしません。
- 元の意味入力、注釈の必要理由、数式と入力値はVisual Planとstructure-reviewから追跡できます。
- 長文は動的高さ測定で明示拒否し、切り捨てや極端な縮小はしません。標準以外のテーマやテンプレートでは再描画・目視が必要です。
- portableテーマはNoto Sans CJK JP/Noto Sansに依存します。PDFは確認時の固定表示。編集可能PPTXはOS・フォントに依存し、ネイティブPowerPointでの目視は未確認です。

## 今回の検証と次の改善

回帰テストは入力範囲・集計条件・未知への数値混入・出典参照・全意味のbindings保持・計算改ざん・対応外pattern・overflow・レビュー陳腐化・ネイティブオブジェクトと出典リンクを検査します。機械検査と独立した読者レビューは別です。

検証記録は `examples/arguments/feedback-2026-10-16.json`。次の候補は、有限な正解数を得た際の事象別評価、一般のcontent+数値根拠で本文を必ず保持する選択契約、および図中の関係を対象にした注釈です。新たなグラフDSLや自動フレームワーク選択はこの変更に含めません。
