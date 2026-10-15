# ボディを分解し、関係を保ったまま図にする

## このマイルストーンの範囲

既存の「意味 → 表現 → 決定論的描画」を保ち、単一階層のボディ分解と六つの関係を扱う。
自然文を入力するだけで分解が正しいと保証する機能や、入れ子の論理ツリーは未実装。
人が対象範囲・分解軸・関係を指定し、その構造を検査可能な PPTX にする段階である。

1. 何を説明するかを `decomposition.scope` に書く。
2. 兄弟要素を分ける一つの基準を `criterion`、分け方の理由を `rationale` に書く。
3. 対象外は `exclusions`、未確認事項は `unresolvedQuestions` に残す。
4. 各要素を `content`、裏付ける事実を `evidence` に分ける。説明だけなら `evidence: []` でよい。
5. 要素間の関係を宣言する。図の見た目から関係を逆算しない。
6. 生成後に構造レビュー、実際の描画、人による意味のレビューを行う。

## 関係と入力の対応

| 情報関係 | 意味モデル | 論理関係 | 選ばれる Pattern |
|---|---|---|---|
| 列挙 | `points.items` | `claim_evidence` など | `three-reasons` |
| 放射 | `distribution.center` と `items` | `distribution` | `radial` |
| フロー | `sequence.steps` | `process` / `timeline` | `process` |
| 背景 | `context.environment` と `items` | `context` | `background` |
| 合流 | `convergence.inputs` と `outcome` | `convergence` | `convergence` |
| 循環 | `cycle.steps`（末尾から先頭へ戻る） | `cycle` | `cycle` |

背景の `context` は共通の環境・基盤という**ボディの関係**であり、スライド全体の役割 `slideRole: context` とは別。
放射の送り先、合流の入力、循環の工程は2〜4件。背景の要素は1〜4件。
列挙の専用 Pattern は2〜4件、フローは2〜5工程。既存のより大きい入力は fallback になる場合がある。
新しい関係の上限超過はスキーマエラー。長文は切り詰めず保持し、`--strict` で配置警告も失敗にする。
汎用の表・複数ページへの自動分割・入れ子の混合関係は今後の課題。

## MECE の確認範囲

各 ContentItem は任意の `id` を持てる。レビュー結果と保存済み bindings の追跡には明示 ID を推奨。
ID がない場合、構造レビューは `items.0` など入力内のパスを識別子として使う。
これらのパスは並べ替えで変わる。ID の重複と未定義 evidenceRefs は生成前に拒否する。

対象が有限集合なら、`decomposition.universe` と兄弟要素ごとの `members` を指定できる。

```json
{
  "decomposition": {
    "scope": "今回の対象機能",
    "criterion": "主担当の部門",
    "rationale": "各機能に一つの主担当を割り当てる",
    "universe": ["検索", "申請", "承認"]
  },
  "content": {
    "kind": "points",
    "items": [
      {"id": "operations", "label": "運用部門", "members": ["検索"]},
      {"id": "business", "label": "業務部門", "members": ["申請", "承認"]}
    ]
  }
}
```

構造レポートは重複するメンバーとノード ID、未分類のメンバー、範囲外のメンバーを個別に報告する。
全条件を満たしても `declared-sets-pass` は「宣言された集合の分割が成立した」だけを意味する。
対象集合自体が正しいか、同じ基準で分解しているか、重要な情報が漏れていないかは人の判断が必要。
集合または所属がない場合は `not-verifiable`、`semanticMECE` は常に `unverified`。
中心・結果・基盤は兄弟要素ではないため、集合の重複検査には混ぜない。

## 実行とレビュー

```bash
npm ci
npm run check
npm run example:relationships
npm run example:research
npm run preview -- examples/relationships/generated/six-relationships.pptx .preview/relationships
npm run preview -- examples/research/generated/nec-disaster-vlm-2026-10-14.pptx .preview/research
```

生成は Node.js 20以上だけで動く。AIキーは不要。プレビューのみ LibreOffice (`soffice`) と Poppler (`pdftoppm`) が必要。サンプルは `portable` テーマを使い、Noto Sans CJK JP / Noto Sans のインストールが必要。既存の `default` と `midnight` テーマは変更しない。
Windows + PowerPoint では既存の `scripts/preview-pptx.ps1` も使える。
このマイルストーンの画像確認は LibreOffice 出力で行った。PowerPoint ネイティブ表示は別途確認が必要。

`--report` は以下を出力する。

- `.pptx`: 編集可能な図形・テキスト。根拠全文と URL はスピーカーノート
- `.visual-plan.json`: Pattern選択理由、検証済みbindings、出典
- `.render-report.json`: 配置・重なり・文字の収まり
- `.structure-review.json`: 元の入力、分解軸、ノードと関係、宣言集合の検査、未確認事項

構造レビューのノードと、描画された要素を一つずつ照合する。特に矢印の向き、循環の戻り、基盤の所属を確認する。
`--strict` の成功は意味の正しさも実画面の見やすさも証明しない。

## 日々の調査から改善する

`examples/research/` は実在する発表を使った資料、`examples/relationships/` は人工的な関係図テスト。
調査結果に不自然な関係を足して六種類を使い切らない。

毎回、入力JSONと出典日・取得日を残し、生成・構造チェック・全スライドの目視確認を行う。
フィードバックには関係の忠実性、漏れ・重複の疑い、レイアウト不具合、修正にかかった時間を記録する。
確認できなかった項目は空欄または未確認とし、数値を創作しない。
再現可能な不具合は小さな入力と期待する図の関係を添え、テストと修正に戻す。
このPRでは定期実行・自動公開・自動マージを設定しない。

## 学びの反映と未対応

要件と実装の対応は `learning-coverage.json` に記録する。
第1段階は単一階層の分解を実装した。親子ツリー、任意のエッジ、複合図、強調対象、表/ガント/四象限、五つのグラフ分類は次段階。
任意エッジを入力できないため、ダングリングエッジや非連結な循環を表現する経路は現在ない。
単純な循環は `steps` の順番と末尾から先頭への戻りで定義する。

既存の数値図については、無順序カテゴリから変化率を作らない、異単位を前後比較しない、相関や構成を誤って棒グラフにしない最低限のガードのみ加えた。
