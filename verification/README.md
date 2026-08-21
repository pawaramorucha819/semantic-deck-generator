# Verification Kit（実機 PPTX での検証用）

PowerPoint で人が作った実際のスライドを使って、
**PPTX → Template / Pattern → 再生成** の往復で情報が失われないことを確認するための一式です。

```text
semantic-deck-v2-import-test.pptx   入力（実機で作成した検証用 PPTX）
  Slide 1                           表紙（Title Layout）
  Slide 2                           3 カード（背景 / アクセントバー / 区切り線 / 見出し / 数値）
  Master / Layout                   ロゴ・フッター・水平線・ページ番号

interpretation-slide2.json          Slide 2 の要素の意味を人が指定した解釈
reuse-presentation.json             取り込んだ Pattern を別内容で再利用する意味構造
reuse-generation.config.json        再利用時の生成設定
templates/verification-source/      Template Import の結果（再生成可能）
captured-patterns/test-three-column/ Pattern Capture の結果（人間レビュー後 approved）
knowledge/                          knowledge ingest の結果
generated/                          生成物（Git 管理外）
```

## 再現手順

```bash
# 1) 既存 PPTX から企業テンプレートを取り込む
npm run sdg -- template import verification/semantic-deck-v2-import-test.pptx \
  --id verification-source --out-dir verification/templates \
  --organization "SDG V2 検証用サンプル" --confidential "CONFIDENTIAL"

# 2) スライド 2 を Pattern として取り込む（出力は必ず draft）
#    --interpretation を付けると、各要素の意味を人が指定した解釈で取り込む
npm run sdg -- pattern import verification/semantic-deck-v2-import-test.pptx \
  --slide 2 --id test-three-column --out-dir verification/captured-patterns --scope company \
  --interpretation verification/interpretation-slide2.json

# 3) 人間レビュー: capture-report.json の inventory を確認し、
#    pattern.json の metadata.status を "approved" にする

# 4) 同じ Template 上で別内容へ再利用する
npm run sdg -- generate verification/reuse-presentation.json \
  --config verification/reuse-generation.config.json
```

## 自動テスト

`tests/verification-e2e.test.ts` がこの PPTX を固定資産として使い、次を回帰として固定しています。

- Content Master にフッターの水平線（h = 0）が残ること
- Slide 2 の区切り線（h = 0）が Pattern に残ること
- 元 PPTX の重なり順（カード背景 → アクセント → 区切り線）が保たれること
- 再利用生成がレイアウトエラーなく完了すること

Phase 2 で追加した回帰:

- 解釈 JSON の各要素が役割どおり（index / label / metric / description / accent / divider）に入ること
- `01` が literal として焼き付かず、再利用先の項目順から `01 / 02 / 03` が作り直されること
- description がカードごとに 1 回だけ描かれること
- capture-report.json に inventory と採用された解釈が残ること
- 解釈を渡さない従来の取り込みが変わらず動くこと（後方互換）

## 解釈 JSON（interpretation-slide2.json）

`capture-report.json` の inventory を見て人が書いた、Slide 2 の要素の意味です。

| 要素 | 役割 |
| --- | --- |
| 0 / 1 / 2 | eyebrow / headline / subline |
| 3, 10, 17 | カード背景（background） |
| 4, 11, 18 | アクセントバー（accent） |
| 5, 12, 19 | 連番（index）→ 順番から作り直される |
| 6, 13, 20 | 見出し（label） |
| 7, 14, 21 | 数値（metric） |
| 8, 15, 22 | 区切り線（divider, h = 0） |
| 9, 16, 23 | 説明（description） |
| 24 | 出典（sourceText） |
| 25 | ページ番号プレースホルダ（Template の責務なので未割り当て） |

## 既知の未対応（後続 Phase）

Phase 1 は「客観的に読み取れる幾何・重なり順・種別を失わない」こと、
Phase 2 は「正しい解釈を与えれば、そのとおりの Pattern が出てくる」ことが対象です。
**解釈そのものの精度向上（heuristic の改善 / AI 解釈）は Phase 3 です。**

既定のヒューリスティクスは以下をまだ判定できません（`--interpretation` を使えば回避できます）。

- `01 / 02 / 03` のような連番を index として認識しない
- 見出しと説明の区別が付かず、両方が description に入る
- accent / divider を装飾として認識せず、固定デザイン扱いになる
- 色を Theme トークンへ寄せる際の近似（青いアクセントが別トークンへ寄る等）
