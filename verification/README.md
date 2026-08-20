# Verification Kit（実機 PPTX での検証用）

PowerPoint で人が作った実際のスライドを使って、
**PPTX → Template / Pattern → 再生成** の往復で情報が失われないことを確認するための一式です。

```text
semantic-deck-v2-import-test.pptx   入力（実機で作成した検証用 PPTX）
  Slide 1                           表紙（Title Layout）
  Slide 2                           3 カード（背景 / アクセントバー / 区切り線 / 見出し / 数値）
  Master / Layout                   ロゴ・フッター・水平線・ページ番号

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
npm run sdg -- pattern import verification/semantic-deck-v2-import-test.pptx \
  --slide 2 --id test-three-column --out-dir verification/captured-patterns --scope company

# 3) 人間レビュー: pattern.json の metadata.status を "approved" にする

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

## 既知の未対応（後続 Phase）

Phase 1 は「客観的に読み取れる幾何・重なり順・種別を失わない」ことだけが対象です。
次は意味推定の課題として残っています。

- `01 / 02 / 03` のような連番ラベルが item として推定されない
- description が label 側にも入り二重表示になる
- 色を Theme トークンへ寄せる際の近似（青いアクセントが別トークンへ寄る等）
