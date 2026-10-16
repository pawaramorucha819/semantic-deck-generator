# 一次情報から作る日次リサーチ資料

SE・PM・ITコンサルタントを目指す読者向けのサンプルです。4枚の初期版に加え、補足説明なしで読む8枚の解説版を追加しました。NECの[2026年10月14日の公式発表](https://jpn.nec.com/press/202610/20261014_01.html)を2026年10月15日に確認しました。図式を網羅するために事実を加工せず、並列・合流・工程のみを使います。

## 解説版を使う

`npm run example:explanatory` で、`nec-explained-2026-10-15.json` から8ページを生成します。初出語・仕組み・条件・限界を可視の本文に置き、最後に一次URLを示します。詳細は [説明型資料の基準](../../docs/standalone-explanatory-decks.md) を参照してください。

## ファイル

- `nec-disaster-vlm-2026-10-14.json`：意味・根拠・出典のみを持つ本文4枚。座標・配色・Pattern指定は含みません。
- `review-feedback.template.json`：レビュー時にコピーして記入するJSON。機械検査と人の評価を分けた記録用で、生成CLIが読み込む設定ファイルではありません。

全スライドで `supportingMessage` に事実・試算・執筆者の提案を明記しています。根拠は `evidenceRefs` で追跡でき、出典のURL・公開日・参照箇所・確認日と詳しい注意書きはPowerPointのスピーカーノートに残ります。表紙は不要です。

```bash
npm run sdg -- generate examples/research/nec-disaster-vlm-2026-10-14.json \
  --no-title-slide --ai-mode off --report --strict --theme portable \
  --out-dir examples/research/generated
```

## 数値を読むときの境界

- 検索対象の5,035枚は、東日本大震災の記録映像から抽出した車載画像です。
- 確認候補の544枚は、12種類の被災事象についての平均です。正解画像数ではありません。
- 約7時間→約45分は1枚5秒の目視換算です。実運用の総時間や金銭的ROIの実測ではありません。
- 再現率の数値はこの発表の本文・注記にありません。関連論文全体への断定はしていません。
- 実用化は2027年度中という目標です。
- 最終スライドはSE・PM向けの導入提案で、NECが実施済みの業務手順ではありません。

## 毎日の「調査→資料→レビュー」

1. **調査**：一次情報を読み、公開日・確認日・該当箇所を保存。出典事実、計算条件、未記載事項、自分の解釈を分けます。
2. **構造化**：1枚1メッセージで本文を作成。分解の対象範囲・基準・理由・除外・未解決点を `decomposition` に記録します。図は情報の関係に合わせます。
3. **生成・機械検査**：スキーマと参照、ラベル重複、関係の向き、fallback、Render Reportを確認。生成が成功しただけでは品質合格にしません。
4. **人手レビュー**：PowerPointを開くか各ページを画像化して、下記を確認。テンプレートを日付付きでコピーし、スライドID・修正内容・理由を記録します。
5. **修正・再確認**：事実や論点の修正は入力JSONへ、生成方法の変更は設定へ戻します。再生成し、指摘を直したページと関係するページを再確認します。

### 短い目視チェックリスト

- [ ] 見出しだけで主張が伝わり、本文と注記が投影時にも読める
- [ ] 「平均544枚」「12事象」「1枚5秒」が紙面に残り、約45分を実測ROIと誤解しない
- [ ] 並列・合流・工程の違いが図から分かり、出典にない因果・循環・全体網羅性を足していない
- [ ] 分解基準に対する意味の重複・抜けを人が確認し、許容した重複と対象外を説明できる
- [ ] 出典事実・計算条件・SE／PMの提案が区別でき、ノートから原文と計算を追える
- [ ] はみ出し、重なり、細かすぎる文字、切れた注記がなく、修正後も再確認した

### 機械検査と意味上のMECEは別

スキーマ適合、参照整合、重複ラベルがないこと、入力に沿った矢印は機械で確認できます。語義上の重複や範囲の網羅性は、それだけでは証明できません。`decomposition` も執筆者による説明であり、自動MECE認証ではありません。

この例では、画像の元集合と確認候補は排他的ではなく、検索品質と運用負荷の評価も関連します。意味が重なる観点を無理に排他的と呼ばず、対象範囲と限界をレビュー記録に残してください。レビュー前のテンプレートは `semanticMECE: "unverified"` とし、機械検査が通っても自動的に合格へ変更しません。

## フォントと確認環境

サンプルは `portable` テーマ（Noto Sans CJK JP / Noto Sans）を使用します。編集・再描画する環境にもこのフォントが必要です。添付PDFは確認済み表示ですが、PPTXはOSとフォントにより改行や表示が変わります。今回の目視確認はLibreOfficeで行い、PowerPointネイティブ表示は未確認です。

## GitHub AI Scan / Agentic autofix 解説（2026-10-16）

`github-security-explained-2026-10-16.json` は、GitHubの2026年10月15日の一般提供発表と公式Docsを使った8ページの解説です。検出と修正を別の入口として示し、用語・仕組み・対象条件・課金・限界・提案を可視本文に置きました。二つの発表URLと主要DocsのURLはPDFでも読めます。

```sh
npm run sdg -- generate examples/research/github-security-explained-2026-10-16.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir examples/research/generated
npm run preview -- examples/research/generated/github-security-explained-2026-10-16.pptx .preview/github-security-explained-2026-10-16
```

`review-github-security-2026-10-16.json` は実施したアシスタント読者レビューと修正の記録です。ユーザーの品質承認ではありません。PDF内のDocs出典不足と組織単位の有効化範囲の説明不足を修正し、全8ページの再確認を通過しました。`github-security-manifest-2026-10-16.json` に生成リビジョン・入力ハッシュ・再生成コマンドを記録しています。

Strict生成は配置警告0・fallbackなし。LibreOffice出力を確認し、PowerPointネイティブ表示は未確認です。PDFのURLはコピー可能ですが、クリック用注釈はありません。AI機能の有効化・実測評価は行っていません。生成物は既存のignore対象に保存します。

## オブジェクトと説明を対応させる改良版

`github-security-linked-2026-10-16.json` は上記8ページの内容を10ページに再構成し、同じ番号・名称で図と説明を対応づけた版です。`nodeRefs` は意味を確認して明示指定します。費用と長い出典は別ページに分けました。発表日10月15日と確認日10月16日（JST）を区別します。

```sh
npm run sdg -- generate examples/research/github-security-linked-2026-10-16.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir examples/research/generated
npm run preview -- examples/research/generated/github-security-linked-2026-10-16.pptx .preview/linked
```

`review-github-security-linked-2026-10-16.json` はアシスタント検証の記録です。利用機能の有効化や性能の実測は行っていません。旧版は比較のため残しています。


## GPT-LiveとDecisions：判断と操作完了（2026-10-16）

`live-decisions-boundary-2026-10-16.json` は2枚の資料です。1枚目はnative delegationの判断・操作・実状態確認の責任を工程図で示し、2枚目は可逆操作を使う未実施の検証案を示します。GPT-Live-1.1のAPI提供済みや速度改善率は主張しません。

```sh
npm run sdg -- generate examples/research/live-decisions-boundary-2026-10-16.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir examples/research/generated
npm run preview -- examples/research/generated/live-decisions-boundary-2026-10-16.pptx .preview/live-decisions-boundary-2026-10-16
```

`review-live-decisions-boundary-2026-10-16.json` に実際の修正と独立したアシスタント読者レビュー、`live-decisions-boundary-manifest-2026-10-16.json` に生成head・ハッシュ・再現条件を記録しています。コード追加なし。strict警告0・fallbackなし、全2枚をLibreOffice/Popplerで確認しました。PowerPointネイティブ表示とAPI操作試験は未実施。ユーザーの品質承認や意味上のMECE完全性は保証しません。PDF出典はコピー可能な文字で、クリック用注釈はありません。生成物は既存のignore対象に保存します。


## OSSライセンス方針の事前照会とマージ条件（2026-10-17）

`license-gate-2026-10-17.json` は、10月16日一般提供のライセンス適合機能を2枚で説明します。共通方針を依存追加前のAPI照会とPR差分検査へつなぐ運用案、Evaluate / Activeと既存ブランチ保護の違いを示します。API組込みは提案で、試験や有効化はしていません。検出範囲と法的保証の限界を紙面に残します。

```sh
npm run sdg -- generate examples/research/license-gate-2026-10-17.json --no-title-slide --ai-mode off --report --strict --theme portable --out-dir examples/research/generated
npm run preview -- examples/research/generated/license-gate-2026-10-17.pptx .preview/license-gate-2026-10-17
```

`license-gate-manifest-2026-10-17.json` と `review-license-gate-2026-10-17.json` に生成head・再現条件・独立アシスタント評価を記録。コード変更なし、708テスト、strict配置警告0。全2頁をLibreOfficeで目視しました。意味上のMECE・PowerPointネイティブ表示・API実試験は未検証。PDF出典は文字として表示し、クリック用注釈はありません。
