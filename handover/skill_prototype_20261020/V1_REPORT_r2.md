# V1 修正報告（r2）｜スキル版 prototype｜運用日 2026-10-20（火）｜AIスライド 新環境 12往復目
> 判定63〜66 の修正。新フォルダ `handover/skill_prototype_20261020/` に配置（前日フォルダ `skill_prototype_20261018/` は無変更）。作業は作業領域 `/tmp/proto2/`。原本・gpts-package への書込みなし
> 🔴 V1 報告 r1 §3「ランク欄のはみ出し」は判定62 のとおり検査道具の誤検出（AIスライドの見落とし：段落内の文字の大きさの違い）

## ■ 1. 修正の一覧（前日版との差）
| ファイル | 修正 | 判定 |
|---|---|---|
| SKILL.md（2,660字｜前 2,464字） | ランクの5語の1行／URL が読めないときは推測で採点しない1行／点検の合格表示を購入者に見せる1行／「入力だけ直しても PPTX と食い違うと不合格」 | 63・65 |
| scripts/check_output.py | 総評の空・項目なし・末尾「…」を不合格／rank_label の5語点検／**PPTX の中の文字を読んで入力と照合**（C-3 総評・C-1 ランク欄の1・2枚目） | 63・64 |
| assets/sample_input_rankB.json | rank_label を「標準的」に（C-1・C-3） | 63 |
| references/（4ファイル新規） | scoring_rules.md・input_schema.md・dialogue_templates.md・diagram_patterns.md（前日の README_references.md は削除＝本フォルダに含めない） | 65 |
| README.md | 版の表記のみ | — |
| scripts/03_pptx_builder.py | **無改変**（287,389B・47629772eaff） | 58-2 |

### references/ の元にしたもの（lite・PLAIN の全文は入れていない）
| ファイル | 元 | 加工 |
|---|---|---|
| scoring_rules.md | PLAIN D・E・F・G・H・J-8 の要点 | 購入者向けに言い換え・社内の経緯や判定名・GPTs 固有の記述を削除 |
| input_schema.md | PLAIN J-3 のコメント・builder の LIMITS | 表にまとめ直し |
| dialogue_templates.md | PLAIN C（E1〜E4）・I の固定文面 | 固定文面は逐語。E1 の「クローラーを制限」を「内容を読み取れませんでした」に一般化。「📊 ダウンロード」の行を外し、点検の合格表示の行を追加 |
| diagram_patterns.md | 購入者向け抜粋版 r1（69cd7836a014）の本文（逐語・付録は除外） | 冒頭に「`/mnt/data` の読み替え」の注記1つを追加 |
- ⚠️ 統括に確認したい点：抜粋版 r1 は統括の判定（付録 B の確認事項1〜4）がまだ「起案」の扱いのままか。判定済みでなければ diagram_patterns.md もその判定に従う

## ■ 2. 試験結果
### 合格入力（rank_label＝標準的）
| 生成物 | B | md5 | verify_render |
|---|---:|---|---|
| 01_UI_Scorecard_sample_lp.pptx | 33,582 | f60bdf6e475e | NG 0（**ランク欄の NG も消えた**） |
| 02_Proposals_sample_lp.pptx | 34,160 | 6a6c6329fa18 | NG 0 |
| 03_Visual_Board_sample_lp.pptx | 39,744 | 6c8feb1d081a | NG 0（WRAP 3・WARN 2・OVER1 5｜前日と同じ） |
- check_output.py：⭕ 合格（終了コード0）／test_*.pptx：0件

### わざと崩した入力（すべて不合格・終了コード1）
| # | 試験 | 判定 | 結果 |
|:-:|---|---|---|
| t1 | 総評の末尾が「…」 | 64-1 | 🚨「総評の末尾が省略記号」 |
| t2 | 総評が空 | 64-2 | 🚨「総評が空」 |
| t3 | 総評の項目が無い | 64-3 | 🚨「総評の項目が無い」 |
| t4 | JSON の総評だけ書き換え（PPTX は元のまま） | 64-4 | 🚨「C-3 の PPTX に総評が見当たらない」 |
| t5 | rank_label＝標準的（改善余地あり） | 63 | 🚨 5語と違う＋PPTX と不一致 |
| t6 | rank_label＝優良（ランクB） | 63 | 🚨 5語と違う＋PPTX と不一致 |
| t7 | rank_label が無い | 63 | 🚨 項目が無い＋PPTX と不一致 |
| t8 | 総評18字 | （前回） | 🚨 18字＋PPTX と不一致 |
| t9 | JSON の rank_label を直したが C-1 を作り直さない | 64-4 | 🚨「C-1 の PPTX のランク欄が入力と一致しない」 |
- 前日の試験（件数・test_*.pptx・日本語ファイル名・リンク）も同じ処理のまま残している

## ■ 3. 統括に審査をお願いする点（差分のみ）
1. check_output.py の追加分（§2 の t1〜t9）
2. SKILL.md の追加3点
3. references/ の4ファイル（とくに dialogue_templates.md の E1 の一般化）
4. 手順書 r1 の追加5点（`v2_v6_procedure_draft_20261020_r1.md`）
