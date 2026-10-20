# r4 の実施記録（判定77〜80）｜運用日 2026-10-21（水）｜AIスライド 新環境 16往復目
> 作業領域 `/tmp/proto4/`。r3（`skill_prototype_20261020_r3/`・zip e82ba6f995ba）は無変更で保存（判定77）。builder は無改変（47629772eaff）

## ■ 1. 直した範囲（判定78）
| ファイル | 修正 |
|---|---|
| references/input_schema.md | highlights を `{title, target_area, priority, effort, before, after}` に／「target_area・priority・effort は C-2 の対応する改善提案と同じ値」／target_area 20字／`flow_summary` を**必須**で追加（§2）／「area ではなく target_area」を明記 |
| assets/sample_input_rankB.json | highlights を C-2 の提案1〜3 と一致（ファーストビュー・高・小／申込フォーム・高・小／メインコピー・中・小）／flow_summary「FVとフォームで行動が止まり申込に届かない」（21字） |
| scripts/check_output.py | 点検7 を追加：各件の target_area（area のときは専用の文面）・PPTX に「箇所：{target_area}」・priority（高中低）・effort（小中大）／flow_summary（必須・40字以内・PPTX に「→ {まとめ}」） |
| SKILL.md・README・references 3件 | 先頭の版表記のみ（r3 → r4） |

## ■ 2. flow_summary の過去の出力の確認（判定78-2）
| 過去の GPTs の出力（7月のドッグフーディング `_reference/v3.4_dogfooding_20260720/output/`） | 行動フローの見出しの右の「→ まとめ」 |
|---|---|
| case01_lp_diagnosis | なし |
| case02_banner_diagnosis | **あり**「→ CTAで問い合わせ理由を補強」 |
| case03_quick_score | **あり**「→ FV・料金後のCTA不足が主な離脱点」 |
| case04_fv_wireframe | なし |
| case05_designer_brief | **あり**「→ 問い合わせ前の不安解消とCTA増加が鍵」 |
- 診断サンプルの元データ `SAMPLES_DATA_20260716.md` にも3か所（150・307・465行）
- → **GPTs の出力にあった欄**と確認。判定78-2 に従い**必須に上げ、点検にも入れた**（無い・空・40字超・PPTX に無い → 不合格）

## ■ 3. 試験（判定78-5）
| 試験 | 結果 |
|---|---|
| u1 r3 の見本（area のまま・flow_summary なし）で生成 → 点検 | 🚨 不合格10件（「area ではなく target_area」×3・priority×3・effort×3・flow_summary） |
| u2 r4 の見本 | ⭕ 合格（終了コード0） |
| u3 PPTX は r3 の見本で作ったまま・JSON だけ target_area 等を足す | 🚨 不合格4件（PPTX に「箇所：…」×3・まとめ） |
| 追加：priority＝最優先／flow_summary なし／41字／target_area なし | 🚨 いずれも不合格 |
| t1〜t11（r4 の見本を基準に作り直し） | すべて不合格（終了コード1）。test_*.pptx の混入も不合格 |
| verify_render（r4 の見本の生成物） | C-1 NG 0／C-2 NG 0／C-3 NG 0（WRAP 3・WARN 2・OVER1 5） |
| 生成物 | 01 33,582B（962714929c37）／02 34,160B（51700fb0c6e6）／03 **39,851B**（060e0ebcdb57｜r3 は 39,744B＝箇所・優先度・まとめの文字が増えた分） |
| 3枚目の表示 | 「高・工数：小・箇所：ファーストビュー」「高・工数：小・箇所：申込フォーム」「中・工数：小・箇所：メインコピー」／1枚目「→ FVとフォームで行動が止まり申込に届かない」 |
- 判定73 の方式（判定78-6）：すべての試験がとおったため、統括の再審査を待たずに zip を作成（統括は事後確認）

## ■ 4. zip（r4）
| 項目 | 値 |
|---|---|
| ファイル | `ui-diagnosis-director_skill_prototype_r4_20261021.zip` |
| B／md5 | 101,828B・`e8677c8c772f` |
| 形 | いちばん上に `ui-diagnosis-director/`（判定72） |
| 展開しての確認 | 読込み（版 17.2.17）→ 生成 → 点検合格 ⭕ |

| ファイル | B | md5 |
|---|---:|---|
| `ui-diagnosis-director/README.md` | 1149 | `4ccdf4426788` |
| `ui-diagnosis-director/SKILL.md` | 5398 | `45d345ba4566` |
| `ui-diagnosis-director/assets/sample_input_rankB.json` | 8997 | `6867e42ff10f` |
| `ui-diagnosis-director/references/diagram_patterns.md` | 13832 | `a96600a6536e` |
| `ui-diagnosis-director/references/dialogue_templates.md` | 5414 | `280131f2fb09` |
| `ui-diagnosis-director/references/input_schema.md` | 3407 | `b3ffc69a71fb` |
| `ui-diagnosis-director/references/scoring_rules.md` | 4502 | `db7a6a2da0d9` |
| `ui-diagnosis-director/scripts/03_pptx_builder.py` | 287389 | `47629772eaff` |
| `ui-diagnosis-director/scripts/check_output.py` | 10783 | `9af2fdecb7d2` |
