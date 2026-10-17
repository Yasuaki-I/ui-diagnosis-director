# V1 結果報告｜スキル版 prototype｜運用日 2026-10-18（日）｜AIスライド 新環境 9往復目
> 判定61 の報告項目どおり。本フォルダは試作（prototype）であり配布物ではない。本物の GPTs・販売ページ・掲載一式には触れていない
> 作業は作業領域 `/tmp/proto/` で行い、結果を本フォルダに write5.py（新規）で配置。原本・gpts-package への書込みなし

## ■ 1. 構成（ui-diagnosis-director/）
| ファイル | 中身 | 状態 |
|---|---|---|
| SKILL.md | 3ステップ・採点・builder 読込み3段・点検・やらないこと（2,464字） | 起案（統括審査待ち） |
| README.md | 導入・使い方の骨子（1番目＝新しいチャット／2番目＝名前で呼び出す） | 骨子 |
| scripts/03_pptx_builder.py | 採用版の写し **287,389B・47629772eaff（一致）**・無改変 | — |
| scripts/check_output.py | 点検スクリプト（新規） | 起案（統括審査待ち） |
| references/README_references.md | references/ の予定4ファイルの一覧（中身は審査後） | 骨子 |
| assets/sample_input_rankB.json | V1 用の入力（ランクB・架空のサンプル） | 試験用 |
- 再配布禁止の表記：SKILL.md・README・check_output.py・references の先頭に入れた。**builder 本体には入れていない**（判定58-7）

## ■ 2. V1 の結果
| 項目 | 結果 |
|---|---|
| builder の読込み（3段・2行方式） | ⭕「builder：…/scripts/03_pptx_builder.py｜版 17.2.17」と報告して使用 |
| PPTX 3点 | ⭕ 下表 |
| test_*.pptx | ⭕ 0件（作業領域全体を検索） |
| check_output.py（合格入力） | ⭕ 合格（総評16字・件数3・ファイル名・リンク無し・メイリオのみ・最小14.0pt） |
| verify_render.py | C-2 ⭕ NG 0／C-3 ⭕ NG 0（WRAP 3・WARN 2・OVER1 5）／**C-1 🚨 NG 2**（§3） |

| 生成物 | B | md5 | スライド | フォント | 最小 |
|---|---:|---|:-:|---|---:|
| 01_UI_Scorecard_sample_lp.pptx | 33,612 | 286b8dfa658c | 2 | メイリオ | 14.0pt |
| 02_Proposals_sample_lp.pptx | 34,160 | bad99129d7df | 2 | メイリオ | 14.0pt |
| 03_Visual_Board_sample_lp.pptx | 39,744 | 71b5de8ded15 | 3 | メイリオ | 14.0pt |

### check_output.py の不合格の検出（わざと崩した入力）
| 試験 | 結果 |
|---|---|
| 総評18字「FV内CTA追加で無料体験へ誘導する」 | 🚨 不合格「総評が18字（上限17字）」⭕ 検出。**builder はこの18字を止めずに生成した**（builder の上限は80字のため）＝点検スクリプトでしか止まらない |
| ランクC に強み・課題3件 | 🚨 不合格5件（C-1 強み・課題／C-3 強み・課題／top_issues）⭕ 検出 |
| test_output.pptx・日本語ファイル名・ハイパーリンクを混入 | 🚨 不合格3件 ⭕ すべて検出 |

## ■ 3. 🔴 発見：C-1 の rank_label が長いと NG（builder 本体の既知の性質の可能性）
- 入力 `rank_label`＝「標準的（改善余地あり）」（**builder のデモと同じ文字列**）で、C-1 のランク欄（幅66px）が NG 2件
- 切り分け：「標準的」→ NG 0／空 → NG 0／「改善余地あり」「標準的な水準」「標準的（改善余地）」→ NG 2
- lite・PLAIN に rank_label の字数規定は**見当たらない**（grep で0件）。GPTs 経路でも同じ入力なら同じになるはず（builder は同一）
- 対処の案（判定はしない）：①SKILL.md・references に「rank_label は3字程度（例：標準的）」と規定し、check_output.py でも点検 ②builder 本体は変えない（判定58-2）
- 🔴 V6（PowerPoint 実機の目視）でこの欄の見え方を確かめる必要あり

## ■ 4. V1 で確かめられなかったこと（V2 以降）
- AI が流れ（ヒアリング・確認A/B）を守るか／URL の読み取り／その環境の python-pptx／スキルの置き場所／利用枠での途中停止／PowerPoint 実機の見え方
