<!-- 【購入者限定】本ファイルの再配布・転載・転売は固く禁じます。｜UI診断ディレクター スキル版 prototype r5 案A（2026-10-23） -->
# 入力の項目一覧（builder の3関数に渡す内容）
- 字数の上限を超えると builder は止まる（エラーの文に項目名が出る）。その項目を言い換えて作り直す
- 入力全体を次の形で JSON に保存して、点検（scripts/check_output.py）にかける
```
{"site_slug": "英数字", "diagnosis": {C-1}, "proposals": {C-2}, "visual_data": {C-3}}
```
- 見本：`assets/sample_input_rankB.json`（架空のサイト・ランクB）

## C-1 diagnosis（add_scorecard_onepager｜01_UI_Scorecard_{site_slug}.pptx）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name | サービス名（明示＞サイト名＞画像内のロゴ＞「診断対象サイト」） | 30字 |
| input_type | 「URL」「スクショ」「URL / LP」など短く | 全角10字 |
| total_score | 総合点（整数・50点満点） | — |
| rank／rank_label | ランク（S〜D）／ランクの説明（scoring_rules.md §4 の1語） | — |
| scores | 10件：{category, score（1〜5 または '－'）, max: 5, comment（所見）} | comment 40字 |
| strengths | 強み（**3件**） | 各40字 |
| priority_issues | 最優先の課題（**3件**・ランクに関わらず） | 各20字 |
| conclusion | 結論の1文 | 60字（builder の上限は120字） |

## C-2 proposals（add_proposal_onepager｜02_Proposals_{site_slug}.pptx）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name | C-1 と同じ | 30字 |
| proposals | 5件（scoring_rules.md §6 の項目） | title 36字ほか |
| summary | 最優先の一手 | 120字 |

## C-3 visual_data（add_visual_board｜03_Visual_Board_{site_slug}.pptx）
**漏れがあると、その欄は代わりの文になるか空欄になる。すべて入れる。**（highlights の箇所は `area` ではなく **`target_area`** と書く）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name・input_type・total_score・rank・rank_label | C-1 と同じ値 | — |
| purpose | 診断の狙い（例：無料体験CVを増やす） | 20字 |
| target | 診断対象の範囲（例：LP全体） | 20字 |
| diagnosis_date | 診断日（例：2026年10月20日） | — |
| summary | **総評（1行の標語）** | **全角17字** |
| top_issue | 最優先の課題の見出し | 40字 |
| direction | 改善の方向 | 120字 |
| sections | 構造マップ 5〜9件：{no, label, desc, has_issue（true/false）, status（'good' または 'issue'）} ※5つとも必須 | label 20字 |
| flow_steps | 行動の流れ 6件：{label, status（'✓' または '✕'）, note} | label＋note で25字 |
| flow_summary | **必須**：行動の流れのまとめ（1文・例：FVと料金で判断保留が多発） | 40字 |
| highlights | Before/After 3組：{title, target_area, priority, effort, before, after}。**target_area・priority・effort は、C-2 の対応する改善提案と同じ値を使う**（priority＝高・中・低／effort＝小・中・大） | title 30字・target_area 20字・before/after 各40字 |
| scores・strengths・priority_issues | C-1 から**そのまま写す**（件数も同じ） | — |
| top_issues | priority_issues を**そのまま写す**（同じ件数） | 各20字 |
