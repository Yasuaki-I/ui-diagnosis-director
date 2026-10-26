<!-- 【購入者限定】本ファイルの再配布・転載・転売は固く禁じます。｜UI診断ディレクター スキル版 prototype r6.1（2026-10-27） -->
# 入力の項目一覧（builder の3関数に渡す内容）
- 字数の上限を超えると builder が止まるか、点検が不合格になる。その項目を言い換えて作り直す
- 入力全体を次の形で JSON に保存して、点検（scripts/check_output.py）にかける
```
{"site_slug": "英数字", "record": {...}, "diagnosis": {C-1}, "proposals": {C-2}, "visual_data": {C-3}}
```
- 見本：`assets/sample_input_rankB.json`・`sample_input_rankC.json`・`sample_input_rankD.json`（架空のサイト）

## 字数の数え方（3つの上限をすべて満たす）
| 記号 | 上限 | 数え方 |
|:-:|---|---|
| A | 実文字数 | len（1字＝1） |
| B | builder の切り詰め | 半角（ASCII）＝0.5、それ以外＝1.0。超えると builder が「…」で切る |
| C | 表示の換算 | 半角英数記号＝0.5、全角・かな・漢字・①※→●などの記号・全角スペース＝1.0 |
- **入力に使えない文字**：半角カナ（全角に直す）・改行・タブ・絵文字・異体字セレクタ・結合文字・ゼロ幅文字・制御文字・私用領域・上のどれにも当たらない記号（✓✕⚠ など。記号が必要な欄は builder が描く）
- 「C ちょうど＝必ず合格」ではない。最後は点検の結果と Windows の PowerPoint での目視で確かめる

## record（正式名称・URL｜PPTX には出ない）
| 項目 | 内容 |
|---|---|
| service_name_official | 正式名称（空は不可） |
| target_url | `http://` か `https://` で始まる URL、または null（分からないときは null。架空の URL を作らない） |
| input_origin | `url` か `screenshot` のどちらか。url のときは target_url が必須。screenshot のときは input_type に「スクショ」を含める。input_type に「URL」を含む ⇔ url |
- record には表示の字数・文字種の制約を当てない（PPTX に出ないため）

## C-1 diagnosis（add_scorecard_onepager｜01_UI_Scorecard_{site_slug}.pptx）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name | 表示名（明示＞サイト名＞画像内のロゴ＞「診断対象サイト」）。C-1・C-2・C-3 で同じ | len 30・表示11換算 |
| input_type | 「URL」「スクショ」「URL / LP」など短く | 表示10換算 |
| total_score | 総合点（整数・50点満点） | — |
| rank／rank_label | ランク（S〜D）／ランクの説明（scoring_rules.md §4 の1語） | — |
| scores | 10件：{category, score（1〜5 または '－'）, max: 5, comment（所見）} | comment len 40 |
| strengths | 強み（**3件**） | 各 len 40 かつ 28換算（C-3 1枚目で builder が28換算を超えると切るため） |
| priority_issues | 最優先の課題（**3件**・ランクに関わらず） | 各20（len・換算とも） |
| conclusion | 結論の1文 | len 60（2行まで・114換算） |

## C-2 proposals（add_proposal_onepager｜02_Proposals_{site_slug}.pptx）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name | C-1 と同じ | len 30・表示11換算 |
| proposals | 5件（scoring_rules.md §6 の項目） | title len 30・32換算／target_area len 25・10換算／issue len 25／before・after len 40・24換算 |
| summary | 最優先の一手 | len 120（2行まで・114換算） |

## C-3 visual_data（add_visual_board｜03_Visual_Board_{site_slug}.pptx）
**漏れがあると、その欄は代わりの文になるか空欄になる。すべて入れる。**（highlights の箇所は `area` ではなく **`target_area`** と書く）
| 項目 | 内容 | 上限 |
|---|---|---|
| service_name・input_type・total_score・rank・rank_label | C-1 と同じ値 | — |
| purpose | 診断の狙い（例：無料体験CVを増やす） | len 20・14換算 |
| target | 診断対象の範囲（例：LP全体） | len 20・14換算 |
| diagnosis_date | 診断日（例：2026年10月20日） | 14換算 |
| summary | **総評（1行の標語）** | **len 17** |
| top_issue | 最優先の課題の見出し（top_issues が無いときの代わり） | len 40・20換算 |
| direction | 改善の方向 | len 120（2行まで・112換算） |
| sections | 構造マップ 5〜9件：{no, label, desc, has_issue（true/false）, status（'good' または 'issue'）} ※5つとも必須 | label len 20・9換算／desc 10換算／label＋desc 16換算（同じ1行に出る） |
| flow_steps | 行動の流れ 6件：{label, status（'✓' または '✕'）, note} | label 6換算・note 9換算／label＋note で len 25 |
| flow_summary | **必須**：行動の流れのまとめ（1文・例：FVと料金で判断保留が多発） | len 40・31換算 |
| highlights | Before/After 3組（**C-2 の提案と箇所で1対1に対応。対応する組は priority・effort・target_area・before・after を同じ文字列に**）：{title, target_area, priority, effort, before, after}。**target_area・priority・effort は、C-2 の対応する改善提案と同じ値を使う**（priority＝高・中・低／effort＝小・中・大） | title len 30・27換算／target_area 10換算／before・after 24換算（**C-2 の対応する提案と同じ文字列**） |
| scores・strengths・priority_issues | C-1 から**そのまま写す**（件数も同じ） | — |
| top_issues | priority_issues を**そのまま写す**（同じ件数） | 各20 |
