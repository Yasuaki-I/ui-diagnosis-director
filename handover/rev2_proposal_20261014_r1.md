# PLAIN rev2・lite rev2 起案 r1（R-7 解消｜起案キュー⑧）｜2026-10-14（水）｜AIスライド側 5往復目
> 🔴 本書は起案である（統括判定前）。gpts-package・handover の現行版には一切書き込んでいない（作業領域で作成）
> 正とする文面：`claude_chat_message_20261014_r7_revision_director_r1.md`（5,824B・`ed3d311ce2dd`）の「再掲2」逐語版／判定：`claude_chat_verdict_20261014_r7_director.md`（3,636B・`eda3b0bd6b64`）判定1〜5
> 前版 `rev2_proposal_20261014.md`（3,669B・`3e9305a1155a`）は無変更で残す（確定版文面の受領前の起案のため、本書で置き換える）

## ■ 1. PLAIN rev2（確定版逐語・445行の1か所のみ）
| 項目 | 現行 | rev2（案） |
|---|---|---|
| ファイル名 | `01_Instructions_v3.5_PLAIN_20261007.md` | **`01_Instructions_v3.5_PLAIN_20261014.md`**（日付付番の慣例） |
| 字数／B | 36,721字／77,248B | 36,899字／**77,502B** |
| md5 | `59ebb0757fde` | **`dad52ffe7360`** |
```diff
- exec(open('/mnt/data/03_pptx_builder.py').read())
+ _ns = {'__name__': 'pptx_builder'}; exec(open('/mnt/data/03_pptx_builder.py').read(), _ns)
+ globals().update({k: v for k, v in _ns.items() if not k.startswith('__')})
+ # ※ デモを動かさないための読み込み方（builder 末尾のデモが実行されず、test_*.pptx は作られません）
```
- 冒頭の改稿履歴（「🔴 2026-10-xx 改稿」行）は追加していない（確定文面の範囲外）。追加の要否は判定事項（追加すると B・md5 を再算出）
- 確定版3行で試験：test_*.pptx 0件・関数呼出し可（前回 r7_trial と同結果）

## ■ 2. lite rev2：🚨 J-3 内だけでは 92字を回収できない（判定5 の但し書きに該当）
- コード2行は確定版の逐語のまま（変更なし）。2行差替え直後＝8,092字（92字超過）
### 2-1. J-3 内コメント4行（＋見出し行）の圧縮｜回収 **29字**（最大限）
| 旧文 | 新文 |
|---|---|
| `# 3辞書（詳細はvisual_data_schema.md）` | （削除し、3行目末尾の「(詳細visual_data_schema.md)」へ移す） |
| `# C-1: service_name/total_score/rank/scores10/strengths3/priority_issues3(20字)/conclusion` | `#C-1:service_name/total_score/rank/scores10/strengths3/priority_issues3(20字)/conclusion C-2:proposals5/summary(120字)` |
| `# C-2: proposals5/summary(120字)` | （上の行へ統合） |
| `# C-3: C-1全項＋purpose/target(各20字)/diagnosis_date/summary(17字)/direction(120字)/sections/flow_steps(6)/highlights3` | `#C-3:C-1全項+purpose/target(各20字)/diagnosis_date/summary(17字)/direction(120字)/sections(no/label/desc/has_issue/status必須)/flow_steps(6)/highlights3` |
| `# ★sectionsはno/label/desc/has_issue/status必須。scores/strengths/priority_issues/top_issuesはC-1から全転記（件数もG項に一致）` | `#scores/strengths/priority_issues/top_issuesはC-1全転記・件数G項一致(詳細visual_data_schema.md)` |
- 審査基準の自己点検：必須フィールド名（service_name〜highlights3・sections の5キー・top_issues）と字数指定はすべて残存・**欠落0**。「件数G項一致」を残存（G項本体の `S/A/B=3・C=4・D=5` は §2-2 でも無変更）
- 回収内訳：見出し行の統合 −5／C-1・C-2 行の統合 −4／C-3 行（sections 5キーを括弧で吸収）+2／末行（★sections文の移動・「から」「（件数も…に）」の削除）−22 → 計 **29字**。フィールド名の列挙を残す限り、J-3 内はこれが上限
### 2-2. 他節の圧縮案（併記｜回収 **72字**）＝**太字マーカー `**` の除去のみ・文言は1字も変えない**
| 節 | 行 | 変更 | 回収 |
|---|---|---|---:|
| C. 入力 | 17 | `**URL**`・`**スクショ**`・`**両方**`・`**サービス名**` の太字を外す | 16 |
| C. エラー時テンプレ | 22 | `**E1. URL失敗**`・`**E2. タイプ不明**` | 8 |
| 同 | 23 | `**E3. 読取困難**`・`**E4.**` | 8 |
| D. タイプ自動判定 | 27 | `**LP**`・`**EC**`・`**コーポレート**`・`**SaaS**` | 16 |
| 同 | 28 | `**バナー**`・`**その他**` | 8 |
| G. 強み・課題・結論 | 40 | `**強み◎3項目**`・`**最優先課題⚠**`・`**数**`・`**結論**`（`**20字以内**` は残す） | 16 |
- 選定理由：ラベル（項目名）の強調だけを外す。禁止・必達の強調（「即生成不可」「20字以内」「絶対禁止」等）は全件残す。細則11 第3項「削除条項0件」に触れない（条項・語句の削除なし）
- I項は無変更（細則11 第2項の「I項完全一致」を維持）／J-7 必要3項も無変更

### 2-3. lite rev2 の結果
| 項目 | 現行 | rev2（案：§2-1＋§2-2） |
|---|---|---|
| ファイル名 | `01_Instructions_v3.5.2_lite_20260902.md` | **`01_Instructions_v3.5.4_lite_20261014.md`**（v3.5.3 は handover に既存 `70cf379275c4` のため欠番扱い） |
| 字数／B | 7,976字／16,223B | **7,991字（残り9字）／16,210B** |
| md5 | `61b5f4e7b928` | **`6d53115592be`** |
- 本文1行目の版表記「（v3.5.1）」は既知の食い違い。rev2 で「（v3.5.4）」に直すと +0字（同字数）で直せる → 直すか判定事項（直すと md5 を再算出）
- 残り9字のため、今後の lite 改訂の余地はほぼない（記録事項）

## ■ 3. 判定依頼
1. PLAIN rev2（`dad52ffe7360`・77,502B）の採用候補登録／改稿履歴行の要否
2. lite J-3 圧縮（§2-1・29字）の採否
3. lite 他節圧縮（§2-2・太字除去72字）の採否
4. lite の版番号 v3.5.4・ファイル名／1行目の版表記を直すか
5. 配布は GPTs 実機確認後（判定2）。確認項目：① /mnt/data に test_*.pptx が無い ② 3ファイルが正常に出る（PLAIN・lite の両方で各1診断）
