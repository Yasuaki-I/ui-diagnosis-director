# 📚 事前資料3点｜統括初案（10/7 目安）向け｜2026-10-05（月）AIスライド側 14往復目

- 根拠：10/5 統括判定 第1便 伺い③（事前資料3点の整備のみ許可｜調査・設計・提案は凍結）
- 🔴 本書は **所在と既存記録の抜粋のみ** を置く。新たな測定・原因分析・是正案は含めない
- 抜粋はすべて配布中 builder `gpts-package/03_pptx_builder.py`（v17_216）と PLAIN `gpts-package/01_Instructions_v3.5_PLAIN_20260925.md` から行い、行番号は本日 `grep -n` で取得した
- 🚨 J-7-6-5 規準1 に従い md5・実測B は転記しない（`manifest_20261004.tsv` を参照）

---

## ■ 資料1｜起票 [A]（タイトル折返しとファネル重なり）の境界値記録の有無
### 1-1. 結論
- 🔴 **タイトル幅の境界値（何 px・何字で2行になるか）を定めた記録は、`handover/` に見当たらない。**
- 確認手段：`handover/*.md` と `handover/v37/*.md` に対する `grep`（`_v17_title`・`title_top`・「タイトル」×「折返」「2行」「幅」）。ヒットは「共通関数の一覧」「変更が及ばない箇所」「V37-1 計画書の試算」「10/4 実機目視記録」のみで、境界値を定めた記述はない
- ⚠️ 「見当たらない」は上記の検索語の範囲での結果であり、別表記の記録の不在までは保証しない（判断原理19）

### 1-2. 既存の関連記録（事実のみ）
| # | 記録 | 内容 | 所在 |
|:---:|---|---|---|
| 1 | 10/4 実機目視 | スライド5（43字タイトル）が実機で2行に折返し、2行目が y 126〜151 に描かれ `body_top`（138）を越える | `handover/v37/v37_device_check_record_20261004.md` §2 A-1 |
| 2 | 10/4 builder 推定 | 同タイトルの `_est_text_w`（22pt）推定 1,197px（枠 1,200px に「収まる」と推定） | 同 §2 A-3 |
| 3 | 10/4 `verify_render.py` | 同タイトル：🚨 **NG**／枠w 1200・枠h 34・22pt 太字・実描画 **1218.2px**・**2行**（slide33・v17_215 と v17_216 の両生成物で同一。v17_214r1 の生成物でも同一行を確認） | `handover/v37/v37_1_render_v17215_20261004.txt` 23行目／`v37_2_render_v17216_20261004.txt` 23行目 |
| 4 | 入力 | 固定入力 `v371_title_collide`（タイトル＝43字） | `handover/v37/v37_fixtures_20261004b.py` 73行目付近 |
- ⚠️ #3 は 10/4 の V37-1・V37-2 検証時点で `verify_render.py` が NG として出力していた。AIスライドは当時「通算表記の行」と「既存行の基準出力との一致」のみを判定に用い、この行を報告で取り上げていなかった（新規検出ではなく既存出力の見落とし｜事実として記す）

### 1-3. 関連コード（抜粋）
```python
# gpts-package/03_pptx_builder.py L4029-4033
V17_AREA = {
    'left': 40, 'right': 1240, 'width': 1200,
    'title_top': 90, 'body_top': 138, 'body_bottom': 646,
    'gap': 16,
}
# L4109-4113（全11パターン共通｜v17.2.13 から無変更）
def _v17_title(slide, title, palette):
    """パターン共通のタイトル行（ヘッダ帯とは別のスライド内見出し）"""
    add_text(slide, V17_AREA['left'], V17_AREA['title_top'], V17_AREA['width'],
             str(title), 22, bold=True, color=hex_to_rgb(palette['primary']),
             height_px=34)
```
- `_v17_title` の呼出し：12箇所（L4192・4196・4292・4401・4632・4725・4831・5194・5561・5656・5766・5963｜L4192 は category のフォールバック内）
- 推定器：`_est_text_w`（L4489｜全角＝size×1.34px・半角＝size×0.70px）／`_v17_text_w14`（L5055｜14pt 太字専用）

---

## ■ 資料2｜起票 [B]（注記配置）の footer 現行構造の抜粋
### 2-1. `_add_footer`（L589-594）
```python
def _add_footer(slide, page_num, total, author='紺＆クリーン スライド作成'):
    """フッター帯（ロゴ＋ページ番号のみ・条項8準拠）"""
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 660, CANVAS_W_PX, 1, fill=GRAY_BORDER)
    add_text(slide, 40, 682, 400, author, 14, bold=True, color=NAVY)
    add_text(slide, 1140, 685, 120, f'{page_num} / {total}', 14,
             color=PAGE_NUM, align=PP_ALIGN.RIGHT)
```
### 2-2. 縦方向の現行配置（図解スライド｜実座標 1280×720）
| y | 要素 | 出典 |
|---|---|---|
| 0〜60 | ヘッダ帯（`_add_header`） | L576 |
| 90〜124 | スライド内タイトル（22pt・枠高34） | `_v17_title` L4109 |
| 138〜646 | 本体（`body_top`〜`body_bottom`） | `V17_AREA` L4031 |
| 646〜660 | 空き 14px | 同上と L591 |
| 660 | フッター境界線（高さ1px・GRAY_BORDER） | L591 |
| 682〜706 | 作成者表記（14pt 太字・x40・幅400） | L592 |
| 685〜709 | ページ番号（14pt・x1140・幅120・右揃え） | L593 |
### 2-3. 関連する既存の取り決め（所在のみ）
- フッター帯は「ロゴ＋ページ番号のみ・条項8準拠」（`_add_footer` の docstring｜注釈をフッターに書かない旨は skill `ec` の必達条項にも記載）
- `_add_footer` の呼出し：23箇所（図解スライドは `add_diagram_slide` L4529 から）
- `body_bottom`〜フッター境界の 14px が 14pt 1行に足りないことの記録：`handover/claude_chat_plan_20261004_v37_2.md` §0／v17_216 `draw_funnel` docstring（L5159〜）
- 案L の現行配置（最下段の左・最下段上端＋28px）：v17_216 `draw_funnel` docstring L5159〜／実機画像 `handover/v37/device_check_20261004/slide6.png`・`slide7.png`

---

## ■ 資料3｜PLAIN 改稿用の条文・docstring 所在一覧
### 3-1. PLAIN（`gpts-package/01_Instructions_v3.5_PLAIN_20260925.md`｜handover 側と同一ファイル）
| 節 | 行 | 内容 | 改稿の論点（10/4 判定で示されたもの） |
|---|:---:|---|---|
| J-7（見出し） | 524 | 図解パターン | — |
| J-7-4 | 573 | `diagram_data` 必須キー表（funnel 行は 583｜`title`/`stages`[label,score,description]） | 新規任意キー（`cumulative`・`footnote`、および v17.2.4 の `value`・`unit`・`action`・`total_note`）の記載有無 |
| J-7-5 | 593 | 説明文の推奨字数（全角30字・「30字以内なら1行」） | 「30字は目安、2行折返しを許容する設計」への是正（10/4 判定） |
| J-7-6 | 619 | 文字揃えの適用基準 | — |
| J-7-6-2 | 661 | 文字幅の検証手順（A：推定器変更時／B：描画テキスト追加時） | — |
| J-7-6-3〜5 | 786・853・875 | 切詰基準／数値報告の入力併記／派生値を作らない | — |
| J-7-6-5 追記 | 923〜 | V37 番号の食い違い表（台帳と引継書） | — |
| J-7-7 | 938 | フォールバック挙動 | — |
| J-7-8 | 943 | 既定フロー化の移行判断基準 | — |
- 版管理：改稿は第18条 細則2（日付繰り上げ新版・前版無変更保存）＋ §0 #6-1（配置前の差分確認）
- lite（`gpts-package/01_Instructions_v3.5.2_lite_20260902.md`）：J-7 は 268行目〜（`diagram_data` 必須キーは「builder定義済」と記すのみ）。⚠️ lite の改稿は 9/9 判定で行わない扱い

### 3-2. builder docstring（`gpts-package/03_pptx_builder.py`｜v17_216）
| 箇所 | 行 | 内容 |
|---|:---:|---|
| `draw_funnel` Args | 5096〜 | `stages` の任意キー（value・unit・action）／`total_note`・`cumulative`・`footnote` |
| `[v17.2.4]` 節 | 5120 | 密度強化（①②④⑤）／③⑥見送りの経緯と案2 への改め |
| `[v17.2.14]` 節 | 5131 | V37-3（`_tight` の条件・gap 5・+32・band_h 80） |
| `[v17.2.15]` 節 | 5142 | V37-1（`cumulative`・有効数字ルール・衝突省略） |
| `[v17.2.16]` 節 | 5159 | V37-2（`footnote`・収容行数・太字用推定器による安全側見積もり） |
| `_v17_title` | 4109 | タイトル行（docstring 1行のみ） |
| `_add_footer` | 589 | フッター（docstring 1行のみ） |
| `V17_AREA` | 4029 | 描画領域（コメントなし） |
| `_est_text_w`／`_v17_text_w14`／`_v17_fits_one_line` | 4489／5055／5080 | 推定器 |

### 3-3. 台帳（改稿は 10/5 判定 伺い② 案ア により一括｜所在のみ）
| 台帳 | 所在 |
|---|---|
| SCRIPTS_REGISTRY（§6 N-6） | `handover/scripts/SCRIPTS_REGISTRY_20260924.md` |
| v3.7 候補台帳（§0 方式記述／V37-1・V37-2 形状変更） | `handover/v37_candidates_20260914.md` |
| 下辺余裕 監視台帳（watchlist 追加） | `handover/bottom_margin_watchlist_20260921.md` |
