# 🔧 v17 P1｜実装手順書｜category / breakdown / comparison

- 作成日：**2026-08-23（日）18:00連結③**
- 作成者：AIスライド（実装領域）
- 対象成果物：`v17_p1_draw_patterns_20260823.py`
- 実装担当：**入江さん**（開発環境への投入）
- 位置づけ：⭐ **第16条「段階2（実装済）」までがAIスライドの対象範囲**｜段階3（統合済）は本手順書に沿った投入完了をもって成立

---

## ■ 0. 3行サマリ

1. 添付 `.py` の「▼▼▼ v17 追加ブロック」以降を、`03_pptx_builder_v16_5_20260728.py` の**末尾に追記**する（推奨）。
2. `__version__` を **`'17.0.0'`／`__version_date__ = '2026-08-23'`** に更新する（⭐ 第16条 細則7）。
3. 回帰確認は §5 のチェックリスト（既存C-1〜C-3が壊れていないこと＋新規47項目）で行う。

---

## ■ 1. 投入対象ファイルと版数

| 項目 | 現状 | 投入後 |
|---|---|---|
| ファイル名 | `03_pptx_builder_v16_5_20260728.py` | ⭐ **`03_pptx_builder_v17_20260823.py`**（改名推奨） |
| ⚠️ `__version__`（35行） | ⚠️ `'15.0.0'` | ⭐ **`'17.0.0'`** |
| ⚠️ `__version_date__`（36行） | ⚠️ `'2026-07-12'` | ⭐ **`'2026-08-23'`** |
| サイズ／行数 | 181,647B／3,861行 | 追記分 約22KB／約530行 |

⚠️ **ファイル名だけを変えても段階判定は変わらない。** 第16条 細則7により、判定根拠は `__version__` 宣言である。**必ず両方を書き換える。**

---

## ■ 2. 挿入位置｜2案（どちらでも動作する）

### ⭐ 案A（推奨）｜ファイル末尾に追記

```
（既存 3,861行）
...
if __name__ == '__main__':   ← 既存の末尾ブロックがある場合はその**前**に挿入
# ▼▼▼ v17 追加ブロック ▼▼▼
（本ファイルの該当部分を貼り付け）
# ▲▲▲ v17 追加ブロック ここまで ▲▲▲
```

**根拠**：v17ブロックは `px` / `add_text` / `add_shape` / `add_paragraph_box` / `_add_bg_frame` / `set_run` / `CANVAS_W_PX` / `TEXT` / `WHITE` を**関数の実行時に参照**する設計（import時に評価しない）ため、定義順の制約を受けない。末尾追記が最も安全。

### 案B｜Phase Aブロック直後

`03_pptx_builder_v16_5_20260728.py` の

- **213行**：`# ▲▲▲ Phase A 追加ブロック ここまで ▲▲▲`
- **215行**：`# キャンバス（1280×720 想定）`

の**間**に挿入する。辞書定義（`DIAGRAM_PATTERNS` 181〜194行／`DIAGNOSIS_TO_PATTERN` 197〜209行）の直後に描画層が並ぶため可読性が高い。案Aと動作差はない。

⚠️ **どちらの案でも、既存の C-1〜C-3 描画ロジック（`add_cover` ほか `add_*` 21件）には一切触れない。**

---

## ■ 3. 依存関係

### 3-1. 外部ライブラリ

| ライブラリ | 追加インストール |
|---|:---:|
| `python-pptx` | ⭕ **不要**（既存の import で充足） |
| その他 | ⭕ **なし** |

⭐ **新規の `import` 文は1行も追加していない。** v17ブロック内で必要な `MSO_SHAPE` / `PP_ALIGN` / `RGBColor` は、既存 import に加えて関数内 import でも二重に担保している。

### 3-2. ビルダー既存要素への依存（すべて既存・変更不要）

| 参照先 | ビルダー行 | 用途 |
|---|:---:|---|
| `DIGITAL_AGENCY_PALETTE` | **95〜165行** | 7テーマ×8色階調 |
| `DIAGRAM_PATTERNS` | **181〜194行** | 原本12種の3属性（`ja`/`use`/`shape`） |
| `DIAGNOSIS_TO_PATTERN` | **197〜209行** | 11マッピング（`resolve_pattern` が参照） |
| `px()` | **225行** | px→EMU 変換 |
| `set_run()` | **230行** | ⭐ **メイリオ3スクリプト強制＋14pt下限**（条項2） |
| `add_text()` | **366行** | 1行テキストボックス |
| `add_paragraph_box()` | **398行** | 複数行ボックス |
| `add_shape()` | **455行** | 図形（⭐ 図形内テキストは空＝条項7） |
| `_add_bg_frame()` | **520行** | 角丸カード |
| `CANVAS_W_PX` / `CANVAS_H_PX` | **216〜217行** | 1280×720 |

### 3-3. ⚠️ 原本を書き換えていないこと（第9条・2層構造の維持）

⭐ **`DIAGRAM_PATTERNS` に拡張プロパティを追記していない。** 拡張層は新設の `DIAGRAM_PATTERN_SPEC` に分離した（統括判定①｜選択肢[A]｜「原本＝骨格／実装＝拡張」）。原本辞書は**1文字も変更しない**。

---

## ■ 4. 公開API｜7関数

| 関数 | 用途 |
|---|---|
| `get_theme_palette(theme_id)` | テーマID→8色階調（未知IDは `SolidGray`） |
| `select_theme_by_project_type(project_type, warning_flag=False)` | テーマ自動選定（`warning_flag=True`→`Red` 優先） |
| `resolve_pattern(diagnosis_key)` | 診断カテゴリ→パターンキー（原本 `DIAGNOSIS_TO_PATTERN` 参照｜未定義は `category`） |
| `draw_category(slide, palette, data)` | 分類（grid／3〜6要素） |
| `draw_breakdown(slide, palette, data)` | 分解（vertical／3〜7要素） |
| `draw_comparison(slide, palette, data)` | 比較（horizontal／2〜3要素） |
| `draw_pattern(slide, pattern_key, palette, data)` | ディスパッチャ（未実装キーは `category` に退避） |

### 4-1. 最小の呼び出し例

```python
exec(open('/mnt/data/03_pptx_builder_v17_20260823.py').read())

prs   = create_presentation()
theme = select_theme_by_project_type('lp', warning_flag=False)   # -> 'Orange'
pal   = get_theme_palette(theme)

slide = _blank_slide(prs)
_add_header(slide, '改善提案のカテゴリ分類', 'DIAGRAM')
_add_footer(slide, 1, 3)

report = draw_category(slide, pal, {
    'title': '診断結果｜改善提案のカテゴリ分類',
    'categories': [
        {'label': 'ファーストビュー', 'score': 82, 'description': '訴求と導線が一致'},
        {'label': '情報設計',       'score': 64, 'description': '見出し階層に飛び'},
        {'label': '導線設計',       'score': 38, 'description': 'CTAが画面外'},
        {'label': '可読性',         'score': 71, 'description': '字間に余地'},
    ],
})
prs.save('out.pptx')
print(report)   # {'pattern':'category','elements_drawn':4,'fallback_from':None,'notes':[...]}
```

### 4-2. データスキーマ

```python
# category
{'title': str,
 'categories': [{'label': str, 'score': int|None, 'description': str}, ...]}   # 3〜6

# breakdown
{'title': str,
 'whole': {'label': str, 'value': int|float|None},          # value=None なら合計を自動算出
 'components': [{'label': str, 'value': int|float,
                 'score': int|None, 'note': str}, ...]}     # 3〜7・value は 0以上

# comparison
{'title': str,
 'comparison_axis': str|None,          # 任意（原本 requires_axes=False のため必須にしない）
 'attribute_labels': [str, ...],
 'items': [{'label': str, 'score': int|None,
            'attributes': {str: str}}, ...]}                # 2〜3
```

⚠️ **戻り値は必ず dict（描画レポート）である。** `notes` に丸め込み・フォールバック・省略の記録が入るため、**GPTs側でログ出力することを推奨**する。

---

## ■ 5. 回帰確認手順｜3段

### 段1｜既存機能の後方互換（⭐ 最優先）

| # | 確認 | 期待 |
|:---:|---|---|
| 1 | `add_cover` / `add_agenda` / `add_scorecard_onepager` / `add_proposal_onepager` / `add_visual_board` を従来通り実行 | ⭕ **従来と同一の出力**（v17ブロックは既存関数を一切参照・改変しない） |
| 2 | `DIAGRAM_PATTERNS` / `DIAGNOSIS_TO_PATTERN` の内容 | ⭕ **変更なし**（diff で確認） |
| 3 | `validate_length` / `LIMITS` | ⭕ **変更なし** |
| 4 | 既存の `.pptx` 生成スクリプトをそのまま実行 | ⭕ **例外なし・レイアウト変化なし** |

### 段2｜v17新機能テスト（同梱ハーネス）

`v17_p1_test_harness.py` を同ディレクトリに置いて実行：

```bash
BUILDER=03_pptx_builder_v17_20260823.py \
V17=03_pptx_builder_v17_20260823.py \
OUT=./test_out python3 v17_p1_test_harness.py
```

⚠️ **統合後は `BUILDER` と `V17` に同じファイルを指定すると二重定義になる。** 統合後は `V17=/dev/null` 相当の空ファイルを指定するか、ハーネス冒頭の2行目 `exec` をコメントアウトする。

**AIスライド側の実行結果（統合前・2ファイル構成）：⭕ 47項目 PASS／0 FAIL**

| 群 | 内容 | 項目数 |
|:---:|---|:---:|
| T1 | ⭐ **21組合せ**（3パターン×7テーマ） | **21** |
| T2 | category 要素数境界（2/3/4/5/6/7） | 6 |
| T3 | breakdown 境界＋数値不正＋合計0 | 7 |
| T4 | comparison 要素数（1/2/3/4） | 4 |
| T5 | 警告オーバーライド閾値40 | 1 |
| T6 | `score=None` 許容 | 1 |
| T7 | ⭐ **決定論性**（category 100回／breakdown 50回→出力1種） | 2 |
| T8 | 描画範囲（ヘッダ60〜フッター660px内） | 1 |
| T9 | ⭐ **最小14pt・メイリオ3スクリプト強制** | 1 |
| T10 | `resolve_pattern` マッピング | 1 |
| T11 | 未実装キーの `category` 退避 | 1 |
| T12 | ⭐ **既存C-1〜C-3との同時使用** | 1 |
| **計** | – | ⭐ **47** |

### 段3｜目視確認（LibreOffice / PowerPoint）

AIスライド側では LibreOffice 変換→PNG化で3パターンを目視確認済（Blueテーマ）。**PowerPoint実機での確認は入江さん側で実施をお願いしたい。**

| 確認点 | AIスライド側の結果 |
|---|:---:|
| カード内の文字切れ | ⭕ なし |
| 値ラベルのバー外あふれ | ⭕ **バー幅不足時は自動でバー右外へ退避**（`_est_text_w` による事前見積） |
| 警告色セル（score<40）の可読性 | ⭕ ⭐ **輝度0.5判定で文字色を自動反転**（Green/Cyan/Red/Orange の `warning`=`#CCCCCC` でも判読可） |
| ヘッダ帯／フッターとの干渉 | ⭕ なし（描画域 y=90〜646） |

---

## ■ 6. ⭐ 拡張定義集約表からの実装マッピング（第9条｜参照行明記）

| # | パターン | 集約表の規定 | 実装箇所 |
|:---:|---|---|---|
| P1-2 | `category` | min3／max6／grid／`uniform_parallel` | `DIAGRAM_PATTERN_SPEC['category']`｜全セル `secondary` 基準色で並列性担保 |
| P1-3 | `breakdown` | min3／max7／vertical／`proportional` | `_breakdown_tier(ratio)`｜30%↑=`primary`／20%↑=`secondary`／10%↑=`midtone`／未満=`light` |
| P1-4 | `comparison` | min2／**max3**／horizontal／`discrete_contrast` | `contrast_keys`｜2要素=[`midtone`,`primary`]／3要素=[`midtone`,`secondary`,`primary`]（中間色を挟まない） |

### ⚠️ 集約表と実装記録（8/8）で数値が相違する2点｜集約表を採用した

| 項目 | 実装記録（8/8） | ⭐ 集約表（8/15） | 採用 | 根拠 |
|---|:---:|:---:|:---:|---|
| `breakdown` の要素数 | 親1〜3／子2〜4（階層型） | ⭐ **min3／max7（構成要素の分解）** | ⭐ **集約表** | 集約表は8/15の統括判定②に基づく確定版で、原本`use`「全体を構成要素に**分解**」との整合が明示されている |
| `comparison` の最大要素数 | max4 | ⭐ **max3** | ⭐ **集約表** | ⭐ **原本 `use` が「**2〜3要素**の対比」と明記**（ビルダー**184行**）｜原本に要素数が明記された唯一のパターンであり、原本が優越する |

⚠️ **この2点は第10条の自律判断事項として `ops_rule_article10_case_registry.md` に事例014として登録する。**

### 横断的原則4件の実装対応（集約表「横断的な設計原則」より）

| 原則 | 実装 |
|---|---|
| ① categoryフォールバック・例外を投げない | `draw_breakdown` / `draw_comparison` は範囲外・データ不正時に `draw_category` を呼ぶ。⭐ **`raise` は1箇所もない** |
| ② 警告オーバーライド（score<40→warning色） | `_tier_fill()`｜⭐ **`score=None` は判定スキップ**（network の先例を全パターンに一般化） |
| ③ 決定論性（同一入力→同一出力） | 乱数・時刻・辞書順依存を排除｜T7で100回／50回→出力1種を実測 |
| ④ python-pptx標準の基本図形優先 | ⭐ **使用図形は `ROUNDED_RECTANGLE` と `RECTANGLE` の2種のみ**｜`BLOCK_ARC` 等の `adjustments` 環境依存図形・コネクタは未使用（環境依存リスクなし） |

---

## ■ 7. ⭐ 実装上の追加判断｜2件（集約表に規定がなかった事項）

| # | 事項 | 判断 | 根拠 |
|:---:|---|---|---|
| 1 | ⭐ **カード上の文字色** | 背景色の**WCAG相対輝度0.5**を境に `WHITE` / `TEXT` を自動反転（`_text_color_on`） | ⚠️ Green/Cyan/Red/Orange の `warning` は `#CCCCCC`（明色）。白文字固定では **score<40 の警告セルが判読不能**になる。決定論的関数のため原則③に反しない |
| 2 | ⭐ **値ラベルのバー内外判定** | 概算幅（`_est_text_w`）がバー幅を超える場合はバー右外に配置 | 条項5（内容追従）／文字切れ根絶（v10方針）の踏襲。構成比が小さい要素でラベルが切れる事象を実測で検出し対処 |

⚠️ **いずれも第10条の自律判断範囲として事例014に併記する。**

---

## ■ 8. 未対応事項（v17 P2／P3 で実装）

| 対象 | パターン | 予定 |
|---|---|---|
| P2 | `pyramid` / `sequence` / `framework` | 8/24 |
| P3 | `funnel` / `timeline` / `contrast` / `cycle` / `network` | 8/25〜 |

⚠️ **P2／P3 未実装の間、`draw_pattern()` にこれらのキーを渡すと `category` に退避し、`notes` に「v17 P1 の対象外」と記録される（例外は出ない）。**

⭐ `integration` は原本 `DIAGNOSIS_TO_PATTERN` に対応診断カテゴリが存在しないため**構造的に対象外**（第16条 細則6の「対象外」区分）。

---

## ■ 9. 統合完了時に更新すべき記録（第16条）

| # | 対象 | 更新内容 |
|:---:|---|---|
| 1 | `v17_implementation_log_*.md` | `category` / `breakdown` / `comparison` の**段階3を ✅ に更新** |
| 2 | 同上 | ⭐ **段階判定の根拠として `__version__ = '17.0.0'` を明記**（細則7） |
| 3 | `image_registry_*.md` | 画像03が表現する実装物の統合状態を更新（細則6｜第17条 細則6連携） |
| 4 | GPTs環境 | 配布後に**段階4を ✅**｜⚠️ 配布環境のバージョン番号で確認（第16条 段階4の判定基準） |

⚠️ **AIスライドは段階3・段階4を自己申告で ✅ にしない。** 入江さんの投入完了報告を受けて更新する（8/22 統括指示）。

---

**v17 P1 実装手順書｜2026-08-23（日）18:00連結③｜⭐ 47項目 PASS／0 FAIL｜⚠️ 段階2（実装済）までがAIスライドの対象｜段階3は入江さんの投入をもって成立**
