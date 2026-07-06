# 紺＆クリーン デザイン仕様書

このドキュメントは、GPT が PPTX 生成時に参照する **デザイン仕様** をまとめたものです。Instructions の必達13条項を実装レベルで具体化します。

---

## 1. キャンバス・グリッド

| 項目 | 値 |
|---|---|
| 既定キャンバスサイズ | 1280 × 720 px（16:9）|
| PPTX 換算 | Width = Inches(13.333), Height = Inches(7.5) |
| 本文セーフエリア | 左右各 40px、上 60px（ヘッダ帯）、下 60px（フッター帯） |
| 本文有効領域 | 1200 × 600 px（top=60, height=600） |
| フッター境界線 | y = 660px（width=1280, height=1px, color=#CCCCCC） |

---

## 2. 配色パレット（固定）

| 用途 | カラーコード | RGB | 使用箇所 |
|---|---|---|---|
| メインナビー | `#1C366C` | (28, 54, 108) | ヘッダ帯・カードアクセント・タイトル・表ヘッダ |
| アクセント赤 | `#D0021B` | (208, 2, 27) | キー数値の強調・警告・赤ライン装飾 |
| ライトグレー（背景） | `#F4F5F8` | (244, 245, 248) | カード背景・縞模様セル背景 |
| ボーダーグレー | `#D0D4DC` | (208, 212, 220) | 表罫線・カード枠線 |
| 本文テキスト | `#404040` | (64, 64, 64) | 通常本文 |
| サブテキスト | `#606060` | (96, 96, 96) | 注釈・脚注・補足説明 |
| ナビー薄文字 | `#9DB0D6` | (157, 176, 214) | ヘッダ帯内の補助テキスト（白背景紺帯内） |
| 強調用ゴールド | `#FFD54F` | (255, 213, 79) | 紺帯内でのキーワード強調（赤の代替） |
| 白 | `#FFFFFF` | (255, 255, 255) | 紺帯上のテキスト・本文背景 |

**配色の使い方ルール**:
- 1スライド内で「ナビー」「赤」「ゴールド」を併用しても良いが、各色は **多くて2か所まで**
- 赤＝危機感／優先対応／問題深刻度 ／ ナビー＝規模・成果・整理済みの量
- 蛍光色・ネオン色・パステル系は使用しない

---

## 3. タイポグラフィ

### フォント

- **全テキスト要素**：`メイリオ`（python-pptx では `run.font.name = 'メイリオ'`）
- 強調目的でも他フォントへの切替は禁止

### サイズスケール（pt）

| 用途 | サイズ | 太さ | 行間 |
|---|---|---|---|
| 大見出し（表紙タイトル） | 36〜44pt | Bold | 1.3 |
| ヘッダ帯タイトル | 22〜24pt | Bold | 1.2 |
| 中見出し（カードヘディング） | 20〜22pt | Bold | 1.4 |
| 本文 | 16〜18pt | Regular | 1.6 |
| 表セル本文 | 14〜16pt | Regular | 1.5 |
| 注釈・脚注 | 14pt | Regular | 1.6 |
| ページ番号・著者表記 | 14pt | Regular | 1.0 |
| 強調数値（核心） | 28〜44pt | Bold | 1.0 |
| 強調数値（主要） | 18〜24pt | Bold | 1.0 |

**最小ルール**: いかなる用途でも 14pt 未満は禁止。

---

## 4. レイアウト共通要素

### ヘッダ帯（スライド上端）

- 矩形 shape：left=0, top=0, width=1280, height=60, fill=#1C366C
- タイトルテキスト：left=40, top=18, font_size=22pt, color=#FFFFFF, bold
- サブタイトル（英字ラベル、右端）：left=1080, top=22, width=160, font_size=14pt, color=#9DB0D6, align=right

### フッター帯（スライド下端）

- 境界線 shape：left=0, top=660, width=1280, height=1, fill=#CCCCCC
- 著者表記：left=40, top=682, font_size=14pt, color=#1C366C, bold, letter_spacing=2
- ページ番号：left=1200, top=685, width=50, font_size=14pt, color=#262626, align=right

**フッター帯内に注釈・出所は配置しない**。出所は本文エリア最下部（top<660）に置く。

### 結論帯（オプション、本文末尾）

- 矩形 shape：left=40, width=1200, height=46〜80（内容次第）, fill=#1C366C, border_radius=8
- ラベル：font_size=13pt, color=#9DB0D6, bold, letter_spacing=2（例: "CONCLUSION"）
- 本文：font_size=18〜20pt, color=#FFFFFF, bold, line_height=1.4

---

## 5. 10種レイアウト詳細

### Layout 1: 表紙（add_cover）

**用途**: プレゼンの開始ページ。タイトル・日付・著者を提示。

**構成要素（z-index 順）**:
| z | 要素 | 座標 | サイズ | 内容・スタイル |
|---|---|---|---|---|
| 1 | ナビー背景 | (0, 0) | 1280×660 | fill=#1C366C |
| 5 | 赤ライン装飾 | (80, 360) | 480×3 | fill=#D0021B |
| 10 | プロジェクトラベル | (80, 160) | 900×26 | font=18pt, bold, color=#9DB0D6, letter_spacing=3 |
| 10 | メインタイトル | (80, 200) | 1120×130 | font=40〜48pt, bold, color=#FFFFFF, line_height=1.35 |
| 10 | サブタイトル | (80, 380) | 1120×70 | font=20〜22pt, color=#E6EAF3, line_height=1.6 |
| 10 | 日付＋著者 | (740, 570) | 460×40 | font=22pt, bold, color=#FFFFFF, align=right |
| 1 | 白フッター | (0, 661) | 1280×59 | fill=#FFFFFF |
| 10 | 著者表記 | (40, 682) | 400×20 | font=14pt, bold, color=#1C366C |
| 10 | ページ番号 | (1200, 685) | 50×20 | font=14pt, color=#262626, align=right |

### Layout 2: アジェンダ（add_agenda）

**用途**: 全体目次。章番号付きでスライドの構成を示す。

**構成要素**:
- ヘッダ帯（共通）
- リード文：「本資料の構成」など短い説明（top=90, font=18pt）
- アジェンダ項目（4〜8項目）を縦に並べる：
  - 各項目：「01」（紺、44pt, bold）＋章タイトル（22pt, bold）＋短い説明（16pt, color=#606060）
  - 項目間隔：80px
- フッター（共通）

### Layout 3: 課題整理（add_issue_summary）

**用途**: 現状の課題3点をカード形式で並列に示す。

**構成要素**:
- ヘッダ帯（共通）
- リード文：「現状の○つの課題」（top=90, font=18pt）
- 3カード横並び（各 390×400px、left=40/445/850, top=160）：
  - カード背景：fill=#F4F5F8, border_radius=8
  - 左帯：6px 幅の縦帯、fill=#1C366C
  - 番号ラベル：「01」「02」「03」（font=44pt, color=#1C366C, bold）
  - カード見出し：22pt, bold, color=#1C366C
  - カード本文：16pt, color=#404040, line_height=1.7（4〜6行）
- 結論帯（オプション、top=580）
- フッター（共通）

### Layout 4: 優先度マトリクス（add_priority_matrix）

**用途**: 4象限マトリクスで施策・課題をプロットする。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- 4象限の図：
  - 全体エリア：left=200, top=140, width=900, height=460
  - 縦軸ラベル（上＝高、下＝低）：left=160, font=14pt, color=#1C366C, bold
  - 横軸ラベル（左＝低、右＝高）：top=610, font=14pt, color=#1C366C, bold
  - 4象限の背景：左上＝最優先（fill=#FFE8E8）、右上＝重要（fill=#F4F5F8）、左下＝注視（fill=#F4F5F8）、右下＝後回し（fill=#FAFAFA）
  - 象限ラベル：各象限の隅に小さく（font=14pt, color=#606060）
  - プロット点（円）：直径20〜40px、fill=#1C366C または #D0021B、上に番号
  - プロット点の右にラベル（font=14pt）
- フッター（共通）

### Layout 5: OK・NG例の対比（add_ok_ng_pair）

**用途**: 良い例（OK）と悪い例（NG）を左右に並べて対比。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- 左右2カラム（各 580×460px、left=40/660, top=140）：
  - 左カラム（NG）：上部に赤バッジ「NG ✕」、内側にNG例の説明＋スクリーンショット風モック
  - 右カラム（OK）：上部に緑バッジ「OK ✓」、内側にOK例の説明＋スクリーンショット風モック
  - 各カラム下部に短いキャプション（font=14pt, color=#606060）
- 結論帯（top=620）
- フッター（共通）

### Layout 6: 施策一覧表（add_action_table）

**用途**: 複数施策を表組でリスト化（最大10行程度）。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- テーブル（left=40, top=140, width=1200, height=実測）：
  - ヘッダ行：背景=#1C366C, 文字=#FFFFFF, bold, font=16pt, padding=10
  - データ行：背景=#FFFFFF / #F8F9FB の縞模様、罫線=#D0D4DC, font=14〜15pt, padding=8〜10
  - **訴求数値の強調**：セル内で数字部分だけ font=18pt, color=#D0021B, bold
- 出所注記（テーブル下、フッター帯の上）：font=14pt, color=#606060
- 結論帯（オプション）
- フッター（共通）

### Layout 7: KPI カード（add_kpi_card）

**用途**: 核心数値を大きく強調する。3〜4個の KPI を並べる。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- KPI カード（3〜4個横並び）：
  - 各カード：290〜390×280px、fill=#FFFFFF, border=2px #1C366C, border_radius=8
  - ラベル（上部）：font=14pt, color=#1C366C, bold, letter_spacing=2（例: "売上前年比"）
  - 大きな数値：font=44〜56pt, bold, color=#D0021B（赤）または #1C366C（紺）
  - 単位（数値の右）：font=20pt, color=同色
  - 下部の説明文：font=14pt, color=#606060
- 結論帯（top=550）
- フッター（共通）

### Layout 8: スケジュール／ガント風（add_schedule_gantt）

**用途**: 時系列の計画を月別バーで示す。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- スケジュール表：
  - 上部に月ヘッダ（例：4月／5月／6月／7月／8月／9月）（各列幅 170px）
  - 左列にタスク名（font=16pt）
  - 各タスク行に色付き矩形バーで期間を示す（fill=#1C366C, height=24, border_radius=12）
  - マイルストーンは菱形マーク（◆）で表示
- 出所注記
- フッター（共通）

### Layout 9: 運用フロー比較（add_flow_compare）

**用途**: Before / After のフロー（業務手順）を上下2段で対比。

**構成要素**:
- ヘッダ帯（共通）
- リード文（top=90）
- 上段：Before
  - 左端に「BEFORE」バッジ（fill=#888, color=#FFF, font=14pt, bold）
  - 横方向にステップカードを並べる（4〜6個、各 180×80px、間に矢印 ▸）
  - ステップカード：fill=#F4F5F8, border=1px #D0D4DC, 内側にステップ名（font=15pt）
- 下段：After
  - 左端に「AFTER」バッジ（fill=#1C366C, color=#FFF, font=14pt, bold）
  - 横方向にステップカード（fill=#FFFFFF, border=2px #1C366C, 内側にステップ名）
- 結論帯（top=600）
- フッター（共通）

### Layout 10: クロージング（add_closing）

**用途**: 最終スライド。次のアクション・問い合わせ先・謝辞。

**構成要素**:
- ナビー背景（cover と同じ）：left=0, top=0, width=1280, height=660, fill=#1C366C
- 「Thank you」または「Next Step」見出し：left=80, top=200, font=48pt, bold, color=#FFFFFF
- 赤ライン装飾：left=80, top=300, width=480, height=3, fill=#D0021B
- メインメッセージ：left=80, top=340, width=1120, font=22pt, color=#E6EAF3, line_height=1.6
- 問い合わせ先（右下）：left=740, top=560, font=16pt, color=#FFFFFF, align=right
- フッター（白背景に著者表記、cover と同じ）

---

## 6. 訴求数値の強調実装

### 表セル内での数字強調

```python
cell = table.cell(row, col)
cell.text = ''
para = cell.text_frame.paragraphs[0]
# 通常テキスト
run1 = para.add_run()
run1.text = 'うち '
run1.font.name = 'メイリオ'
run1.font.size = Pt(14)
# 強調数字（同じセル内で別 run）
run2 = para.add_run()
run2.text = '20件'
run2.font.name = 'メイリオ'
run2.font.size = Pt(18)
run2.font.bold = True
run2.font.color.rgb = RGBColor(0xD0, 0x02, 0x1B)
# 続きのテキスト
run3 = para.add_run()
run3.text = ' が要対応'
run3.font.name = 'メイリオ'
run3.font.size = Pt(14)
```

### カード内の核心数値（独立 textbox）

```python
# 数字専用テキストボックス
num_box = slide.shapes.add_textbox(Inches(0.6), Inches(2.1), Inches(2), Inches(1))
num_tf = num_box.text_frame
num_tf.word_wrap = True
num_para = num_tf.paragraphs[0]
num_run = num_para.add_run()
num_run.text = '39.6%'
num_run.font.name = 'メイリオ'
num_run.font.size = Pt(44)
num_run.font.bold = True
num_run.font.color.rgb = RGBColor(0xD0, 0x02, 0x1B)

# 説明文専用テキストボックス（数字の下）
desc_box = slide.shapes.add_textbox(Inches(0.6), Inches(2.8), Inches(2), Inches(0.4))
desc_tf = desc_box.text_frame
desc_para = desc_tf.paragraphs[0]
desc_run = desc_para.add_run()
desc_run.text = '高優先度（S＋A）の割合'
desc_run.font.name = 'メイリオ'
desc_run.font.size = Pt(14)
desc_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
```

---

## 7. PPTX 単位換算チート

| 1280×720 px キャンバスでの座標 | python-pptx での換算 |
|---|---|
| 1 px | Emu(9525) |
| 横方向 1280px | Inches(13.333) |
| 縦方向 720px | Inches(7.5) |
| left=40px | Inches(0.417) または Emu(381000) |
| top=60px | Inches(0.625) |
| left=1200px | Inches(12.5) |
| top=660px | Inches(6.875) |

便利関数:
```python
def px(n):
    """1280×720 想定のピクセル値を Emu に変換"""
    return int(n * 9525)
```

---

## 8. NG パターン集（やってはいけない）

| NG | 理由 |
|---|---|
| `font.name = 'Yu Gothic'` | 条項1違反（メイリオ以外） |
| `font.size = Pt(12)` | 条項2違反（14pt未満） |
| 表セル内に独立 textbox を重ねる | 条項4違反（PPT表セルとして扱えなくなる） |
| カードに `shape.text_frame.text = "..."` で文字を入れる | 条項7違反（編集時に文字が動かせない） |
| `RGBColor(0xFF, 0xFF, 0x00)` 蛍光黄色をテキスト背景に | 条項13違反（蛍光色禁止） |
| ハイパーリンク `slide.shapes.add_hyperlink()` | スライドにリンクを埋め込まない |
| 表紙の日付を `PP_ALIGN.LEFT` に | 条項9違反（右下右揃え必須） |
| フッター帯（top>=660）に注釈を置く | 条項8違反（ロゴ＋ページ番号専用） |

---

このスペックを守って実装すれば、デザインの一貫性と品質が保証されます。
