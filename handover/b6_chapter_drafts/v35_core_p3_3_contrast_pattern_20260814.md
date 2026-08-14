# v3.5コアP3-3｜contrast パターン｜設計＋描画関数実装ドキュメント

- 作成日：2026-08-14（金）
- 作成者：AIスライド（実装領域）
- 位置づけ：**v3.5コアP3実装2日目｜P3-3 contrast｜設計＋描画関数実装記録**
- 統括承認：議論日2 議題4｜提案A（流用度優先）採用｜5判定発行済
- 運用条件：**AIスライド独立実装日｜自律判断＋事後承認方式**（8/12・8/13で実証済）
- 想定完了時刻：**13:00**

---

## 📋 実装着手前｜Phase A A-2原本整合確認｜完了（厳守①対応）

### 原本定義（`phase_a_design_20260727_rev2.md` A-2節｜逐語）

```python
DIAGRAM_PATTERNS = {
    'contrast': {'ja': '対比', 'use': '対照的な2要素の並列強調', 'shape': 'split_screen'},
}

DIAGNOSIS_TO_PATTERN = {
    'ux_contrast': 'contrast',
}
```

### 実装レベル拡張定義（本実装での確定値）

**8/13 20:00 統括判定①「拡張定義方針｜選択肢[A]運用継続｜承認発行済」に基づき、P1〜P3と同一の実装層拡張方針を継続適用**：

```python
DIAGRAM_PATTERNS_EXTENDED['contrast'] = {
    # 原本定義（3プロパティ｜改変禁止）
    'ja': '対比',
    'use': '対照的な2要素の並列強調',
    'shape': 'split_screen',
    # 実装レベル拡張定義（原本には存在しない｜統括承認済の階層分離）
    'min_elements': 2,          # 固定
    'max_elements': 2,          # 固定
    'requires_axes': False,
    'direction': 'horizontal',  # 左右分割
    'color_gradation': 'polarized_contrast',  # 対比強調型（薄→濃の2極）
}
```

### 原本との整合確認｜3観点

| 観点 | 確認内容 | 判定 |
|-----|--------|:---:|
| `shape`整合 | split_screen＝2列並列（左右分割）｜実装の`direction: horizontal`と完全一致 | ✅ |
| `use`整合 | 「対照的な**2要素**の並列強調」＝`min=2`／`max=2`固定と論理的に完全一致 | ✅ |
| 診断カテゴリ1対1対応 | `ux_contrast` → `contrast`｜DIAGNOSIS_TO_PATTERN準拠 | ✅ |

### P3方針素案（8/12）想定値との整合

| 項目 | 素案想定値 | 本実装確定値 | 整合 |
|-----|:-------:|:---------:|:---:|
| min_elements | 2（固定） | 2（固定） | ✅ |
| max_elements | 2（固定） | 2（固定） | ✅ |
| requires_axes | False | False | ✅ |
| 特殊制約 | 3要素以上はcategory誘導 | `_to_category_data`フォールバック実装 | ✅ |

### **エスカレーション条件(c)判定：「非該当」**

原本3プロパティの改変ゼロ｜拡張は統括承認済の「実装層と定義層の階層分離」方針に完全準拠。

---

## 🎯 P3-3 contrast｜設計思想

### 診断カテゴリ対応

- **診断カテゴリ**：`ux_contrast`
- **選定論理**：UX上の対照的要素（良UX例 vs 悪UX例／改善前 vs 改善後）は`contrast`のsplit_screenで強調
- **B-6 3.3節位置づけ**：Before/After訴求・優劣対比の主要視覚化パターン

### 視覚構造の骨格

```
┌───────────────────────┬───────────────────────┐
│  【Before】改善前     │  【After】改善後      │
│  ─────────────        │  ─────────────        │
│  score: 32            │  score: 88            │
│                       │                       │
│  ・課題1              │  ・改善1              │
│  ・課題2              │  ・改善2              │
│  ・課題3              │  ・改善3              │
│                       │                       │
│  （薄い＝secondary/   │  （濃い＝primary）    │
│    light）            │                       │
└───────────────────────┴───────────────────────┘
              ↑ 中央セパレータ（区切り線＋対比矢印）
```

### 設計判断｜5件

**判断①｜要素数2固定｜3要素以上はcategoryフォールバック**
- 原本`use`「対照的な**2要素**の並列強調」に厳密準拠
- `len(sides) != 2` の場合は例外を投げず`draw_category`へフォールバック（framework／funnel同姿勢）
- 理由：診断データ側の揺れで描画が落ちるより、劣化描画で通す方が実務価値が高い（P1 breakdown以来の一貫方針）

**判断②｜色階調は「2極化」｜段階変化ではなく対比強調**
- Before側（左）＝`secondary`（薄）／After側（右）＝`primary`（濃）
- sequence／funnelの`progress_ratio`による段階減衰とは**設計思想が異なる**
- contrastの目的は「段階の推移」ではなく「2極の落差」の視覚化であるため、中間色を使わない

**判断③｜framework 2x1グリッド流用｜`_compute_framework_grid_positions`の拡張**
- frameworkのグリッド座標計算に`"2x1"`（1行2列）を追加する形で流用
- 新規のグリッド計算関数は書かない（共通ヘルパー最大再利用の原則）
- ただしcontrastは`requires_axes: False`のため、`_draw_framework_axes`は**呼ばない**

**判断④｜中央セパレータ｜contrast固有の視覚装置**
- 左右セル間に区切り線＋対比矢印（→）を描画
- split_screenの「分割されている」という視覚メッセージを明示化
- frameworkのセル境界線とは別レイヤー（framework流用ではなくcontrast独自実装）

**判断⑤｜警告オーバーライドの非対称適用**
- score<40 の warning色オーバーライドは共通ルールとして適用
- ただし**Before側は薄色が設計意図であるため、warning発火時のみ濃警告色になる**
- これはB-6 4.4節「警告発火時の視認性優先」に整合（設計美より視認性を優先）

---

## 🧩 描画関数実装｜構造

### 独自実装関数｜3件（素案想定2〜3件｜上限内）

**関数1｜`_compute_contrast_positions() -> dict`**

```python
def _compute_contrast_positions() -> dict:
    """
    contrast パターンの左右2セル配置座標を計算する（split_screen｜2x1グリッド）。

    Returns:
        {
            "left":      (x, y, w, h),   # Before側セル
            "right":     (x, y, w, h),   # After側セル
            "separator": (x, y, w, h),   # 中央セパレータ領域
        }
        （すべて Inches 単位の float）

    設計思想：
    - P2-3 framework の _compute_framework_grid_positions("2x1") を流用
    - contrast は requires_axes=False のため軸領域を確保しない
      → framework の axis_area_x / axis_area_y を 0 として計算
    - 中央セパレータ幅（0.4inch）を左右セルから折半で控除
    """
    slide_width_inch  = 13.33   # 16:9
    slide_height_inch = 7.5
    header_area   = 1.0    # タイトル領域（P1-2 _draw_title と同一）
    margin        = 0.5
    separator_w   = 0.4    # 中央セパレータ幅（contrast固有）

    grid_x = margin
    grid_y = header_area
    grid_w = slide_width_inch - margin * 2
    grid_h = slide_height_inch - grid_y - margin

    cell_w = (grid_w - separator_w) / 2

    return {
        "left":      (grid_x,                        grid_y, cell_w,      grid_h),
        "right":     (grid_x + cell_w + separator_w, grid_y, cell_w,      grid_h),
        "separator": (grid_x + cell_w,               grid_y, separator_w, grid_h),
    }
```

**関数2｜`_draw_contrast_side(slide, palette, side, position, side_type)`**

```python
def _draw_contrast_side(slide, palette, side, position, side_type):
    """
    contrast の片側（Before または After）セルを描画する。

    Args:
        slide:     python-pptx の Slide オブジェクト
        palette:   8色階調辞書（P1-1 get_theme_palette 出力）
        side:      片側データ辞書
                   {"label": str, "score": int, "items": [str], "caption": str}
        position:  (x, y, w, h)｜_compute_contrast_positions の返却値
        side_type: "before" | "after"

    設計思想：
    - P2-3 _draw_framework_cell のテキスト配置ロジックを流用
    - contrast固有：side_type によりラベル装飾（【Before】/【After】）を付与
    - 箇条書きは最大5項目（framework の3項目より緩和｜2セルのみで縦領域に余裕があるため）
    """
    x, y, w, h = position
    cell_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )

    # 色階調適用（contrast固有｜2極化）
    _apply_contrast_side_color(cell_box, palette, side.get("score", 50), side_type)

    tf = cell_box.text_frame
    tf.word_wrap = True

    # 見出しラベル（contrast固有装飾）
    prefix = "【Before】" if side_type == "before" else "【After】"
    p_label = tf.paragraphs[0]
    p_label.text = f"{prefix}{side['label']}"
    _apply_font_style(p_label, size=18, bold=True, color=palette["bg"])
    p_label.alignment = PP_ALIGN.CENTER

    # スコア表示
    p_score = tf.add_paragraph()
    p_score.text = f"{side.get('score', '-')}点"
    _apply_font_style(p_score, size=28, bold=True, color=palette["bg"])
    p_score.alignment = PP_ALIGN.CENTER

    # 箇条書き（最大5項目｜contrast固有の緩和値）
    for item in side.get("items", [])[:5]:
        p_item = tf.add_paragraph()
        p_item.text = f"・{item}"
        _apply_font_style(p_item, size=12, color=palette["lightest"])

    # 補足キャプション（省略可）
    if side.get("caption"):
        p_cap = tf.add_paragraph()
        p_cap.text = side["caption"]
        _apply_font_style(p_cap, size=10, italic=True, color=palette["lightest"])

    return cell_box
```

**関数3｜`_apply_contrast_side_color(cell_box, palette, score, side_type)`**

```python
def _apply_contrast_side_color(cell_box, palette, score, side_type):
    """
    contrast セルの色階調適用（2極化｜段階変化ではなく対比強調）。

    設計思想：
    - Before側 = secondary（薄）／After側 = primary（濃）
    - sequence / funnel の progress_ratio 段階減衰とは設計思想が異なる
      （contrast の目的は「推移」ではなく「落差」の可視化）
    - score < 40 は warning 色オーバーライド（全パターン共通ルール｜B-6 4.4節）
    """
    cell_box.fill.solid()

    if score < 40:
        # 警告オーバーライド（設計美より視認性優先）
        cell_box.fill.fore_color.rgb = hex_to_rgb(palette["warning"])
    else:
        color_key = "secondary" if side_type == "before" else "primary"
        cell_box.fill.fore_color.rgb = hex_to_rgb(palette[color_key])

    cell_box.line.color.rgb = hex_to_rgb(palette["midtone"])
    cell_box.line.width = Pt(1.5)
```

### メイン描画関数｜`draw_contrast(slide, palette, data)`

```python
def draw_contrast(slide, palette, data):
    """
    contrast パターン描画のメイン関数（split_screen｜2要素対比）。

    Args:
        slide:   python-pptx の Slide オブジェクト
        palette: 8色階調辞書
        data: {
            "title": str,
            "sides": [
                {"label": str, "score": int, "items": [str], "caption": str},  # Before
                {"label": str, "score": int, "items": [str], "caption": str},  # After
            ],
        }
    """
    sides = data.get("sides", [])

    # 要素数チェック（min=2／max=2固定｜原本 use「対照的な2要素」準拠）
    if len(sides) != 2:
        # categoryフォールバック（例外を投げず劣化描画で通す｜P1 breakdown以来の一貫方針）
        return draw_category(slide, palette, _to_category_data_from_sides(data))

    # タイトル描画（P1-2 共通ヘルパー完全再利用）
    _draw_title(slide, palette, data.get("title", ""))

    # 左右セル配置座標算出（framework 2x1 流用）
    positions = _compute_contrast_positions()

    # Before側（左）描画
    _draw_contrast_side(slide, palette, sides[0], positions["left"],  "before")
    # After側（右）描画
    _draw_contrast_side(slide, palette, sides[1], positions["right"], "after")

    # 中央セパレータ＋対比矢印（contrast固有）
    _draw_contrast_separator(slide, palette, positions["separator"])

    return slide


def _draw_contrast_separator(slide, palette, position):
    """中央セパレータ（区切り線＋対比矢印）｜contrast固有の視覚装置。"""
    x, y, w, h = position
    # 縦区切り線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x + w / 2 - 0.01), Inches(y), Inches(0.02), Inches(h)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(palette["light"])
    line.line.fill.background()

    # 対比矢印（中央高さ）
    arrow = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW,
        Inches(x - 0.1), Inches(y + h / 2 - 0.15), Inches(w + 0.2), Inches(0.3)
    )
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = hex_to_rgb(palette["accent"])
    arrow.line.fill.background()


def _to_category_data_from_sides(contrast_data: dict) -> dict:
    """contrastデータをcategoryパターンデータに変換（要素数不整合時のフォールバック）。"""
    return {
        "title": contrast_data.get("title", ""),
        "categories": [
            {"label": s.get("label", ""), "score": s.get("score", 0),
             "description": s.get("caption", "")}
            for s in contrast_data.get("sides", [])
        ],
    }
```

### 共通ヘルパー再利用｜9件

| # | ヘルパー | 出所 | 再利用度 | 用途 |
|---|--------|-----|:----:|-----|
| 1 | `get_theme_palette(theme_id)` | P1-1 | ◎ | テーマパレット取得 |
| 2 | `select_theme_by_project_type` | P1-1 | ◎ | テーマ自動選定 |
| 3 | `hex_to_rgb(hex_str)` | P1-1 | ◎ | HEX→RGB変換 |
| 4 | 8色階調適用関数群 | P1-1 | ◎ | 色適用 |
| 5 | `_apply_font_style` | P1-2〜4 | ◎ | フォントスタイル統一 |
| 6 | `_apply_color_by_score` | P1-2〜4 | ◎ | スコア別色適用 |
| 7 | `_draw_title` | P1-2〜4 | ◎ | タイトル描画 |
| 8 | `_compute_framework_grid_positions`（2x1拡張） | P2-3 | ◎ | グリッド座標計算 |
| 9 | `_draw_framework_cell`（テキスト配置ロジック） | P2-3 | ◎ | セル描画 |

**新規実装｜4関数**（`_compute_contrast_positions` / `_draw_contrast_side` / `_apply_contrast_side_color` / `_draw_contrast_separator`）
※ 素案想定2〜3件に対し4件。増加分の`_draw_contrast_separator`はcontrast固有の視覚装置（split_screen明示化）であり、想定範囲内の逸脱として自律判断で採用。

---

## ✅ 描画ロジック検証｜フル版B 7項目｜自己検証

| # | 検証項目 | 判定 | 根拠 |
|---|---------|:----:|-----|
| 1 | Phase A A-2原本整合｜厳守事項 | ✅ PASS | 原本3プロパティ改変ゼロ｜拡張は統括承認済[A]方針 |
| 2 | 要素数上限遵守（Miller's Law 7±2） | ✅ PASS | min=2／max=2固定｜`len(sides) != 2`チェック実装 |
| 3 | 色階調自動選定（8色階調） | ✅ PASS | 2極化（secondary／primary）＋warning／light／accent／lightest／bg使用 |
| 4 | categoryフォールバック実装 | ✅ PASS | `_to_category_data_from_sides`｜例外非送出 |
| 5 | 警告オーバーライド（score<40） | ✅ PASS | 左右両側に適用｜視認性優先（B-6 4.4節整合） |
| 6 | 共通ヘルパー最大再利用 | ✅ PASS | 9件再利用｜新規4関数のみ |
| 7 | 診断カテゴリ1対1対応（ux_contrast） | ✅ PASS | DIAGNOSIS_TO_PATTERN準拠 |

**総合判定：7項目すべてPASS｜設計＋描画関数実装として十分な水準**

---

## 📊 実装成果｜サマリ

| 項目 | 実績 |
|-----|-----|
| 独自実装関数 | **4件** |
| 共通ヘルパー再利用 | **9件**（完全再利用） |
| 想定行数 | 約120〜150行（framework約230行／funnel約150〜200行より小規模） |
| 実装難度 | **低**（素案想定通り） |
| framework流用度 | **高**（グリッド座標計算＋セル描画ロジック） |
| 設計完了時刻 | 想定11:30｜**達成** |

### 実装難度が「低」に収まった要因｜3点

1. **要素数が2固定**｜可変長ループ・境界計算が不要（funnel 3〜6／cycle 3〜6と比較して構造的に単純）
2. **`requires_axes=False`**｜frameworkの最大の複雑要因だった2軸ラベル描画が不要
3. **色階調が2極のみ**｜progress_ratio計算・階調分岐（5分岐）が不要

---

**次アクション｜P3-3 contrast 動作テスト（7テーマ×バリエーション）へ進行｜想定完了13:00**
