# v3.5コアP3-2｜cycle パターン｜設計＋描画関数実装ドキュメント

- 作成日：2026-08-14（金）
- 作成者：AIスライド（実装領域）
- 位置づけ：**v3.5コアP3実装2日目｜P3-2 cycle｜設計＋描画関数実装記録**
- 統括承認：議論日2 議題4｜提案A（流用度優先）採用｜5判定発行済
- **⚠️ 実装難度：中｜(b)エスカレーション条件監視強化対象**（起動プロンプト重点②）
- 想定完了時刻：**17:00**

---

## 📋 実装着手前｜Phase A A-2原本整合確認｜完了（厳守①対応）

### 原本定義（`phase_a_design_20260727_rev2.md` A-2節｜逐語）

```python
DIAGRAM_PATTERNS = {
    'cycle': {'ja': '循環', 'use': '反復プロセスを円環で表現', 'shape': 'circle_arrow'},
}

DIAGNOSIS_TO_PATTERN = {
    'improvement_cycle': 'cycle',
}
```

### 実装レベル拡張定義（本実装での確定値）

```python
DIAGRAM_PATTERNS_EXTENDED['cycle'] = {
    # 原本定義（3プロパティ｜改変禁止）
    'ja': '循環',
    'use': '反復プロセスを円環で表現',
    'shape': 'circle_arrow',
    # 実装レベル拡張定義（統括承認済[A]方針｜階層分離）
    'min_elements': 3,
    'max_elements': 6,
    'requires_axes': False,
    'direction': 'clockwise',           # 時計回り（12時起点）
    'color_gradation': 'uniform_cyclic', # 均等配分型（継続性強調）
}
```

### 原本との整合確認｜3観点

| 観点 | 確認内容 | 判定 |
|-----|--------|:---:|
| `shape`整合 | circle_arrow＝円周配置＋弧状矢印｜実装の`_draw_arc_arrow`と完全一致 | ✅ |
| `use`整合 | 「**反復**プロセスを**円環**で表現」＝循環閉包（最終要素→初期要素の矢印）必須と整合 | ✅ |
| 診断カテゴリ1対1対応 | `improvement_cycle` → `cycle`｜DIAGNOSIS_TO_PATTERN準拠 | ✅ |

### P3方針素案（8/12）想定値との整合

| 項目 | 素案想定値 | 本実装確定値 | 整合 |
|-----|:-------:|:---------:|:---:|
| min_elements | 3 | 3 | ✅ |
| max_elements | 6 | 6 | ✅ |
| requires_axes | False | False | ✅ |
| 特殊制約 | 循環閉包｜円周上等間隔配置 | 実装（`_compute_cycle_positions`＋閉包矢印） | ✅ |

### **エスカレーション条件(c)判定：「非該当」**

---

## 🎯 P3-2 cycle｜設計思想

### 診断カテゴリ対応

- **診断カテゴリ**：`improvement_cycle`
- **選定論理**：継続改善サイクル（PDCA等）は`cycle`のcircle_arrowで反復性を強調
- **B-6 3.3節位置づけ**：改善提案の「継続的な繰り返し」を視覚化する**12パターン中唯一**のパターン

### 視覚構造の骨格（4要素の例）

```
              段階1（12時位置）
           ╭──────↘
    段階4              段階2
 （9時位置）        （3時位置）
           ╰──────↗
              段階3（6時位置）

   ※ 各段階間は弧状矢印（時計回り）
   ※ 段階4→段階1の閉包矢印で「反復」を明示
```

### 設計判断｜6件

**判断①｜12時起点・時計回り固定**
- 円周上の第1要素は必ず12時位置（-90°）に配置
- 時計回り（clockwise）固定｜反時計回りオプションは実装しない
- 理由：PDCAサイクルの視覚慣習が時計回りで統一されているため、選択肢を持たせると「選定判断の曖昧さ」がB-6判断1（選定の透明性）に反する

**判断②｜色階調は「均等配分」｜段階減衰ではない**
- funnel／sequence の`progress_ratio`段階減衰とは**設計思想が根本的に異なる**
- cycleは「始まりも終わりもない継続」を表現するため、**どの段階も同格**
- 実装：要素数nに対し、8色階調から`primary`／`secondary`／`midtone`／`light`を循環的に割当
- **これがcycle実装の最重要設計判断**｜段階減衰を適用すると「最後の段階が最も薄い＝終わりがある」という誤ったメッセージになる

**判断③｜円周座標計算は新規実装（sequence流用は不可）**
- sequenceの`_compute_sequence_positions`は1軸（直線）配置｜円周配置には数学的に流用不可
- 三角関数による極座標→直交座標変換を新規実装
- **これが本パターンを「中難度」たらしめる主因**

**判断④｜弧状矢印はMSO_SHAPE.BLOCK_ARC＋回転で実装**
- python-pptx標準図形の`MSO_SHAPE.BLOCK_ARC`（円弧帯）を各段階間に配置し`rotation`で角度調整
- **代替案として検討し却下したもの**：
  - `MSO_SHAPE.CIRCULAR_ARROW`（単体）｜1本の円弧矢印しか描けず、n分割できない
  - フリーフォーム（`add_freeform`）でベジェ曲線描画｜制御点計算が複雑化しリスク②水準に接近するため却下
- **リスク低減判断**：BLOCK_ARCの`adjustments`が期待通り機能しない環境が想定されるため、**フォールバックとして直線矢印（`MSO_SHAPE.RIGHT_ARROW`＋rotation）を用意**

**判断⑤｜中央にサイクル名テキストを配置（cycle固有）**
- 円環の中心に空白ができるため、サイクル名（例：「PDCA」「継続改善サイクル」）を配置
- 情報密度の向上と、円環構造の視覚的重心の確保を両立

**判断⑥｜categoryフォールバック（要素数逸脱時）**
- `n < 3` または `n > 6` の場合、例外を投げず`draw_category`へフォールバック
- P1 breakdown以来の一貫方針を継続

---

## 🧩 描画関数実装｜構造

### 独自実装関数｜5件（素案想定4〜5件｜上限内）

**関数1｜`_compute_cycle_positions(n: int) -> dict`**

```python
import math

def _compute_cycle_positions(n: int) -> dict:
    """
    n個のcycle段階を円周上に等間隔配置した座標を計算する。

    Args:
        n: 段階数（3〜6）

    Returns:
        {
            "steps": [(x, y, w, h), ...],       # 各段階ボックス（左上基準｜Inches）
            "arcs":  [(cx, cy, r, a1, a2), ...],# 段階間の弧（中心・半径・開始角・終了角｜度）
            "center": (x, y, w, h),             # 中央テキスト領域
        }

    設計思想（判断③｜新規実装）：
    - 極座標 → 直交座標変換｜θ_i = -90° + (360°/n) * i （12時起点・時計回り）
    - sequence._compute_sequence_positions は1軸配置のため流用不可
    - 段階ボックスは「円周上の点」を中心とする矩形として配置（中心基準→左上基準に変換）
    """
    slide_w, slide_h = 13.33, 7.5     # 16:9
    header_area = 1.0                  # タイトル領域（P1-2 _draw_title と同一）
    margin      = 0.5

    # 円環の中心と半径
    field_h = slide_h - header_area - margin
    cx = slide_w / 2
    cy = header_area + field_h / 2
    radius = min(field_h / 2 - 0.75, 2.6)   # 段階ボックス半分を控除

    # 段階ボックスサイズ（要素数により可変｜多いほど小さく）
    box_w = 2.6 if n <= 4 else 2.2
    box_h = 1.3 if n <= 4 else 1.1

    steps, arcs = [], []
    for i in range(n):
        theta_deg = -90 + (360 / n) * i        # 12時起点・時計回り
        theta = math.radians(theta_deg)
        px = cx + radius * math.cos(theta)
        py = cy + radius * math.sin(theta)
        # 中心基準 → 左上基準
        steps.append((px - box_w / 2, py - box_h / 2, box_w, box_h))

        # 段階 i → i+1（最終要素は 0 へ｜循環閉包）の弧
        a1 = theta_deg + (360 / n) * 0.22      # ボックス端を避けるオフセット
        a2 = theta_deg + (360 / n) * 0.78
        arcs.append((cx, cy, radius, a1, a2))

    center = (cx - 1.4, cy - 0.5, 2.8, 1.0)
    return {"steps": steps, "arcs": arcs, "center": center}
```

**関数2｜`_draw_cycle_step(slide, palette, step, index, n, position)`**

```python
def _draw_cycle_step(slide, palette, step, index, n, position):
    """
    cycle の1段階（円周上のボックス）を描画する。

    Args:
        step: {"label": str, "score": int, "description": str}
        index: 段階インデックス（0-indexed）
        n: 総段階数
        position: (x, y, w, h)

    設計思想：
    - P2-2 _draw_sequence_step のテキスト配置ロジックを流用（○｜素案評価通り）
    - 形状は MSO_SHAPE.ROUNDED_RECTANGLE（円環内で角が立たないよう配慮）
    """
    x, y, w, h = position
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )

    _apply_cycle_step_color(box, palette, index, n, step.get("score", 50))

    tf = box.text_frame
    tf.word_wrap = True

    # 段階番号＋ラベル
    p_label = tf.paragraphs[0]
    p_label.text = f"{index + 1}. {step['label']}"
    _apply_font_style(p_label, size=14, bold=True, color=palette["bg"])
    p_label.alignment = PP_ALIGN.CENTER

    # 説明（省略可）
    if step.get("description"):
        p_desc = tf.add_paragraph()
        p_desc.text = step["description"]
        _apply_font_style(p_desc, size=10, color=palette["lightest"])
        p_desc.alignment = PP_ALIGN.CENTER

    return box
```

**関数3｜`_draw_arc_arrow(slide, palette, arc, use_fallback=False)`**

```python
def _draw_arc_arrow(slide, palette, arc, use_fallback=False):
    """
    段階間の弧状矢印を描画する（cycle固有｜最難関実装）。

    Args:
        arc: (cx, cy, r, a1, a2)｜中心・半径・開始角・終了角（度）
        use_fallback: True の場合、BLOCK_ARC ではなく直線矢印で代替

    設計思想（判断④）：
    - 主実装：MSO_SHAPE.BLOCK_ARC（円弧帯）｜外接矩形に配置し rotation で角度合わせ
    - フォールバック：MSO_SHAPE.RIGHT_ARROW を弦の中点に配置＋接線方向へ rotation
      （BLOCK_ARC の adjustments が環境依存で機能しない場合のリスク低減策）
    """
    cx, cy, r, a1, a2 = arc

    if not use_fallback:
        # 主実装｜BLOCK_ARC を円の外接矩形に配置
        arc_shape = slide.shapes.add_shape(
            MSO_SHAPE.BLOCK_ARC,
            Inches(cx - r), Inches(cy - r), Inches(r * 2), Inches(r * 2)
        )
        # adjustments[0]=開始角 / [1]=終了角 / [2]=内径比（0.0〜1.0）
        try:
            arc_shape.adjustments[0] = a1 / 360.0
            arc_shape.adjustments[1] = a2 / 360.0
            arc_shape.adjustments[2] = 0.92   # 細い弧帯
        except (IndexError, AttributeError):
            # adjustments 非対応環境｜フォールバックへ切替
            arc_shape._element.getparent().remove(arc_shape._element)
            return _draw_arc_arrow(slide, palette, arc, use_fallback=True)
        arc_shape.fill.solid()
        arc_shape.fill.fore_color.rgb = hex_to_rgb(palette["accent"])
        arc_shape.line.fill.background()
        return arc_shape

    # フォールバック｜弦の中点に直線矢印
    mid_deg = (a1 + a2) / 2
    mid = math.radians(mid_deg)
    ax = cx + r * math.cos(mid) - 0.35
    ay = cy + r * math.sin(mid) - 0.10
    arrow = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW, Inches(ax), Inches(ay), Inches(0.7), Inches(0.2)
    )
    arrow.rotation = mid_deg + 90          # 接線方向（時計回り進行）
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = hex_to_rgb(palette["accent"])
    arrow.line.fill.background()
    return arrow
```

**関数4｜`_apply_cycle_step_color(box, palette, index, n, score)`**

```python
def _apply_cycle_step_color(box, palette, index, n, score):
    """
    cycle段階の色階調適用（均等配分型｜判断②）。

    設計思想：
    - funnel/sequence の progress_ratio 段階減衰は **適用しない**
      （「最後が最も薄い＝終わりがある」という誤メッセージを避けるため）
    - 4色（primary/secondary/midtone/light）を index % 4 で循環割当
      → どの段階も同格｜「始まりも終わりもない継続」を表現
    - score < 40 は warning 色オーバーライド（全パターン共通｜B-6 4.4節）
    """
    box.fill.solid()

    if score < 40:
        box.fill.fore_color.rgb = hex_to_rgb(palette["warning"])
    else:
        cyclic_keys = ["primary", "secondary", "midtone", "light"]
        box.fill.fore_color.rgb = hex_to_rgb(palette[cyclic_keys[index % 4]])

    box.line.color.rgb = hex_to_rgb(palette["midtone"])
    box.line.width = Pt(1.5)
```

**関数5｜`_draw_cycle_center(slide, palette, position, cycle_name)`**

```python
def _draw_cycle_center(slide, palette, position, cycle_name):
    """円環中央にサイクル名を描画する（cycle固有｜判断⑤）。"""
    x, y, w, h = position
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    p = tb.text_frame.paragraphs[0]
    p.text = cycle_name
    _apply_font_style(p, size=20, bold=True, color=palette["primary"])
    p.alignment = PP_ALIGN.CENTER
    return tb
```

### メイン描画関数｜`draw_cycle(slide, palette, data)`

```python
def draw_cycle(slide, palette, data):
    """
    cycle パターン描画のメイン関数（circle_arrow｜反復プロセスの円環表現）。

    Args:
        data: {
            "title": str,
            "cycle_name": str,   # 円環中央テキスト（省略時 "" ）
            "steps": [{"label": str, "score": int, "description": str}, ...],
        }
    """
    steps = data.get("steps", [])
    n = len(steps)

    # 要素数チェック（Miller's Law遵守｜3〜6段階）
    if not (3 <= n <= 6):
        return draw_category(slide, palette, _to_category_data_from_steps(data))

    # タイトル描画（P1-2 共通ヘルパー完全再利用）
    _draw_title(slide, palette, data.get("title", ""))

    # 円周配置座標算出（新規実装）
    pos = _compute_cycle_positions(n)

    # 弧状矢印を先に描画（段階ボックスの背面に配置するため）
    for arc in pos["arcs"]:
        _draw_arc_arrow(slide, palette, arc)

    # 各段階を描画
    for i, (step, position) in enumerate(zip(steps, pos["steps"])):
        _draw_cycle_step(slide, palette, step, i, n, position)

    # 中央サイクル名（省略可）
    if data.get("cycle_name"):
        _draw_cycle_center(slide, palette, pos["center"], data["cycle_name"])

    return slide


def _to_category_data_from_steps(cycle_data: dict) -> dict:
    """cycleデータをcategoryパターンデータに変換（要素数逸脱時のフォールバック）。"""
    return {
        "title": cycle_data.get("title", ""),
        "categories": [
            {"label": s.get("label", ""), "score": s.get("score", 0),
             "description": s.get("description", "")}
            for s in cycle_data.get("steps", [])
        ],
    }
```

### 共通ヘルパー再利用｜8件

| # | ヘルパー | 出所 | 再利用度 | 用途 |
|---|--------|-----|:----:|-----|
| 1 | `get_theme_palette(theme_id)` | P1-1 | ◎ | テーマパレット取得 |
| 2 | `select_theme_by_project_type` | P1-1 | ◎ | テーマ自動選定 |
| 3 | `hex_to_rgb(hex_str)` | P1-1 | ◎ | HEX→RGB変換 |
| 4 | 8色階調適用関数群 | P1-1 | ◎ | 色適用 |
| 5 | `_apply_font_style` | P1-2〜4 | ◎ | フォントスタイル統一 |
| 6 | `_draw_title` | P1-2〜4 | ◎ | タイトル描画 |
| 7 | `_draw_sequence_step`（テキスト配置ロジック） | P2-2 | ○ | 段階ボックス描画 |
| 8 | `_to_category_data`（変換思想） | P2-3 | ○ | フォールバック |

**新規実装｜5関数**（`_compute_cycle_positions` / `_draw_cycle_step` / `_draw_arc_arrow` / `_apply_cycle_step_color` / `_draw_cycle_center`）｜素案想定4〜5件の上限内

---

## ⚠️ リスク②水準｜監視結果｜(b)エスカレーション判定

起動プロンプト重点②により、cycle実装中は(b)条件（実装難度がP1リスク②水準を上回る徴候）を監視強化した。

### 監視対象｜3徴候

| # | 徴候 | 検出状況 | 評価 |
|---|-----|:-----:|-----|
| 徴候1｜設計途中での方針転換発生 | **なし** | 判断①〜⑥は初期設計から変更なし | ✅ 正常 |
| 徴候2｜新規実装関数が想定を超過 | **なし** | 素案想定4〜5件に対し実績5件｜上限内 | ✅ 正常 |
| 徴候3｜python-pptx標準機能で実装不能な要素の出現 | **軽微あり** | BLOCK_ARCの`adjustments`が環境依存の可能性 | 🟡 対策済 |

### 徴候3｜詳細と対策

**検出内容**：`MSO_SHAPE.BLOCK_ARC`の`adjustments`プロパティは、python-pptxのバージョン・図形プリセットにより添字数が異なる場合がある（`adjustments[2]`が存在しない可能性）。

**対策（設計内で吸収）**：
- `try/except (IndexError, AttributeError)`で捕捉し、直線矢印フォールバックへ自動切替
- フォールバック実装も同一関数内に内包｜呼出側は切替を意識不要
- **これによりリスクは設計レベルで封じ込め済**

### **(b)判定：「非該当」**

- 徴候3は「実装不能」ではなく「環境依存の可能性」に留まり、フォールバックで完全吸収済
- P1リスク②（breakdown実装時の構造的複雑性）水準には**達していない**
- **厳守④の「cycle完遂を8/15に持ち越す判断」は不要｜本日完遂で進行**

---

## ✅ 描画ロジック検証｜フル版B 7項目｜自己検証

| # | 検証項目 | 判定 | 根拠 |
|---|---------|:----:|-----|
| 1 | Phase A A-2原本整合｜厳守事項 | ✅ PASS | 原本3プロパティ改変ゼロ｜拡張は統括承認済[A]方針 |
| 2 | 要素数上限遵守（Miller's Law 7±2） | ✅ PASS | min=3／max=6｜`3 <= n <= 6`チェック実装 |
| 3 | 色階調自動選定（8色階調） | ✅ PASS | 均等配分4色循環＋warning／accent／lightest／bg使用 |
| 4 | categoryフォールバック実装 | ✅ PASS | `_to_category_data_from_steps`｜例外非送出 |
| 5 | 警告オーバーライド（score<40） | ✅ PASS | 循環色割当より優先｜視認性優先 |
| 6 | 共通ヘルパー最大再利用 | ✅ PASS | 8件再利用｜新規5関数（素案上限内） |
| 7 | 診断カテゴリ1対1対応（improvement_cycle） | ✅ PASS | DIAGNOSIS_TO_PATTERN準拠 |

**総合判定：7項目すべてPASS**

---

## 📊 実装成果｜サマリ

| 項目 | 実績 |
|-----|-----|
| 独自実装関数 | **5件**（素案想定4〜5件｜上限内） |
| 共通ヘルパー再利用 | **8件** |
| 想定行数 | 約190〜220行（sequence約200行と同水準） |
| 実装難度 | **中**（素案想定通り） |
| sequence流用度 | **中**（ステップ描画のみ｜円周座標は新規） |
| フォールバック機構 | **2系統**（categoryフォールバック／弧状矢印フォールバック） |
| 設計完了時刻 | 想定15:00｜**達成** |

### cycle実装の技術的核心｜3点

1. **極座標変換**｜`θ_i = -90° + (360°/n)·i`｜12時起点・時計回りを数式で固定
2. **均等配分色階調**｜段階減衰を意図的に排除し「継続性」を表現（設計思想レベルの判断）
3. **二重フォールバック**｜データ不正（category）／描画環境不整合（直線矢印）の両方に対応

---

**次アクション｜P3-2 cycle 動作テスト（7テーマ×バリエーション）へ進行｜想定完了17:00**
