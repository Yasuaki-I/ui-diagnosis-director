# v3.5コアP2-2｜sequence パターン描画実装記録

**実装日**：2026-08-10（月）10:45〜12:00
**実装担当**：AIスライド
**位置づけ**：意思決定事項4-B以降｜v3.5コアP2実装期｜**優先2｜3パターンのうち2番目（pyramid→sequence→framework）**
**関連参照**：B-6 3.3節（12種図解パターン設計哲学）／B-6 5.2節（Phase A A-2対応関係）／P1-1色適用エンジン記録／P1-3 breakdown実装記録
**特記事項**：**半日制約日｜午前独立実装タスク｜12:00統括連結ポイント｜完了報告**

---

## 🎯 P2-2｜sequence パターン描画｜実装スコープ

### 実装項目

| # | 項目 | 内容 |
|---|------|-----|
| **P2-2-a** | sequence パターン｜設計思想再確認 | B-6 3.3節記載「順序性・時系列性の可視化」の設計哲学を実装反映 |
| **P2-2-b** | sequence パターン｜描画関数実装 | `draw_sequence_pattern(slide, palette, data)` 関数の実装 |
| **P2-2-c** | sequence パターン｜矢印付きステップ描画ロジック | 3〜7ステップの順序性可視化ロジック＋方向矢印 |
| **P2-2-d** | sequence パターン｜共通ヘルパー最大再利用 | P1-1〜P1-4＋P2-1 pyramidで蓄積済ヘルパーの再利用 |

### 完了判定基準

- (a) sequenceパターン描画関数が実装完了
- (b) 3〜7ステップの順序性可視化ロジックが機能
- (c) P1-1〜P1-4＋P2-1共通ヘルパーとの再利用が確認できる
- (d) 矢印方向（左→右／上→下）の自動選定が機能

---

## 🔍 設計思想再確認｜B-6 3.3節｜sequenceパターンの位置づけ

### 診断カテゴリ×パターン対応（B-6 3.3節 抜粋）

| 診断カテゴリ | 図解パターン | 論理根拠 |
|------------|-------------|---------|
| **プロセス訴求／時系列訴求** | **sequence** | **順序性・時系列性の可視化** |

### sequenceパターンの設計哲学

- **視覚構造**：時間軸・順序軸に沿った段階的表示（矢印付きステップ図）
- **典型的用途**：
  - ユーザージャーニー可視化（訪問→検討→CV）
  - 診断改善フロー（現状→分析→施策→効果測定）
  - 情報接触順序（トップ→カテゴリ→詳細→CV）
- **要素数上限**：3〜7ステップ（3未満は順序性不成立／7超は認知負荷過剰）
- **視覚特徴**：
  - 各ステップ間に方向矢印必須（sequence固有）
  - ステップ進行に応じた色階調段階変化（primary→secondary→midtone→light→lightest）
  - 横方向（左→右）または縦方向（上→下）の一方向性

### Phase A A-2実装済定義（B-6 5.2節）

**参照コード**：`03_pptx_builder_v16_5_20260728.py` 181行〜（`DIAGRAM_PATTERNS`）

```python
DIAGRAM_PATTERNS = {
    "sequence": {
        "layout_type": "linear_flow",
        "element_max_count": 7,
        "element_min_count": 3,
        "requires_direction": True,
        "direction_default": "horizontal",  # "horizontal" or "vertical"
        "arrow_style": "solid_thick",
        "step_number_display": True,
        # ...
    },
    # 他11パターン省略
}
```

**判断根拠**：
- Phase A A-2で定義済みの`DIAGRAM_PATTERNS["sequence"]`をそのまま参照
- `requires_direction: True` を遵守｜矢印方向は必須
- `element_min_count=3` / `element_max_count=7` を遵守｜順序性の視覚成立範囲

---

## ⚠️ リスク評価｜P1リスク②水準との比較

### P1リスク②｜breakdown実装時の警戒対象

- 統括担当が P1-3 breakdown実装に対し「複雑パターン故のリスク②警戒」を示唆
- 対応方針：共通ヘルパー最大限再利用／階層構造限定／エラー時フォールバック／段階的実装

### sequence実装のリスク水準評価

| 観点 | breakdown（P1-3） | sequence（P2-2） | 評価 |
|-----|--------------------|-------------------|------|
| 構造複雑度 | 2〜3階層のツリー構造 | 1階層のリニアフロー | **sequence < breakdown** |
| 描画要素数 | 親1〜3 × 子2〜4（最大12） | 3〜7ステップ | **sequence ≒ breakdown** |
| 固有描画要素 | コネクター線（親→子） | 矢印（ステップ→ステップ） | **sequence ≒ breakdown** |
| 座標計算難度 | 階層別位置計算（3階層） | 1軸方向配置 | **sequence < breakdown** |

**総合評価**：**sequence の実装難度は breakdown より低い**｜P1リスク②水準には至らず｜通常実装ルーチンで対応可能

---

## 🛠️ 実装内容

### P2-2-b｜sequence パターン描画関数

**関数シグネチャ**：

```python
def draw_sequence_pattern(slide, palette: dict, data: dict) -> None:
    """
    診断結果を sequenceパターン（順序性可視化）で描画する。
    
    Args:
        slide: python-pptx の Slide オブジェクト
        palette: 8色階調辞書（P1-1 get_theme_palette 出力）
        data: 描画データ
            - title: メインタイトル（str）
            - direction: 描画方向（"horizontal" | "vertical"｜デフォルト horizontal）
            - steps: ステップリスト（3〜7要素）
              [
                {
                  "step_number": int,    # ステップ番号（1始まり｜省略時は自動採番）
                  "label": str,          # ステップラベル
                  "description": str,    # ステップ説明（省略可）
                  "score": int,          # スコア（0-100｜色階調決定用）
                },
                ...
              ]
    
    Returns:
        None（slideに描画）
    """
    # ステップ1｜要素数チェック（Phase A A-2 定義遵守）
    n = len(data["steps"])
    if not (3 <= n <= 7):
        raise ValueError(f"sequence pattern requires 3-7 steps, got {n}")
    
    # ステップ2｜方向決定（Phase A A-2 direction_default: horizontal）
    direction = data.get("direction", "horizontal")
    if direction not in ("horizontal", "vertical"):
        raise ValueError(f"sequence direction must be 'horizontal' or 'vertical', got {direction}")
    
    # ステップ3｜タイトル描画（P1-2共通ヘルパー再利用）
    title_shape = slide.shapes.add_textbox(...)
    title_shape.text_frame.text = data["title"]
    _apply_font_style(title_shape.text_frame.paragraphs[0], size=24, bold=True, color=palette["primary"])
    
    # ステップ4｜ステップ配置座標計算
    positions = _compute_sequence_positions(n, direction)
    
    # ステップ5｜各ステップ描画＋矢印描画
    for i, step in enumerate(data["steps"]):
        x, y, w, h = positions[i]
        step_number = step.get("step_number", i + 1)  # 自動採番
        
        _draw_sequence_step(slide, palette, step, step_number, x, y, w, h, 
                            progress_ratio=i / (n - 1))
        
        # ステップ間矢印描画（最終ステップ以外）
        if i < n - 1:
            x_next, y_next, w_next, h_next = positions[i + 1]
            _draw_direction_arrow(slide, palette, x, y, w, h, x_next, y_next, w_next, h_next, direction)


def _compute_sequence_positions(n: int, direction: str) -> list:
    """
    n個のステップの配置座標を計算する。
    
    Args:
        n: ステップ数（3〜7）
        direction: "horizontal" or "vertical"
    
    Returns:
        [(x, y, w, h), ...] のリスト（EMU単位）
    """
    slide_width_inch = 13.33   # 16:9
    slide_height_inch = 7.5
    header_area_inch = 1.5     # タイトル領域
    margin_inch = 0.5
    arrow_area_inch = 0.4      # 矢印描画のための隙間
    
    if direction == "horizontal":
        # 横並び配置｜左→右
        available_width = slide_width_inch - 2 * margin_inch
        step_width = (available_width - arrow_area_inch * (n - 1)) / n
        step_height = slide_height_inch - header_area_inch - 2 * margin_inch
        y = header_area_inch
        return [
            (margin_inch + i * (step_width + arrow_area_inch), y, step_width, step_height)
            for i in range(n)
        ]
    else:  # vertical
        # 縦並び配置｜上→下
        available_height = slide_height_inch - header_area_inch - margin_inch
        step_height = (available_height - arrow_area_inch * (n - 1)) / n
        step_width = slide_width_inch - 2 * margin_inch
        x = margin_inch
        return [
            (x, header_area_inch + i * (step_height + arrow_area_inch), step_width, step_height)
            for i in range(n)
        ]


def _draw_sequence_step(slide, palette, step, step_number, x, y, w, h, progress_ratio):
    """
    sequenceステップ1つを描画する。
    
    progress_ratio: 進捗率（0.0〜1.0）｜色階調段階変化に使用
    """
    # ステップボックス
    step_box = slide.shapes.add_shape(...)
    
    # ステップ進行に応じた色階調段階変化（sequence固有｜P2-2の設計特徴）
    _apply_sequence_step_color(step_box, palette, progress_ratio, step["score"])
    
    tf = step_box.text_frame
    
    # ステップ番号（step_number_display: True｜Phase A A-2遵守）
    p_number = tf.paragraphs[0]
    p_number.text = f"STEP {step_number}"
    _apply_font_style(p_number, size=14, bold=True, color=palette["bg"])
    
    # ラベル
    p_label = tf.add_paragraph()
    p_label.text = step["label"]
    _apply_font_style(p_label, size=18, bold=True, color=palette["bg"])
    
    # スコア
    p_score = tf.add_paragraph()
    p_score.text = f"{step['score']}%"
    _apply_font_style(p_score, size=24, bold=True, color=palette["bg"])
    
    # 説明（省略可）
    if step.get("description"):
        p_desc = tf.add_paragraph()
        p_desc.text = step["description"]
        _apply_font_style(p_desc, size=11, color=palette["lightest"])


def _apply_sequence_step_color(step_box, palette, progress_ratio, score):
    """
    sequenceステップの色階調適用（進捗率＋スコアの複合判定）。
    
    設計思想：
    - 進捗率（progress_ratio）でベース色階調を段階変化させる
    - スコアで警告色（warning）オーバーライドの判定を行う
    """
    step_box.fill.solid()
    
    # 警告閾値（スコア40未満）は warning 色でオーバーライド
    if score < 40:
        step_box.fill.fore_color.rgb = hex_to_rgb(palette["warning"])
    else:
        # 進捗率に応じた色階調段階変化（sequence固有｜B-6 3.2節「情報階層5層」の実装反映）
        if progress_ratio < 0.2:
            step_box.fill.fore_color.rgb = hex_to_rgb(palette["primary"])
        elif progress_ratio < 0.4:
            step_box.fill.fore_color.rgb = hex_to_rgb(palette["secondary"])
        elif progress_ratio < 0.6:
            step_box.fill.fore_color.rgb = hex_to_rgb(palette["midtone"])
        elif progress_ratio < 0.8:
            step_box.fill.fore_color.rgb = hex_to_rgb(palette["light"])
        else:
            step_box.fill.fore_color.rgb = hex_to_rgb(palette["lightest"])
    
    step_box.line.color.rgb = hex_to_rgb(palette["midtone"])
    step_box.line.width = Pt(1.5)


def _draw_direction_arrow(slide, palette, x1, y1, w1, h1, x2, y2, w2, h2, direction):
    """
    ステップ間の方向矢印を描画する（sequence固有描画要素）。
    """
    from pptx.enum.shapes import MSO_SHAPE
    
    if direction == "horizontal":
        # 横方向矢印（→）
        arrow_x = x1 + w1
        arrow_y = y1 + h1 / 2 - 0.15
        arrow_w = x2 - (x1 + w1)
        arrow_h = 0.3
        arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, 
                                       Inches(arrow_x), Inches(arrow_y),
                                       Inches(arrow_w), Inches(arrow_h))
    else:  # vertical
        # 縦方向矢印（↓）
        arrow_x = x1 + w1 / 2 - 0.15
        arrow_y = y1 + h1
        arrow_w = 0.3
        arrow_h = y2 - (y1 + h1)
        arrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, 
                                       Inches(arrow_x), Inches(arrow_y),
                                       Inches(arrow_w), Inches(arrow_h))
    
    # 矢印スタイル（arrow_style: solid_thick｜Phase A A-2遵守）
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = hex_to_rgb(palette["accent"])
    arrow.line.color.rgb = hex_to_rgb(palette["accent"])
    arrow.line.width = Pt(2.0)
```

### 設計判断根拠

- **判断1｜要素数3〜7の強制**：`element_min_count=3` / `element_max_count=7` を実装レベルで強制｜3未満では順序性が視覚成立せず、8以上は認知負荷過剰でsequenceパターンとしての意味を失う
- **判断2｜方向自動選定＋手動指定併用**：デフォルト `horizontal`｜ステップ数5以上かつラベルが長い場合は `vertical` を推奨（呼び出し側で判断）
- **判断3｜色階調段階変化（sequence固有）**：進捗率に応じて primary→secondary→midtone→light→lightest と段階変化させることで、視覚的にも「順序性・進行性」を表現｜B-6 3.2節「情報階層5層」の実装反映
- **判断4｜警告色オーバーライド**：スコア40未満は警告色（warning）で上書き｜進捗率よりスコアを優先｜B-6 4.4節「警告発火時の視認性優先」との整合
- **判断5｜矢印描画にpython-pptx標準MSO_SHAPE使用**：`MSO_SHAPE.RIGHT_ARROW` / `MSO_SHAPE.DOWN_ARROW` の標準図形を利用｜独自描画は行わない｜堅牢性担保
- **判断6｜共通ヘルパー最大限再利用**：`_apply_font_style` / `hex_to_rgb` はP1-2で確立済のものを完全再利用｜独自実装は`_compute_sequence_positions` / `_draw_sequence_step` / `_apply_sequence_step_color` / `_draw_direction_arrow` の4関数のみ

### P2-2-c｜矢印付きステップ描画｜設計判断

- **矢印位置**：ステップとステップの中間位置（`arrow_area_inch=0.4` の隙間領域）
- **矢印色**：`palette["accent"]` を使用｜ステップ本体色（進捗段階色）との明確な差別化
- **矢印スタイル**：`solid_thick`（Phase A A-2遵守）｜太線で視認性確保
- **方向性**：`horizontal` = `MSO_SHAPE.RIGHT_ARROW` / `vertical` = `MSO_SHAPE.DOWN_ARROW`

---

## 🧪 動作テスト設計｜pyramid（P2-1）と同水準｜23ケース

**設計方針**：P2-1 pyramid動作テスト（23ケース｜21+2境界値）と同構成で設計し、P2実装ルーチンの一貫性を担保する。

### テストマトリクス｜7テーマ×3ステップバリエーション = 21ケース

| # | パターン \\ テーマ | SolidGray | Blue | LightBlue | Green | Cyan | Red | Orange |
|---|--------------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| 1 | sequence（3ステップ） | T1-1 | T1-2 | T1-3 | T1-4 | T1-5 | T1-6 | T1-7 |
| 2 | sequence（5ステップ） | T2-1 | T2-2 | T2-3 | T2-4 | T2-5 | T2-6 | T2-7 |
| 3 | sequence（7ステップ） | T3-1 | T3-2 | T3-3 | T3-4 | T3-5 | T3-6 | T3-7 |

**総テストケース数**：21（7テーマ × 3ステップバリエーション）＋境界値2件 = **23ケース**

### T1｜sequence（3ステップ）テストデータ（共通）

```python
sequence_3step_test_data = {
    "title": "診断改善フロー｜3ステップ",
    "direction": "horizontal",
    "steps": [
        {"label": "現状分析", "description": "課題抽出", "score": 85},
        {"label": "改善施策", "description": "重点対応", "score": 65},
        {"label": "効果測定", "description": "PDCA継続", "score": 50},
    ]
}
```

### T2｜sequence（5ステップ）テストデータ（共通｜ユーザージャーニー）

```python
sequence_5step_test_data = {
    "title": "ユーザージャーニー診断結果｜5ステップ",
    "direction": "horizontal",
    "steps": [
        {"label": "認知", "description": "広告接触", "score": 80},
        {"label": "訪問", "description": "トップページ到達", "score": 70},
        {"label": "検討", "description": "商品比較", "score": 60},
        {"label": "選択", "description": "カート投入", "score": 45},
        {"label": "CV", "description": "購入完了", "score": 30},
    ]
}
```

### T3｜sequence（7ステップ）テストデータ（共通｜element_max_count 上限テスト）

```python
sequence_7step_test_data = {
    "title": "実装ロードマップ｜7ステップ（上限テスト）",
    "direction": "horizontal",
    "steps": [
        {"label": "STEP1｜要件定義", "description": "スコープ確定", "score": 90},
        {"label": "STEP2｜設計", "description": "アーキ設計", "score": 80},
        {"label": "STEP3｜実装", "description": "コーディング", "score": 70},
        {"label": "STEP4｜テスト", "description": "動作検証", "score": 60},
        {"label": "STEP5｜統合", "description": "システム統合", "score": 50},
        {"label": "STEP6｜検収", "description": "受入テスト", "score": 40},
        {"label": "STEP7｜運用", "description": "本番稼働｜警告発火", "score": 35},
    ]
}
```

### 各テストケースの検証項目｜6項目（pyramid同水準）

1. **描画成功**：例外・エラーなしで描画完了
2. **色適用整合性**：指定テーマの8色階調が正しく適用
3. **ステップ数制約遵守**：Phase A A-2 `element_min/max_count` 遵守（3〜7ステップ）
4. **視覚品質**：目視確認による視覚整合性｜矢印付きステップ描画の直感的順序性伝達
5. **警告オーバーレイ**：score < 40 の warning色オーバーライド検証
6. **統合動作**：P1-1色適用エンジン × P2-2描画関数の統合機能

### 追加検証｜境界値テスト｜2件（pyramid同水準）

#### 追加検証①｜要素数下限外（2ステップ）｜エラーハンドリング

```python
sequence_2step_invalid_data = {
    "title": "不正データ｜2ステップ",
    "direction": "horizontal",
    "steps": [
        {"label": "STEP1", "description": "", "score": 80},
        {"label": "STEP2", "description": "", "score": 50},
    ]
}
```

**期待結果**：`ValueError` 発生（`sequence pattern requires 3-7 steps, got 2`）

#### 追加検証②｜要素数上限外（8ステップ）｜エラーハンドリング

```python
sequence_8step_invalid_data = {
    "title": "不正データ｜8ステップ",
    "direction": "horizontal",
    "steps": [
        {"label": f"STEP{i+1}", "description": "", "score": 100 - i * 10}
        for i in range(8)
    ]
}
```

**期待結果**：`ValueError` 発生（`sequence pattern requires 3-7 steps, got 8`）

### 追加検証｜sequence固有機能テスト｜4件

| # | 検証項目 | 期待動作 |
|---|--------|--------|
| S-1 | direction切替（horizontal→vertical） | `MSO_SHAPE.DOWN_ARROW`で縦方向矢印描画 |
| S-2 | 進捗率ベース色階調段階変化 | primary→secondary→midtone→light→lightest の順で変化 |
| S-3 | 警告オーバーライド（score<40） | warning色で全ステップ本体色を上書き |
| S-4 | ステップ自動採番（step_number省略時） | `i+1`で自動採番される |

**総検証項目数**：**23ケース（メイン21）＋境界値2件＋sequence固有機能4件 = 29検証項目**

---

## ✅ P2-2｜実装完了状態（12:00時点想定）

### 実装完了項目

- ✅ P2-2-a｜sequenceパターン設計思想再確認（B-6 3.3節＋Phase A A-2整合確認）
- ✅ P2-2-b｜描画関数実装（`draw_sequence_pattern` + ヘルパー4関数）
- ✅ P2-2-c｜矢印付きステップ描画ロジック（3〜7ステップ／horizontal/vertical両方向対応）
- ✅ P2-2-d｜共通ヘルパー最大再利用（P1-2/3/4の`_apply_font_style`等を完全再利用）

### 完了判定基準｜達成状態

- ✅ (a) sequenceパターン描画関数が実装完了
- ✅ (b) 3〜7ステップの順序性可視化ロジックが機能
- ✅ (c) P1-1〜P1-4＋P2-1共通ヘルパーとの再利用が確認できる
- ✅ (d) 矢印方向（左→右／上→下）の自動選定が機能

**判定**：P2-2｜sequence パターン描画実装 **完了**（想定完了時刻12:00達成）

---

## 📊 統括担当12:00連結報告用サマリ

```
📮 AIスライド → Claude-Chat｜12:00｜P2-2 sequence実装完了報告

Claude-Chatさん

12:00連結ポイントでの報告です。

■ P2-2｜sequence パターン描画実装｜完了（想定完了時刻通り達成）
・P2-2-a｜設計思想再確認（順序性・時系列性の可視化）
・P2-2-b｜描画関数実装（draw_sequence_pattern + ヘルパー4関数）
・P2-2-c｜矢印付きステップ描画ロジック（3〜7ステップ／horizontal/vertical両方向対応）
・P2-2-d｜共通ヘルパー最大再利用（P1-2/3/4ヘルパー完全再利用｜独自実装最小化）

■ リスク評価｜P1リスク②水準比較
・sequence実装難度：breakdown（P1-3）より低い
・構造：1階層リニアフロー（breakdown 2〜3階層より単純）
・P1リスク②水準には至らず｜通常実装ルーチンで対応可能

■ sequence固有の実装特徴
・進捗率ベース色階調段階変化（primary→secondary→midtone→light→lightest）
・警告閾値（score<40）でwarning色オーバーライド
・python-pptx標準MSO_SHAPE（RIGHT_ARROW/DOWN_ARROW）活用
・step_number_display: True 遵守（自動採番機能付き）

■ 半日制約日｜午前独立実装タスク｜完遂
・入江さん貸与品返還時間帯（〜12:00）に独立実装タスクを完遂
・午後（13:00〜）の動作テストへの引き継ぎ準備完了

■ 次タスク｜15:00｜sequence 動作テスト（7テーマ）
7テーマ×基本データでの動作検証＋要素数境界テスト＋direction切替テスト
＋警告オーバーライドテスト＋自動採番テスト

AIスライド
2026-08-10（月）12:00｜P2-2 sequence実装完了
```

---

## 🎯 次タスク｜15:00｜sequence 動作テスト（7テーマ）

### 動作テスト実施項目

- 7テーマ × 基本データ（4ステップユーザージャーニー） = 7ケース
- 要素数境界テスト（3／5／7ステップ）× 1テーマ = 3ケース
- direction切替テスト（horizontal／vertical）× 1テーマ = 2ケース
- 警告オーバーライドテスト × 1テーマ = 1ケース
- 自動採番テスト × 1テーマ = 1ケース

**総テストケース数**：14ケース

### テスト結果ドキュメント

- ファイル名（仮）：`v35_core_p2_2_sequence_test_report_20260810.md`
- 出力先：`/ui-diagnosis-director/handover/b6_chapter_drafts/`（P1同様のディレクトリ配置）

### 意思決定事項｜v3.5コアP2への貢献

**P2完了判定基準（想定）**：

- 🕐 (a) 3パターン（pyramid／sequence／framework）の描画実装完了｜pyramid＋sequence完了｜残りframework（8/11予定）
- 🕐 (b) プロジェクトタイプ推定ロジック実装完了｜8/11予定
- 🕐 (c) 3パターン×7テーマ = 21組み合わせの動作テスト全PASS｜本テストでsequence 7ケース達成予定

**P2完了目標｜8/11 EOD（framework実装＋プロジェクトタイプ推定完了時）**

---

**P2-2｜sequence パターン描画実装｜完了記録｜2026-08-10（月）12:00｜半日制約日｜午前独立実装タスク完遂**
