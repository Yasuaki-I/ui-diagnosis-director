# v3.5コアP3-5｜network パターン｜設計ドキュメント（設計先行版｜09:00〜10:00枠）

- 作成日：2026-08-15（土）
- 作成者：AIスライド（実装領域）
- 位置づけ：**v3.5コア最後の未実装パターン｜12パターン中最高難度｜設計先行（描画関数の骨格まで）**
- 統括承認：8/12議題4｜段階的実装（3→5→7ノード）｜categoryフォールバック必須｜エッジ交差回避はv3.5範囲外
- **⚠️ (b)エスカレーション条件｜最警戒レベル**
- 想定完了時刻：**10:00（設計先行）｜実装本体は見積り後に配置確定**

---

## 📋 実装着手前｜Phase A A-2原本整合確認｜完了

### 原本定義（`phase_a_design_20260727_rev2.md` A-2節｜逐語）

```python
DIAGRAM_PATTERNS = {
    'network': {'ja': 'ネットワーク', 'use': 'ノード間の関係性', 'shape': 'node_edge'},
}

DIAGNOSIS_TO_PATTERN = {
    'site_structure': 'network',
}
```

### 実装レベル拡張定義（本実装での確定値）

```python
DIAGRAM_PATTERNS_EXTENDED['network'] = {
    # 原本定義（3プロパティ｜改変禁止）
    'ja': 'ネットワーク',
    'use': 'ノード間の関係性',
    'shape': 'node_edge',
    # 実装レベル拡張定義（統括承認済[A]方針）
    'min_elements': 3,
    'max_elements': 7,
    'requires_axes': False,
    'direction': 'hierarchical_top_down',   # 階層型（上→下）
    'color_gradation': 'depth_hierarchical', # 階層深度別（pyramid流用）
}
```

### 原本との整合確認｜3観点

| 観点 | 確認内容 | 判定 |
|-----|--------|:---:|
| `shape`整合 | node_edge＝ノード（矩形）＋エッジ（コネクタ線）の二重要素を両方描画 | ✅ |
| `use`整合 | 「ノード**間**の関係性」＝**エッジ描画が本質的必須要件**｜ノードのみの描画は原本違反 | ✅ |
| 診断カテゴリ1対1対応 | `site_structure` → `network` | ✅ |

### **エスカレーション条件(c)判定：「非該当」**

---

## 🎯 P3-5 network｜設計思想

### 診断カテゴリ対応

- **診断カテゴリ**：`site_structure`
- **選定論理**：サイト構造はノード間関係であり、`network`のnode_edgeで表現
- **B-6 3.3節位置づけ**：ノード＋エッジによる構造可視化｜**12パターン中最複雑**
- **想定活用シーン**：サイト構造診断／ページ階層可視化／リンク関係性

### 視覚構造の骨格（7ノードの例）

```
              [1｜TOP]              ← depth 0（primary）
             ╱        ╲
    [2｜サービス]  [3｜会社情報]      ← depth 1（secondary）
        ╱    ╲          │
  [4｜料金] [5｜事例]  [6｜採用]      ← depth 2（midtone）
        ╲    ╱
      [7｜お問い合わせ]              ← depth 3（light）※共通合流ノード
```

### 設計判断｜7件

**判断①｜階層型レイアウト（hierarchical_top_down）を採用｜力学的配置は採らない**
- ノード配置は「depth（階層深度）× 同depth内の順序」で決定論的に算出
- **却下した代替案**：力学モデル（force-directed）配置｜反復計算が必要で、同一データで毎回同じ図が出ない
- **採用根拠**：B-6判断1「選定の透明性」＝同一データから同一の図が再現されること。診断書は再現性が必須のため、決定論的レイアウト以外は選択肢にならない

**判断②｜depth はデータ側で明示｜自動推定はしない**
- 各ノードは`depth`（0〜3）を明示的に持つ
- **却下した代替案**：エッジ関係からdepthを自動推定（BFS探索）｜循環参照時に無限ループのリスク
- **採用根拠**：リスク②水準の回避｜自動推定は「実装難度が構造的に上がる」典型例

**判断③｜色階調は pyramid の階層深度別ロジックを流用**
- depth 0=primary／1=secondary／2=midtone／3=light
- P2-1 pyramid の`_apply_pyramid_level_color`と**同一思想**（素案の再利用評価「○｜階層深度別」通り）

**判断④｜エッジは MSO_CONNECTOR.STRAIGHT で実装**
- `slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, begin_x, begin_y, end_x, end_y)`
- **却下した代替案**：`MSO_CONNECTOR.ELBOW`（直角コネクタ）｜サイト構造図として一般的だが、始点/終点の自動ルーティングがpython-pptxでは制御困難
- **cycle実装の教訓を適用**：BLOCK_ARCで環境依存リスクを踏んだため、networkでは**最も基本的な図形（直線コネクタ）を第一選択**とする

**判断⑤｜エッジ交差回避ロジックは実装しない（統括承認済｜v3.5コア範囲外）**
- 8/12議題4で保留判断が承認済
- 7ノード上限（Miller's Law）の範囲では、階層レイアウトなら交差は稀
- v3.6以降のバックログへ記録

**判断⑥｜categoryフォールバック｜3起動条件（統括指示）**
1. **ノード数超過**：`n < 3` または `n > 7`
2. **エッジ関係不整合**：エッジの`from`/`to`が存在しないノードIDを参照
3. **階層構造データ不正**：`depth`が欠損／負値／depth 0（ルート）が存在しない

**判断⑦｜ノードラベルは短縮表示｜スコアは省略可**
- ノードボックスは他パターンより小さい（7ノード時 1.8×0.75 inch）ため、ラベルは12ptで最大2行
- `score`は省略可（サイト構造図では全ページにスコアが付かないケースが多い）｜省略時は警告判定をスキップ

---

## 🧩 描画関数実装｜骨格（設計先行版）

### 独自実装関数｜6件（素案想定5〜6件｜上限内）

**関数1｜`_compute_network_positions(nodes: list) -> dict`**

```python
def _compute_network_positions(nodes: list) -> dict:
    """
    ノードの階層型レイアウト座標を計算する（hierarchical_top_down）。

    Args:
        nodes: [{"id": str, "label": str, "depth": int, "score": int}, ...]

    Returns:
        {
            "nodes": {node_id: (x, y, w, h), ...},   # Inches
            "centers": {node_id: (cx, cy), ...},      # エッジ接続用の中心座標
        }

    設計思想（判断①②）：
    - depth ごとにノードをグルーピングし、各 depth を1行として上から配置
    - 同一 depth 内は左右等間隔（中央揃え）
    - 力学的配置は採らない｜同一データ→同一図の再現性を担保（B-6判断1）
    """
    slide_w, slide_h = 13.33, 7.5
    header_area, margin = 1.0, 0.5

    # depth ごとにグルーピング
    depth_groups = {}
    for node in nodes:
        depth_groups.setdefault(node["depth"], []).append(node)
    max_depth = max(depth_groups.keys())

    # ノードサイズ（総数により可変｜判断⑦）
    n = len(nodes)
    box_w = 2.4 if n <= 5 else 1.8
    box_h = 0.9 if n <= 5 else 0.75

    field_h = slide_h - header_area - margin
    row_h = field_h / (max_depth + 1)

    positions, centers = {}, {}
    for depth, group in depth_groups.items():
        cnt = len(group)
        total_w = cnt * box_w + (cnt - 1) * 0.5      # ノード間隔 0.5inch
        start_x = (slide_w - total_w) / 2             # 中央揃え
        y = header_area + row_h * depth + (row_h - box_h) / 2
        for i, node in enumerate(group):
            x = start_x + i * (box_w + 0.5)
            positions[node["id"]] = (x, y, box_w, box_h)
            centers[node["id"]] = (x + box_w / 2, y + box_h / 2)

    return {"nodes": positions, "centers": centers}
```

**関数2｜`_draw_network_node(slide, palette, node, position)`**

```python
def _draw_network_node(slide, palette, node, position):
    """
    ノード1つを描画する。

    設計思想：
    - P2-3 _draw_framework_cell のセル描画ロジックを流用（素案評価「○」通り）
    - 形状は MSO_SHAPE.ROUNDED_RECTANGLE
    - ラベルは12pt最大2行（判断⑦）｜score は省略可
    """
    x, y, w, h = position
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    _apply_network_node_color(box, palette, node["depth"], node.get("score"))

    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = node["label"]
    _apply_font_style(p, size=12, bold=True, color=palette["bg"])
    p.alignment = PP_ALIGN.CENTER

    # score は省略可（判断⑦）
    if node.get("score") is not None:
        p_score = tf.add_paragraph()
        p_score.text = f"{node['score']}"
        _apply_font_style(p_score, size=10, color=palette["lightest"])
        p_score.alignment = PP_ALIGN.CENTER

    return box
```

**関数3｜`_draw_network_edge(slide, palette, edge, centers)`**

```python
def _draw_network_edge(slide, palette, edge, centers):
    """
    エッジ1本を描画する（原本 use「ノード間の関係性」の本質要件）。

    Args:
        edge: {"from": str, "to": str, "label": str}
        centers: {node_id: (cx, cy)}

    設計思想（判断④）：
    - MSO_CONNECTOR.STRAIGHT を第一選択（cycle の BLOCK_ARC 環境依存の教訓）
    - ELBOW（直角コネクタ）は始点/終点の自動ルーティングが制御困難なため却下
    - 線色は light｜ノードより後退させ、ノードの可読性を優先
    """
    bx, by = centers[edge["from"]]
    ex, ey = centers[edge["to"]]
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(bx), Inches(by), Inches(ex), Inches(ey)
    )
    conn.line.color.rgb = hex_to_rgb(palette["light"])
    conn.line.width = Pt(1.5)
    return conn
```

**関数4｜`_apply_network_node_color(box, palette, depth, score)`**

```python
def _apply_network_node_color(box, palette, depth, score):
    """
    ノードの色階調適用（階層深度別｜pyramid流用｜判断③）。

    設計思想：
    - depth 0=primary／1=secondary／2=midtone／3=light（以降は light 固定）
    - score が None の場合は警告判定をスキップ（判断⑦）
    - score < 40 は warning オーバーライド（全パターン共通｜B-6 4.4節）
    """
    box.fill.solid()

    if score is not None and score < 40:
        box.fill.fore_color.rgb = hex_to_rgb(palette["warning"])
    else:
        depth_keys = ["primary", "secondary", "midtone", "light"]
        key = depth_keys[min(depth, 3)]
        box.fill.fore_color.rgb = hex_to_rgb(palette[key])

    box.line.color.rgb = hex_to_rgb(palette["midtone"])
    box.line.width = Pt(1.5)
```

**関数5｜`_validate_network_data(data) -> tuple`**

```python
def _validate_network_data(data) -> tuple:
    """
    network データを検証する（categoryフォールバック3起動条件｜判断⑥｜統括指示）。

    Returns:
        (is_valid: bool, reason: str)
    """
    nodes = data.get("nodes", [])
    edges = data.get("edges", [])
    n = len(nodes)

    # 条件1｜ノード数超過
    if not (3 <= n <= 7):
        return False, f"node_count_out_of_range({n})"

    node_ids = {nd.get("id") for nd in nodes}

    # 条件2｜エッジ関係不整合
    for e in edges:
        if e.get("from") not in node_ids or e.get("to") not in node_ids:
            return False, f"edge_reference_invalid({e.get('from')}->{e.get('to')})"

    # 条件3｜階層構造データ不正
    depths = [nd.get("depth") for nd in nodes]
    if any(d is None or not isinstance(d, int) or d < 0 for d in depths):
        return False, "depth_invalid"
    if 0 not in depths:
        return False, "root_node_missing"

    return True, "valid"
```

**関数6｜`_to_category_data_from_nodes(network_data) -> dict`**

```python
def _to_category_data_from_nodes(network_data: dict) -> dict:
    """networkデータをcategoryパターンデータに変換（フォールバック用）。"""
    return {
        "title": network_data.get("title", ""),
        "categories": [
            {"label": nd.get("label", ""), "score": nd.get("score") or 0,
             "description": f"depth {nd.get('depth', '-')}"}
            for nd in network_data.get("nodes", [])
        ],
    }
```

### メイン描画関数｜`draw_network(slide, palette, data)`

```python
def draw_network(slide, palette, data):
    """
    network パターン描画のメイン関数（node_edge｜サイト構造の関係性可視化）。

    Args:
        data: {
            "title": str,
            "nodes": [{"id": str, "label": str, "depth": int, "score": int|None}, ...],
            "edges": [{"from": str, "to": str, "label": str}, ...],
        }
    """
    # データ検証（3起動条件｜統括指示）
    is_valid, reason = _validate_network_data(data)
    if not is_valid:
        # categoryフォールバック（例外は投げない｜P1 breakdown以来の一貫方針）
        return draw_category(slide, palette, _to_category_data_from_nodes(data))

    # タイトル描画（P1-2 共通ヘルパー完全再利用）
    _draw_title(slide, palette, data.get("title", ""))

    # 階層レイアウト座標算出
    pos = _compute_network_positions(data["nodes"])

    # エッジを先に描画（ノードの背面に配置）
    for edge in data.get("edges", []):
        _draw_network_edge(slide, palette, edge, pos["centers"])

    # ノードを描画
    for node in data["nodes"]:
        _draw_network_node(slide, palette, node, pos["nodes"][node["id"]])

    return slide
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
| 7 | `_apply_pyramid_level_color`（思想流用） | P2-1 | ○ | 階層深度別色階調 |
| 8 | `_draw_framework_cell`（セル描画ロジック） | P2-3 | ○ | ノード描画 |

**新規実装｜6関数**（素案想定5〜6件の上限内）

---

## ⏱️ 実装本体｜想定所要時間｜暫定見積り（厳守①｜10:00連結①発信用）

### 段階的実装｜3段階（8/12統括承認済）

| 段階 | 内容 | 統括想定 | **AIスライド見積り** | 差分 | 根拠 |
|-----|-----|:------:|:--------------:|:---:|-----|
| 段階1｜3ノード | `_compute_network_positions`＋`_draw_network_node`＋`_apply_network_node_color`＋`_draw_network_edge` の基本動作確立 | 60分 | **50分** | **-10分** | 座標計算がdepth走査のみ（cycleの三角関数より単純）｜pyramid色階調を思想流用 |
| 段階2｜5ノード | 同一depth内の複数ノード配置＋エッジ分岐（1親→2子） | 45分 | **35分** | **-10分** | 段階1の`depth_groups`ロジックがそのまま多列対応｜追加実装は中央揃え計算のみ |
| 段階3｜7ノード | max境界＋ノードサイズ可変（1.8×0.75）＋合流ノード（2親→1子）＋`_validate_network_data` | 75分 | **65分** | **-10分** | 検証関数（3条件）は独立実装可｜合流エッジは`_draw_network_edge`の再呼出のみ |
| – | 動作テスト（7テーマ×3ノード数＋境界値＋固有機能） | – | **120分** | – | contrast 29／cycle 38項目の実績ベース｜network想定40項目前後 |

### ⭐ **実装本体の暫定見積り｜合計 150分（2.5時間）＋動作テスト 120分**

**内訳**：段階1（50分）＋段階2（35分）＋段階3（65分）＝**150分**

### 見積り根拠｜cycleとの比較（本日の判断材料）

| 観点 | cycle（8/14実績） | **network（本日見積り）** | 評価 |
|-----|:-------------:|:-------------------:|-----|
| 座標計算の複雑度 | 三角関数（極座標変換） | depth走査＋中央揃え（四則演算のみ） | **networkの方が単純** |
| 描画図形の種類 | ROUNDED_RECTANGLE＋BLOCK_ARC（環境依存リスク） | ROUNDED_RECTANGLE＋STRAIGHT connector（基本図形） | **networkの方が低リスク** |
| 要素の種類 | 単一（段階ボックスのみ） | **二重（ノード＋エッジ）** | **networkの方が複雑** |
| データ検証の必要性 | 要素数チェックのみ | **3条件検証（ノード数／エッジ参照／階層構造）** | **networkの方が複雑** |
| 新規実装関数数 | 5件 | 6件 | ほぼ同等 |

**総合評価**：**「二重要素＋データ検証」で複雑度は上がるが、「座標計算＋図形選択」ではcycleより低リスク**。差し引きでcycle（実測4時間＝設計＋実装＋テスト）と**同水準〜やや上**と判断。

### ⚠️ 統括担当への判断依頼｜13:00〜14:00統合確認バッファの転用可否

**AIスライド側の推奨判断**：**転用は不要**

**根拠**：
- 実装本体150分＋テスト120分＝**計270分（4.5時間）**
- 本日の残余稼働枠：10:00〜21:00＝11時間｜うち推定ロジック実装3時間＋55シナリオ総合テスト（14:00以降）
- **network実装本体は14:00以降の枠に収まる見込み**であり、13:00〜14:00バッファは本来目的（推定ロジック↔network↔テーマ選定の接続確認）に充てる方が全体最適
- ただし、段階3（7ノード）で(b)徴候を検出した場合は**即座に転用要請を発報**する

---

## ✅ 描画ロジック検証｜フル版B 7項目｜自己検証（設計先行時点）

| # | 検証項目 | 判定 | 根拠 |
|---|---------|:----:|-----|
| 1 | Phase A A-2原本整合 | ✅ PASS | 原本3プロパティ改変ゼロ｜エッジ描画で`use`「ノード間の関係性」を担保 |
| 2 | 要素数上限遵守（Miller's Law 7±2） | ✅ PASS | min=3／max=7｜`_validate_network_data`条件1で実装 |
| 3 | 色階調自動選定（8色階調） | ✅ PASS | depth別4色＋warning／light／lightest／bg使用 |
| 4 | categoryフォールバック実装 | ✅ PASS | **3起動条件（統括指示）を`_validate_network_data`で実装** |
| 5 | 警告オーバーライド（score<40） | ✅ PASS | score=None時はスキップ（判断⑦） |
| 6 | 共通ヘルパー最大再利用 | ✅ PASS | 8件再利用｜新規6関数（素案上限内） |
| 7 | 診断カテゴリ1対1対応（site_structure） | ✅ PASS | DIAGNOSIS_TO_PATTERN準拠 |

**総合判定：7項目すべてPASS｜設計先行として十分な水準**

---

## 🚦 (b)エスカレーション条件｜設計先行時点の監視結果

| 徴候 | 検出状況 | 評価 |
|-----|:-----:|-----|
| 徴候1｜設計途中での方針転換 | **なし** | 判断①〜⑦は初期設計から変更なし |
| 徴候2｜新規実装関数が想定超過 | **なし** | 素案5〜6件に対し6件｜上限内 |
| 徴候3｜python-pptx標準機能で実装不能な要素 | **なし** | ROUNDED_RECTANGLE／STRAIGHT connector はいずれも基本図形｜cycleのBLOCK_ARCのような環境依存リスクなし |
| 徴候4｜**却下判断が3件発生**（力学配置／depth自動推定／ELBOWコネクタ） | **あり（正常）** | いずれも**リスク②回避のための意図的な却下**｜設計難度を下げる方向の判断であり、徴候ではなく成熟指標 |

### **(b)判定：「非該当」**｜設計先行時点

高難度パターンながら、**「複雑な代替案を意図的に却下して単純な決定論的実装に寄せる」判断を3件行ったことで、実装難度を設計段階で引き下げた**。cycleでBLOCK_ARC環境依存を踏んだ教訓が、networkの図形選択（基本図形優先）に直接反映されている。

---

**次アクション｜10:00 統括連結①｜network設計先行完了報告＋実装本体見積り（150分＋テスト120分）発信｜厳守①履行**
