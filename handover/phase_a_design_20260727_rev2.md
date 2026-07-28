# Phase A 設計ドラフト rev2 — UI診断ディレクター機能強化

**作成日**：2026-07-27（月）初版 → 2026-07-28（火）rev2
**作成者**：AIスライド
**位置づけ**：Claude-Chat先行調査（`CLAUDE_CHAT_REPLY_ADDENDUM_20260727.md` ＋ `builder_ v16_sample.py`）の反映版
**進行方式**：[A] 独断ドラフト → 入江さんレビュー → 必要時Claude-Chat相談

---

## 📝 rev2 改訂サマリ

初版（2026-07-27作成）に対する改訂点：

| 項目 | 初版 | rev2 |
|---|---|---|
| **A-1 辞書構造** | 7色×3階調＝21色（暫定転記） | **7テーマ×8色階調＝56色＋共通閾値色**（Claude-Chat先行調査で公式JSON確定） |
| **A-1 増加行数** | 約30行 | **約80行** |
| **A-1 ライセンス表記** | CC BY 4.0想定 | **PDL1.0（公共データ利用規約 第1.0版）** |
| **A-1 応用パス** | 「業種別カラー選択」の余地 | **10項目スコア5段階への色分け応用パス**を追記 |
| **A-2 / A-3** | 初版から変更なし | 変更なし |

---

## 🎯 Phase Aのスコープ（変更なし）

Phase Aは以下の3タスクで構成する。**GPTs Instructions lite版（7,986字・余裕14字）には手を入れず**、builder側（Python）に組み込む方針。

| タスク | 内容 | 期待効果 |
|---|---|---|
| **A-1** | デジタル庁7テーマ×8色階調パレットのbuilder取り込み | 診断結果に「デジタル庁ガイドブック準拠」の権威付け＋10項目スコア色分けへの応用パス |
| **A-2** | パーツ図鑑12種図解パターン辞書化 | v3.5(a)バナー掲載媒体診断の視覚化基盤／C-2改善提案の説明力向上 |
| **A-3** | web-director.skillの位置づけ整理 | UI診断ディレクターの10項目評価軸の背骨として、Web案件全体フローの中での位置づけを明確化 |

Phase Aは**v3.5コア（バナー掲載媒体診断）実装前の"下ごしらえ"**として位置づける。

---

## A-1：デジタル庁7テーマ×8色階調パレットのbuilder取り込み（rev2）

### 現状（builder v16 L49-L68）— 既存ブランドカラー

```python
NAVY        = RGBColor(0x1C, 0x36, 0x6C)   # メインナビー
RED         = RGBColor(0xD0, 0x02, 0x1B)   # アクセント赤
LIGHT_GRAY  = RGBColor(0xF4, 0xF5, 0xF8)   # カード背景
BORDER_GRAY = RGBColor(0xD0, 0xD4, 0xDC)   # 罫線
TEXT        = RGBColor(0x40, 0x40, 0x40)   # 本文
SUB_TEXT    = RGBColor(0x60, 0x60, 0x60)   # 注釈
NAVY_LIGHT  = RGBColor(0x9D, 0xB0, 0xD6)
NAVY_E6     = RGBColor(0xE6, 0xEA, 0xF3)
GOLD        = RGBColor(0xFF, 0xD5, 0x4F)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_BORDER = RGBColor(0xCC, 0xCC, 0xCC)
PAGE_NUM    = RGBColor(0x26, 0x26, 0x26)
STRIPE      = RGBColor(0xF8, 0xF9, 0xFB)
ORANGE      = RGBColor(0xF9, 0x73, 0x16)
ORANGE_LIGHT = RGBColor(0xFE, 0xF1, 0xE3)
RED_LIGHT   = RGBColor(0xFD, 0xEC, 0xEC)
PRIO_RED    = RGBColor(0xD0, 0x02, 0x1B)
PRIO_ORANGE = RGBColor(0xE8, 0x8B, 0x1F)
PRIO_GRAY   = RGBColor(0x88, 0x88, 0x88)
```

**運用実態**：NAVY / RED / ORANGE の3色をブランドカラーとして全PPTXで統一使用。既存アイキャッチ・note投稿等で公開済み。**これらは維持する**。

### 改修案（rev2）— Claude-Chat先行調査反映

Claude-Chatが公式JSON（デジタル庁 GitHub `powerbi-templates/powerbi-theme-json/`）から確認した正確な値を取り込む。**7テーマ×各8色階調セット＋共通閾値色**の構造。

```python
# ============================================================
# デジタル庁 公式カラーパレット（Phase A-1）
# 出典：デジタル庁 ダッシュボードデザインの実践ガイドブックとデザインテンプレート
# https://www.digital.go.jp/resources/dashboard-guidebook
# ライセンス：PDL1.0（公共データ利用規約 第1.0版）
# GitHub：https://github.com/digital-go-jp/policy-dashboard-assets
# 取得日：2026-07-27（Claude-Chat先行調査）
# ============================================================
DIGITAL_AGENCY_PALETTE = {
    "SolidGray": {
        "primary":  "#4D4D4D",
        "secondary":"#767676",
        "midtone":  "#999999",
        "light":    "#CCCCCC",
        "lightest": "#F2F2F2",
        "accent":   "#3460FB",  # 強調用（青）
        "warning":  "#FE3939",  # 警告用（赤）
        "bg":       "#F8F8FB",
    },
    "Blue": {
        "primary":  "#0017C1",
        "secondary":"#3460FB",
        "midtone":  "#7096F8",
        "light":    "#C5D7FB",
        "lightest": "#E8F1FE",
        "accent":   "#FE3939",
        "warning":  "#FFBBBB",
        "bg":       "#F8F8FB",
    },
    "LightBlue": {
        "primary":  "#0055AD",
        "secondary":"#008BF2",
        "midtone":  "#57B8FF",
        "light":    "#C0E4FF",
        "lightest": "#F0F9FF",
        "accent":   "#FE3939",
        "warning":  "#FFBBBB",
        "bg":       "#F8F8FB",
    },
    "Green": {
        "primary":  "#115A36",
        "secondary":"#259D63",
        "midtone":  "#51B883",
        "light":    "#9BD4B5",
        "lightest": "#E6F5EC",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Cyan": {
        "primary":  "#006F83",
        "secondary":"#00A3BF",
        "midtone":  "#2BC8E4",
        "light":    "#99F2FF",
        "lightest": "#E9F7F9",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Red": {
        "primary":  "#CE0000",
        "secondary":"#FE3939",
        "midtone":  "#FF7171",
        "light":    "#FFBBBB",
        "lightest": "#FDEEEE",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
    "Orange": {
        "primary":  "#AC3E00",
        "secondary":"#FB5B01",
        "midtone":  "#FF8D44",
        "light":    "#FFC199",
        "lightest": "#FFEEE2",
        "accent":   "#666666",
        "warning":  "#CCCCCC",
        "bg":       "#F8F8FB",
    },
}

# 全テーマ共通の閾値色（good/bad判定用）
DIGITAL_AGENCY_THRESHOLD = {
    "center":  "#E6E6E6",  # 中央値・中立表示
    # maximum / minimum は各テーマの4番目色に準ずる（テーマ依存）
}
```

### rev2で追加：10項目スコア色分けへの応用パス

Claude-Chatの提案を踏まえ、**10項目5段階評価と色階層の対応マップ**を将来設計として明記する（Phase B以降で実装）：

| スコア | 判定 | 使用色（案） |
|---|---|---|
| 5点 | 優秀 | `Green.primary` #115A36 |
| 4点 | 良好 | `Green.midtone` #51B883 |
| 3点 | 標準 | `SolidGray.primary` #4D4D4D or `Blue.midtone` #7096F8 |
| 2点 | 課題 | `Orange.midtone` #FF8D44 |
| 1点 | 要改善 | `Red.primary` #CE0000 |

この応用パスの実装は **Phase B以降** とする。理由：
- v3.5コアはバナー掲載媒体診断の実装がメインスコープ
- スコア色分けを組み込む場合、Visual Board（C-3）のスコア図表描画関数 `add_visual_board()` 内の色指定ロジック改修が必要（約50〜80行）
- Phase Aでは辞書追加のみで**既存描画結果に一切影響しない**ことを優先

### 適用範囲

**v3.5時点では"内部保持のみ"とし、利用ロジックは実装しない**。理由：

- 現行のNAVY/RED/ORANGEブランドカラーは既にβ募集アイキャッチ・note投稿等で公開済み。色を変えるとブランド一貫性が崩れる
- パレット選択機能を実装するには**GPTs Instructions側でヒアリング項目追加が必要**（業種選択→パレット自動選択）。Instructions lite版は7,986字で余裕14字しかないため実装不可
- v3.5→v3.6でInstructions PLAIN版の再構成タイミングで、パレット選択UIを追加検討

**Phase Aでは以下のみ実施：**
- `DIGITAL_AGENCY_PALETTE` 辞書定義追加（約70行）
- `DIGITAL_AGENCY_THRESHOLD` 辞書定義追加（約5行）
- 冒頭に出典コメント追加（約8行）
- **既存のC-1〜C-3描画ロジックには一切触れない**（後方互換完全維持）

### 増加行数見積もり（rev2更新）

- 追加：約80行（辞書 + 出典コメント）
- 変更：0行
- 削除：0行
- **リスク**：低（既存ロジック非改修のため）

---

## A-2：パーツ図鑑12種図解パターン辞書化（初版から変更なし）

### 対象パターン（うちた氏保管庫より）

パーツ図鑑・図解集で確認できた12種の図解パターン。**うちた氏より商用利用許諾済み**（2026-07-26 スクショで確認・`UCHITA_LICENSE_RECORD_20260726.md`で保全予定）。

| # | パターン名 | 用途 | UI診断ディレクターでの活用シーン |
|---|---|---|---|
| 1 | 分類（Category） | 要素を並列カテゴリで整理 | C-2改善提案5件を「デザイン系／情報設計系／CV系」等に分類表示 |
| 2 | ピラミッド（Pyramid） | 階層・優先順位を上下で表現 | C-1優先課題3件を高→低で階層可視化 |
| 3 | 比較（Comparison） | 2〜3要素の対比 | 現状 vs 改善後のBefore/After（既存C-3で使用中） |
| 4 | 順序（Sequence） | ステップ・時系列を左→右で表現 | ユーザー行動フロー（既存C-3で使用中） |
| 5 | 循環（Cycle） | 反復プロセスを円環で表現 | 改善→測定→再改善のPDCAサイクル説明 |
| 6 | 絞り込み（Funnel） | 上から下へ絞り込むファネル型 | サイト流入→CV率の各段階可視化 |
| 7 | 時間軸（Timeline） | 期間別のマイルストーン | 改善実施スケジュール提案 |
| 8 | 分解（Breakdown） | 全体を構成要素に分解 | 総合スコア50点の10項目内訳可視化 |
| 9 | 対比（Contrast） | 対照的な2要素の並列強調 | 良UX例 vs 悪UX例 |
| 10 | 統合（Integration） | 複数要素の統合結果 | 診断結果→統合改善方針 |
| 11 | フレームワーク（Framework） | 既知のマトリクス型（4象限等） | 影響度×実装コスト マトリクス |
| 12 | ネットワーク（Network） | ノード間の関係性 | ページ間リンク構造・ユーザー導線 |

### 改修案

builder側に**図解パターン辞書**を定義し、C-2改善提案の描画時に「このパターンで表現すると効果的」というメタ情報を持たせる。ただしv3.5では**辞書定義のみ・描画ロジック実装なし**とする（Phase Bで実装）。

```python
# ============================================================
# パーツ図鑑・図解集 由来 図解パターン辞書（Phase A-2）
# 出典：うちた氏「パーツ図鑑_120種」「図解集_50種」「テンプレ大全_100枚」
# 商用利用許諾：2026-07-26 note返信にて確認済み
# ライセンス記録：UCHITA_LICENSE_RECORD_20260726.md 参照
# ============================================================
DIAGRAM_PATTERNS = {
    'category':     {'ja': '分類',           'use': '要素を並列カテゴリで整理',      'shape': 'grid'},
    'pyramid':      {'ja': 'ピラミッド',     'use': '階層・優先順位を上下で表現',    'shape': 'triangle'},
    'comparison':   {'ja': '比較',           'use': '2〜3要素の対比',                'shape': 'side_by_side'},
    'sequence':     {'ja': '順序',           'use': 'ステップ・時系列を左→右',       'shape': 'arrow_chain'},
    'cycle':        {'ja': '循環',           'use': '反復プロセスを円環で表現',      'shape': 'circle_arrow'},
    'funnel':       {'ja': '絞り込み',       'use': '上から下へ絞り込むファネル型',  'shape': 'trapezoid'},
    'timeline':     {'ja': '時間軸',         'use': '期間別のマイルストーン',        'shape': 'horizontal_bar'},
    'breakdown':    {'ja': '分解',           'use': '全体を構成要素に分解',          'shape': 'tree'},
    'contrast':     {'ja': '対比',           'use': '対照的な2要素の並列強調',       'shape': 'split_screen'},
    'integration':  {'ja': '統合',           'use': '複数要素の統合結果',            'shape': 'merge'},
    'framework':    {'ja': 'フレームワーク', 'use': '4象限マトリクス等',             'shape': 'quadrant'},
    'network':      {'ja': 'ネットワーク',   'use': 'ノード間の関係性',              'shape': 'node_edge'},
}

# 診断結果→推奨図解パターンのマッピング
DIAGNOSIS_TO_PATTERN = {
    'proposal_categorization': 'category',
    'priority_ranking':        'pyramid',
    'before_after':            'comparison',
    'user_flow':               'sequence',
    'improvement_cycle':       'cycle',
    'conversion_funnel':       'funnel',
    'schedule':                'timeline',
    'score_breakdown':         'breakdown',
    'ux_contrast':             'contrast',
    'impact_cost_matrix':      'framework',
    'site_structure':          'network',
}
```

### 適用範囲

**v3.5時点では"辞書のみ・描画未実装"**。理由：
- 12種の描画ロジックを全部実装すると1,500〜2,000行の追加が必要
- v3.5コア（バナー掲載媒体診断）実装と競合するリソース
- Phase Bで、βFB集約後に「実務層が実際に必要とするパターン」を絞り込んでから実装

### 増加行数見積もり

- 追加：約35行
- 変更：0行
- 削除：0行
- **リスク**：低

---

## A-3：web-director.skillの位置づけ整理（初版から変更なし）

### 現状把握

**web-director.skill**（入江さんご自身の知見の集約）は既にAnthropic Claude Skills形式でパッケージ化済み。「マーケティングオーケストレーター」プロジェクトで活用中。UI診断ディレクターとの併用は未実施。

### 10項目評価軸との対応マッピング

| UI診断ディレクター10項目 | web-director.skill 6大機能領域 | 対応度 |
|---|---|---|
| ①FV（ファーストビュー） | ④ワイヤーフレーム計画 | 🟢 高 |
| ②キャッチコピー | ⑥クライアント提案書（メッセージング原則） | 🟢 高 |
| ③CTA設計 | ⑤UI/UXガイダンス（インタラクションパターン） | 🟢 高 |
| ④信頼性 | ⑥クライアント提案書（社会的証明の原則） | 🟡 中 |
| ⑤フォーム | ⑤UI/UXガイダンス（フォーム設計原則） | 🟢 高 |
| ⑥レスポンシブ | ⑤UI/UXガイダンス（モバイルUX） | 🟢 高 |
| ⑦読みやすさ | ⑤UI/UXガイダンス（タイポグラフィ／情報階層） | 🟢 高 |
| ⑧情報設計 | ③ユーザーフロー設計＋④ワイヤーフレーム計画 | 🟢 高 |
| ⑨ブランド一貫性 | ⑤UI/UXガイダンス（一貫性原則） | 🟢 高 |
| ⑩表示速度 | ⑤UI/UXガイダンス（パフォーマンス最適化） | 🟡 中 |

**結論**：10項目のうち8項目が「高」対応。web-director.skillは**UI診断ディレクターの診断根拠の背骨として完全に成立する**。

### 組み込み方針（Phase Aでの決定）

**Phase Aでは"参照経路の設計＋コード内コメントによる位置づけ明記"を実施**。実際のGPTs Knowledge Base への登録はPhase Bで実施。

builderの冒頭に位置づけコメントを追加：

```python
# ============================================================
# UI診断ディレクターの10項目評価軸の背景知識（Phase A-3）
# 出典：入江さんご自身の知見の集約（web-director.skill）
# 併用プロジェクト：マーケティングオーケストレーター
# 対応マッピング：10項目中8項目が web-director.skill 6大機能領域と「高」対応
# 詳細：phase_a_design_20260727_rev2.md A-3セクション参照
# ============================================================
```

### 適用範囲

- Phase A：位置づけ整理と対応マッピング文書化＋コード内コメント（本ドキュメントに記載）
- Phase B：Knowledge Base用要約版の作成＋実装

**Phase Aでのコード改修**：**約8行のコメント追加のみ**。

---

## 📊 Phase A 全体まとめ（rev2）

| 項目 | A-1 | A-2 | A-3 |
|---|---|---|---|
| **コード追加行数** | 約80行 | 約35行 | 約8行 |
| **既存ロジック変更** | なし | なし | なし |
| **後方互換性** | 完全維持 | 完全維持 | 完全維持 |
| **リスクレベル** | 低 | 低 | 極低 |
| **実装工数見積** | 45分 | 30分 | 15分 |
| **v3.5への影響** | 内部保持のみ・利用ロジック未実装 | 内部保持のみ・描画ロジック未実装 | 位置づけコメントのみ |

**Phase A完了後のbuilder状態**：
- 既存builder v16と機能的に**完全同一**（描画結果は1ピクセルも変わらない）
- 内部データとして**デジタル庁7テーマ×8色階調＋閾値色＋図解パターン12種**を保持
- Phase B以降での本格実装の**基盤整備完了**

---

## 🚦 Phase A→Phase B→v3.5コア の依存関係（変更なし）

```
Phase A（本設計） ──── 内部辞書定義・位置づけ整理
   │
   ├─ A-1（7テーマ×8色階調＋閾値色辞書）
   ├─ A-2（12種図解パターン辞書）
   └─ A-3（web-director.skill位置づけコメント）
        ↓
Phase B ──── Knowledge Base実装・要約版作成
   │
   ├─ B-1（web-director-condensed.md作成）
   ├─ B-2（digital-agency-guidebook-summary.md作成）
   ├─ B-3（diagram-patterns-catalog.md作成）
   ├─ B-4（ppt-visual-catalog スキル新規作成）
   └─ B-5【新規】10項目スコア色分けの描画実装
        ↓
v3.5コア ──── バナー掲載媒体診断の本実装
```

---

## 📅 実装スケジュール提案（rev2更新）

| 日程 | 内容 | 主担当 |
|---|---|---|
| 7/27（月） | Phase A設計ドラフト初版作成 | AIスライド |
| 7/27（月）夜 | Claude-Chat先行調査受領（`CLAUDE_CHAT_REPLY_ADDENDUM_20260727.md` ＋ `builder_ v16_sample.py`） | Claude-Chat |
| **7/28（火）本日** | **Phase A設計ドラフトrev2改訂＋Phase A実装（builder v16.5作成）** | **AIスライド** |
| 7/29（水） | 自己レビュー・入江さんへの提出・必要ならClaude-Chat相談 | AIスライド |
| 7/30（木） | 予備日（改修が入る場合の対応日） | AIスライド |
| 7/31（木）21:00頃 | X軸締切告知投稿 | 入江さん |
| 8/1〜8/5 | Phase B設計 | AIスライド |
| 8/6〜8/15 | v3.5コア実装 | AIスライド＋Claude-Chat |
| 8月中旬 | βFB集約→3者レビュー | 3者 |
| 8月下旬 | v3.5統合→builder v17リリース | 3者 |
| 9月上旬 | β二次募集＋Brain販売開始 | 入江さん |

---

## ❓ 入江さん・Claude-Chatへの確認事項

### 入江さんへの確認事項（rev2で更新）
1. **A-1「7テーマ×8色階調＋閾値色を辞書として内部保持のみ」の判断**でよいか？
2. **A-1「10項目スコア色分けはPhase B以降」の判断**でよいか？
3. **A-2「12種図解パターン辞書のみで描画実装なし」の判断**でよいか？
4. **A-3「web-director.skillはコード内コメントのみで位置づけ明記」の判断**でよいか？

### Claude-Chatへの相談事項
初版で挙げていた2件の相談事項（デジタル庁公式JSON取得・うちた氏出典表記の実装場所）は、**Claude-Chat先行調査で既に回答済み**。追加相談事項は現時点なし。

---

## 📌 参考資料

- Hub `/PROJECT_STATE.md`（正典・7/26最終更新）
- Hub `/ui-diagnosis-director/_reference/pptx_template/`（うちた氏保管庫9ファイル）
- Hub `/ui-diagnosis-director/_reference/design_system/【デジタル庁】ダッシュボードデザインの実戦ガイドブック.pdf`
- Hub `/ui-diagnosis-director/_reference/skills/web-director.skill`
- Hub `/ui-diagnosis-director/handover/CLAUDE_CHAT_REPLY_ADDENDUM_20260727.md`（Claude-Chat先行調査）
- Hub `/ui-diagnosis-director/handover/builder_ v16_sample.py`（Claude-Chat提供実装用サンプルコード）
- デジタル庁公式：https://www.digital.go.jp/resources/dashboard-guidebook
- デジタル庁GitHub：https://github.com/digital-go-jp/policy-dashboard-assets

---

**Phase A設計ドラフト rev2 以上**
