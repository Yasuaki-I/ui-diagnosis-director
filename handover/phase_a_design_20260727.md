# Phase A 設計ドラフト — UI診断ディレクター機能強化

**作成日**：2026-07-27（月）
**作成者**：AIスライド（入江さん主担当日として）
**位置づけ**：7/26共有の3参考資料（PPTXテンプレート／デジタル庁ガイドブック／web-director.skill）のうち、**即着手可能な直接組み込み領域＝Phase A**の設計ドラフト
**進行方式**：[A] 独断ドラフト → 入江さんレビュー → 必要時Claude-Chat相談

---

## 🎯 Phase Aのスコープ

Phase Aは以下の3タスクで構成する。**GPTs Instructions lite版（7,986字・余裕14字）には手を入れず**、builder側（Python）に組み込む方針。

| タスク | 内容 | 期待効果 |
|---|---|---|
| **A-1** | デジタル庁7色パレットのbuilder取り込み | 診断結果に「デジタル庁ガイドブック準拠」の権威付け＋業種別カラー選択の余地 |
| **A-2** | パーツ図鑑12種図解パターン辞書化 | v3.5(a)バナー掲載媒体診断の視覚化基盤／C-2改善提案の説明力向上 |
| **A-3** | web-director.skillの位置づけ整理 | UI診断ディレクターの10項目評価軸の背骨として、Web案件全体フローの中での位置づけを明確化 |

Phase Aは**v3.5コア（バナー掲載媒体診断）実装前の"下ごしらえ"**として位置づける。バナー掲載媒体診断機能そのものの実装はv3.5コアで行う。

---

## A-1：デジタル庁7色パレットのbuilder取り込み

### 現状（builder v16 L49-L68）

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

**運用実態**：NAVY / RED / ORANGE の3色をブランドカラーとして全PPTXで統一使用（PROJECT_STATE Section 3参照）。

### 改修案

デジタル庁7色パレット（Solid Gray／Blue／Light Blue／Cyan／Green／Orange／Red）を**追加辞書**として定義し、**既存のNAVY/RED/ORANGEはデフォルト維持**する。

```python
# ─── デジタル庁ダッシュボードデザインガイドブック準拠 7色パレット ───
# 出典：デジタル庁「ダッシュボードデザインの実践ガイドブック」（2026年3月31日更新）
# https://www.digital.go.jp/resources/dashboard-guidebook

DIGITAL_AGENCY_PALETTES = {
    'solid_gray': {
        'primary':   RGBColor(0x40, 0x40, 0x40),
        'secondary': RGBColor(0x80, 0x80, 0x80),
        'accent':    RGBColor(0x60, 0x60, 0x60),
        'note':      'モノトーン・情報密度重視の実務資料向け',
    },
    'blue': {
        'primary':   RGBColor(0x00, 0x5B, 0xAC),
        'secondary': RGBColor(0x4D, 0x8F, 0xC9),
        'accent':    RGBColor(0x00, 0x3D, 0x7A),
        'note':      '信頼性・公共性重視（デフォルト推奨）',
    },
    'light_blue': {
        'primary':   RGBColor(0x00, 0x9E, 0xE5),
        'secondary': RGBColor(0x66, 0xC7, 0xEF),
        'accent':    RGBColor(0x00, 0x7A, 0xB3),
        'note':      'IT・SaaS系プロダクト向け',
    },
    'cyan': {
        'primary':   RGBColor(0x00, 0xA9, 0xA7),
        'secondary': RGBColor(0x4D, 0xC7, 0xC5),
        'accent':    RGBColor(0x00, 0x80, 0x7E),
        'note':      '医療・ヘルスケア系向け',
    },
    'green': {
        'primary':   RGBColor(0x1A, 0x8C, 0x5A),
        'secondary': RGBColor(0x5A, 0xB0, 0x8A),
        'accent':    RGBColor(0x0F, 0x66, 0x42),
        'note':      'エコ・環境・成長系向け',
    },
    'orange': {
        'primary':   RGBColor(0xE8, 0x6B, 0x1B),
        'secondary': RGBColor(0xF0, 0x95, 0x50),
        'accent':    RGBColor(0xB3, 0x4F, 0x0F),
        'note':      'エネルギー・活動系向け',
    },
    'red': {
        'primary':   RGBColor(0xCC, 0x1F, 0x2E),
        'secondary': RGBColor(0xE0, 0x5A, 0x66),
        'accent':    RGBColor(0x99, 0x0F, 0x1B),
        'note':      '緊急性・強調系向け（多用注意）',
    },
}
# ※各色コードは正式版リリース時（2026-03-31）の公開値を暫定転記。実装時に公式JSONで検証。

# ─── デフォルトパレット選択関数 ───
def get_palette(palette_name='default'):
    """
    パレット名を指定して色セットを返す。
    - 'default': 既存のNAVY / RED / ORANGE（後方互換）
    - 'solid_gray' 〜 'red': デジタル庁7色パレットから選択
    """
    if palette_name == 'default':
        return {'primary': NAVY, 'secondary': RED, 'accent': ORANGE}
    return DIGITAL_AGENCY_PALETTES.get(palette_name, {'primary': NAVY, 'secondary': RED, 'accent': ORANGE})
```

### 適用範囲

**v3.5時点では"内部保持のみ"とし、利用ロジックは実装しない**。理由：

- 現行のNAVY/RED/ORANGEブランドカラーは既にβ募集アイキャッチ・note投稿等で公開済み。色を変えるとブランド一貫性が崩れる
- パレット選択機能を実装するには**GPTs Instructions側でヒアリング項目追加が必要**（業種選択→パレット自動選択）。Instructions lite版は7,986字で余裕14字しかないため実装不可
- v3.5→v3.6でInstructions PLAIN版の再構成タイミングで、パレット選択UIを追加検討

**Phase Aでは以下のみ実施：**
- `DIGITAL_AGENCY_PALETTES` 辞書定義追加（コード追加15行程度）
- `get_palette()` 関数定義追加（10行程度）
- コード冒頭に出典コメント追加（3行）
- **既存のC-1〜C-3描画ロジックには一切触れない**（後方互換完全維持）

### 増加行数見積もり

- 追加：約30行（辞書 + 関数 + コメント）
- 変更：0行
- 削除：0行
- **リスク**：低（既存ロジック非改修のため）

---

## A-2：パーツ図鑑12種図解パターン辞書化

### 対象パターン（うちた氏保管庫より）

パーツ図鑑・図解集で確認できた12種の図解パターン。**うちた氏より商用利用許諾済み**（7/26スクショで確認）。

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
# ─── パーツ図鑑（うちた氏保管庫）由来 図解パターン辞書 ───
# 出典：うちた氏「パーツ図鑑_120種」「図解集_50種」（商用利用許諾済み・2026-07-26）

DIAGRAM_PATTERNS = {
    'category':     {'ja': '分類',           'use': '要素を並列カテゴリで整理', 'shape': 'grid'},
    'pyramid':      {'ja': 'ピラミッド',     'use': '階層・優先順位を上下で表現', 'shape': 'triangle'},
    'comparison':   {'ja': '比較',           'use': '2〜3要素の対比', 'shape': 'side_by_side'},
    'sequence':     {'ja': '順序',           'use': 'ステップ・時系列を左→右', 'shape': 'arrow_chain'},
    'cycle':        {'ja': '循環',           'use': '反復プロセスを円環で表現', 'shape': 'circle_arrow'},
    'funnel':       {'ja': '絞り込み',       'use': '上から下へ絞り込むファネル型', 'shape': 'trapezoid'},
    'timeline':     {'ja': '時間軸',         'use': '期間別のマイルストーン', 'shape': 'horizontal_bar'},
    'breakdown':    {'ja': '分解',           'use': '全体を構成要素に分解', 'shape': 'tree'},
    'contrast':     {'ja': '対比',           'use': '対照的な2要素の並列強調', 'shape': 'split_screen'},
    'integration':  {'ja': '統合',           'use': '複数要素の統合結果', 'shape': 'merge'},
    'framework':    {'ja': 'フレームワーク', 'use': '4象限マトリクス等', 'shape': 'quadrant'},
    'network':      {'ja': 'ネットワーク',   'use': 'ノード間の関係性', 'shape': 'node_edge'},
}

# ─── 診断結果→推奨図解パターンのマッピング ───
DIAGNOSIS_TO_PATTERN = {
    'proposal_categorization': 'category',      # C-2提案の分類表示
    'priority_ranking':        'pyramid',        # C-1優先課題の階層化
    'before_after':            'comparison',     # C-3 Before/After
    'user_flow':               'sequence',       # C-3ユーザー行動フロー
    'improvement_cycle':       'cycle',          # 改善サイクル説明
    'conversion_funnel':       'funnel',         # CVファネル
    'schedule':                'timeline',       # 改善スケジュール
    'score_breakdown':         'breakdown',      # 総合スコア内訳
    'ux_contrast':             'contrast',       # UX良/悪対比
    'impact_cost_matrix':      'framework',      # 影響度×コスト
    'site_structure':          'network',        # サイト構造
}
```

### 適用範囲

**v3.5時点では"辞書のみ・描画未実装"**。理由：

- 12種の描画ロジックを全部実装すると1,500〜2,000行の追加が必要（既存builder v16は約3,440行）
- v3.5コア（バナー掲載媒体診断）実装と競合するリソース
- Phase Bで、βFB集約後に「実務層が実際に必要とするパターン」を絞り込んでから実装

**Phase Aでは以下のみ実施：**
- `DIAGRAM_PATTERNS` 辞書定義追加（15行）
- `DIAGNOSIS_TO_PATTERN` マッピング定義追加（15行）
- 出典コメント追加（3行）
- 既存C-1〜C-3描画ロジックには一切触れない

### 増加行数見積もり

- 追加：約35行
- 変更：0行
- 削除：0行
- **リスク**：低（既存ロジック非改修のため）

---

## A-3：web-director.skillの位置づけ整理

### 現状把握

**web-director.skill**（入江さんご自身の知見の集約）は既にAnthropic Claude Skills形式でパッケージ化済み。内容：

- SKILL.md（メイン仕様書）
- references/templates.md（サイトマップ／要件定義／ユーザーフロー／ワイヤーフレーム／提案書のテンプレート集）
- references/ui-ux-best-practices.md（UI/UX原則・情報アーキテクチャ・カラー理論・タイポグラフィ）
- scripts/generate_sitemap.py（サイトマップ自動生成）

**現在の運用**：「マーケティングオーケストレーター」プロジェクトで活用中。UI診断ディレクターとの併用は未実施。

### 位置づけの整理

UI診断ディレクターの10項目評価軸と、web-director.skillの6大機能領域の対応マッピング：

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

**Phase Aでは"参照経路の設計のみ"を実施**。実際のGPTs Knowledge Base への登録はPhase Bで実施。

具体的には：

1. **web-director.skillを`_reference/skills/`に配置済み**（7/26入江さん実施済み）
2. **UI診断ディレクターのGPTsからは"参考知識"として位置づける**。GPTs Instructions内で明示的に呼び出す機能追加はしない（8000字制限のため）
3. **将来的にKnowledge Baseに登録する場合の設計案**を本ドキュメントに残す（下記）

### 将来のKnowledge Base登録設計案（Phase B以降）

```
UI診断ディレクターGPTs
  ├── Instructions（8000字以内）
  ├── Knowledge Base
  │   ├── web-director-condensed.md         # web-director.skillの要約版（15KB以内）
  │   ├── digital-agency-guidebook-summary.md  # デジタル庁ガイドブック要約版
  │   └── diagram-patterns-catalog.md        # パーツ図鑑12種の使い分け早見表
  └── Actions（builder呼び出し）
```

### 適用範囲

- Phase A：位置づけ整理と対応マッピング文書化のみ（本ドキュメントに記載）
- Phase B：Knowledge Base用要約版の作成＋実装

**Phase Aでのコード改修**：**なし**。

---

## 📊 Phase A 全体まとめ

| 項目 | A-1 | A-2 | A-3 |
|---|---|---|---|
| **コード追加行数** | 約30行 | 約35行 | 0行 |
| **既存ロジック変更** | なし | なし | なし |
| **後方互換性** | 完全維持 | 完全維持 | 完全維持 |
| **リスクレベル** | 低 | 低 | 極低 |
| **実装工数見積** | 30分 | 30分 | 0（文書のみ） |
| **v3.5への影響** | 内部保持のみ・利用ロジック未実装 | 内部保持のみ・描画ロジック未実装 | 参照経路の位置づけ整理のみ |

**Phase A完了後のbuilder状態**：
- 既存builder v16と機能的に**完全同一**（描画結果は1ピクセルも変わらない）
- 内部データとして**デジタル庁7色パレット＋図解パターン12種**を保持
- Phase B以降での本格実装の**基盤整備完了**

---

## 🚦 Phase A→Phase B→v3.5コア の依存関係

```
Phase A（本設計） ──── 内部辞書定義・位置づけ整理
   │                    ↓
   ├─ A-1（7色パレット辞書）
   ├─ A-2（12種図解パターン辞書）
   └─ A-3（web-director.skill位置づけ整理）
                        ↓
Phase B ──── Knowledge Base実装・要約版作成
   │
   ├─ B-1（web-director-condensed.md作成）
   ├─ B-2（digital-agency-guidebook-summary.md作成）
   ├─ B-3（diagram-patterns-catalog.md作成）
   └─ B-4（ppt-visual-catalog スキル新規作成）
                        ↓
v3.5コア ──── バナー掲載媒体診断の本実装
   │
   ├─ ヒアリング項目追加（掲載媒体・位置）
   ├─ C-2改善提案の媒体別粒度設計
   └─ builder改修（add_proposal_onepager内部ロジック）
```

**依存関係の要点**：
- **Phase A**は**v3.5コアの前提条件ではない**（並行可）が、下ごしらえとして完了させておくと Phase B・v3.5コアの実装が加速する
- **Phase B**は**v3.5コアと並行可能**だが、Knowledge Base反映は**v3.5リリース時のGPTs更新と同期**するのが自然

---

## 📅 実装スケジュール提案

| 日程 | 内容 | 主担当 |
|---|---|---|
| 7/27（月）**本日** | Phase A設計ドラフト作成（本書） | AIスライド |
| 7/28（火） | 入江さんレビュー／必要ならClaude-Chat相談 | 入江さん・3者 |
| 7/29（水）〜 7/30（木） | Phase A実装（builder v16.5作成） | AIスライド |
| 7/31（木） | X軸締切告知投稿（既定） | 入江さん |
| 8/1〜8/5 | Phase B設計（Knowledge Base要約版準備） | AIスライド |
| 8/6〜8/15 | v3.5コア（バナー掲載媒体診断）実装 | AIスライド＋Claude-Chat |
| 8月中旬 | βFB集約→3者レビュー | 3者 |
| 8月下旬 | v3.5統合→builder v17リリース | 3者 |
| 9月上旬 | β二次募集＋Brain販売開始 | 入江さん |

---

## ❓ 入江さん・Claude-Chatへの確認事項

### 入江さんへの確認事項（3件）
1. **A-1「v3.5時点では辞書内部保持のみ」の判断**でよいか？（既存ブランドカラーとの共存を優先する方針）
2. **A-2「12種図解パターン辞書のみで描画実装なし」の判断**でよいか？（Phase Bで絞り込んで実装）
3. **A-3「web-director.skillはKnowledge Base登録前段の位置づけ整理のみ」の判断**でよいか？

### Claude-Chatへの相談事項（Phase A実装前に投げたい2件）
1. デジタル庁7色パレットの**正式JSONファイル**（GitHub公開）から実際のカラーコードを取得すべきか？（本設計書は暫定転記のため）
2. うちた氏保管庫の**出典クレジット表記**をコード内コメントに含めるか、別途`ATTRIBUTIONS.md`として管理するか？

---

## 📌 参考資料

- Hub `/PROJECT_STATE.md`（正典・7/26最終更新）
- Hub `/ui-diagnosis-director/_reference/pptx_template/`（うちた氏保管庫9ファイル）
- Hub `/ui-diagnosis-director/_reference/design_system/【デジタル庁】ダッシュボードデザインの実戦ガイドブック.pdf`
- Hub `/ui-diagnosis-director/_reference/skills/web-director.skill`
- デジタル庁公式：https://www.digital.go.jp/resources/dashboard-guidebook
- デジタル庁GitHub：https://github.com/digital-go-jp/policy-dashboard-assets

---

**Phase A設計ドラフト 以上**
