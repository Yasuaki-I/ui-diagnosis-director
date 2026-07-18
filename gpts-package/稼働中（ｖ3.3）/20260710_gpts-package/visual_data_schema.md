# visual_data スキーマ仕様書

`add_visual_board()` および `build_full_report()` 関数に渡す `visual_data` 辞書の完全スキーマ。

このドキュメントは「正典」として扱い、Instructions・関数実装・GPTs側辞書生成ロジックが一貫してこれを参照する。

---

## ◆ 全体構造

```python
visual_data = {
    # === ヘッダー情報（再掲用、diagnosis から複製） ===
    'service_name': str,
    'purpose': str,
    'target': str,
    'diagnosis_date': str,
    'total_score': int,
    'rank': str,
    'rank_label': str,

    # === スライド1：LP構造マップ ===
    # [推奨スキーマ 2026-07-01] label + status を主キーとする
    # 後方互換: name/desc + has_issue 形式も引き続き読み込み可能
    'sections': [
        {'no': int, 'label': str, 'status': str},
        # status='✓' 課題なし（紺丸）、'!' or '✕' 課題あり（赤丸）
        # ... 5-9件
    ],

    # === スライド1：総評・最重要課題 ===
    'summary': str,
    # top_issues (推奨・複数) / top_issue (後方互換・単数) 両対応
    'top_issues': [str, str, str],  # 3件固定推奨

    # === スライド1：行動フロー（6ステップ固定） ===
    'flow_steps': [
        {'label': str, 'status': str, 'note': str},
        # ... 6件固定
    ],
    'flow_summary': str,

    # === スライド2：スコア視覚化 ===
    # [推奨スキーマ 2026-07-01] C-1のscoresと統一（`name`キー）
    # 後方互換: `category` キーも引き続き読み込み可能
    'scores': [
        {'name': str, 'score': int | str, 'max': int},
        # ... 5-10件
    ],

    # === スライド2：強み・課題（再掲） ===
    'strengths': [str, str, str],         # 3件固定
    'priority_issues': [str, str, str],   # 3件固定

    # === スライド3：Before/After Top3 ===
    'highlights': [
        {'no': int, 'title': str, 'target_area': str,
         'before': str, 'after': str, 'priority': str, 'effort': str},
        # ... 3件固定
    ],

    # === 全体：診断方向性 ===
    'direction': str,
}
```

---

## ◆ フィールド詳細

### ヘッダー情報

| フィールド | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `service_name` | str | ◯ | 全角20字 | 診断対象のサービス名。例：`'スピークアップ英会話'` |
| `purpose` | str | ◯ | 全角20字 | LP目的。例：`'無料体験予約・問い合わせ誘導'` |
| `target` | str | ◯ | 全角20字 | 対象タイプ。例：`'スクール紹介LP（英会話）'` |
| `diagnosis_date` | str | ◯ | 形式固定 | `'2026年6月4日'` 形式 |
| `total_score` | int | ◯ | 0-50 | diagnosis['total_score'] と一致 |
| `rank` | str | ◯ | 1文字 | `'S'/'A'/'B'/'C'/'D'` |
| `rank_label` | str | ◯ | 全角20字 | 例：`'標準的（改善余地あり）'` |

### スライド1：LP構造マップ

| フィールド | 型 | 必須 | 件数 | 説明 |
|---|---|:---:|:---:|---|
| `sections` | list | ◯ | 5〜9件可変 | LPの構造要素を上から順に |

各セクション要素：
| キー | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `no` | int | ◯ | - | 1〜9の通し番号 |
| `name` | str | ◯ | 全角10字 | セクション名。例：`'ファーストビュー'` |
| `desc` | str | ◯ | 全角18字 | セクション説明。例：`'メインコピー・説明'` |
| `has_issue` | bool | △ | - | True なら課題マーカー🚩付与（任意、デフォルトFalse） |

### スライド1：総評・最重要課題

| フィールド | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `summary` | str | ◯ | 全角40字 | LP全体の総評を1文で |
| `top_issues` | list[str] | ◯ | 各全角30字 | 3件固定。diagnosis['priority_issues'] と一致推奨 |

### スライド1：行動フロー

| フィールド | 型 | 必須 | 件数 | 説明 |
|---|---|:---:|:---:|---|
| `flow_steps` | list | ◯ | 6件固定 | ユーザー行動の6ステップ |
| `flow_summary` | str | ◯ | 全角40字 | フロー全体の総括 |

各ステップ要素：
| キー | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `label` | str | ◯ | 全角6字 | ステップ名。例：`'見る'`, `'興味'`, `'信頼'`, `'料金確認'`, `'迷う'`, `'離脱'` |
| `status` | str | ◯ | 1文字 | `'✓'` or `'✕'`（成立/不成立） |
| `note` | str | △ | 全角14字 | ✕時の理由。例：`'行動できるボタンがない'`（任意） |

**flow_steps 推奨6ステップ（LP診断時のデフォルト）：**
```python
[
    {'label': '見る', 'status': '✓', 'note': ''},
    {'label': '興味', 'status': '✓', 'note': ''},
    {'label': '信頼', 'status': '✕', 'note': '実績根拠が不足'},
    {'label': '料金確認', 'status': '✕', 'note': 'CTAが見つからない'},
    {'label': '迷う', 'status': '✕', 'note': '次のアクションが不明'},
    {'label': '離脱', 'status': '✕', 'note': '最終CTA文言が弱い'},
]
```

### スライド2：スコア視覚化

| フィールド | 型 | 必須 | 件数 | 説明 |
|---|---|:---:|:---:|---|
| `scores` | list | ◯ | 5〜10件 | diagnosis['scores'] から複製 |

各スコア要素：
| キー | 型 | 必須 | 説明 |
|---|---|:---:|---|
| `category` | str | ◯ | 項目名。例：`'ファーストビュー'` |
| `score` | int or str | ◯ | 1〜5 の整数、または N/A項目は `'－'` |
| `max` | int | ◯ | 通常 5 |

### スライド2：強み・課題

| フィールド | 型 | 必須 | 件数 | 文字数上限 | 説明 |
|---|---|:---:|:---:|---:|---|
| `strengths` | list[str] | ◯ | 3件固定 | 各全角40字 | diagnosis['strengths'] と一致 |
| `priority_issues` | list[str] | ◯ | 3件固定 | 各全角50字 | top_issues と同じでよい |

### スライド3：Before/After Top3

| フィールド | 型 | 必須 | 件数 | 説明 |
|---|---|:---:|:---:|---|
| `highlights` | list | ◯ | 3件固定 | proposals['proposals'][:3] と紐付け |

各ハイライト要素：
| キー | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `no` | int | ◯ | - | 1〜3 |
| `title` | str | ◯ | 全角30字 | 提案タイトル |
| `target_area` | str | ◯ | 全角20字 | 改善箇所 |
| `before` | str | ◯ | 全角40字 | 現状の状態 |
| `after` | str | ◯ | 全角40字 | 改善後の状態 |
| `priority` | str | ◯ | 1文字 | `'高'`/`'中'`/`'低'` |
| `effort` | str | ◯ | 1文字 | `'小'`/`'中'`/`'大'` |

### 全体共通

| フィールド | 型 | 必須 | 文字数上限 | 説明 |
|---|---|:---:|---:|---|
| `direction` | str | ◯ | 全角40字 | 改善の方向性。全スライド共通でフッター上に表示 |

---

## ◆ diagnosis / proposals からの複製マッピング

GPTs は `diagnosis` と `proposals` を生成済み。`visual_data` を作るときは以下を複製する：

```python
visual_data = {
    # diagnosis から複製
    'service_name': diagnosis['service_name'],
    'total_score': diagnosis['total_score'],
    'rank': diagnosis['rank'],
    'rank_label': diagnosis['rank_label'],
    'scores': diagnosis['scores'],
    'strengths': diagnosis['strengths'],
    'priority_issues': diagnosis['priority_issues'],
    'top_issues': diagnosis['priority_issues'],   # 同じ内容
    'summary': diagnosis['conclusion'],            # 結論を総評に再利用

    # proposals から複製（上位3件）
    'highlights': [
        {'no': p['no'], 'title': p['title'], 'target_area': p['target_area'],
         'before': p['before'], 'after': p['after'],
         'priority': p['priority'], 'effort': p['effort']}
        for p in proposals['proposals'][:3]
    ],

    # C-3固有（GPTsが新規生成）
    'purpose': '...',                  # LP/EC等の目的
    'target': '...',                   # 診断対象タイプ
    'diagnosis_date': '...',           # 診断日
    'sections': [...],                 # LP構造マップ 5-9件
    'flow_steps': [...],               # 行動フロー 6件
    'flow_summary': '...',             # フロー総括
    'direction': '...',                # 改善方向性
}
```

---

## ◆ バリデーション規則

`add_visual_board()` 呼び出し前に、以下を確認すること：

| ルール | 違反時の挙動 |
|---|---|
| `sections` の件数が 5〜9 | エラー（5未満→警告、9超→9件で切り捨て） |
| `top_issues` / `strengths` / `priority_issues` が 3件 | 3件未満→空文字で埋める、3件超→3件で切り捨て |
| `flow_steps` が 6件固定 | 6件未満→デフォルト値で補完、6件超→6件で切り捨て |
| `highlights` が 3件固定 | 3件未満→空辞書で埋める、3件超→3件で切り捨て |
| `score` が `'－'` または 1-5 | それ以外→`'－'` 扱い |
| 各文字列の上限超過 | 上限まで切り詰めて `…` 付与 |

---

## ◆ バナー診断時の特例

バナー診断時は以下のフィールドを以下のように扱う：

| フィールド | バナー時の扱い |
|---|---|
| `sections` | バナー要素分解（例：`'メインビジュアル'`/`'コピー'`/`'CTA'`/`'装飾'`）を 3〜5件 |
| `flow_steps` | バナー用6ステップ（`'見る'`/`'読む'`/`'共感'`/`'クリック判断'`/`'迷う'`/`'スルー'`） |
| `scores` | N/A項目は `'－'` 表示（フォーム設計/レスポンシブ/読みやすさ/情報設計/表示速度・技術） |

---

## ◆ 変更履歴

- v1.0（フェーズ1）：表形式スコア視覚化、Before/After Top3、6ステップ行動フロー
- v1.1（予定・フェーズ2）：matplotlib によるレーダーチャート差し替え
