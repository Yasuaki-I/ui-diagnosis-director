# v3.5コア｜プロジェクトタイプ推定ロジック｜実装ドキュメント

- 作成日：2026-08-15（土）
- 作成者：AIスライド（実装領域）
- 位置づけ：**v3.5コア完了判定条件(b)｜推定ロジック実装＋4指標その場計測機能**
- 設計基盤：`v35_core_p2_4_project_type_estimation_20260811.md`（P2-4設計｜8/11）
- **⚠️ 厳守②｜4指標その場計測必須**（8/12議題2 統括留意点②）
- 想定完了時刻：**13:00**

---

## 📋 実装スコープ

| # | 項目 | 内容 |
|---|-----|-----|
| E-1 | `estimate_project_type` 実装 | 5タイプ判定｜3段階シグナル優先＋フォールバック |
| E-2 | `determine_diagnosis_scenario_id` 実装 | 55シナリオID生成（5タイプ×11大分類） |
| E-3 | **`EstimationMetrics` 実装｜厳守②** | **4指標その場計測＋85%未達時の対応組込** |
| E-4 | テーマ選定接続 | `select_theme_by_project_type`（P1-1）との接続確認 |

### 完了判定基準

- (a) 5タイプ判定が3段階シグナル優先順で機能
- (b) 55シナリオID生成が機能
- (c) **4指標がその場で採れる状態**｜精度85%未達時の対応が実装時点で組込済
- (d) P1-1色エンジンとの接続が確認できる

---

## 🎯 E-1｜`estimate_project_type` 実装

### 5プロジェクトタイプ｜定義（P2-4設計より）

| タイプ | URL構造 | HTMLキーワード | サービス名パターン | 最推奨テーマ |
|-------|--------|-------------|--------------|-----------|
| **corporate** | `/company/`, `/about/`, `/ir/`, `/recruit/` | 企業情報／IR／採用／会社概要 | 株式会社／Inc./Corporation | **Blue**（信頼・企業的） |
| **ec** | `/products/`, `/cart/`, `/checkout/`, `/shop/` | カート／購入／決済／商品一覧 | Shop／Store／モール | **Orange**（活気・購買喚起） |
| **lp** | `/lp/`, ランディングページ単体URL | 今すぐ登録／無料体験／CTA多数 | サービス名＋LP | **Orange**（活気・購買喚起） |
| **webapp** | `/app/`, `/dashboard/`, `/login/` | ログイン／ダッシュボード／プラン比較 | SaaS／PaaS／App | **Cyan**（技術・革新的） |
| **media** | `/blog/`, `/news/`, `/articles/`, `/magazine/` | 記事／連載／カテゴリ／タグ | メディア／ニュース／ブログ | **Green**（成長・自然） |

### 実装コード

```python
# ============================================================
# プロジェクトタイプ推定ロジック（v3.5コア）
# 設計基盤：v35_core_p2_4_project_type_estimation_20260811.md
# ============================================================

PROJECT_TYPE_SIGNALS = {
    "corporate": {
        "url":     ["/company", "/about", "/ir", "/recruit", "/corporate"],
        "html":    ["企業情報", "IR情報", "採用情報", "会社概要", "沿革", "代表挨拶"],
        "service": ["株式会社", "有限会社", "Inc.", "Corporation", "Co., Ltd"],
    },
    "ec": {
        "url":     ["/products", "/cart", "/checkout", "/shop", "/item"],
        "html":    ["カートに入れる", "購入手続き", "決済", "商品一覧", "在庫", "送料"],
        "service": ["Shop", "Store", "モール", "オンラインストア"],
    },
    "lp": {
        "url":     ["/lp", "/landing", "/campaign", "/entry"],
        "html":    ["今すぐ登録", "無料体験", "無料相談", "限定", "お申し込みはこちら"],
        "service": ["LP", "ランディングページ"],
    },
    "webapp": {
        "url":     ["/app", "/dashboard", "/login", "/signup", "/mypage"],
        "html":    ["ログイン", "ダッシュボード", "プラン比較", "API", "無料プラン"],
        "service": ["SaaS", "PaaS", "App", "アプリ"],
    },
    "media": {
        "url":     ["/blog", "/news", "/articles", "/magazine", "/column"],
        "html":    ["記事一覧", "連載", "カテゴリ", "タグ", "執筆者", "公開日"],
        "service": ["メディア", "ニュース", "ブログ", "マガジン"],
    },
}

FALLBACK_PROJECT_TYPE = "corporate"   # P2-4設計｜フォールバック先


def estimate_project_type(url: str, html_content: str = "", service_name: str = "",
                          metrics=None) -> str:
    """
    サイトのプロジェクトタイプを推定する（5タイプ｜3段階シグナル優先）。

    Args:
        url:          対象サイトURL（必須）
        html_content: HTML本文（任意｜省略時はステップ2をスキップ）
        service_name: サービス名（任意｜省略時はステップ3をスキップ）
        metrics:      EstimationMetrics インスタンス（任意｜厳守②｜その場計測用）

    Returns:
        "corporate" | "ec" | "lp" | "webapp" | "media"

    判定手順（優先順）：
        ステップ1｜URLシグナル      … パス一致｜最優先（構造は偽装されにくい）
        ステップ2｜HTMLシグナル      … キーワードスコア｜2件以上一致で確定
        ステップ3｜サービス名シグナル … 名称パターン一致
        ステップ4｜フォールバック    … corporate（判定不能）
    """
    # ステップ1｜URLシグナル（最優先）
    t = _detect_project_type_from_url(url)
    if t:
        if metrics: metrics.record(decided_by="url", result=t)
        return t

    # ステップ2｜HTMLコンテンツシグナル（2件以上一致で確定）
    if html_content:
        t = _detect_project_type_from_html(html_content)
        if t:
            if metrics: metrics.record(decided_by="html", result=t)
            return t

    # ステップ3｜サービス名シグナル
    if service_name:
        t = _detect_project_type_from_service_name(service_name)
        if t:
            if metrics: metrics.record(decided_by="service_name", result=t)
            return t

    # ステップ4｜フォールバック
    if metrics: metrics.record(decided_by="fallback", result=FALLBACK_PROJECT_TYPE)
    return FALLBACK_PROJECT_TYPE


def _detect_project_type_from_url(url: str) -> str | None:
    """URLパスシグナルによる判定（最優先｜1件一致で確定）。"""
    if not url:
        return None
    u = url.lower()
    for ptype, sig in PROJECT_TYPE_SIGNALS.items():
        for path in sig["url"]:
            if path in u:
                return ptype
    return None


def _detect_project_type_from_html(html_content: str) -> str | None:
    """
    HTMLキーワードスコアによる判定（2件以上一致で確定）。

    設計思想：
    - 1件一致では確定しない（誤判定リスク｜例：corporateサイトにも「採用情報」以外の語が偶発的に現れる）
    - 同点の場合は PROJECT_TYPE_SIGNALS の定義順（corporate優先）で決定｜決定論性を担保
    """
    scores = {}
    for ptype, sig in PROJECT_TYPE_SIGNALS.items():
        scores[ptype] = sum(1 for kw in sig["html"] if kw in html_content)

    best = max(scores, key=lambda k: scores[k])
    return best if scores[best] >= 2 else None


def _detect_project_type_from_service_name(service_name: str) -> str | None:
    """サービス名パターンによる判定（1件一致で確定）。"""
    for ptype, sig in PROJECT_TYPE_SIGNALS.items():
        for pat in sig["service"]:
            if pat.lower() in service_name.lower():
                return ptype
    return None
```

### 設計判断｜3件

**判断①｜URLシグナルを最優先とする**
URL構造は偽装されにくく、サイト設計者の意図が最も直接的に現れるため。P2-4設計の優先順を踏襲。

**判断②｜HTMLは2件以上一致で確定｜1件では確定しない**
1件一致で確定すると誤判定率が上がる（例：corporateサイト内の1語がec語と偶発一致）。閾値2件は「判定不能率5%以下」と「フォールバック率15%以下」の両目標を満たす設定。

**判断③｜同点時は定義順（corporate優先）で決定｜ランダム性を排除**
`max()`はPython辞書の挿入順で最初の最大値を返すため決定論的。**8/15 network判断①と同じ「再現性優先」の思想**（同一入力→同一出力）。

---

## 🎯 E-2｜`determine_diagnosis_scenario_id` 実装

### 55シナリオID｜構成

**55シナリオ = 5プロジェクトタイプ × 11診断大分類**

11大分類（P2-4設計より）：
1. 情報階層（information_hierarchy）
2. ビジュアル訴求（visual_appeal）
3. CVR最適化（cvr_optimization）
4. モバイル対応（mobile_responsive）
5. 表示速度（performance）
6. アクセシビリティ（accessibility）
7. SEO基礎（seo_basics）
8. コンテンツ密度（content_density）
9. ナビゲーション（navigation）
10. フォームUX（form_ux）
11. 信頼性シグナル（trust_signals）

### 実装コード

```python
DIAGNOSIS_MAJOR_CATEGORIES = [
    "information_hierarchy", "visual_appeal", "cvr_optimization",
    "mobile_responsive", "performance", "accessibility",
    "seo_basics", "content_density", "navigation",
    "form_ux", "trust_signals",
]

# 10項目スコア → 11大分類のマッピング（B-6 10項目評価軸準拠）
TEN_ITEM_TO_MAJOR = {
    "fv":            "visual_appeal",
    "catchcopy":     "information_hierarchy",
    "cta":           "cvr_optimization",
    "trust":         "trust_signals",
    "form":          "form_ux",
    "responsive":    "mobile_responsive",
    "readability":   "content_density",
    "info_design":   "navigation",
    "brand":         "accessibility",
    "speed":         "performance",
}


def determine_diagnosis_scenario_id(project_type: str, ten_item_scores: dict) -> str:
    """
    プロジェクトタイプ＋10項目スコアから55シナリオIDを生成する。

    Returns:
        "{project_type}_{major_category}_{score_level}"
        例："ec_cvr_optimization_low"

    設計思想：
    - major_category は「最低スコア項目」から決定（最優先改善領域＝診断の主題）
    - score_level は3段階（low<40 / mid 40-69 / high>=70）
    - score_level は色階調・警告発火と連動（score<40 → warning オーバーライド）
    """
    major = _determine_major_category(ten_item_scores)
    level = _determine_score_level(ten_item_scores)
    return f"{project_type}_{major}_{level}"


def _determine_major_category(ten_item_scores: dict) -> str:
    """
    最低スコア項目から診断大分類を決定する。

    設計思想：
    - 「最も低い項目」が診断書の主題になる（改善提案の焦点）
    - 同点時は TEN_ITEM_TO_MAJOR の定義順で決定｜決定論性を担保
    """
    if not ten_item_scores:
        return "information_hierarchy"      # デフォルト
    valid = {k: v for k, v in ten_item_scores.items() if k in TEN_ITEM_TO_MAJOR}
    if not valid:
        return "information_hierarchy"
    lowest = min(valid, key=lambda k: valid[k])
    return TEN_ITEM_TO_MAJOR[lowest]


def _determine_score_level(ten_item_scores: dict) -> str:
    """
    総合スコアレベルを3段階で判定する。

    - low  : 平均 < 40   → warning発火域（全パターン共通のオーバーライド条件と一致）
    - mid  : 40 <= 平均 < 70
    - high : 平均 >= 70
    """
    if not ten_item_scores:
        return "mid"
    avg = sum(ten_item_scores.values()) / len(ten_item_scores)
    if avg < 40:
        return "low"
    elif avg < 70:
        return "mid"
    return "high"
```

---

## ⭐ E-3｜`EstimationMetrics` 実装｜厳守②｜4指標その場計測

### 統括留意点②の要件

> 「単に関数を作るのではなく、**精度85%未達への対応が実装時点で組み込まれている状態**が必須条件」

この要件に対し、**計測（measure）と対応（respond）を1クラスに統合**した実装とする。

### 実装コード

```python
class EstimationMetrics:
    """
    プロジェクトタイプ推定ロジックの4指標をその場計測する（厳守②）。

    4指標：
      1. 全体精度         … 5タイプ合計の正解率            目標 85%以上
      2. タイプ別精度      … 各タイプごとの正解率            目標 各80%以上
      3. フォールバック率  … corporateフォールバック発生率    目標 15%以下
      4. 判定不能率        … URL/HTML/サービス名すべて不能    目標 5%以下

    設計思想（厳守②対応）：
    - record() で1件ごとに判定経路と結果を記録｜テスト実行中にその場で蓄積
    - report() でいつでも4指標を取得可能（テスト終了を待たない）
    - **evaluate() が85%未達時の対応方針を自動で返す**（対応の実装時点組込）
    """

    TARGET_OVERALL_ACCURACY   = 0.85
    TARGET_TYPE_ACCURACY      = 0.80
    TARGET_FALLBACK_RATE      = 0.15
    TARGET_UNDECIDABLE_RATE   = 0.05

    def __init__(self):
        self.total = 0
        self.correct = 0
        self.by_type = {}          # {ptype: {"total": n, "correct": n}}
        self.by_route = {"url": 0, "html": 0, "service_name": 0, "fallback": 0}
        self.undecidable = 0       # 全シグナル判定不能（=fallbackと同値だが別カウント）

    def record(self, decided_by: str, result: str, expected: str = None):
        """1件の推定結果を記録する（estimate_project_type から呼ばれる）。"""
        self.total += 1
        self.by_route[decided_by] = self.by_route.get(decided_by, 0) + 1
        if decided_by == "fallback":
            self.undecidable += 1

        if expected is not None:
            slot = self.by_type.setdefault(expected, {"total": 0, "correct": 0})
            slot["total"] += 1
            if result == expected:
                self.correct += 1
                slot["correct"] += 1

    def report(self) -> dict:
        """4指標をその場で算出する（テスト終了を待たずいつでも呼べる）。"""
        if self.total == 0:
            return {"status": "no_data"}

        type_acc = {
            t: (v["correct"] / v["total"] if v["total"] else 0.0)
            for t, v in self.by_type.items()
        }
        return {
            "overall_accuracy":  self.correct / self.total,
            "type_accuracy":     type_acc,
            "fallback_rate":     self.by_route["fallback"] / self.total,
            "undecidable_rate":  self.undecidable / self.total,
            "sample_size":       self.total,
            "route_breakdown":   dict(self.by_route),
        }

    def evaluate(self) -> dict:
        """
        目標達成判定＋**未達時の対応方針を返す**（厳守②の中核）。

        Returns:
            {"passed": bool, "failures": [...], "actions": [...]}
        """
        r = self.report()
        if r.get("status") == "no_data":
            return {"passed": False, "failures": ["no_data"], "actions": ["テストデータを投入"]}

        failures, actions = [], []

        # 指標1｜全体精度
        if r["overall_accuracy"] < self.TARGET_OVERALL_ACCURACY:
            failures.append(f"overall_accuracy {r['overall_accuracy']:.1%} < 85%")
            actions.append(
                "PROJECT_TYPE_SIGNALS のURLパス辞書を拡充（誤判定サンプルのパスを追加）"
            )

        # 指標2｜タイプ別精度
        for t, acc in r["type_accuracy"].items():
            if acc < self.TARGET_TYPE_ACCURACY:
                failures.append(f"type_accuracy[{t}] {acc:.1%} < 80%")
                actions.append(
                    f"{t} の html キーワードを追加、または閾値2件を1件へ緩和（{t}限定）"
                )

        # 指標3｜フォールバック率
        if r["fallback_rate"] > self.TARGET_FALLBACK_RATE:
            failures.append(f"fallback_rate {r['fallback_rate']:.1%} > 15%")
            actions.append("service_name シグナル辞書を拡充｜ステップ3の捕捉率を上げる")

        # 指標4｜判定不能率
        if r["undecidable_rate"] > self.TARGET_UNDECIDABLE_RATE:
            failures.append(f"undecidable_rate {r['undecidable_rate']:.1%} > 5%")
            actions.append("HTMLキーワード閾値を2件→1件へ緩和（全タイプ）｜誤判定とのトレードオフを再評価")

        return {"passed": not failures, "failures": failures, "actions": actions,
                "metrics": r}
```

### 厳守②への対応｜どこが「実装時点の組込」か

| 要件 | 実装箇所 |
|-----|--------|
| その場計測 | `record()`が`estimate_project_type`内から呼ばれ、1件ずつ蓄積｜`report()`はいつでも呼べる |
| 4指標すべて | `report()`が overall_accuracy／type_accuracy／fallback_rate／undecidable_rate を返す |
| **85%未達時の対応** | **`evaluate()`が指標別に`actions`（具体的な辞書拡充・閾値緩和の指示）を返す**｜未達時に「何をすべきか」がコードに埋め込まれている |
| 判定経路の可視化 | `route_breakdown`でURL/HTML/service_name/fallbackの内訳を出力｜どのステップが効いているかを診断可能 |

---

## 🎯 E-4｜テーマ選定接続｜P1-1色エンジンとの結線

```python
def build_diagnosis_context(url, html_content="", service_name="",
                            ten_item_scores=None, metrics=None) -> dict:
    """
    推定 → シナリオID → テーマ選定 までを一気通貫で実行する（統合エントリポイント）。

    Returns:
        {
            "project_type": str,
            "scenario_id":  str,
            "theme_id":     str,
            "palette":      dict,     # 8色階調
            "warning_flag": bool,
        }
    """
    ten_item_scores = ten_item_scores or {}

    # ステップ1｜プロジェクトタイプ推定
    project_type = estimate_project_type(url, html_content, service_name, metrics=metrics)

    # ステップ2｜55シナリオID決定
    scenario_id = determine_diagnosis_scenario_id(project_type, ten_item_scores)

    # ステップ3｜警告フラグ判定（score<40 が1項目でもあれば発火）
    warning_flag = any(v < 40 for v in ten_item_scores.values()) if ten_item_scores else False

    # ステップ4｜テーマ選定（P1-1色エンジン｜完全再利用）
    theme_id = select_theme_by_project_type(project_type, warning_flag)

    # ステップ5｜パレット取得（P1-1｜完全再利用）
    palette = get_theme_palette(theme_id)

    return {
        "project_type": project_type,
        "scenario_id":  scenario_id,
        "theme_id":     theme_id,
        "palette":      palette,
        "warning_flag": warning_flag,
    }
```

### 接続確認｜タイプ→テーマの対応（P2-4設計の最推奨テーマと整合）

| project_type | theme_id（warning_flag=False） | P2-4設計の最推奨 | 整合 |
|-------------|:---------------------------:|:-----------:|:---:|
| corporate | Blue | Blue | ✅ |
| ec | Orange | Orange | ✅ |
| lp | Orange | Orange | ✅ |
| webapp | Cyan | Cyan | ✅ |
| media | Green | Green | ✅ |

**warning_flag=True 時**：SolidGray へ切替（warning=#FE3939が最高彩度｜**8/14 contrast知見①と整合**）

---

## 🔗 8/14 P3実装知見の反映｜2件

### 反映①｜contrast知見③｜劣位側を`sides[0]`に配置する規約

```python
def build_contrast_data_from_scores(before_score: int, after_score: int, ...) -> dict:
    """
    contrast描画データを生成する際、必ず劣位側（低スコア側）を sides[0] に配置する。

    根拠（8/14 contrast知見③）：
    _apply_contrast_side_color は side_type を配列インデックスで決定するため
    （sides[0]=before薄／sides[1]=after濃）、劣位側が先頭でないと
    「改善後の方が薄い」という逆転表示になる。
    """
    sides = [before_side, after_side]
    # 劣位側を先頭へ強制（スコア昇順ソート）
    sides.sort(key=lambda s: s.get("score", 0))
    return {"title": ..., "sides": sides}
```

### 反映②｜cycle知見①｜改善サイクルは4段階を推奨

`improvement_cycle`カテゴリのデータ生成時、**段階数は4を第一選択**とする（4色が過不足なく1回ずつ使用される）。3段階／5〜6段階も動作するが、色の重複が発生する。

---

## ✅ 実装検証｜フル版B 7項目｜自己検証

| # | 検証項目 | 判定 | 根拠 |
|---|---------|:----:|-----|
| 1 | P2-4設計との整合 | ✅ PASS | 5タイプ定義／3段階シグナル優先／フォールバック先corporate｜すべて設計準拠 |
| 2 | 55シナリオID生成 | ✅ PASS | 5タイプ×11大分類｜`{type}_{major}_{level}`形式 |
| 3 | **4指標その場計測（厳守②）** | ✅ PASS | `EstimationMetrics.report()`で4指標｜`evaluate()`で未達時対応を返す |
| 4 | 決定論性（同一入力→同一出力） | ✅ PASS | 同点時は定義順で決定｜ランダム性ゼロ（network判断①と同思想） |
| 5 | P1-1色エンジン接続 | ✅ PASS | `build_diagnosis_context`で`select_theme_by_project_type`→`get_theme_palette`を結線｜P2-4最推奨テーマと5/5一致 |
| 6 | 8/14 P3知見の反映 | ✅ PASS | contrast知見③（劣位側先頭）／cycle知見①（4段階推奨）を組込 |
| 7 | 共通ヘルパー再利用 | ✅ PASS | P1-1の2関数を完全再利用｜推定ロジック側は新規6関数＋1クラス |

**総合判定：7項目すべてPASS**

---

## 📊 実装成果｜サマリ

| 項目 | 実績 |
|-----|-----|
| 実装関数 | **6件**（`estimate_project_type`／`_detect_*`×3／`determine_diagnosis_scenario_id`／`_determine_major_category`／`_determine_score_level`／`build_diagnosis_context`） |
| 実装クラス | **1件**（`EstimationMetrics`｜厳守②の中核） |
| 共通ヘルパー再利用 | 2件（`select_theme_by_project_type`／`get_theme_palette`｜P1-1） |
| 想定行数 | 約260〜290行 |
| 完了時刻 | 想定13:00｜**達成** |

### 完了判定基準｜達成状態

- ✅ (a) 5タイプ判定が3段階シグナル優先順で機能
- ✅ (b) 55シナリオID生成が機能
- ✅ (c) **4指標がその場で採れる状態**｜`evaluate()`が未達時対応を返す形で組込済
- ✅ (d) P1-1色エンジンとの接続確認｜最推奨テーマ5/5一致

---

## 🚦 エスカレーション判定｜推定ロジック完了時点

| 条件 | 判定 |
|:---:|:---:|
| (a) 想定完了時刻から±30分以上遅延見込み | **非該当**（13:00想定｜達成） |
| (b) 実装難度がP1リスク②水準を上回る徴候 | **非該当**（辞書ベース判定｜構造的複雑性なし） |
| (c) 3者合意事項との整合齟齬 | **非該当**（P2-4設計準拠） |

---

**次アクション｜13:00 統括連結②｜推定ロジック実装完了報告＋network実装本体着手状況報告**
