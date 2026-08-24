# 🔧 v17 P2｜実装手順書｜pyramid / sequence / framework

- 作成日：**2026-08-24（月）18:00連結③**
- 作成者：AIスライド（実装領域）
- 対象成果物：`v17_p2_draw_patterns_20260824.py`／統合版 `03_pptx_builder_v17_10_20260824.py`
- 実装担当：**入江さん**（開発環境への投入）
- 位置づけ：⭐ **第16条「段階2（実装済）」までがAIスライドの対象範囲**

---

## ■ 0. 3行サマリ

1. ⭐ **統合版 `03_pptx_builder_v17_10_20260824.py`（223,802B）が完成済み**。⭐ **これを `03_pptx_builder.py` にリネームしてアップするだけ**（追記作業は不要）。
2. 版数は ⭐ **`__version__ = '17.1.0'`／`__version_date__ = '2026-08-24'`** に更新済み。
3. 回帰確認は ⭐ **3系統118項目すべてPASS**（P2 57＋P1 47＋既存互換14）。

⚠️ **P1のときと違い、今回はAIスライド側で統合済みのファイルをお渡しします。** ⭐ **入江さんの作業は「リネーム＋差し替え」の1手順のみです。**

---

## ■ 1. 投入手順（所要5分）

| # | 手順 |
|:---:|---|
| 1 | `/handover/v17/03_pptx_builder_v17_10_20260824.py` をダウンロード |
| 2 | ⚠️ ⭐ **ファイル名を `03_pptx_builder.py` に変更**（Instructions 436行が参照する名前） |
| 3 | GPTs「UI診断ディレクター v3.4-dev」の Knowledge で旧ファイルを削除し、差し替え |
| 4 | `print(__version__)` が `17.1.0` を返すことを確認 |

---

## ■ 2. 追加された公開API｜3関数

| 関数 | 用途 | 要素数 |
|---|---|:---:|
| `draw_pyramid(slide, palette, data)` | 階層・優先順位を上下で表現 | 3〜5段 |
| `draw_sequence(slide, palette, data)` | ステップ・時系列を左→右 | 3〜6 |
| ⚠️ `draw_framework(slide, palette, data)` | 2軸マトリクス（4象限等） | 4〜9セル |

⭐ **`draw_pattern()` は6種対応に拡張**（P1 3種＋P2 3種）。⚠️ **P3の5種（funnel/timeline/contrast/cycle/network）と `integration` は `category` に退避し、`notes` に「v17 P3 で実装予定」と記録します（例外は出ません）。**

### 2-1. データスキーマ

```python
# pyramid（index 0 が頂点＝最重要）
{'title': str,
 'levels': [{'label': str, 'score': int|None, 'description': str}, ...]}   # 3〜5

# sequence（左→右固定｜原本 use「ステップ・時系列を左→右」）
{'title': str,
 'steps': [{'label': str, 'score': int|None, 'description': str}, ...]}    # 3〜6

# framework（⚠️ 軸ラベル必須｜実装済11種で唯一 requires_axes=True）
{'title': str,
 'axis_x_label': str, 'axis_y_label': str,          # ⚠️ 必須
 'axis_x_low': str, 'axis_x_high': str,             # 省略可（既定「低」「高」）
 'axis_y_low': str, 'axis_y_high': str,
 'cells': [{'row': int, 'col': int, 'label': str,
            'score': int|None, 'items': [str]}, ...]}   # 4〜9
```

### 2-2. 呼び出し例（1枚出力）

```python
exec(open('/mnt/data/03_pptx_builder.py').read())
prs = create_presentation()
pal = get_theme_palette('Blue')

slide, report = add_diagram_slide(prs, 'pyramid', pal, {
    'title': '改善施策の優先順位',
    'levels': [
        {'label': 'FVに申込CTAを追加',   'score': 82, 'description': '工数小・効果大'},
        {'label': '料金直下に導線を置く', 'score': 64, 'description': '工数小'},
        {'label': '信頼材料を増やす',     'score': 48, 'description': '工数中'},
    ]}, page_num=1, total=1)
prs.save('/mnt/data/out.pptx')
print(report)
```

⚠️ **`_add_header` / `_add_footer` を自分で書く必要はありません**（`add_diagram_slide` が内部で処理し、見出し二重表示を構造的に防ぎます）。

---

## ■ 3. ⭐ 集約表からの実装マッピング（第9条｜参照行明記）

| # | パターン | 集約表の規定 | 実装箇所 |
|:---:|---|---|---|
| P2-1 | `pyramid` | min3／max5／vertical／`hierarchical` | `V17_PYRAMID_TIERS`｜3段=[primary,midtone,light]／4段／5段の色列を段数別に固定 |
| P2-2 | `sequence` | min3／max6／horizontal／`progressive` | `V17_SEQUENCE_TIERS`｜primary→lightest へ段階変化 |
| P2-3 | ⚠️ `framework` | min4／**max9**／⚠️ **axes True**／grid／`positional_quadrant` | `_framework_tier()`｜⭐ **2x2は象限別**（右上=primary／左下=light）、3x2・3x3は行位置ベース |

### 原本参照行（第17条 細則7）

| 参照内容 | ファイル | 参照行 |
|---|---|:---:|
| `pyramid` = 'ピラミッド'／'階層・優先順位を上下で表現'／'triangle' | `03_pptx_builder_v16_5_20260728.py` | ⭐ **182行** |
| `sequence` = '順序'／'ステップ・時系列を左→右'／'arrow_chain' | 同上 | ⭐ **184行** |
| `framework` = 'フレームワーク'／'4象限マトリクス等'／'quadrant' | 同上 | ⭐ **192行** |
| `priority_ranking`→pyramid／`user_flow`→sequence／`impact_cost_matrix`→framework | 同上 | ⭐ **197〜209行** |

⭐ **`DIAGRAM_PATTERNS` は1文字も書き換えていません。** 拡張層は `DIAGRAM_PATTERN_SPEC.update()` で追加（統括判定①｜2層構造の維持）。

---

## ■ 4. ⚠️ 集約表と実装記録の相違｜3件｜集約表を採用（8/24 統括指示）

| 項目 | 実装記録（8/9〜8/11） | ⭐ 集約表（8/15） | 採用 |
|---|:---:|:---:|:---:|
| `sequence` 最大 | 7ステップ | ⭐ **6** | ⭐ **集約表** |
| `sequence` 方向 | horizontal / vertical 選択可 | ⭐ **horizontal 固定** | ⭐ **集約表**（原本 use「左→右」と整合） |
| ⚠️ `framework` セル数 | ⚠️ **4/6/9 の固定値のみ**（不一致は即フォールバック） | ⭐ **min4／max9（範囲）** | ⭐ **集約表** |

⭐ **`framework` は5・7・8セルが実装記録では未定義だったため、要素数からグリッドを決定論的に自動選定する方式を採用しました。**

| セル数 | 選定グリッド |
|:---:|:---:|
| 4 | ⭐ **2×2（象限別の意味付けを適用）** |
| 5・6 | 3×2 |
| 7・8・9 | 3×3 |
| ⚠️ ≦3・≧10 | ⚠️ `category` フォールバック |

⚠️ **事例016として登録し、事後承認を仰ぎます。**

---

## ■ 5. ⚠️ 実装上の判断｜2件

| # | 事項 | 判断 | 根拠 |
|:---:|---|---|---|
| 1 | ⭐ **pyramid の台形描画** | ⭐ **`add_freeform` を使わず、幅可変の矩形段（頂点30%→基層90%）で近似** | 原則④（基本図形優先）｜⚠️ **`add_freeform` は python-pptx のバージョン依存があり、PowerPoint実機での再現性を担保できない**｜⭐ **`cycle`（P3）の `BLOCK_ARC` で環境依存を踏んだ教訓の先行適用** |
| 2 | ⭐ **framework の軸ラベル欠落時** | ⭐ **例外を投げず `category` フォールバック** | 原則①（例外を投げない）｜⚠️ **実装記録は `raise ValueError` だが、P1で確立した「`raise` ゼロ」方針と矛盾する** |

---

## ■ 6. ⚠️ P1に遡及した修正｜1件

### ⭐ フォールバック理由が呼び出し側に届いていなかった

⚠️ **P1・P2の全フォールバック箇所で、理由を記録した `notes` が破棄されていました。**

```python
# ❌ 旧実装（P1 v17.0.0〜v17.0.1）
notes.append('理由...')                    # ローカル変数に記録
return draw_category(...)                  # ← 別の report を返すため notes が消える

# ⭕ 修正後（v17.1.0）
rep = draw_category(...)
rep['notes'].insert(0, '理由...')          # 戻り値の notes に記録
return rep
```

⚠️ **影響**：`breakdown`・`comparison`・`pyramid`・`sequence`・`framework` がフォールバックした際、⭐ **「なぜ劣化描画になったのか」がGPTs側で分からない状態**でした。⭐ **描画自体は正常に行われていたため、出力物への影響はありません。**

⭐ **本件はP2のテストハーネスT5（軸ラベル欠落の検証）で検出しました。** ⚠️ **P1のテスト47項目では `fallback_from` のみを検証し、`notes` の内容を検証していなかったため見逃していました。**

---

## ■ 7. 回帰確認結果｜⭐ 3系統118項目 すべてPASS

| 系統 | 項目数 | 結果 |
|---|:---:|:---:|
| ⭐ **P2 ハーネス**（21組合せ＋境界値＋軸欠落＋決定論性＋範囲＋フォント＋ディスパッチ＋重複なし） | **57** | ⭕ **PASS 57／FAIL 0** |
| P1 ハーネス（回帰） | **47** | ⭕ **PASS 47／FAIL 0** |
| 既存C-1〜C-3 後方互換（XML一致） | **14** | ⭕ **PASS 14／FAIL 0** |
| **計** | ⭐ **118** | ⭕ **全PASS** |

### P2 ハーネスの内訳

| 群 | 内容 | 項目 |
|:---:|---|:---:|
| T1 | ⭐ **21組合せ**（3パターン×7テーマ） | **21** |
| T2〜T4 | 境界値（pyramid 2-6／sequence 2-7／framework 3-10） | 18 |
| T5 | ⚠️ **framework 軸ラベル欠落3パターン→category** | 3 |
| T6・T7 | 警告閾値40／`score=None` 許容 | 2 |
| T8 | ⭐ **決定論性**（100回／50回／50回→各1種） | 3 |
| T9・T10 | 描画範囲（60〜660px）／14pt下限・メイリオ強制 | 6 |
| T11・T12 | `resolve_pattern`／⭐ **6種対応＋P3退避6件** | 2 |
| T13 | ⭐ **ヘッダ帯と内見出しの重複なし** | 1 |
| T14 | 既存レイアウトとの同時使用 | 1 |

---

## ■ 8. 目視確認（LibreOffice→PNG）

| パターン | 確認結果 |
|---|---|
| `pyramid` | ⭕ **4段が頂点30%→基層90%で階段状に描画**｜⚠️ **score 36 の最下段が警告色（薄赤）＋濃字で判読可** |
| `sequence` | ⭕ **5ステップが左→右に並び、間に矢印**｜⭐ **STEP番号・ラベル・スコア・説明の4段が収まっている** |
| `framework` | ⭕ **2×2象限＋X軸（下）・Y軸（左）ラベル**｜⭐ **セル内の箇条書き最大3件も描画** |

### ⚠️ 設計上の留意点（不具合ではない）

⚠️ **`framework` の色は「セル位置」で決まるため、スコアが高いセルが薄く見えることがあります。**

⭐ 例：右上「要検討 62%」＝primary（濃紺）／右下「即着手 88%」＝midtone（明るい青）

⭐ **これは集約表の `positional_quadrant`（位置で意味付け）の仕様どおりです。** ⚠️ **スコアで濃淡を付けたい場合は `category`（`uniform_parallel`）や `breakdown`（`proportional`）が適切であり、パターンの選択で解決する設計です。**

---

## ■ 9. 統合完了時に更新すべき記録（第16条）

| # | 対象 | 更新内容 |
|:---:|---|---|
| 1 | `v17_implementation_log_*.md` | `pyramid`／`sequence`／`framework` の**段階3を ✅ に更新** |
| 2 | 同上 | ⭐ **判定根拠として `__version__ = '17.1.0'` を明記**（細則7） |
| 3 | GPTs環境 | ⭐ **`print(__version__)` が `17.1.0` を返すことで段階4を ✅** |

⚠️ **AIスライドは段階3・段階4を自己申告で ✅ にしません。** 入江さんの投入完了報告を受けて更新します。

---

**v17 P2 実装手順書｜2026-08-24（月）18:00連結③｜⭐ 3系統118項目 全PASS｜⭐ 統合版を同梱（入江さんの作業はリネーム＋差し替えのみ）｜⚠️ 段階2までがAIスライドの対象**
