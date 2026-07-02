# UI診断ディレクター 配布パッケージ — 2026-07-02（水）

## 変更概要（v7）

7/1 いただいた FB「C-1 結論帯の文字切れ」への根絶対応と、C-2 / C-3 への予防展開を実施。

### ハイライト
- **結論帯 / POINT帯 / 改善方向帯を 62px 化**（従来 32-38px）
- **文字数上限 120字・2行折返し対応**（`word_wrap=True`）
- **色ルール明確化**：ORANGE = 行動喚起（結論・POINT・改善方向）専用
- C-2 改善カード枠を **NAVY_LIGHT** に変更
- C-3 slide1 フロー総括を **ヘッダ右端 GOLD** に移設（重なり解消）

詳細は `HANDOVER_20260702.md` を参照。

---

## パッケージ内容

```
ui-diagnosis-director_20260702/
├── README.md                              ← 本ファイル
├── gpts-package/                          ← GPTs アップロード用ファイル一式
│   ├── 01_Instructions_WebDiagnosis_lite.md
│   ├── 03_pptx_builder.py                 ← v7（帯リサイズ・120字対応）
│   ├── design_system.md                   ← v1.2（色ルール §3.4 追加）
│   └── visual_data_schema.md
└── _verify/
    ├── v7/                                ← 通常データ生成PPTX（3種）
    │   ├── 01_UI診断スコアカード_スピークアップ英会話.pptx
    │   ├── 02_改善提案リスト_スピークアップ英会話.pptx
    │   └── 03_ビジュアル診断ボード_スピークアップ英会話.pptx
    └── v7_stress/                         ← 120字ストレステスト用PPTX（3種）
        ├── 01_C1_120char.pptx
        ├── 02_C2_120char.pptx
        └── 03_C3_120char.pptx
```

## GPTs 側の反映手順

1. OpenAI GPTs 編集画面を開く（**UI診断ディレクター**）。
2. **Instructions** タブに `gpts-package/01_Instructions_WebDiagnosis_lite.md` の全文を貼付（現在 8000 字以内で調整済）。
3. **Knowledge** タブに以下 3 ファイルを差し替えアップロード：
   - `gpts-package/03_pptx_builder.py`
   - `gpts-package/design_system.md`
   - `gpts-package/visual_data_schema.md`
4. **保存** して簡易動作テスト（URL 診断 1 件 → PPTX 3 種納品確認）。

## Obsidian への反映

- `HANDOVER_20260702.md` を「【UI診断ディレクター】/handover/」配下に配置。
- `PROJECT_STATE.md` を「【UI診断ディレクター】/」配下に上書き。

---

（配布日：2026-07-02 / パッケージバージョン：v7）
