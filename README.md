# UI診断ディレクター 配布パッケージ — 2026-07-04（金）

## 変更概要（v9）

7/3 いただいた FB 2件への完全対応。

### ハイライト

- **C-3 slide1・slide2 メタヘッダの左縦バーを NAVY 統一**（診断日・総合スコアが ORANGE/RED から NAVY へ変更）
- **C-3 slide1 Top3 レイアウト崩れ根本解消**（3件×2行折返しでも枠内に完全表示）
  - SUMMARY_H 94→66 / ISSUE_H 172→200 に再配分

詳細は `HANDOVER_20260704.md` を参照。

---

## パッケージ内容

```
ui-diagnosis-director_20260704/
├── README.md                              ← 本ファイル
├── HANDOVER_20260704.md                   ← 本日の詳細進捗
├── PROJECT_STATE.md                       ← プロジェクト現状
├── gpts-package/                          ← GPTs アップロード用ファイル一式
│   ├── 01_Instructions_WebDiagnosis_lite.md
│   ├── 03_pptx_builder.py                 ← v9（メタヘッダNAVY統一＋Top3根本解消）
│   ├── design_system.md
│   └── visual_data_schema.md
└── _verify/
    ├── v9/                                ← 通常データ生成PPTX（3種）
    └── v9_stress/                         ← 120字ストレステスト用PPTX（3種）
```

## GPTs 側の反映手順

1. OpenAI GPTs 編集画面を開く（**UI診断ディレクター**）。
2. **Knowledge** タブで `gpts-package/03_pptx_builder.py` を差し替えアップロード（他ファイルは v8 と同一なので上書き不要）。
3. **保存** して簡易動作テスト（URL 診断 1 件 → PPTX 3 種納品確認）。

---

（配布日：2026-07-04 / パッケージバージョン：v9）
