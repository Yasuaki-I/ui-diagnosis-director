# UI診断ディレクター 配布パッケージ — 2026-07-03（木）

## 変更概要（v8）

7/2 いただいた FB「C-3 レイアウトの微修正」への対応。C-3 の 3つの表示課題を解消。

### ハイライト

- **C-3 slide2 フッター重複解消**：スコア表 10項目が改善方向帯と重ならないよう縦領域を再計算
- **C-3 slide1 Top3 全表示**：Top3 3件が枠内にすべて収まる（切り詰め22字＋line_height 1.35）
- **C-3 slide1 状態伝達の明確化**：記号（✓/!/○）を廃し、サークル色（NAVY=良好/RED=要改善）＋ヘッダ右端の凡例チップで意味を伝える設計へ統一

詳細は `HANDOVER_20260703.md` を参照。

---

## パッケージ内容

```
ui-diagnosis-director_20260703/
├── README.md                              ← 本ファイル
├── HANDOVER_20260703.md                   ← 本日の詳細進捗
├── PROJECT_STATE.md                       ← プロジェクト現状（v8まで反映）
├── gpts-package/                          ← GPTs アップロード用ファイル一式
│   ├── 01_Instructions_WebDiagnosis_lite.md
│   ├── 03_pptx_builder.py                 ← v8（C-3 レイアウト微修正）
│   ├── design_system.md                   ← v1.2
│   └── visual_data_schema.md
└── _verify/
    ├── v8/                                ← 通常データ生成PPTX（3種）
    │   ├── 01_UI診断スコアカード_スピークアップ英会話.pptx
    │   ├── 02_改善提案リスト_スピークアップ英会話.pptx
    │   └── 03_ビジュアル診断ボード_スピークアップ英会話.pptx
    └── v8_stress/                         ← 120字ストレステスト用PPTX（3種）
        ├── 01_C1_120char.pptx
        ├── 02_C2_120char.pptx
        └── 03_C3_120char.pptx
```

## GPTs 側の反映手順

1. OpenAI GPTs 編集画面を開く（**UI診断ディレクター**）。
2. **Knowledge** タブで `gpts-package/03_pptx_builder.py` を差し替えアップロード（他ファイルは v7 と同一なので上書き不要）。
3. **保存** して簡易動作テスト（URL 診断 1 件 → PPTX 3 種納品確認）。

## Obsidian への反映

- `HANDOVER_20260703.md` を「【UI診断ディレクター】/handover/」配下に配置。
- `PROJECT_STATE.md` を「【UI診断ディレクター】/」配下に上書き。

---

（配布日：2026-07-03 / パッケージバージョン：v8）
