# AI-Drive アップロード指針 2026-07-16（木）終業版

**作成日**: 2026-07-16（木）夜
**目的**: 本日生成した全ファイルの、AI-Driveへのアップロード可否・アップロード先を明示

---

## ✅ 既にアップロード済み（対応不要）

| # | ファイル | AI-Drive共有URL |
|---|---|---|
| 1 | フォーマット見本ZIP | https://www.genspark.ai/aidrive/shared/d16fd89d-b4e9-4b3b-ac68-5d6b0f540876 |
| 2 | 実診断サンプルZIP | https://www.genspark.ai/aidrive/shared/79493b13-324b-4b95-bd44-0dc3356d26e3 |
| 3 | 使い方ガイド v1.0 PDF | https://www.genspark.ai/aidrive/shared/42ed457e-9f86-4b23-aa36-946f39f4fcb9 |
| GPTs | UI診断ディレクター本体 | https://chatgpt.com/g/g-6a21047a2d00819191d1192ed70214b1-uizhen-duan-teirekuta |

---

## 🚀 アップロード推奨ファイル（優先順）

### 【最優先】GPTs管理画面へのアップロード（入江さん作業）

以下はChatGPT GPTs管理画面から**UI診断ディレクター GPTs本体の設定へ**アップロードすべきファイル：

| # | ファイル | Project内パス | アップロード先 |
|---|---|---|---|
| G-1 | Instructions v3.3 | `gpts-package/01_Instructions_v3.3_PLAIN_20260716.md` | ChatGPT GPTs設定「Instructions欄」 |
| G-2 | PPTX生成コード v15.5 | `gpts-package/03_pptx_builder.py` | ChatGPT GPTs設定「Knowledge」 |
| G-3 | 業種別診断データ | `handover/SAMPLES_DATA_20260716.md` | ChatGPT GPTs設定「Knowledge」（任意） |

**要確認**: 既にアップロード済みか、入江さんへ状態確認をお願いします。

---

### 【推奨】AI-Drive `/ui-diagnosis-director/deliver/` 配下への追加アップロード

以下は現在Project内のみに保管されており、AI-Driveにアップロードすると入江さんが手元でも閲覧しやすくなるファイル：

| # | ファイル | Project内パス | AI-Driveアップロード推奨先 |
|---|---|---|---|
| A-1 | 使い方ガイド v1.0 PPTX（原本） | `_deliver/guide_v1_20260716/UI診断ディレクター_使い方ガイド_v1.0.pptx` | `/ui-diagnosis-director/deliver/guide/` |
| A-2 | GPTs Instructions v3.3 PLAIN | `gpts-package/01_Instructions_v3.3_PLAIN_20260716.md` | `/ui-diagnosis-director/gpts-package/` |
| A-3 | PPTX生成コード v15.5 | `gpts-package/03_pptx_builder.py` | `/ui-diagnosis-director/gpts-package/` |
| A-4 | build_guide.py | `gpts-package/build_guide.py` | `/ui-diagnosis-director/gpts-package/` |

**用途**: バックアップ・入江さん手元での閲覧・将来のv2.0/v6稿制作時の参照

---

### 【任意】AI-Drive `/ui-diagnosis-director/handover/` 配下への申し送りドキュメント

本日作成した重要ドキュメント。Projectに保管済のため必須ではないが、AI-Driveに置くと入江さんが手元で開けて便利：

| # | ファイル | Project内パス | 用途 |
|---|---|---|---|
| H-1 | HANDOVER_20260716.md | `handover/HANDOVER_20260716.md` | AIスライド→AIチャット申し送り |
| H-2 | CLAUDE_CHAT_20260716.md | `handover/CLAUDE_CHAT_20260716.md` | Claude-Chatへの申し送り |
| H-3 | VIOLET_20260716.md | `handover/VIOLET_20260716.md` | スケジュール関連 |
| H-4 | BETA_RECRUITMENT_2AXIS_20260716.md | `handover/BETA_RECRUITMENT_2AXIS_20260716.md` | 2軸β募集プラン |
| H-5 | BRAIN_INJECTION_RESULT_20260716.md | `handover/BRAIN_INJECTION_RESULT_20260716.md` | 販売ページ差込完了レポート |
| H-6 | BRAIN_PUBLISH_TODO_SEP_20260716.md | `handover/BRAIN_PUBLISH_TODO_SEP_20260716.md` | 9月公開前TODOリスト |
| H-7 | DELIVER_LINKS_20260716.md | `handover/DELIVER_LINKS_20260716.md` | 全共有URL控え |
| H-8 | DELIVER_LIST_20260716.md | `handover/DELIVER_LIST_20260716.md` | 配布物一覧 |
| H-9 | AIDRIVE_UPLOAD_20260716.md | 本ファイル | AI-Driveアップロード指針 |

**推奨アップロード先**: `/ui-diagnosis-director/handover/` （既存構造に追加）

---

## ❌ アップロード不要ファイル

以下はProject内のみで完結する開発用ファイル。AI-Driveへのアップロード**不要**：

| 分類 | ファイル群 | 理由 |
|---|---|---|
| 検証用テスト | `_verify/` 配下すべて | Projectで参照可能・入江さん手元不要 |
| Instructions中間版 | `01_Instructions_v3.1_PLAIN` / `v3.2_PLAIN` 等 | 最新v3.3のみで十分 |
| DIFFファイル | `01_Instructions_v3.X_DIFF_*.md` / `03_pptx_builder_v15_X_DIFF_*.md` | 開発履歴・アップ不要 |
| 実診断原本 | `_deliver/samples_final_20260716/real_diagnosis_openai/*_raw.pptx` | 匿名化前ファイル・社外配布禁止・入江さん管理下Projectのみで保持 |
| プレビューPNG | `_verify/guide_v1_preview/*.png` | 開発確認用 |
| 中間ZIP | `_deliver/ui_diagnosis_samples_3types_20260716.zip` | 4事例版（`ui_diagnosis_samples_4cases_20260716.zip`）が最終形 |

---

## 📋 入江さんの明日以降の作業手順（推奨）

### Step 1: 明日 7/17(金) 午前

- [ ] GPTs管理画面確認：Instructions v3.3 / コード v15.5 が反映済か確認
- [ ] 未反映であれば、`gpts-package/` からアップロード
- [ ] AIチャットのX軸原稿ドラフトを受領・確認

### Step 2: 7/17(金) 午後

- [ ] X予告ツイート予約セット
- [ ] GMB note記事公開予約セット
- [ ] 【任意】AI-Drive `/ui-diagnosis-director/handover/` に本日の申し送りドキュメント一式をアップロード（H-1〜H-9）
- [ ] 【任意】AI-Drive `/ui-diagnosis-director/deliver/guide/` にPPTX原本をアップロード（A-1）

### Step 3: 7/18(土) 09:45〜10:00

- [ ] GMB note公開の最終チェック
- [ ] X本告知メイン投稿→セルフリプライ1（URL）→セルフリプライ2（ハッシュタグ）の投稿

---

## 🗂 アップロード先の全体構造（推奨）

```
/ui-diagnosis-director/
├── deliver/                                    ← 配布物
│   ├── samples/                                ← フォーマット見本ZIP（配置済）
│   │   └── ui_diagnosis_samples_4cases_20260716.zip
│   ├── real_diagnosis/                         ← 実診断（配置済）
│   │   └── real_diagnosis_global_ai_20260716.zip
│   └── guide/                                  ← 使い方ガイド
│       ├── UI診断ディレクター_使い方ガイド_v1.0.pdf（配置済）
│       └── UI診断ディレクター_使い方ガイド_v1.0.pptx（推奨追加）
├── gpts-package/                               ← GPTs管理用（推奨追加）
│   ├── 01_Instructions_v3.3_PLAIN_20260716.md
│   ├── 03_pptx_builder.py
│   └── build_guide.py
├── handover/                                   ← 申し送り（既存構造に追加）
│   ├── HANDOVER_20260716.md
│   ├── CLAUDE_CHAT_20260716.md
│   ├── VIOLET_20260716.md
│   ├── BETA_RECRUITMENT_2AXIS_20260716.md
│   ├── BRAIN_INJECTION_RESULT_20260716.md
│   ├── BRAIN_PUBLISH_TODO_SEP_20260716.md
│   ├── DELIVER_LINKS_20260716.md
│   ├── DELIVER_LIST_20260716.md
│   └── AIDRIVE_UPLOAD_20260716.md（本ファイル）
└── _verify/                                    ← アップ不要
```

---

## ⚠️ セキュリティ注意事項

- 実診断原本（`*_openai_raw.pptx`）は**社外配布禁止**。AI-DriveへアップロードするならフォルダをPrivate設定に。
- 使い方ガイドPDFは既に共有URL発行済のため、AI-Drive上のURLを変更しないよう注意。
- GPTs本体URLはChatGPT側の公開設定（Public）で運用中。設定変更に注意。

---

**AI-Drive アップロード指針 以上**
