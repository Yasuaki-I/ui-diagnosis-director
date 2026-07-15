# AI-Drive アップロード指示書 2026-07-15

**作成日**: 2026-07-15（水）
**位置づけ**: Instructions v3.0 制作完了に伴う、入江さんによる手動アップロード指示
**対象アップロード先**: AI-Drive `/ui-diagnosis-director/` 配下

---

## 🎯 アップロードの目的

1. **チャット環境で私（もう一人の私）がv3.0成果物を参照できるようにする**
2. **7/16-17のGPTs反映作業時に、Configure貼付用ファイルへ即アクセスできるようにする**
3. **プロジェクト履歴として保存する**

---

## 📦 アップロード対象ファイル一覧

### 【最優先】Configure貼付作業に必要（2ファイル）

| # | Projectパス | アップロード先 | 用途 |
|---|---|---|---|
| 1 | `gpts-package/01_Instructions_v3.0_PLAIN_20260715.md` | `/ui-diagnosis-director/gpts-package/` | ★GPTs Instructions欄への貼付用 |
| 2 | `gpts-package/03_pptx_builder.py` | `/ui-diagnosis-director/gpts-package/` | GPTs Knowledge欄への差替用（v15.0.0） |

### 【重要】履歴保持・チャット環境参照用（4ファイル）

| # | Projectパス | アップロード先 | 用途 |
|---|---|---|---|
| 3 | `gpts-package/01_Instructions_v3.0_DIFF_20260715.md` | `/ui-diagnosis-director/gpts-package/` | v2→v3.0の差分メモ |
| 4 | `handover/GPTS_UPDATE_v15_GUIDE.md`（更新版） | `/ui-diagnosis-director/handover/`（上書き） | 反映手順書（v3.0同時反映対応版） |
| 5 | `handover/CHAT_ONBOARDING.md`（更新版） | `/ui-diagnosis-director/handover/`（上書き） | 呼称ルール明文化版 |
| 6 | `handover/HANDOVER_20260715.md`（本日終業時作成予定） | `/ui-diagnosis-director/handover/` | 本日進捗の申し送り |

### 【任意】完全な履歴保持を希望する場合（1ファイル）

| # | Projectパス | アップロード先 | 用途 |
|---|---|---|---|
| 7 | `gpts-package/01_Instructions_v3.0_20260715.md` | `/ui-diagnosis-director/gpts-package/` | 本体版（コメント含む・履歴用） |

---

## 📋 アップロード手順

### Step 1: AI-Drive Webインターフェイスを開く

Genspark AI-Drive の Web インターフェイスにアクセスし、`/ui-diagnosis-director/` フォルダを開いてください。

### Step 2: `gpts-package/` サブフォルダに移動

- 既存フォルダ：`/ui-diagnosis-director/gpts-package/`
- 既存内容：`01_Instructions.md`, `01_Instructions_WebDiagnosis.md`, `01_Instructions_WebDiagnosis_lite.md`, `03_pptx_builder.py`

### Step 3: 新規ファイルをアップロード（3〜4ファイル）

以下のファイルをProject（Genspark AI Slidesのファイルブラウザ）からダウンロード → AI-Driveへアップロード：

**最優先2ファイル**：
- ✅ `01_Instructions_v3.0_PLAIN_20260715.md`
- ✅ `03_pptx_builder.py`（v15.0.0・既存を上書き）

**重要1ファイル**：
- ✅ `01_Instructions_v3.0_DIFF_20260715.md`

**任意1ファイル**（履歴保持を希望する場合）：
- ⚪ `01_Instructions_v3.0_20260715.md`

### Step 4: `handover/` サブフォルダを更新

以下のファイルをProject側から取得 → AI-Drive `/ui-diagnosis-director/handover/` へアップロード（既存を上書き）：

- ✅ `GPTS_UPDATE_v15_GUIDE.md`（v3.0同時反映対応版）
- ✅ `CHAT_ONBOARDING.md`（呼称ルール明文化版）
- ✅ `HANDOVER_20260715.md`（本日終業時に作成予定）

### Step 5: アップロード確認

AI-Drive上で以下のファイル一覧を確認：

**`/ui-diagnosis-director/gpts-package/`**:
- 既存：01_Instructions.md, 01_Instructions_WebDiagnosis.md, 01_Instructions_WebDiagnosis_lite.md
- 更新：03_pptx_builder.py（v15.0.0）
- **新規**：01_Instructions_v3.0_PLAIN_20260715.md ★
- **新規**：01_Instructions_v3.0_DIFF_20260715.md
- 新規（任意）：01_Instructions_v3.0_20260715.md

---

## 🎯 アップロード優先順位まとめ

**時間がない場合、最低限これだけアップロード**：
1. `01_Instructions_v3.0_PLAIN_20260715.md` ← GPTs反映に必須
2. `03_pptx_builder.py`（v15.0.0） ← GPTs反映に必須
3. `HANDOVER_20260715.md`（終業時作成後） ← チャット環境用

**時間があれば、全ファイルアップロード**：
- 上記3ファイル + 4〜7の全ファイル
- 履歴保持・チャット環境からの参照性向上

---

## 💡 配布用ZIPを活用する簡略化案

`_deliver/ui-diagnosis-director_20260715.zip` に本日の全成果物が格納されています（終業時に作成予定）。

### 簡略化されたアップロード手順

1. `_deliver/ui-diagnosis-director_20260715.zip` を Projectからダウンロード
2. AI-Drive上で解凍（`aidrive_tool` の decompress機能を使用）
3. 解凍後のフォルダから必要なファイルを `/ui-diagnosis-director/` 配下へ移動

**ただし**：既存ファイルの上書き扱いに注意。既存の `handover/` 配下ファイルを維持しつつ、新規ファイルだけを追加する形になります。

---

## ✅ アップロード完了後の確認

以下すべてがAI-Driveに存在すれば、7/16-17のGPTs反映作業の準備完了です：

- [ ] `/ui-diagnosis-director/gpts-package/01_Instructions_v3.0_PLAIN_20260715.md`
- [ ] `/ui-diagnosis-director/gpts-package/03_pptx_builder.py`（v15.0.0）
- [ ] `/ui-diagnosis-director/handover/GPTS_UPDATE_v15_GUIDE.md`（更新版）
- [ ] `/ui-diagnosis-director/handover/HANDOVER_20260715.md`

---

## 🔗 関連ドキュメント

- 反映手順書：`handover/GPTS_UPDATE_v15_GUIDE.md`
- 発注文：`handover/INSTRUCTIONS_V3_BRIEF_20260715.md`
- 差分メモ：`gpts-package/01_Instructions_v3.0_DIFF_20260715.md`

---

**アップロード指示書 以上**

**作業のご負担があれば、私に相談してください。時間がなければ最優先2ファイルだけでOKです。**
