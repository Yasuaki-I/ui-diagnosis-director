# AI-Drive アップロード指針 2026-07-18（土）

**作成日**: 2026-07-18（土）終業時
**目的**: 本日AIスライド側で生成したファイルの、AI-Driveへの適正配置を明示する

---

## 📋 本日生成ファイル一覧（7ファイル）

### Genspark project ルートに生成されたファイル

| # | ファイル名 | サイズ | 種別 |
|---|---|---|---|
| 1 | `01_Instructions_v3.4_PLAIN_20260721.md` | 27,450 bytes | GPTs Instructions参照用フル版 |
| 2 | `01_Instructions_v3.4_lite_20260721.md` | 16,624 bytes | GPTs Configure貼付用 |
| 3 | `BACKLOG_v3.5_20260718.md` | 3,396 bytes | v3.5候補管理台帳 |
| 4 | `HANDOVER_20260718.md` | 7,460 bytes | AIチャット申し送り |
| 5 | `CLAUDE_CHAT_20260718.md` | 5,264 bytes | Claude-Chat申し送り |
| 6 | `VIOLET_20260718.md` | 3,816 bytes | ヴァイオレット申し送り |
| 7 | `AIDRIVE_UPLOAD_20260718.md` | 本ファイル | 本アップロード指針 |

---

## 🎯 AI-Driveアップロード先マッピング

### `/ui-diagnosis-director/gpts-package/` へアップロード（2ファイル）

**入江さん側で既にアップロード完了済み**（本日の途中で実施頂きました）。以下2ファイルが該当：

| ファイル | 用途 |
|---|---|
| `01_Instructions_v3.4_lite_20260721.md` | GPTs v3.4-dev環境の指示欄に貼付済 |
| `01_Instructions_v3.4_PLAIN_20260721.md` | 参照用フル版・変更履歴保持 |

**アップロード状態**：✅ **完了確認済**（8往復目時点）

### `/ui-diagnosis-director/handover/` へアップロード（5ファイル）

以下5ファイルは**入江さん手動アップロード推奨**：

| # | ファイル | 優先度 | アップロード先 | 備考 |
|---|---|---|---|---|
| 1 | `BACKLOG_v3.5_20260718.md` | ★★★ 高 | `/ui-diagnosis-director/handover/` | v3.5候補判定基準・8月中旬参照 |
| 2 | `HANDOVER_20260718.md` | ★★★ 高 | `/ui-diagnosis-director/handover/` | AIチャット申し送り本体 |
| 3 | `CLAUDE_CHAT_20260718.md` | ★★★ 高 | `/ui-diagnosis-director/handover/` | Claude-Chat向け相談事項 |
| 4 | `VIOLET_20260718.md` | ★★★ 高 | `/ui-diagnosis-director/handover/` | ヴァイオレット向けスケジュール |
| 5 | `AIDRIVE_UPLOAD_20260718.md` | ★★ 中 | `/ui-diagnosis-director/handover/` | 本ファイル自身・アップロード履歴 |

### 追加でアップロードが必要なファイル（次回セッションプロンプト）

ステップ5で作成予定：

| # | ファイル | アップロード先 |
|---|---|---|
| 6 | `NEXT_SESSION_PROMPT_20260718.md` | `/ui-diagnosis-director/handover/` |

---

## 🚀 入江さん側のアップロード作業手順

### 全体像

Genspark project ルートには本日生成の7ファイル（うち2ファイルはgpts-package用として既にAI-Drive反映済）が保管されています。残り5ファイル＋次回プロンプト1ファイル = **6ファイル**を`/ui-diagnosis-director/handover/`にアップロードお願いします。

### 手順（推奨）

1. Genspark project ルートの各ファイルをダウンロード（右パネルのダウンロードボタン）
2. AI-Drive `/ui-diagnosis-director/handover/` を開く
3. 一括アップロード（ドラッグ&ドロップ推奨）
4. アップロード完了後、ファイル数を確認

### アップロード後のファイル数確認

`/ui-diagnosis-director/handover/` の想定ファイル数：
- 本日終業前時点：47ファイル
- 本日追加分：6ファイル
- **本日終業後想定**：**53ファイル**

もし数が合わない場合は、7/17までのファイルとの重複や、アップロード漏れの可能性があります。

---

## 📌 補足：既にアップロード完了済みの本日ファイル

以下2ファイルは、入江さん側で本日途中（v3.4-dev貼付タイミング）にアップロード完了済：

- ✅ `/ui-diagnosis-director/gpts-package/01_Instructions_v3.4_lite_20260721.md`
- ✅ `/ui-diagnosis-director/gpts-package/01_Instructions_v3.4_PLAIN_20260721.md`

これらの重複アップロードは不要です。

---

## 🎯 明日以降のAI-Drive利用予定

### 7/19（日）

- **AIチャット側で通読レビュー支援時**：`/ui-diagnosis-director/gpts-package/01_Instructions_v3.4_lite_20260721.md` を参照
- **通読フィードバック整理時**：新規ファイル `handover/V3.4_FEEDBACK_20260719.md` 等を作成する可能性あり

### 7/20（月）

- **セルフドッグフーディング時**：新規ファイル `handover/V3.4_DOGFOODING_20260720.md` 等を作成する可能性あり

### 7/21（火）

- **文言微調整完了時**：Instructions v3.4 最終版として `handover/gpts-package/01_Instructions_v3.4_lite_20260721.md` を上書き更新の可能性あり
- **v3.4本番昇格判定時**：Claude-Chat判断に応じて、v3.3本体との入替え or 併存判断

---

## ✅ アップロードチェックリスト（入江さん確認用）

- [ ] `HANDOVER_20260718.md` → `/handover/`
- [ ] `CLAUDE_CHAT_20260718.md` → `/handover/`
- [ ] `VIOLET_20260718.md` → `/handover/`
- [ ] `AIDRIVE_UPLOAD_20260718.md` → `/handover/`
- [ ] `BACKLOG_v3.5_20260718.md` → `/handover/`
- [ ] `NEXT_SESSION_PROMPT_20260718.md` → `/handover/`（ステップ5で作成）

すべて `/ui-diagnosis-director/handover/` 配下にフラットアップロードでOK（サブフォルダ化不要）。

---

**AIDRIVE_UPLOAD指針 以上**
