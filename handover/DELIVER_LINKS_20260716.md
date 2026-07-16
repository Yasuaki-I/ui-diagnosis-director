# 配布物 共有URL 控え

**作成日**: 2026-07-16（木）
**発行者**: 入江 靖章
**用途**: v5.2_ready 販売ページ差込用URLの一元管理

---

## 🔗 発行済み共有URL 一覧（5本）

| # | ファイル | 共有URL | 用途 |
|---|---|---|---|
| 1 | フォーマット見本ZIP（3業種9ファイル・約363KB） | https://www.genspark.ai/aidrive/shared/d16fd89d-b4e9-4b3b-ac68-5d6b0f540876 | §7 フォーマット見本ブロック |
| 2A | 実診断 UI Scorecard（Global AI 匿名化・PPTX） | https://www.genspark.ai/aidrive/shared/b3620204-39f9-410f-9147-d1689fced603 | 内部運用・差替用（販売ページ非掲載） |
| 2B | 実診断 Proposals（Global AI 匿名化・PPTX） | https://www.genspark.ai/aidrive/shared/53e98cea-0311-4bb1-8c86-1a4d1c57bf42 | 内部運用・差替用（販売ページ非掲載） |
| 2C | 実診断 Visual Board（Global AI 匿名化・PPTX） | https://www.genspark.ai/aidrive/shared/9d075467-2c76-4900-9b12-a3346ee0fcc5 | 内部運用・差替用（販売ページ非掲載） |
| 3 | 使い方ガイド v1.0 PDF（11P・約425KB） | https://www.genspark.ai/aidrive/shared/e1ba0b2a-7a69-4d88-9b99-476dc03d0cf5 | §13 有料エリア（購入者DL用） |

---

## 🆕 追加発行が必要なURL（1本）

Q2確定事項「実診断はZIP1本化で販売ページ掲載」に伴い、以下のZIPを追加アップロード＆共有URL発行してください。

| # | ファイル | AI-Driveアップロード先 | Project内取得元 |
|---|---|---|---|
| **2Z** | **実診断ZIP（3ファイル一式・約87KB）** | `/ui-diagnosis-director/deliver/real_diagnosis/` | `_deliver/samples_final_20260716/real_diagnosis_openai/real_diagnosis_global_ai_20260716.zip` |

### ファイル名（そのまま使用可）
```
real_diagnosis_global_ai_20260716.zip
```

### ZIP内訳
- `01_UI_Scorecard_global_ai.pptx`
- `02_Proposals_global_ai.pptx`
- `03_Visual_Board_global_ai.pptx`

### 発行URL控え欄

| 発行日 | URL |
|---|---|
| ______ | ______________________________________ |

---

## 🎯 販売ページ掲載方針（Q2確定事項の反映）

### 販売ページ本文掲載URL（無料エリア）

| # | ファイル | 掲載URL |
|---|---|---|
| 1 | フォーマット見本ZIP | ✅ (1) |
| 2 | 実診断ZIP（3ファイル一式） | ✅ (2Z) ← **追加発行待ち** |

### 販売ページ本文掲載URL（有料エリア §13）

| # | ファイル | 掲載URL |
|---|---|---|
| 3 | 使い方ガイド v1.0 PDF | ✅ (3) |

### 内部運用のみ保持

| # | ファイル | 用途 |
|---|---|---|
| 2A, 2B, 2C | 実診断個別3URL | 差替・サポート時の個別配布用（販売ページ非掲載） |

---

## ⚠️ 公開前 疎通確認チェックリスト

未ログイン状態またはシークレットウィンドウで、以下5本のURLが実際に開けるか確認してください。

- [ ] (1) フォーマット見本ZIP → ダウンロード成功
- [ ] (2Z) 実診断ZIP ← 追加発行後にチェック
- [ ] (3) 使い方ガイド v1.0 PDF → プレビュー / ダウンロード成功
- [ ] (参考: 2A/2B/2C 個別URL) → 開けることを確認（内部運用用）

### 疎通確認手順

1. Chrome または Safari で **シークレットウィンドウ**を開く
2. 各URLをアドレスバーへペースト → Enter
3. ログイン画面が出ずにダウンロード or プレビューが始まればOK
4. モバイル環境（iPhone Safari / Android Chrome）でも同様に確認

### GPTs本体URLの確認ポイント

- シークレットウィンドウではChatGPTログイン画面が出るのが正常（購入者はChatGPT Plusにログインした状態で使用するため）
- ログイン後に「UI診断ディレクター」のチャット画面が開ければOK
- 404・Errorが出る場合はGPTsの公開設定（Public）を確認

---

## 🔄 URL差し替え運用（v1.1以降のアップデート時）

同一URLで新バージョンに差し替える運用：

1. AI-Drive上で古いファイルを右クリック → 「置き換え」
2. 新ファイルをアップロード（同一パス）
3. 共有URLは自動的に維持される（購入者は都度リンクを控える必要なし）

### 適用対象
- 使い方ガイド v1.0 → v1.1（Section 3-4追加時）
- フォーマット見本ZIP（内容更新時）
- 実診断ZIP（新しい実診断事例に差替える場合）

---

## 📌 保管方針

- 本ファイル `handover/DELIVER_LINKS_20260716.md` は **一次情報**として保持
- v5.2 の販売ページ以外にも、note記事内リンクや今後の告知でも本ファイルのURLを参照
- 販売終了時は本ファイルの各URLに「販売終了（YYYY-MM-DD）」を追記

---

**共有URL控え 以上**
