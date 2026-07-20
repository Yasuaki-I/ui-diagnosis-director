# 【Phase2 差し替え手順書】03_pptx_builder.py v15.6 修正版

**作成日**: 2026-07-20（月）
**対象**: 入江さん
**目的**: Visual Board描画欠落バグ（case01/04）の恒久対応。作業ミス防止のため1枚に集約。
**昇格判定**: 7/21（火）夜。本手順完了 + 追加ドッグフーディング2件（後述）合格で本番昇格。

---

## 1. 修正版ファイル

| 項目 | 内容 |
|---|---|
| ファイル名 | `03_pptx_builder_v16_20260720.py` |
| バージョン | v15.6（内部コメント標記） |
| サイズ | 175,003 bytes（旧版 173,549 bytes → +1,454 bytes） |
| 変更行 | 5ハンク、20行追加（機能追加のみ、既存ロジック破壊なし） |
| 配置先（GPTs） | `/mnt/data/03_pptx_builder.py` として上書き |

---

## 2. 修正内容サマリ（5点）

Claude-ChatのFB「キー名・プレースホルダ・空値時フォールバック実装必須」に対応。

| # | 修正箇所（関数） | 症状 | 修正内容 |
|---|---|---|---|
| ① | `add_visual_board` L2907付近 sections | スライド1のLP構造マップで「良好/要改善」マーク未反映 | `sections`空時に代替1件を挿入し空白描画を防止 |
| ② | `add_visual_board` L3200付近 scores | スライド2の10項目スコア図表が空白 | `scores`空時に⚠警告文＋原因説明を2行描画 |
| ③ | `add_visual_board` L3257付近 scores描画 | 同上のバー描画部 | `else`分岐でフォールバック描画（バー無し・警告文のみ） |
| ④ | `add_visual_board` L3282付近 strengths | 右上の強み欄が空白 | `strengths`空時に警告代替文を挿入 |
| ⑤ | `add_visual_board` L3312付近 priority_issues | コメント欄（最優先課題）が空白 | `priority_issues`空時に警告代替文を挿入 |

**設計方針**：
- **後方互換性維持**：既存の正常データは一切影響なし。空値時のみフォールバックが発火。
- **原因の可視化**：単に代替文を出すのではなく「visual_dataに`◯◯`キーが必須」と明記し、次回運用でGPTs側の欠落を検知可能に。
- **二重防御**：ビルダー側フォールバック + Instructions J-3の必須キー明記により、片方が破綻しても復旧可能。

---

## 3. 差し替え手順（GPTs Configure）

### Step 1. 現行ファイルのバックアップ（推奨・任意）

GPTsのKnowledgeから `03_pptx_builder.py` を一度ダウンロードして手元保存。ロールバック時に使用。

### Step 2. Knowledge差し替え

1. GPTs Configure画面を開く
2. **Knowledge**セクションから旧`03_pptx_builder.py`を**削除**
3. `03_pptx_builder_v16_20260720.py` をアップロード
4. **ファイル名を`03_pptx_builder.py`へリネーム**（GPTs側のexec()パスが固定のため必須）
5. Save

### Step 3. Instructions差し替え

同時に、Instructions v3.4-lite/PLAIN（本日更新版）も入れ替え：

- lite版：`01_Instructions_v3.4_lite_20260721.md`（**7,986字**、上限内）
- PLAIN版：`01_Instructions_v3.4_PLAIN_20260721.md`（参考用・アップロード不要）

**Instructions側の主な更新点**：
- **I項ステップ3**：「PPTX 3ファイル同時生成（**必ず3ファイル納品。他成果物への置き換え不可**）」を追記（スコープ・ドリフト防止／Case03対応）
- **J-3 visual_data必須フィールド**：`sections`各要素に`has_issue`(bool)と`status`('good'|'issue')の両方必須と明記。`scores/strengths/priority_issues`は`C-1から完全転記必須`を強調。

### Step 4. 動作確認（下記「4. 追加ドッグフーディング」参照）

---

## 4. 追加ドッグフーディング（昇格判定・7/21夜実施）

Claude-Chatより追加指示。所要合計10分。

### ① J-6 PNG完全禁止化の再確認（5分）

**手順**：
1. GPTsに「PC版FVのワイヤーフレームを個別画像で出してください」等、画像単体を要求
2. **期待動作**：ZIPリンクのみ提示。PNG直リンクは提示しない
3. さらに「PNGで直接欲しい」と押した場合の期待動作：
   - 「PNGはフルスクリーンから戻れない不具合事例のためZIP経由でお渡しします」と応答し、**ZIPのみ提示**

**合格基準**：PNG直リンクが1つも表示されない。

### ② Case03 PPTX必達動作確認（5分）

**手順**：
1. GPTsに「speakup-eikaiwa.com のスコアだけ手早く教えて」と依頼（Case03シナリオ再現）
2. **期待動作**：
   - ステップ1・2を短縮でも通過
   - ステップ3で**必ず3つのPPTXファイル**（Scorecard/Proposals/Visual Board）を納品
   - Markdown・ZIP・PNGでの代替出力は行わない

**合格基準**：3つのPPTXファイルが確実に生成・提示される。

### ③ Visual Board描画確認（case01再現）

**手順**：
1. GPTsに case01 と同じLP診断を依頼
2. 生成された `03_Visual_Board_{site_slug}.pptx` を開く
3. **期待動作**：
   - スライド1：LP構造マップの各セクションに「良好●」「要改善▲」マークが描画
   - スライド2：10項目スコアが数値＋バーで描画
   - スライド2 右下：最優先課題コメントが3件描画

**合格基準**：3項目すべて空白なし。

---

## 5. ロールバック手順（万が一の場合）

修正版で新たな不具合が出た場合：

1. GPTs Configure → Knowledge から `03_pptx_builder.py` を削除
2. Step 1でバックアップした旧版をアップロード（`03_pptx_builder.py`名）
3. Instructions は lite版のみ v3.3（7/16版）にロールバック
4. `HANDOVER_PHASE2_20260720.md` に事象を追記して報告

---

## 6. AI-Drive アップロード（入江さん作業）

以下2ファイルをAI-Driveへアップロード：

| ファイル | 保存先 | 用途 |
|---|---|---|
| `03_pptx_builder_v16_20260720.py` | `/ui-diagnosis-director/gpts-package/` | 差し替え用・履歴保管 |
| `01_Instructions_v3.4_lite_20260721.md` | `/ui-diagnosis-director/gpts-package/` | GPTs Instructions貼付元 |
| `HANDOVER_PHASE2_20260720.md`（本書） | `/ui-diagnosis-director/handover/` | 手順書保管 |

---

## 7. v3.5 繰越確認事項

Case05（6項目化）はv3.5繰越で確定。β運用中に「4項目のバラつき」が顕在化するかを観察してから判断（→BACKLOG_v3.5更新済み）。

---

## 8. サマリ

- ✅ 03_pptx_builder.py v15.6：Visual Board描画欠落5箇所すべてフォールバック実装
- ✅ Instructions lite版：Case03スコープドリフト防止文言追加、visual_data必須キー要件強化、7,986字で上限内
- ⏭ 7/21夜：追加ドッグフーディング2件合格→本番昇格
- ⏭ β運用開始後：Case05観察→v3.5仕様判断

以上。ご確認のほどよろしくお願いいたします。
