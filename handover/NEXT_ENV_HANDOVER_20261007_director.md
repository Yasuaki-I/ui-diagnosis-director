# 新環境引継書 NEXT_ENV_HANDOVER_20261007_director.md
統括担当向け｜発行：2026-10-07（シリーズ締め：10/6・26往復で完走）

## §0 本書の位置づけ
- 本書を正とする。前版 20261004 版は無変更保存（md5・バイト数不変）。
- 新環境での往復数は 1 から数え直す。冒頭明記は「X往復目／（新シリーズの往復枠）」。
- 必達ルール9（日次照合）は通算継続。次回照合は **34日目**（manifest_20261007.tsv 照合済みであればその翌回）。
- 照合番号の規準（10/6裁定）：番号＝照合の通算実施回数。1日1回以上、事象駆動の追加照合も番号を進める。

## §1 環境差の注意
- 新環境（Genspark Hub）に `/mnt/aidrive` マウントなし。verify_md5.py 相当は gsk aidrive download で実体DLし md5先頭12桁＋バイト数を照合して再現（26日目から実績あり、規準変更なし）。
- 統括はドライブ直接書込禁止（第8条）。判定・申し送りは本文提示 → AIスライドが write5.py 5段階で配置。

## §2 現行ファイル（確定値）
- builder 採用版：`gpts-package/03_pptx_builder.py`＝v17_216（md5 332199ad7dc2／286637B）
- ロールバック：直近 `handover/03_pptx_builder_v17_215_20261004.py`／最終 `..._v17_213_20260922.py`
- PLAIN：`01_Instructions_v3.5_PLAIN_20261007.md`（77248B・59ebb0757fde、handover/・gpts-package/ 2か所）。前版 20260925（md5 93095a862614）無変更保存
- 台帳：SCRIPTS_REGISTRY_20261007／v37_candidates_20261007／bottom_margin_watchlist_20261007（20586B・61ecef1ee82e、B-3 は案(b2) 線なし＋観測で確定）
- 判定書点検：`find_prefix.py claude_chat_verdict_` → 直近 `claude_chat_verdict_20261007_restored_by_aislide.md`（12690B・1b1678d6606d、D-1〜D-8 解消済み）

## §3 完了事項（10/4〜10/6 シリーズ）
1. V37-1〜V37-3 全件通過：v17_214r1（n=6 下辺余裕・案(i)）／v17_215（cumulative・通算指標）／v17_216（footnote・注記）
2. v17_216 配布完了（write5 5段階通過・照合 NG 0件・verify_frozen 終了コード0）
3. PLAIN 改稿：J-7-5 部分改訂＋J-7-9（cumulative）・J-7-10（footnote）新設。`_est_text_w` は条文に0件（grep確認済）
4. 台帳改稿5件同時発行。宿題6（J-7-4 funnel任意キー棚卸し）を追加
5. 運用確定：verify_render 件数並記義務化（NG/WRAP/WARN/OVER1、既存NG行一覧化）／PYTHONDONTWRITEBYTECODE=1／__pycache__ クローズ
6. 事業方針：販売基点＝Brain／提供方法＝ハイブリッド型C（記事 sales/brain_ui_diagnosis_draft_v5_2_public.md＋実装パッケージ＋手順書）／価格＝原稿§12 の基準線維持（4,980／6,980／9,800円）
7. β二次募集 → Brain販売前・先行レビュー募集へ再フレーミング。企画案 beta2_plan_draft_20261007_rev1.md（5670B・19eaccf992c8）承認済み

## §4 未決事項・宿題（新環境の最初の確認対象）
1. **手順書の作成（新環境 最初のタスク候補）**：環境構築〜初回生成〜cumulative/footnote の使い方。AIスライド起案 → 統括判定
2. 更新ポリシー：builder 改版時の購入者再配布方針 → **入江さん決定待ち**
3. 購入者前提（ローカル Python 環境）・サポート範囲（手順書記載範囲のみ）の明文化 → 手順書に含める
4. 先行レビュー企画案 §5 の未決（記念日使用の要否／8月中未実施の説明文面／告知対象／告知文の扱い）。告知文全文起案 beta2_announcement_draft_20261007.md は保留中
5. PowerPoint 実機目視：代表ケース確認（LibreOffice 画素走査の代替にならない旨の留保）。B-3(b2) の観測を含む。連絡は入江さん経由
6. v3.8 系候補（棚上げ）：[A] タイトル折返し（推定器は実描画より 21.2px 小さく未信頼・第1段は PLAIN 運用ガイド＋生成前警告）／[B] 注記配置改編（body_bottom 14px 制約・footer 境界線 660 含む）
7. watchlist：v17_216 注記の最下インク y=638（下辺余裕 8px）継続観測
8. 台帳宿題6：J-7-4 の funnel 任意キー棚卸し（次回改稿時）

## §5 新環境で最初にやること
1. 本書の読込 → 照合実行（34日目）→ 判定書着到点検
2. §4-1 手順書の起案着手可否を入江さんに確認
3. PLAIN 20261007版・builder v17_216・台帳3点の現行性を manifest の md5 と突合
4. 販売方針C・価格基準線・β再フレーミングの前提を崩す外部変化の有無を確認

## §6 訂正記録（本シリーズ）
- 10/4 EOD「__pycache__ 3箇所」は誤り → 正しくは10/4新規2箇所（handover/__pycache__/ は9/14から既存）
- 統括側の書き写し誤り：企画案の正は `beta2_plan_draft_20261007_rev1.md`（`_rev1_rev1.md` は存在しない）
- システム時刻は実日付より1日前に表示される留保あり（stat 提示時）
