# 🚀 引継書｜AIスライド（実装領域）｜2026-10-14（水）発行｜次セッションの起点
> 🔴 次セッションの冒頭で最初に読む。§4 の手順で冒頭照合（**47日目**・`manifest_20261014.tsv`）を行い、作業計画を提示して入江さんの承認を得てから着手する
> 前版：`NEXT_ENV_HANDOVER_20261013_aislide.md`（無変更で保存）。環境移行の注意は前版「★」をすべて引き継ぐ（往復数は入江さんの指定値／日付は翌日運用／`PYTHONDONTWRITEBYTECODE=1`／上書き前の事前照合／FUSE 層は待機・再実行／更新時刻を根拠にしない）
> ⚠️ FUSE 層：10/14 の初回照合は NG 111 → 60秒×2 の待機で NG 0。初回の NG が「不在」だけなら待機して再実行する

## ■ 1. 現在の状態（10/14 終業時点）
### 1-1. 採用版3点（変更なし）
| 対象 | md5 | B |
|---|---|---:|
| builder v17_217 `gpts-package/03_pptx_builder.py` | `47629772eaff` | 287,389 |
| PLAIN `01_Instructions_v3.5_PLAIN_20261007.md`（gpts-package／handover） | `59ebb0757fde` | 77,248 |
| lite `gpts-package/01_Instructions_v3.5.2_lite_20260902.md` | `61b5f4e7b928` | 16,223 |
### 1-2. 🆕 採用候補（判定10 待ち。GPTs 実機確認の後に、採用版の切替えを別の判定で行う）
| 対象 | 配置先 | md5 | B |
|---|---|---|---:|
| PLAIN `01_Instructions_v3.5_PLAIN_20261014.md` | gpts-package／handover | `338e8465b214` | 77,622 |
| lite `01_Instructions_v3.5.4_lite_20261014.md` | gpts-package／handover | `54cd8cb41f07` | 16,210 |
- 内容：J-3 の読込みを2行化（R-7 解消・test_*.pptx を作らない）。lite は J-3 コメントの圧縮と太字マーカーの除去で 7,991字（**残り9字**）
- GPTs の割当て：Instructions＝lite（8,000字上限）／Knowledge＝PLAIN（判定13）
- 判定の所在：判定1〜5＝`claude_chat_verdict_20261014_r7_director.md`／判定6〜15＝`claude_chat_verdict_20261014_restored_by_aislide.md`／確定版の差替文面＝`claude_chat_message_20261014_r7_revision_director_r1.md`
### 1-3. 商品一式・先行レビュー・インシデント
- 前版 §1-2・§1-3・§1-4 から変更なし。作成者不明の書込みは 10/12 以降なし（10/14 の終業時の同名2件は、入江さんの判断で対象外）。Genspark は回答待ち
- manifest の正：`handover/scripts/manifest_20261014.tsv`（10/14 終業時に発行）

## ■ 2. 次セッションのタスク
| # | 内容 |
|:---:|---|
| 0 | 冒頭手順（§4）＝47日目（manifest_20261014）。判定書の着到は **53件**との差 |
| 1 | 起案キュー②：公開ルール（7/26 §7）の改記案。**現行文・改記案・影響範囲の対照表**で示す（判定15 依頼2で着手承認済み） |
| 2 | 入江さんの作業の記録：Brain 販売ページの掲載（PJ 10/19（月）期限）／GPTs 実機確認（販売開始後で可。lite・PLAIN で各1診断、test_*.pptx が作られないこと・3ファイルが正常に出ること）→ 結果を復元記録に逐語で残す → 採用版切替えの判定依頼 → 変更履歴表 §1 に「読込み手順の2行化による R-7 解消」を追記（判定4） |
| 3 | 宿題（EOD 20261014 §5）から、余力に応じて |
| 4 | Genspark の回答が届いたら逐語で記録する |

## ■ 3. 検証スクリプト
前版 §3 と同じ（`handover/scripts/`・`export PYTHONDONTWRITEBYTECODE=1`）

## ■ 4. セッション冒頭手順
1. `verify_md5.py manifest_20261014.tsv`（47日目）
2. `find_prefix.py claude_chat_verdict_`（前回 53件）
3. 採用版3点と採用候補4か所が変わっていないこと
4. 🚨 handover/・scripts/・v37/・gpts-package/・sales/ に manifest にない新しいファイルがないか（あれば触れずに報告）
5. `__pycache__`（handover/ 直下6件・ほかは不在）
6. 作業計画を申告して承認を得る

## ■ 5. 「やらない」（前版 §6 を継続し、本日分を追加）
- 🚨 GPTs 実機確認の前に採用版を切り替える／GPTs 本体へ反映する
- 🚨 字数計画の審査をせずに lite の改訂を起案する（残り9字）
