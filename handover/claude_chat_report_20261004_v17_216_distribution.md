# 📦 配布報告｜builder v17_216 を gpts-package へ配布｜2026-10-04（日）AIスライド側 9往復目

- 根拠：10/4 統括判定 第7便（配布依頼 条件付き承認）
- 🔴 状態：**配布完了・照合 NG 0件**

## 1. 配布
| 項目 | 結果 | 確認手段（入力） |
|---|---|---|
| 書込 | ⭕ 5段階全通過（旧 md5 先頭12桁 c11ab9642b8e を削除→書込→照合→40秒待機→再読込） | `write5.py handover/03_pptx_builder_v17_216_20261004.py gpts-package/03_pptx_builder.py`（src は絶対パスで指定） |
| 候補版との一致 | ⭕ md5・バイト数とも一致（NG 0件） | `verify_md5.py --check gpts-package/03_pptx_builder.py <候補版のmd5> <候補版のB>`（期待値は候補版 `handover/03_pptx_builder_v17_216_20261004.py` の実測値を入力） |
- ⚠️ 1回目の実行で src を相対パスで渡し「src が存在しない」で停止した（①削除の前に停止するため書込なし）。src を絶対パスに直して再実行した
- 同時変更なし：PLAIN・lite・台帳・他スクリプトには触れていない

## 2. 添付事項（判定 第7便の指示）
| 事項 | 結果 | 確認手段（入力） |
|---|---|---|
| v17_213→v17_216 の verify_frozen | ⭕ **終了コード0**（承認外差分0件・差分は draw_funnel のみ） | `verify_frozen.py handover/03_pptx_builder_v17_213_20260922.py handover/03_pptx_builder_v17_216_20261004.py --allow draw_funnel` |
| 新規キー未指定時の同一性 | ⭕ 初版固定入力27ケース（cumulative・footnote を含まない）のうち **23ケースは v17.2.13 と差分行0** | 両版で同一入力から生成し TSV 行比較（`v37/v37_dist_tsv_diff_v17213_v17216_20261004.tsv`） |
| 差分ケース | ⭕ **差分は V37-3 適用の4ケースのみ**（funnel_n6・funnel_n6_long・funnel_n6_dense_long・funnel_n6_noscore）。差分ありの集合と V37-3 適用（notes）の集合が一致 | 同上 |

## 3. ロールバック版の保持
| 版 | 状態 | 確認手段 |
|---|---|---|
| 直近 v17_215（`handover/03_pptx_builder_v17_215_20261004.py`） | ⭕ 無変更 | `verify_md5.py --check`（9往復目の作業前に取得した md5 と一致） |
| 最終 v17_213（`handover/03_pptx_builder_v17_213_20260922.py`） | ⭕ 無変更 | `verify_md5.py --check`（manifest_20261003.tsv の値と一致） |

## 4. 実機目視の準備（判定 第7便）
- 代表8ケースの確認用ファイルを、**配布版 builder** で生成して配置した：`handover/v37/v37_device_check_v17216_20261004.pptx`
  - ① n=6 短文（V37-3 非該当） ② n=6 30字（V37-3） ③ n=6 長文・密度強化（V37-3） ④ 通算（auto） ⑤ 通算の衝突省略 ⑥ 注記 n=3 ⑦ 注記＋通算＋密度強化 n=6 ⑧ 注記の収容超過省略
- 🔴 **PowerPoint 実機目視は未実施（「実機目視待ち」）**。入江さんの実機での確認が必要。完了まで v3.7 完了の外部告知・β募集での言及は保留する

## 5. 記録待ち（台帳改稿時に実施｜同時変更禁止のため今回は行わない）
- 台帳改稿時の必須確認：① SCRIPTS_REGISTRY §6 の N-6 表記 ② v37_candidates §0 の方式記述 ③ V37-1 の形状変更 ④ V37-2 の形状変更 ＋ watchlist へ「funnel n=6・footnote 併用」を追加
- `manifest_20261003.tsv` の採用版 builder 行は旧 md5 のため、今後この照合では NG になる（配布による正当な変更）。終業時に `manifest_20261004.tsv` を発行し、新しい md5 で照合する
