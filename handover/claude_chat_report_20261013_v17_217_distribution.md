# 📦 配布報告｜builder v17_217 を gpts-package へ配布｜2026-10-13（火）｜AIスライド側 24往復目
- 根拠：第24便 判定2（配布実施の承認｜条件A 完全充足）
- 🔴 状態：**配布完了**

## 1. 配布
| 手順 | 結果 | 確認手段（入力） |
|---|---|---|
| 事前照合（第18便 判定2 の再発防止＝初適用） | ⭕ 現行 `gpts-package/03_pptx_builder.py`＝`332199ad7dc2`・286,637B（v17_216）で期待値と一致 | `md5sum`・`stat` |
| 複写元の照合 | ⭕ `handover/03_pptx_builder_v17_217_20261010.py`＝`47629772eaff`・287,389B | 同上 |
| 書込前の手元複写 | ⭕ 現行版を作業領域に複写（`332199ad7dc2`） | `cp`・`md5sum` |
| 書込 | ⭕ 5段階全通過（①旧 `332199ad7dc2` を削除 →②書込 →③照合一致 →④40秒待機 →⑤再読込 287,389B・`47629772eaff` 一致） | `write5.py`（src は作業領域の複写） |
| 単発照合 | ⭕ `gpts-package/03_pptx_builder.py`＝`47629772eaff`・287,389B | `verify_md5.py --check`（番号を進めない） |
| 版の表示 | ⭕ `__version__`＝**17.2.17**・`__version_date__`＝2026-10-10 | 手順書 rev1 §3 の確認コマンド（作業領域に複写して `importlib` で読込・`PYTHONDONTWRITEBYTECODE=1`） |
| `__pycache__` | ⭕ gpts-package に作られていない | `find` |

## 2. ロールバック
- 直近：v17_216＝`handover/03_pptx_builder_v17_216_20261004.py`（`332199ad7dc2`｜保存済み・照合で無変更）
- 最終：v17_213（不変）

## 3. 採用版3点の「正」（配布後｜第24便 判定2）
| # | 対象 | md5 | B |
|:---:|---|---|---:|
| 1 | builder v17_217（`gpts-package/03_pptx_builder.py`） | `47629772eaff` | 287,389 |
| 2 | PLAIN（`01_Instructions_v3.5_PLAIN_20261007.md`） | `59ebb0757fde` | 77,248 |
| 3 | lite v3.5.2（`01_Instructions_v3.5.2_lite_20260902.md`） | `61b5f4e7b928` | 16,223 |
- 同時変更なし：PLAIN・lite・他スクリプトには触れていない
