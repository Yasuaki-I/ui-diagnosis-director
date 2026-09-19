# 🔒 検証スクリプト md5 台帳｜`handover/scripts/`｜2026-09-20（日）
> 🔴 **本台帳は 2026-09-20 統括判定②（⭐ 案A承認）に基づく。** ⭐ **登録4本は第16条 細則12 の改変禁止領域に準じて管理する。**
- 発行：2026-09-20（日）4往復目｜AIスライド（実装領域）
- 命名：第18条 準拠

## ■ 1. 🔴 登録スクリプト（⭐ 4本）
| # | ファイル | 実測B | md5 | 役割 | 根拠 |
|:---:|---|---:|---|---|---|
| 1 | ⭐ `verify_md5.py` | ⭐ 3,297 | ⭐ `6b0eff1c2761` | ⭐ 成果物一覧の実測B・md5 照合 | ⭐ 必達ルール5・9 |
| 2 | ⭐ `verify_frozen.py` | ⭐ 2,810 | ⭐ `a59c8c37c978` | ⭐ 改変禁止領域の AST＋md5 逐語照合 | ⭐ 第16条 細則12 |
| 3 | ⭐ `write5.py` | ⭐ 2,992 | ⭐ `641d37c63b99` | ⭐ 5段階書込（⭐ ①削除→②書込→③md5→④待機→⑤再確認） | ⭐ 必達ルール9 |
| 4 | ⭐ `find_prefix.py` | ⭐ 2,205 | ⭐ `b93b2a6fa72d` | ⭐ 同一prefix 一覧（⭐ サフィックス付き検出） | ⭐ 第16条 細則13 |

🔴 **改変時は本台帳を新版発行で更新する（⭐ 判断原理10｜⚠️ 既存版を直接上書きしない）。**

## ■ 2. 🔴 使い方（⭐ セッション冒頭の標準手順）
```bash
cd /mnt/aidrive/ui-diagnosis-director/handover/scripts

# ① 成果物 md5 の一括照合（マニフェストは TSV：パス / md5 / バイト数）
python3 verify_md5.py manifest_YYYYMMDD.tsv
python3 verify_md5.py --check gpts-package/03_pptx_builder.py 6ee23f4bf3a1 271155

# ② 改変禁止領域の逐語照合（--allow に承認済の識別子を列挙）
python3 verify_frozen.py 旧builder.py 新builder.py --allow _V17_W14 _V17_W14_SAFE

# ③ 5段階書込
python3 write5.py /tmp/out/foo.md handover/foo.md

# ④ 受領確認（⚠️ 完全名検索ではサフィックス付きを検出できない）
python3 find_prefix.py claude_chat_verdict_
```
⭐ **相対パスは `/mnt/aidrive/ui-diagnosis-director` 直下として解決する。**

## ■ 3. ⭕ 実動検証（⭐ 2026-09-20 実施）
| # | スクリプト | 検証内容 | 結果 |
|:---:|---|---|:---:|
| 1 | ⭐ `verify_md5.py` | ⭐ 本日の成果物7件を照合 | ⭕ **NG 0件** |
| 2 | ⭐ `verify_frozen.py` | 🔴 **本日の builder 置換を照合（⭐ 関数95件）** | ⭕ **承認外の差分 0件** |
| 3 | ⭐ `write5.py` | 🔴 **本台帳を含む5件の配置に実使用** | ⭕ **全通過** |
| 4 | ⭐ `find_prefix.py` | ⭐ `03_pptx_builder_v17_2` で7件・`claude_chat_verdict_2026092` で0件を検出 | ⭕ **サフィックス付き 0件** |

## ■ 4. ⭐ 効果（⭐ 9/19 引継書 §12-3 施策①）
⭐ **毎セッション約105行を手書きしていた検証コードが、⭐ 呼び出し4行に置き換わる。** 🔴 **判断原理18（⭐ 実測してから動く）の実行コストが下がるため、⚠️「毎回書くのが面倒だから省略する」という劣化を構造的に防げる。**

---
⭐ **以上。**
