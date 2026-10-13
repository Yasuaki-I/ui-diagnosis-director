<!-- 配置：AIスライド側 3往復目（2026-10-14 水）が write5.py 5段階で handover/ へ配置（第8条・統括 依頼1）。本文は入江さんが中継した統括の申し送りを逐語で転記 -->
<!-- ⚠️ 中継時の本文に「上記「PLAIN rev2」「lite rev2」」の差替文面そのものは含まれていなかった。r7_trial_20261014.md §1 の推奨2行を指すものとして扱う（AIスライド注記） -->

統括（新環境）からの申し送り｜2026-10-14（水）｜統括側 4往復目

【依頼1】判定便の配置（第8条・write5.py 5段階） 以下2ファイルを handover/ へ配置し、配置後の B・md5 先頭12桁を manifest_20261014 へ登録すること。

claude_chat_verdict_20261014_r7_director.md：R-7 解消案の判定（判定1＝2行方式で起案可・v17_218 不作成／判定2＝GPTs 実機確認を採用条件／判定3＝PLAIN は注記1行付き・lite は2行のみ／判定4＝変更履歴の対象に含める）→ 統括 1往復目の本文を逐語転記
本申し送りを claude_chat_message_20261014_r7_revision_director.md として配置（改訂文面の確定版として保全）
【依頼2】PLAIN rev2・lite rev2 の起案（R-7 解消・起案キュー⑧ の本線化） r7_trial_20261014.md §1 の推奨2行を承認済み。差替文面は以下の通り確定（統括が PLAIN 445行・lite 239行の現行を実測確認済み）。

PLAIN J-3：現行 exec(open('/mnt/data/03_pptx_builder.py').read()) を、上記「PLAIN rev2」の2行＋注記1行に差替え
lite J-3：同1行を「lite rev2」の2行のみに差替え（注記なし）
新版ファイル名・版番号・manifest 登録値（B・md5）は AIスライド側で算出・提案のうえ、統括の判定を受けること（配布は GPTs 実機確認後）
【依頼3】起動時作業の補完

find_prefix.py claude_chat_verdict_ の正規実行（10/13 時点＝該当51件との差分確認。統括側 sandbox では gsk ls の表示制約により件数を確定できず unconfirmed のまま）
handover/__pycache__/ の件数確認（6件目の削除可否は未判定のまま・経過観察）
r7_trial_20261014.md の B・md5 先頭12桁を算出し manifest_20261014 に登録
【報告】統括側で確定済みの事項

必達ルール9（46日目）：verify_md5.py manifest_20261013.tsv を統括側 sandbox で代替実行 → 照合202件・NG 0・終了コード0（データ行202 が正・awk NR 237 は空行・コメント行を含む全行カウントだった）
採用版3点の無変更：builder 47629772eaff・PLAIN 59ebb0757fde・lite 61b5f4e7b928（すべて一致）
手順書 rev1・抜粋版 r1：GPTs 経路の exec 読込行を含まないため無改訂で確定

---
<!-- 補遺（_r1）：AIスライド側 5往復目（2026-10-14 水）。原本 claude_chat_message_20261014_r7_revision_director.md（2,801B・1b68cfd6ef34）は無変更で残す。以下は統括 6往復目の申し送り（再掲2・申し送り本文）を逐語で転記。以後の起案はこの逐語版を正とする -->

■ 再掲2：統括 5往復目の「【確定版・逐語】R-7 差替文面」節【逐語】
PLAIN rev2（01_Instructions_v3.5_PLAIN_20261007.md J-3・445行目を差替え）
差替え前（現行・この1行を削除）：

Copyexec(open('/mnt/data/03_pptx_builder.py').read())
差替え後（以下の3行を挿入）：

Copy_ns = {'__name__': 'pptx_builder'}; exec(open('/mnt/data/03_pptx_builder.py').read(), _ns)
globals().update({k: v for k, v in _ns.items() if not k.startswith('__')})
# ※ デモを動かさないための読み込み方（builder 末尾のデモが実行されず、test_*.pptx は作られません）
lite rev2（01_Instructions_v3.5.2_lite_20260902.md J-3・239行目を差替え）
差替え後（以下の2行のみ・注記行は付けない）：

Copy_ns = {'__name__': 'pptx_builder'}; exec(open('/mnt/data/03_pptx_builder.py').read(), _ns)
globals().update({k: v for k, v in _ns.items() if not k.startswith('__')})
変更しないもの（再確認）
builder v17_217（47629772eaff）…17行の ⚠️ 注記も含め無変更
手順書 rev1（661391b40ae9）…importlib 記述のみのため無改訂
PLAIN 抜粋版 r1（69cd7836a014）…exec 読込行を含まないため無改訂

■ AIスライドへの申し送り（コピー用）
統括（新環境）からの申し送り｜2026-10-14（水）｜統括側 6往復目

【中継不備への対応】 差替文面の節が2回続けて届いていない事象を認識。本申し送りに、要求された2点（①統括 1往復目の判定本文 判定1〜4、②統括 5往復目の【確定版・逐語】R-7 差替文面 節）を省略なく同梱した。以後は本同梱版を正とすること。

【新判定：判定5（lite 8,000字超過 92字）】 コード2行は逐語維持・圧縮禁止。超過92字は lite J-3 コードブロック内のコメント4行（# C-1:〜# ★sectionsは…）の圧縮で回収。差分対照表と回収字数の内訳を添えて統括の判定を請うこと。J-3 内で回収不可の場合は他節の圧縮案を併記して再判定。上限超過の黙認は認めない。

【処理順の確認】 提示された順序（①判定書の配置 write5.py 5段階 → ②申し送りファイルの _r1 新規配置・原本無変更で残存 → ③PLAIN rev2・lite rev2 の起案＋B・md5 添付）を承認。lite rev2 は判定5 の圧縮案を含めて起案されたい。

（同梱1）統括 1往復目の判定本文（判定1〜4）＝本往復冒頭の「再掲1」を逐語転記 （同梱2）統括 5往復目の【確定版・逐語】R-7 差替文面＝本往復の「再掲2」を逐語転記
