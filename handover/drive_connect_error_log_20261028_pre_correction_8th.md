# AIドライブ接続エラー記録｜運用日 2026-10-28（水）｜判定143(c)
> 原因は未確認。表示された文言はそのまま写す。時刻はコンピューターのシステム時刻（UTC・運用日の前日付）

| # | システム時刻（UTC） | 操作 | 表示された文言（そのまま） |
|:-:|---|---|---|
| 1 | 2026-10-27 00:4x（1往復目） | `ls /mnt/aidrive/ui-diagnosis-director` | `ls: cannot access '/mnt/aidrive/ui-diagnosis-director': No such file or directory`（45秒・60秒待って再確認しても同じ） |
| 2 | 2026-10-27 00:49:17（2往復目・ログファイルの時刻） | 読み取り専用での接続を試行（ro_mount.py） | `ERROR:aidrive_fuse.fuse_driver:Authentication failed: [401] Authentication failed. This appears to be an unauthorized runtime environment.` |
| 3 | 同上 | `aidrive ls /` | `✗ AIDrive API Error: [401] Authentication failed. This appears to be an unauthorized runtime environment.` |
| 4 | 2026-10-27 01:28:11 | `os.path.ismount('/mnt/aidrive')` | `False`（未接続のまま） |

- 読み取り用の窓口（ドライブ一覧・一時URL発行）は同じ時間帯に使えた
- 試行した接続プロセスは停止済み。ドライブへの書込みは行っていない

## 追記｜入江さんによる AIスライド再起動の後（8往復目・判定146(a)）
| # | システム時刻（UTC） | 操作 | 表示された文言（そのまま） |
|:-:|---|---|---|
| 5 | 2026-10-27 01:50:48 | `os.path.ismount('/mnt/aidrive')` | `False`（コンピューターの起動からの時間は 1:13＝コンピューター自体は再起動されていない） |
| 6 | 2026-10-27 01:51:34（40秒待機後） | `ls /mnt/aidrive/ui-diagnosis-director/handover` | `ls: cannot access '/mnt/aidrive/ui-diagnosis-director/handover': No such file or directory` |
| 7 | 同上 | `aidrive ls /` | `✗ AIDrive API Error: [401] Authentication failed. This appears to be an unauthorized runtime environment.` |
- 判定146(a) の確認は1回で終了。接続は戻らなかった。原因は未確認
