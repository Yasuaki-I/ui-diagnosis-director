#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""find_prefix.py｜同一prefix 一覧（サフィックス付き検出）（第16条 細則13）

2026-09-20 統括判定②（案A承認）に基づく常設スクリプト。
毎セッション手書きしていた prefix 検出コード（約10行）を置き換える。

背景:
    完全名での存在確認では、同名衝突時に自動付与される "(1)" 付きの
    ファイルを検出できない（9/17 に実際に見落とした）。
    受領確認は必ず「同一 prefix での一覧」で行う。

表示について（2026-09-24 統括判定 諮問③ 案A で是正）:
    従来は長いファイル名を n[-58:] で「左から」切り詰めていた。この形式は
    先頭が欠落するため、目視・二次集計で誤読を招く。2026-09-24 に実際に
    「claude」で始まらない行が grep から漏れ、35件を29件と誤集計した。
    以後は「右側を省略し末尾に … を付す」形式とする。

    なお総数は冒頭の「該当 N 件」を正とする。表示行から数え直さないこと
    （PLAIN J-7-6-5 規準2）。

検出ロジックについて:
    N-2（正規表現の検出漏れリスク）は 2026-09-24 に 5 パターンへの直接適用で
    検証し、検出漏れが存在しないことを確認してクローズした（統括判定 諮問③）。
    DOTALL は本件に無関係（単一行のファイル名に対する re.search のため）。

使い方:
    python3 find_prefix.py claude_chat_verdict_
    python3 find_prefix.py claude_chat_verdict_ --dir handover
"""
import os
import re
import sys
import hashlib

ROOT = '/mnt/aidrive/ui-diagnosis-director'


def _ellip(name, width):
    """右側を省略して末尾に … を付す（左は必ず残す）。

    2026-09-24 統括判定 諮問③ 案A。従来の n[-width:] は先頭が欠落し、
    目視・grep による二次集計で誤読を招いた。
    """
    if len(name) <= width:
        return name
    return name[:width - 1] + '…'


def md5_of(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()[:12]


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    prefix = argv[0]
    sub = 'handover'
    if '--dir' in argv:
        sub = argv[argv.index('--dir') + 1]
    d = os.path.join(ROOT, sub)

    names = sorted(n for n in os.listdir(d) if n.startswith(prefix))
    print('探索: %s  prefix="%s"' % (d, prefix))
    print('該当 %d 件\n' % len(names))
    print('%-58s %10s  %-12s %s'
          % ('ファイル', 'バイト', 'md5', 'サフィックス'))
    suffixed = []
    for n in names:
        p = os.path.join(d, n)
        if not os.path.isfile(p):
            continue
        m = re.search(r'\((\d+)\)(?=\.[^.]+$|$)', n)
        tag = '🚨 (%s) 付き' % m.group(1) if m else ''
        if m:
            suffixed.append(n)
        print('%-58s %10d  %-12s %s'
              % (_ellip(n, 58), os.path.getsize(p), md5_of(p), tag))
    print()
    if suffixed:
        print('🚨 サフィックス付き %d 件を検出 → 第18条（別名で並立させない）'
              'に従い名称衝突を解消すること' % len(suffixed))
    else:
        print('サフィックス付き 0 件 ⭕')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
