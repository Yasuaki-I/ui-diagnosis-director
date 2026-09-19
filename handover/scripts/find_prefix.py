#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""find_prefix.py｜同一prefix 一覧（サフィックス付き検出）（第16条 細則13）

2026-09-20 統括判定②（案A承認）に基づく常設スクリプト。
毎セッション手書きしていた prefix 検出コード（約10行）を置き換える。

背景:
    完全名での存在確認では、同名衝突時に自動付与される "(1)" 付きの
    ファイルを検出できない（9/17 に実際に見落とした）。
    受領確認は必ず「同一 prefix での一覧」で行う。

使い方:
    python3 find_prefix.py claude_chat_verdict_
    python3 find_prefix.py claude_chat_verdict_ --dir handover
"""
import os
import re
import sys
import hashlib

ROOT = '/mnt/aidrive/ui-diagnosis-director'


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
              % (n[-58:], os.path.getsize(p), md5_of(p), tag))
    print()
    if suffixed:
        print('🚨 サフィックス付き %d 件を検出 → 第18条（別名で並立させない）'
              'に従い名称衝突を解消すること' % len(suffixed))
    else:
        print('サフィックス付き 0 件 ⭕')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
