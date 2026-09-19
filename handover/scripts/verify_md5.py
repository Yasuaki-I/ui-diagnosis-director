#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_md5.py｜成果物一覧の実測B・md5 照合（必達ルール5・9）

2026-09-20 統括判定②（案A承認）に基づく常設スクリプト。
毎セッション手書きしていた md5 照合コード（約40行）を置き換える。

使い方:
    python3 verify_md5.py manifest.tsv
    python3 verify_md5.py --check path/to/file.py d466ef702fe7 271150

manifest.tsv の書式（タブ区切り・# 始まりはコメント）:
    相対パス<TAB>期待md5(先頭12桁可)<TAB>期待バイト数

判定:
    ⭕ OK  = md5・バイト数ともに一致
    🚨 NG  = いずれかが不一致、またはファイルが存在しない
終了コード: NG が1件でもあれば 1
"""
import hashlib
import os
import sys

ROOT = '/mnt/aidrive/ui-diagnosis-director'


def md5_of(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def resolve(path):
    """相対パスは AI ドライブのプロジェクト直下として解決する。"""
    return path if os.path.isabs(path) else os.path.join(ROOT, path)


def check_one(path, exp_md5, exp_size):
    """1件を照合し (状態, 実測md5, 実測B, メッセージ) を返す。"""
    full = resolve(path)
    if not os.path.isfile(full):
        return 'MISSING', '-', '-', 'ファイルが存在しない'
    actual = md5_of(full)
    size = os.path.getsize(full)
    ng = []
    if exp_md5 and not actual.startswith(exp_md5.strip().lower()):
        ng.append('md5 不一致（期待 %s）' % exp_md5)
    if exp_size is not None and size != int(exp_size):
        ng.append('バイト数 不一致（期待 %s）' % exp_size)
    return ('NG' if ng else 'OK'), actual[:12], size, ' / '.join(ng)


def load_manifest(path):
    rows = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.split('\t')
            p = parts[0]
            m = parts[1] if len(parts) > 1 and parts[1] not in ('-', '') else None
            s = parts[2] if len(parts) > 2 and parts[2] not in ('-', '') else None
            rows.append((p, m, s))
    return rows


def main(argv):
    if len(argv) >= 2 and argv[0] == '--check':
        rows = [(argv[1],
                 argv[2] if len(argv) > 2 else None,
                 argv[3] if len(argv) > 3 else None)]
    elif argv:
        rows = load_manifest(argv[0])
    else:
        print(__doc__)
        return 2

    ng = 0
    print('%-4s %-52s %-14s %10s  %s'
          % ('判定', 'ファイル', '実測md5', '実測B', '備考'))
    for path, exp_md5, exp_size in rows:
        st, actual, size, msg = check_one(path, exp_md5, exp_size)
        mark = '⭕' if st == 'OK' else '🚨'
        if st != 'OK':
            ng += 1
        print('%-4s %-52s %-14s %10s  %s'
              % (mark, path[-52:], actual, size, msg))
    print()
    print('照合 %d 件 ／ NG %d 件 → %s'
          % (len(rows), ng, 'NG 0件 ⭕' if ng == 0 else '🚨 要確認'))
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
