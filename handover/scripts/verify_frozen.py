#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_frozen.py｜改変禁止領域の AST＋md5 逐語照合（第16条 細則12）

2026-09-20 統括判定②（案A承認）に基づく常設スクリプト。
毎セッション手書きしていた逐語照合コード（約35行）を置き換える。

使い方:
    python3 verify_frozen.py 旧builder.py 新builder.py
    python3 verify_frozen.py 旧builder.py 新builder.py --allow _V17_W14 _V17_W14_SAFE

判定:
    ・関数／クラスの AST ダンプ md5 を全件比較し、差分・追加・削除を列挙する
    ・モジュール直下の代入も値の AST md5 で比較する
    ・--allow で明示した識別子の差分のみ「承認済の変更」として NG から除く
終了コード: 承認外の差分が1件でもあれば 1
"""
import ast
import hashlib
import sys


def sig_defs(tree):
    out = {}
    for n in ast.walk(tree):
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            out[n.name] = hashlib.md5(ast.dump(n).encode()).hexdigest()[:12]
    return out


def sig_assigns(tree):
    out = {}
    for n in tree.body:
        if isinstance(n, ast.Assign):
            for t in n.targets:
                if isinstance(t, ast.Name):
                    out[t.id] = hashlib.md5(
                        ast.dump(n.value).encode()).hexdigest()[:12]
    return out


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    before, after = argv[0], argv[1]
    allow = set()
    if '--allow' in argv:
        allow = set(argv[argv.index('--allow') + 1:])

    tb = ast.parse(open(before, encoding='utf-8').read())
    ta = ast.parse(open(after, encoding='utf-8').read())

    db, da = sig_defs(tb), sig_defs(ta)
    ab, aa = sig_assigns(tb), sig_assigns(ta)

    changed = sorted(k for k in db if k in da and db[k] != da[k])
    added = sorted(k for k in da if k not in db)
    removed = sorted(k for k in db if k not in da)
    var_changed = sorted(k for k in ab if ab.get(k) != aa.get(k))

    print('関数・クラス総数   : before %d / after %d' % (len(db), len(da)))
    print('AST差分のある定義  : %s' % (changed or 'なし ⭕'))
    print('追加された定義     : %s' % (added or 'なし ⭕'))
    print('削除された定義     : %s' % (removed or 'なし ⭕'))
    print('モジュール変数差分 : %s' % (var_changed or 'なし ⭕'))
    if allow:
        print('承認済の変更（--allow）: %s' % sorted(allow))

    ng = [k for k in changed + added + removed + var_changed if k not in allow]
    print()
    print('承認外の差分 %d 件 → %s'
          % (len(ng), 'NG 0件 ⭕' if not ng else '🚨 %s' % ng))
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
