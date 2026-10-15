#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""write5.py｜5段階書込（必達ルール9）

2026-09-20 統括判定②（案A承認）に基づく常設スクリプト。
毎セッション手書きしていた5段階書込コード（約20行）を置き換える。

5段階:
    ① 既存を削除
    ② 書込
    ③ md5 照合（src ↔ dst）
    ④ 40秒待機
    ⑤ ドライブAPI相当の再読込で内容一致を確認

使い方:
    python3 write5.py /tmp/out/foo.md handover/foo.md
    python3 write5.py /tmp/out/foo.md handover/foo.md --wait 40

相対の宛先は AI ドライブのプロジェクト直下として解決する。
終了コード: いずれかの段階で不一致なら 1
"""
import hashlib
import os
import shutil
import sys
import time

ROOT = '/mnt/aidrive/ui-diagnosis-director'


def md5_of(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    src = argv[0]
    dst = argv[1]
    if not os.path.isabs(dst):
        dst = os.path.join(ROOT, dst)
    wait = 40
    if '--wait' in argv:
        wait = int(argv[argv.index('--wait') + 1])

    if not os.path.isfile(src):
        print('🚨 src が存在しない: %s' % src)
        return 1
    src_md5 = md5_of(src)
    src_size = os.path.getsize(src)
    print('src: %s / %dB / %s' % (src, src_size, src_md5[:12]))
    print('dst: %s' % dst)
    print()

    # ① 削除
    if os.path.exists(dst):
        prev = md5_of(dst)
        os.remove(dst)
        print('① 既存を削除（旧 md5 %s）' % prev[:12])
    else:
        print('① 既存なし')

    # ② 書込
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)
    print('② 書込完了')

    # ③ md5 照合
    dst_md5 = md5_of(dst)
    ok3 = (dst_md5 == src_md5) and (os.path.getsize(dst) == src_size)
    print('③ md5 照合: src %s / dst %s → %s'
          % (src_md5[:12], dst_md5[:12], '一致 ⭕' if ok3 else '🚨 不一致'))
    if not ok3:
        return 1

    # ④ 待機
    print('④ %d 秒待機…' % wait)
    time.sleep(wait)

    # ⑤ 再読込で確認
    if not os.path.isfile(dst):
        print('⑤ 🚨 待機後に dst が消失')
        return 1
    final_md5 = md5_of(dst)
    final_size = os.path.getsize(dst)
    ok5 = (final_md5 == src_md5) and (final_size == src_size)
    print('⑤ 再読込: %dB / %s → %s'
          % (final_size, final_md5[:12], '一致 ⭕' if ok5 else '🚨 不一致'))
    print()
    print('5段階書込 %s' % ('全通過 ⭕' if ok5 else '🚨 失敗'))
    print('※ FUSE層が空に見える場合も消失と断定しないこと'
          '（引継書 §7-5｜ドライブAPI で実在を確認する）')
    return 0 if ok5 else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
