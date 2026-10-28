#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""write5.py｜5段階書込（必達ルール9）｜改修案（2026-10-16・判定17・25）＋接続確認 v3（2026-10-28・判定139｜接続環境での1件確認までは試用）

5段階:
    ① 既存の確認（既定：既存があれば書かずに停止）
    ② 書込
    ③ md5 照合（src ↔ dst）
    ④ 40秒待機
    ⑤ 再読込で内容一致を確認

使い方:
    python3 write5.py /tmp/out/foo.md handover/foo.md
    python3 write5.py /tmp/out/foo.md handover/foo.md --wait 40
    python3 write5.py /tmp/out/foo.md handover/foo.md --replace <旧md5先頭12桁> <旧バイト数>
      … 上書きは --replace で旧値を明示した場合だけ。旧値が実測と一致したときに限り置き換える（事前照合の組込み）

終了コード: 0＝全通過／1＝照合不一致／2＝引数誤り／3＝既存ありで停止（書いていない）
           4＝AIドライブ未接続で停止（⓪②で止まった場合は書いていない／⑤で止まった場合は書込済みの可能性あり）｜2026-10-28

⓪ 接続確認（2026-10-28 追加案）:
    /mnt/aidrive が実際にドライブとして接続されていることを、書込の前に確かめる。
    接続されていないまま書くと、コンピューター内にだけフォルダとファイルができ、
    ③⑤の照合も同じ場所を見るため「全通過」に見えてしまう（10/28 に確認した危険）。
    接続の確認は ⓪書込前・②書込直前・⑤再読込の直前 の3回行う（判定139-3）。
    照合した場所の実際のパスを表示する（判定139-3）。
"""
import hashlib
import os
import shutil
import sys
import time

ROOT = '/mnt/aidrive/ui-diagnosis-director'
MOUNT = '/mnt/aidrive'


def drive_connected():
    """⓪ /mnt/aidrive が接続済み（マウント）で、プロジェクトの直下が見えるかを返す。"""
    if not os.path.ismount(MOUNT):
        return False, '%s が接続（マウント）されていない' % MOUNT
    if not os.path.isdir(os.path.join(ROOT, 'handover')):
        return False, '%s/handover が見えない' % ROOT
    return True, ''


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

    # ⓪ 接続確認（2026-10-28 追加案）
    ok0, why = drive_connected()
    if not ok0:
        print('⓪ 🛑 AIドライブ未接続 → 書かずに停止: %s' % why)
        print('   待って再実行してもよいが、直らなければ処理を中断して報告する（引継書 §0-10）')
        return 4
    print('⓪ AIドライブ接続 確認 ⭕')
    if not os.path.isabs(dst):
        dst = os.path.join(ROOT, dst)
    wait = 40
    if '--wait' in argv:
        wait = int(argv[argv.index('--wait') + 1])
    replace = None
    if '--replace' in argv:
        i = argv.index('--replace')
        try:
            replace = (argv[i + 1].lower(), int(argv[i + 2]))
        except (IndexError, ValueError):
            print('🚨 --replace には <旧md5先頭12桁> <旧バイト数> の2つが必要')
            return 2
        if len(replace[0]) < 12:
            print('🚨 --replace の md5 は先頭12桁以上を指定する')
            return 2

    if not os.path.isfile(src):
        print('🚨 src が存在しない: %s' % src)
        return 1
    src_md5 = md5_of(src)
    src_size = os.path.getsize(src)
    print('src: %s / %dB / %s' % (src, src_size, src_md5[:12]))
    print('dst: %s' % dst)
    print()

    # ① 既存の確認
    if os.path.exists(dst):
        prev_md5 = md5_of(dst)
        prev_size = os.path.getsize(dst)
        if replace is None:
            print('① 🛑 既存あり → 書かずに停止: %s（%dB / %s）' % (dst, prev_size, prev_md5[:12]))
            print('   上書きが意図どおりなら --replace %s %d を付けて再実行する。'
                  '意図しないファイルなら触れずに判定3-5（停止・保全・報告）' % (prev_md5[:12], prev_size))
            return 3
        if not (prev_md5.startswith(replace[0]) and prev_size == replace[1]):
            print('① 🛑 事前照合 不一致 → 書かずに停止: 実測 %dB / %s ≠ 指定 %dB / %s'
                  % (prev_size, prev_md5[:12], replace[1], replace[0][:12]))
            print('   判定3-5（停止・保全・報告）')
            return 3
        os.remove(dst)
        print('① 事前照合 一致（%dB / %s）→ 既存を削除して置き換える' % (prev_size, prev_md5[:12]))
    else:
        if replace is not None:
            print('① 🛑 --replace 指定だが既存がない → 書かずに停止（宛先の誤りの可能性）')
            return 3
        print('① 既存なし')

    # ② 書込（直前にもう一度接続を確かめる）
    if not drive_connected()[0]:
        print('② 🛑 書込直前にAIドライブ未接続 → 書かずに停止')
        return 4
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)
    print('② 書込完了')

    # ③ md5 照合
    dst_md5 = md5_of(dst)
    ok3 = (dst_md5 == src_md5) and (os.path.getsize(dst) == src_size)
    print('③ 照合した場所の実際のパス: %s' % os.path.realpath(dst))
    print('③ md5 照合: src %s / dst %s → %s'
          % (src_md5[:12], dst_md5[:12], '一致 ⭕' if ok3 else '🚨 不一致'))
    if not ok3:
        return 1

    # ④ 待機
    print('④ %d 秒待機…' % wait)
    time.sleep(wait)

    # ⑤ 再読込で確認（直前にもう一度接続を確かめる｜判定139-3）
    if not drive_connected()[0]:
        print('⑤ 🛑 再読込の直前にAIドライブ未接続 → 照合できないため停止（書込済みの可能性あり。'
              '消失と断定せず、接続が戻ってから確認・報告する）')
        return 4
    print('⑤ 照合した場所の実際のパス: %s' % os.path.realpath(dst))
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
