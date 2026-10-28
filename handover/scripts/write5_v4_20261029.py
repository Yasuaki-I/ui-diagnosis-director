#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""write5_v4｜5段階書込（必達ルール9）を gsk aidrive 方式に置き換えた版（2026-10-29・判定155の4点を反映・統括の差分確認待ち）

背景:
    2026-10-29 Genspark サポート回答：AIスライドの新しい実行基盤では /mnt/aidrive の自動マウントと
    旧 aidrive コマンドは提供しない。読み込み・保存は gsk aidrive download / upload で行う。
    現行 write5.py・write5_v3 は /mnt/aidrive 前提のため、新基盤では使えない（v3 は⓪で必ず止まる）。

5段階（各段の照合先はすべて AI ドライブ上の実体。コンピューター内の場所は照合に使わない）:
    ⓪ gsk の版を表示し、1.10 未満なら書かずに停止（判定155-3）
    ① 既存の確認：宛先フォルダを gsk aidrive ls で一覧し、同名・同名(n) があれば書かずに停止
    ② 書込：gsk aidrive upload --on-conflict error（同名があればサーバー側でも失敗させる。上書き手段は持たない）
       応答の実際の保存先（data.path）が宛先と完全一致しなければ、③へ進まず停止（判定155-1）
    ③ md5 照合：ドライブから別の一時フォルダへ download し、src と照合
    ④ 待機：既定40秒。40秒未満は受け付けない（判定137-3・142-4）
    ⑤ 再取得：もう一度 download し、src と照合。ドライブ上の実際のパスを表示
    ⑥ 宛先フォルダを一覧し直し、宛先の名前がちょうど1件・同名(n)なしを確認して表示（判定155-2）

使い方:
    python3 write5_v4_gsk_20261029_candidate.py <src> <ドライブ上の宛先パス> [--wait 40] [--dry-run]
      宛先は /ui-diagnosis-director/ からの相対パス（例 handover/foo.md）か、/ で始まる絶対パス
      --dry-run：①まで行い、②以降は行わない（書込なし）

終了コード: 0＝全通過／1＝照合不一致・取得失敗・保存先の不一致（書込済みの可能性あり）
           2＝引数誤り・gsk の版が古い（書いていない）／3＝既存ありで停止（書いていない）
           5＝gsk の呼び出し失敗（①で失敗＝書いていない／②以降で失敗＝書込済みの可能性あり）
"""
import hashlib, json, os, re, subprocess, sys, tempfile, time

ROOT = '/ui-diagnosis-director'
MIN_GSK = (1, 10)


def gsk_version():
    """⓪ gsk の版を (表示文字列, タプル) で返す。読めなければ (文字列, None)。"""
    try:
        r = subprocess.run(['gsk', '--version'], capture_output=True, text=True, timeout=60)
        v = (r.stdout or r.stderr).strip().splitlines()[0].strip()
    except Exception as e:
        return '取得失敗: %s' % e, None
    m = re.match(r'(\d+)\.(\d+)', v)
    return v, ((int(m.group(1)), int(m.group(2))) if m else None)


def md5_of(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for c in iter(lambda: f.read(1 << 20), b''):
            h.update(c)
    return h.hexdigest()


def gsk(args, timeout=180):
    r = subprocess.run(['gsk', 'aidrive'] + args, capture_output=True, text=True, timeout=timeout)
    try:
        return json.loads(r.stdout), r
    except Exception:
        return None, r


def listing(folder):
    d, r = gsk(['ls', '--path', folder])
    if not d or d.get('status') != 'ok':
        return None
    t = d['data']['result']
    return [l.split(' | ')[0] for l in t.split('\n')
            if ' | ' in l and not l.startswith('Name') and not l.startswith('---')]


def fetch(dpath, tag):
    out = os.path.join(tempfile.mkdtemp(prefix='w5v4_%s_' % tag), os.path.basename(dpath))
    d, r = gsk(['download', '--path', dpath, '--local_file', out])
    if not d or d.get('status') != 'ok' or not os.path.isfile(out):
        return None
    return out


def main(argv):
    if len(argv) < 2:
        print(__doc__); return 2
    src, dst = argv[0], argv[1]
    dpath = dst if dst.startswith('/') else ROOT + '/' + dst
    wait = 40
    if '--wait' in argv:
        i = argv.index('--wait')
        if i + 1 >= len(argv) or not re.fullmatch(r'\d+', argv[i + 1]):
            print('🚨 --wait の後に秒数（整数）が必要（判定155-4）'); return 2
        wait = int(argv[i + 1])
        if wait < 40:
            print('🚨 --wait は40秒未満にできない（判定137-3・142-4）'); return 2
    dry = '--dry-run' in argv
    if '--replace' in argv:
        print('🚨 本案に上書き機能はない（判定142-3）'); return 2
    # ⓪ gsk の版（判定155-3）
    gv, gt = gsk_version()
    print('⓪ gsk の版: %s' % gv)
    if gt is None or gt < MIN_GSK:
        print('⓪ 🛑 gsk が 1.10 未満、または版を読めない → 書かずに停止'); return 2
    if not os.path.isfile(src):
        print('🚨 src が存在しない: %s' % src); return 1
    sm, ss = md5_of(src), os.path.getsize(src)
    print('src: %s / %dB / %s' % (src, ss, sm[:12]))
    print('dst（AIドライブ）: %s' % dpath)

    # ① 既存の確認
    folder, name = os.path.dirname(dpath), os.path.basename(dpath)
    names = listing(folder)
    if names is None:
        print('① 🛑 宛先フォルダの一覧を取得できない → 書かずに停止: %s' % folder); return 5
    stem, ext = os.path.splitext(name)
    hit = [n for n in names if n == name or re.fullmatch(re.escape(stem) + r'\(\d+\)' + re.escape(ext), n)]
    if hit:
        print('① 🛑 既存あり → 書かずに停止: %s' % hit); return 3
    print('① 既存なし（宛先フォルダ %d 件を確認）' % len(names))
    if dry:
        print('--dry-run：②以降は行わない（書込なし）'); return 0

    # ② 書込
    d, r = gsk(['upload', '--local_file', src, '--upload_path', dpath, '--on-conflict', 'error'], timeout=600)
    if not d or d.get('status') != 'ok':
        print('② 🛑 upload 失敗（書込済みかは不明。消失・未書込と断定せず報告）: %s'
              % (r.stdout or r.stderr)[-400:]); return 5
    print('② 書込完了（応答: %s）' % json.dumps(d.get('data', {}), ensure_ascii=False)[:300])
    landed = (d.get('data') or {}).get('path')
    print('② 実際の保存先（応答）: %s' % landed)
    if landed != dpath:
        print('② 🛑 実際の保存先が宛先と一致しない、または応答に保存先がない → ③へ進まず停止'
              '（書込済みの可能性あり。別名で保存された可能性も含め報告）（判定155-1）'); return 1
    print('② 保存先 一致 ⭕')

    # ③ md5 照合（ドライブから取得して照合）
    f3 = fetch(dpath, 's3')
    if not f3:
        print('③ 🚨 ドライブから取得できない（書込済みの可能性あり）'); return 1
    m3, s3 = md5_of(f3), os.path.getsize(f3)
    ok3 = (m3 == sm and s3 == ss)
    print('③ 照合した場所（AIドライブ）: %s' % dpath)
    print('③ md5 照合: src %s / drive %s（%dB）→ %s' % (sm[:12], m3[:12], s3, '一致 ⭕' if ok3 else '🚨 不一致'))
    if not ok3:
        return 1

    # ④ 待機
    print('④ %d 秒待機…' % wait); time.sleep(wait)

    # ⑤ 再取得で確認
    f5 = fetch(dpath, 's5')
    if not f5:
        print('⑤ 🚨 待機後に取得できない（消失と断定せず報告）'); return 1
    m5, s5 = md5_of(f5), os.path.getsize(f5)
    ok5 = (m5 == sm and s5 == ss)
    print('⑤ 照合した場所（AIドライブ）: %s' % dpath)
    print('⑤ 再取得: %dB / %s → %s' % (s5, m5[:12], '一致 ⭕' if ok5 else '🚨 不一致'))
    if not ok5:
        print('\n5段階書込 🚨 失敗'); return 1

    # ⑥ 宛先フォルダの一覧を取り直す（判定155-2）
    names6 = listing(folder)
    if names6 is None:
        print('⑥ 🚨 宛先フォルダの一覧を取り直せない（書込済み。報告する）'); return 1
    same = [n for n in names6 if n == name]
    dup = [n for n in names6 if re.fullmatch(re.escape(stem) + r'\(\d+\)' + re.escape(ext), n)]
    print('⑥ 宛先フォルダ再一覧: %d 件／宛先の名前 %d 件／同名(n) %d 件 %s'
          % (len(names6), len(same), len(dup), dup if dup else ''))
    ok6 = (len(same) == 1 and not dup)
    print('\n5段階書込 %s' % ('全通過 ⭕' if ok6 else '🚨 失敗（⑥）'))
    return 0 if ok6 else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
