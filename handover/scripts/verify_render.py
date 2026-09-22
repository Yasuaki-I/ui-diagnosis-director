#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_render.py｜生成済 PPTX の全テキスト段落を実測し折返しを検出する

（2026-09-22 統括判定 第2便 判定② 案A で常設化｜第3便で verify_labels.py を統合）

■ なぜ静的解析ではなく生成物を測るのか

前身の `verify_labels.py` は builder のソースから「固定文字列 × 固定枠幅」の
`add_text` を抽出して検査していた。しかしこの方式は原理的に取りこぼす。

    add_text(slide, x, y, 50, f'{sv}/5', 14, bold=score_bold)   ← 検出不能
    add_text(slide, 56, y, 180, cat_disp, 14, bold=True)        ← 検出不能

f-string・変数・関数戻り値で組み立てられる文字列は、ソースを読むだけでは
最終的に何が描かれるか分からない。2026-09-22 の実機目視で発見された2件
（太字スコア値 55.5px>50px／項目名 186.7px>180px）はいずれもこの死角にあり、
`verify_labels.py` は「NG 0件」を返していた。

本スクリプトは **実際に生成された .pptx を開いて全テキスト段落を実測する**。
描画された文字列そのものを測るため、組み立て方によらず漏れない。

■ 実測の一次情報

Meiryo W53（本文）／ Meiryo W53 Bold（太字）の hmtx アドバンス幅。
builder 内の推定器 `_v17_text_w14` はこの実測から導出した近似表であり、
本スクリプトはフォントを直接引くため近似を経由しない。

■ 安全係数について（PLAIN J-7-6）

折返しの判定に安全係数（`_V17_W14_SAFE`）は適用しない。安全係数は
「枠を広く確保する」ための安全側バイアスであり、判定に用いると1行に収まる
文字列を2行と誤判定する（9/17 実機24点：安全係数あり 22/24 → 素の幅 24/24）。

■ 判定

    NG    … 折返しが起き、かつ必要高が枠高を超える（文字が潰れる）
    WRAP  … 折返すが枠高に収まる（設計上許容されうる。要目視）
    WARN  … 1行だが余裕 < 6px（些細な字形差で符号が反転しうる）
    OVER1 … 1文字が枠幅超過（折返せないため横にはみ出すのみ）
    OK    … 1行で余裕 >= 6px

■ フォント取得（2026-09-23 統括判定② 案A で追加｜N-4）

2026-09-23 の起動時、フォント取得が HTTP 403 Forbidden で停止し検査を
実行できなかった。原因は `urllib.request.urlretrieve` が User-Agent を
送らないため CDN 側で拒否されること（同一URLへ UA 付与で成功することを
実測して切り分けた）。環境差ではなく実装側の欠落である。

`_fetch()` で3経路を順に試行する。

    ① User-Agent を付与して取得（403 の直接原因への対処）
    ② 素の urlretrieve（UA 不要な環境への後方互換）
    ③ ドライブ内 scripts/fonts/ の退避コピー（CDN 停止時の最終手段）

全経路が失敗した場合は「本検査は未実施である（NG 0件を意味しない）」を
明示して停止する。例外で落ちるだけでは「検査を飛ばして実機目視へ進む」
誘因になり、2026-09-22 に3日連続で発生した「測定していないものを測定した
と報告する」構造へ再接近するためである（第16条 細則15 の趣旨の機械化）。

経路③の退避コピーの作成自体は本スクリプトの責務に含めない。フォントの
常設化はライセンス判断を伴うため別途諮問事項として保留されている
（2026-09-23 統括判定②）。既定では不在であり、経路①②のみで動作する。

■ 使い方

    python3 verify_render.py foo.pptx
    python3 verify_render.py --check '次の一手' 90 14 --bold   # 単発（旧 verify_labels.py 相当）

終了コード: NG が1件でもあれば 1／フォント取得不能なら 3（検査未実施）
"""
import os
import re
import sys
import zipfile

FONT_B = '/tmp/meiryo-bold.woff2'
FONT_R = '/tmp/meiryo.woff2'
FONT_URL_B = ('https://cdn1.genspark.ai/user-upload-image/fonts/'
              'ms-japanese/meiryo-bold.woff2')
FONT_URL_R = ('https://cdn1.genspark.ai/user-upload-image/fonts/'
              'ms-japanese/meiryo.woff2')
FONT_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              'fonts')
UNVERIFIED_MARK = '検査未実施'   # 機械可読な明示文字列
WARN_PX = 6.0
EMU = 9525.0          # 1px（実座標 1280x720 基準）


def _font(path, url):
    from fontTools.ttLib import TTFont
    if not os.path.isfile(path):
        _fetch(url, path)
    return TTFont(path)


def _fetch(url, path):
    """フォントを3経路で取得する。全失敗なら検査未実施を明示して停止する。"""
    import urllib.request
    errs = []

    # ① User-Agent 付与（403 の直接原因への対処）
    try:
        req = urllib.request.Request(
            url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        if len(data) < 100000:
            raise OSError('取得サイズが小さすぎる: %dB' % len(data))
        with open(path, 'wb') as f:
            f.write(data)
        return
    except Exception as e:
        errs.append('① User-Agent 付与: %s' % e)

    # ② 素の urlretrieve（後方互換）
    try:
        urllib.request.urlretrieve(url, path)
        return
    except Exception as e:
        errs.append('② urlretrieve: %s' % e)

    # ③ ドライブ内の退避コピー（作成は本スクリプトの責務外）
    cache = os.path.join(FONT_CACHE_DIR, os.path.basename(path))
    try:
        if os.path.isfile(cache):
            import shutil
            shutil.copyfile(cache, path)
            return
        errs.append('③ 退避コピー: 不在 (%s)' % cache)
    except Exception as e:
        errs.append('③ 退避コピー: %s' % e)

    sys.stderr.write(
        '\n🚨 フォントを取得できないため実測を実行できない。\n'
        + '\n'.join('  ' + e for e in errs)
        + '\n\n⚠️ 本検査は %s である（NG 0件を意味しない）。\n'
          '🚨 「実測で確認した」と報告してはならない（第16条 細則15）。\n'
          '⭐ 対処: 下記2点を手動配置してから再実行する。\n'
          '  %s\n  %s\n' % (UNVERIFIED_MARK, FONT_B, FONT_R))
    raise SystemExit(3)


def _mk(f):
    upm = f['head'].unitsPerEm
    hm = f['hmtx']
    cm = f.getBestCmap()

    def w(t, pt):
        s = 0.0
        for c in str(t):
            g = cm.get(ord(c))
            s += (hm[g][0] / upm) if g else 1.0
        return s * pt / 72.0 * 96.0
    return w


_WB = _WR = None


def _metrics():
    global _WB, _WR
    if _WB is None:
        _WB = _mk(_font(FONT_B, FONT_URL_B))
        _WR = _mk(_font(FONT_R, FONT_URL_R))
    return _WB, _WR


def text_px(text, pt, bold=True):
    """実描画幅（px｜実座標1280x720）を返す。"""
    wb, wr = _metrics()
    return (wb if bold else wr)(text, pt)


def scan(path):
    """.pptx の全スライドの全テキスト段落を実測して行を返す。"""
    wb, wr = _metrics()
    z = zipfile.ZipFile(path)
    out = []
    for name in sorted(n for n in z.namelist()
                       if re.match(r'ppt/slides/slide\d+\.xml$', n)):
        x = z.read(name).decode('utf8')
        page = name.split('/')[-1]
        for m in re.finditer(r'<p:sp>.*?</p:sp>', x, re.S):
            s = m.group(0)
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', s)
            ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', s)
            if not off or not ext:
                continue
            for p in re.findall(r'<a:p>.*?</a:p>', s, re.S):
                txt = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', p))
                if not txt.strip():
                    continue
                szl = re.findall(r'sz="(\d+)"', p)
                pt = int(szl[0]) / 100.0 if szl else 18.0
                bold = 'b="1"' in p
                lh = re.search(r'<a:lnSpc><a:spcPct val="(\d+)"/>', p)
                lhv = int(lh.group(1)) / 100000.0 if lh else 1.2
                wrap = re.search(r'wrap="(\w+)"', s)
                w = int(ext.group(1)) / EMU
                h = int(ext.group(2)) / EMU
                px = (wb if bold else wr)(txt, pt)
                if wrap and wrap.group(1) == 'none':
                    st, lines = 'OK', 1          # 折返さない設定
                elif len(txt.strip()) <= 1:
                    # 1文字は折返せない（横にはみ出すのみ）
                    st, lines = ('OVER1' if px > w else 'OK'), 1
                else:
                    lines = max(1, int(-(-px // w)) if w > 0 else 1)
                    need = lines * pt * lhv / 72.0 * 96.0
                    if lines >= 2 and need > h + 0.5:
                        st = 'NG'
                    elif lines >= 2:
                        st = 'WRAP'
                    elif (w - px) < WARN_PX:
                        st = 'WARN'
                    else:
                        st = 'OK'
                out.append((st, page, int(off.group(2)) / EMU, w, h,
                            pt, bold, px, lines, txt))
    return out


def _single(argv):
    """旧 verify_labels.py の --check 相当（単発測定）。"""
    text, w, pt = argv[0], float(argv[1]), float(argv[2])
    bold = '--bold' in argv or '-b' in argv
    px = text_px(text, pt, bold)
    room = w - px
    st = 'NG 🚨' if px > w else ('WARN ⚠️' if room < WARN_PX else 'OK ⭕')
    print('テキスト : %s' % text)
    print('条件     : 枠 %.1fpx / %.0fpt / %s' % (w, pt, '太字' if bold else '標準'))
    print('実描画幅 : %.1fpx（余裕 %+.1fpx）' % (px, room))
    print('判定     : %s' % st)
    return 1 if px > w else 0


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    if argv[0] == '--check':
        return _single(argv[1:])
    rows = scan(argv[0])
    cnt = {k: [r for r in rows if r[0] == k]
           for k in ('NG', 'WRAP', 'WARN', 'OVER1')}
    print('検査対象: %s' % argv[0])
    print('テキスト段落 %d 件\n' % len(rows))
    print('%-6s %-11s %6s %6s %6s %5s %4s %8s %4s  %s'
          % ('判定', 'slide', 'y', '枠w', '枠h', 'pt', 'B', '実描画',
             '行', 'テキスト'))
    order = {'NG': 0, 'WRAP': 1, 'WARN': 2, 'OVER1': 3, 'OK': 4}
    for r in sorted(rows, key=lambda r: order[r[0]]):
        if r[0] == 'OK':
            continue
        st, pg, y, w, h, pt, b, px, ln, txt = r
        mark = {'NG': '🚨 NG', 'WRAP': '⚠️ WRAP', 'WARN': '⚠️ WARN',
                'OVER1': '⚠️ OVER1'}[st]
        print('%-6s %-11s %6.0f %6.0f %6.0f %5.0f %4s %8.1f %4d  %s'
              % (mark, pg, y, w, h, pt, 'B' if b else '-', px, ln, txt[:34]))
    ok = len(rows) - sum(len(v) for v in cnt.values())
    print()
    print('NG %d 件 ／ WRAP %d 件 ／ WARN %d 件 ／ OVER1 %d 件 ／ OK %d 件 → %s'
          % (len(cnt['NG']), len(cnt['WRAP']), len(cnt['WARN']),
             len(cnt['OVER1']), ok,
             'NG 0件 ⭕' if not cnt['NG'] else '🚨 要是正'))
    print('※ OVER1 = 1文字が枠幅超過（折返し不能・横にはみ出すのみ）')
    print('※ 本検査は PowerPoint 実機の画素確認を代替しない'
          '（LibreOffice も同様｜PLAIN J-7-6-2 B）')
    return 1 if cnt['NG'] else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
