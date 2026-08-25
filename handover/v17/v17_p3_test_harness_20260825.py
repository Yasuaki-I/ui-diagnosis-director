#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v17 P3｜テストハーネス（funnel / cycle / contrast / timeline / network）

検証範囲
  T1  35組合せ（5パターン × 7テーマ）が例外なく描画されること
  T2  要素数の境界値（min-1 / min / max / max+1）
  T3  警告オーバーライド（score < 40 → warning 色）
  T4  決定論性（同一入力 → 同一 XML）
  T5  timeline の requires_axes=True（軸欠落 → category フォールバック＋notes）
  T6  network のフォールバック3起動条件（統括指示｜8/12 議題4）
  T7  描画範囲（V17_AREA の 60〜660px を超えないこと）
  T8  フォント制約（メイリオ／14pt 下限）
  T9  cycle の均等配分（段階減衰にしないこと＝色の巡回）
  T10 contrast の2極化（中間色を使わないこと）
  T11 draw_pattern ディスパッチ（11パターン＋未実装キー退避）
  T12 後方互換（P1・P2 の描画が変化しないこと）

  作成 : 2026-08-25（火）18:00連結③｜AIスライド（実装領域）
  対象 : builder_v17_20.py（v17.2.0）
"""

import os
import sys
import copy

BUILDER_CANDIDATES = [
    'builder_v17_20.py',
    '03_pptx_builder.py',
    '/mnt/data/03_pptx_builder.py',
    '03_pptx_builder_v17_20_20260825.py',
]

PASS, FAIL = [], []


def ok(tid, msg):
    PASS.append((tid, msg))
    print('  [PASS] %-5s %s' % (tid, msg))


def ng(tid, msg):
    FAIL.append((tid, msg))
    print('  [FAIL] %-5s %s' % (tid, msg))


def load():
    for p in BUILDER_CANDIDATES:
        if os.path.exists(p):
            ns = {'__name__': 'builder_v17'}
            exec(compile(open(p, encoding='utf-8').read(), p, 'exec'), ns)
            print('ビルダー: %s  __version__=%r' % (p, ns.get('__version__')))
            return ns
    print('[NG] ビルダーが見つかりません')
    sys.exit(1)


NS = load()
P = NS['create_presentation']
BLANK = NS['_blank_slide']
PAL = NS['get_theme_palette']
DRAW = NS['draw_pattern']
SPEC = NS['DIAGRAM_PATTERN_SPEC']
THEMES = list(NS['DIGITAL_AGENCY_PALETTE'].keys())
AREA = NS['V17_AREA']
WARN = NS['V17_WARNING_SCORE']


def mk(pattern, n=None, score_base=70, axis=True, ids=True, edges=True):
    """パターン別のテストデータを生成する（決定論的）"""
    def items(k, prefix='要素'):
        return [{'label': '%s%d' % (prefix, i + 1),
                 'score': score_base - i * 5,
                 'description': '説明テキスト%d' % (i + 1)} for i in range(k)]

    if pattern == 'funnel':
        k = n if n is not None else 4
        return {'title': 'ファネル検証', 'stages': items(k, '段階')}
    if pattern == 'cycle':
        k = n if n is not None else 4
        # ⚠️ cycle の正式キーは 'phases'（'stages' は funnel 側）
        return {'title': 'サイクル検証', 'cycle_name': 'PDCA',
                'phases': items(k, '工程')}
    if pattern == 'contrast':
        k = n if n is not None else 2
        return {'title': '対比検証',
                'sides': [{'label': '側%d' % (i + 1), 'score': score_base - i * 30,
                           'items': ['項目A', '項目B']} for i in range(k)]}
    if pattern == 'timeline':
        k = n if n is not None else 4
        ms = items(k, 'マイルストーン')
        for i, m in enumerate(ms):
            m['axis'] = ('2026/%d' % (8 + i)) if axis else ''
        return {'title': '時間軸検証',
                'axis_label': '実施時期' if axis else '',
                'milestones': ms}
    if pattern == 'network':
        k = n if n is not None else 4
        nodes = []
        for i in range(k):
            nd = {'label': 'ノード%d' % (i + 1), 'score': score_base - i * 5,
                  'depth': 0 if i == 0 else (1 if i < 3 else 2)}
            if ids:
                nd['id'] = 'n%d' % (i + 1)
            nodes.append(nd)
        eg = []
        if edges and k >= 2:
            eg = [{'from': 'n1', 'to': 'n%d' % (i + 1)} for i in range(1, min(k, 3))]
        return {'title': 'ネットワーク検証', 'nodes': nodes, 'edges': eg}
    raise ValueError(pattern)


P3 = ['funnel', 'cycle', 'contrast', 'timeline', 'network']
ALL11 = ['category', 'breakdown', 'comparison', 'pyramid', 'sequence',
         'framework', 'funnel', 'cycle', 'contrast', 'timeline', 'network']


def shapes_xml(slide):
    from lxml import etree
    return etree.tostring(slide.shapes._spTree)


# ---------------------------------------------------------------- T1
print('\n=== T1｜35組合せ（5パターン × 7テーマ）===')
print('テーマ数: %d  %s' % (len(THEMES), THEMES))
for pat in P3:
    for th in THEMES:
        prs = P()
        sl = BLANK(prs)
        try:
            rep = DRAW(sl, pat, PAL(th), mk(pat))
            if rep.get('fallback_from'):
                ng('T1', '%s/%s 想定外のフォールバック: %s' % (pat, th, rep.get('notes')))
            elif len(sl.shapes) < 2:
                ng('T1', '%s/%s 図形数不足 %d' % (pat, th, len(sl.shapes)))
            else:
                ok('T1', '%-9s / %-9s shapes=%d' % (pat, th, len(sl.shapes)))
        except Exception as e:
            ng('T1', '%s/%s 例外 %s: %s' % (pat, th, type(e).__name__, e))

# ---------------------------------------------------------------- T2
print('\n=== T2｜要素数の境界値 ===')
for pat in P3:
    mn, mx = SPEC[pat]['min_elements'], SPEC[pat]['max_elements']
    for n, expect_fb in ((mn - 1, True), (mn, False), (mx, False), (mx + 1, True)):
        if n < 0:
            continue
        prs = P(); sl = BLANK(prs)
        try:
            rep = DRAW(sl, pat, PAL('Blue'), mk(pat, n=n))
            got_fb = bool(rep.get('fallback_from'))
            if got_fb == expect_fb:
                ok('T2', '%-9s n=%d fallback=%s（期待どおり）' % (pat, n, got_fb))
            else:
                ng('T2', '%-9s n=%d fallback=%s 期待=%s' % (pat, n, got_fb, expect_fb))
        except Exception as e:
            ng('T2', '%s n=%d 例外 %s' % (pat, n, e))

# ---------------------------------------------------------------- T3
print('\n=== T3｜警告オーバーライド（score < %d）===' % WARN)
for pat in P3:
    for th in ('Blue', 'SolidGray'):
        pal = PAL(th)
        prs = P(); sl = BLANK(prs)
        DRAW(sl, pat, pal, mk(pat, score_base=30))
        xml = shapes_xml(sl).decode('utf-8', 'ignore')
        wa = pal['warning'].lstrip('#').upper()
        if wa in xml.upper():
            ok('T3', '%-9s/%-9s warning色 %s を検出' % (pat, th, wa))
        else:
            ng('T3', '%-9s/%-9s warning色 %s 未検出' % (pat, th, wa))

# ---------------------------------------------------------------- T4
print('\n=== T4｜決定論性（同一入力 → 同一XML）===')
for pat in P3:
    d = mk(pat)
    outs = []
    for _ in range(2):
        prs = P(); sl = BLANK(prs)
        DRAW(sl, pat, PAL('Green'), copy.deepcopy(d))
        outs.append(shapes_xml(sl))
    if outs[0] == outs[1]:
        ok('T4', '%-9s 2回の出力XMLが同一' % pat)
    else:
        ng('T4', '%-9s 出力XMLが不一致' % pat)

# ---------------------------------------------------------------- T5
print('\n=== T5｜timeline requires_axes=True（統括厳守事項）===')
# 5-1 axis_label 欠落
prs = P(); sl = BLANK(prs)
rep = DRAW(sl, 'timeline', PAL('Blue'), mk('timeline', axis=False))
if rep.get('fallback_from') == 'timeline' and rep.get('notes'):
    ok('T5', 'axis 欠落 → category フォールバック＋notes 記録')
else:
    ng('T5', 'axis 欠落時の挙動が不正: %r' % rep)
# 5-2 axis_label はあるが milestone.axis が空
d = mk('timeline')
d['milestones'][1]['axis'] = ''
prs = P(); sl = BLANK(prs)
rep = DRAW(sl, 'timeline', PAL('Blue'), d)
if rep.get('fallback_from') == 'timeline':
    ok('T5', 'milestone.axis 欠落 → category フォールバック')
else:
    ng('T5', 'milestone.axis 欠落時の挙動が不正: %r' % rep)
# 5-3 requires_axes フラグ自体
if SPEC['timeline']['requires_axes'] is True:
    ok('T5', "SPEC['timeline']['requires_axes'] is True")
else:
    ng('T5', 'requires_axes が True でない')
# 5-4 正常時は軸ラベルが描画されること
prs = P(); sl = BLANK(prs)
DRAW(sl, 'timeline', PAL('Blue'), mk('timeline'))
# ⚠️ 日本語は XML 内で数値文字参照になるため、テキスト走査で検査する
texts = []
for sh in sl.shapes:
    if sh.has_text_frame:
        texts.append(sh.text_frame.text)
joined = '\n'.join(texts)
if '実施時期' in joined and '2026/8' in joined:
    ok('T5', '軸ラベル・時間軸値の両方が描画されている')
else:
    ng('T5', '軸ラベルまたは時間軸値が描画されていない: %r' % texts[:4])

# ---------------------------------------------------------------- T6
print('\n=== T6｜network フォールバック3起動条件 ===')
cases = [
    ('ノード数超過', mk('network', n=SPEC['network']['max_elements'] + 1)),
    ('エッジ参照不整合', None),
    ('階層構造データ不正', None),
]
d2 = mk('network')
d2['edges'] = [{'from': 'n1', 'to': 'n99'}]
cases[1] = ('エッジ参照不整合', d2)
d3 = mk('network')
d3['nodes'][1]['id'] = d3['nodes'][0]['id']   # ID重複
cases[2] = ('階層構造データ不正（ID重複）', d3)
for name, d in cases:
    prs = P(); sl = BLANK(prs)
    rep = DRAW(sl, 'network', PAL('Blue'), d)
    if rep.get('fallback_from') == 'network' and rep.get('notes'):
        ok('T6', '%s → category フォールバック＋notes' % name)
    else:
        ng('T6', '%s のフォールバックが機能しない: %r' % (name, rep))
# 正常時はエッジが描かれること
prs = P(); sl = BLANK(prs)
rep = DRAW(sl, 'network', PAL('Blue'), mk('network'))
if not rep.get('fallback_from') and any('エッジ' in s for s in (rep.get('notes') or [])):
    ok('T6', '正常時はエッジ描画が notes に記録される')
else:
    ng('T6', '正常時のエッジ描画記録がない: %r' % rep)

# ---------------------------------------------------------------- T7
print('\n=== T7｜描画範囲（%d〜%dpx）===' % (AREA['title_top'], AREA['body_bottom']))
EMU = 9525  # 1px
for pat in P3:
    prs = P(); sl = BLANK(prs)
    DRAW(sl, pat, PAL('Blue'), mk(pat, n=SPEC[pat]['max_elements']))
    bad = []
    for sh in sl.shapes:
        try:
            t = sh.top / EMU
            b = (sh.top + sh.height) / EMU
            l = sh.left / EMU
            r = (sh.left + sh.width) / EMU
        except TypeError:
            continue
        rot = float(getattr(sh, 'rotation', 0) or 0)
        # 回転図形は外接矩形が変わるため許容幅を持たせる
        tol = 60 if rot else 2
        if t < 60 - tol or b > AREA['body_bottom'] + tol or l < 0 - tol or r > 1280 + tol:
            bad.append('%s(t=%.0f b=%.0f l=%.0f r=%.0f rot=%.0f)'
                       % (sh.shape_type, t, b, l, r, rot))
    if bad:
        ng('T7', '%-9s 範囲外 %d件: %s' % (pat, len(bad), bad[:2]))
    else:
        ok('T7', '%-9s 全図形が描画範囲内（n=max）' % pat)

# ---------------------------------------------------------------- T8
print('\n=== T8｜フォント制約（メイリオ／14pt 下限）===')
import re as _re
for pat in P3:
    prs = P(); sl = BLANK(prs)
    DRAW(sl, pat, PAL('Blue'), mk(pat))
    xml = shapes_xml(sl).decode('utf-8', 'ignore')
    # ⚠️ 'メイリオ' は XML 内で '&#12513;...' と数値文字参照になるため復号する
    import html as _html
    fonts = {_html.unescape(f) for f in _re.findall(r'typeface="([^"]+)"', xml)}
    bad_f = {f for f in fonts if f and f not in ('Meiryo', 'メイリオ')}
    sizes = [int(s) / 100.0 for s in _re.findall(r'sz="(\d+)"', xml)]
    small = [s for s in sizes if s < 14]
    if bad_f:
        ng('T8', '%-9s メイリオ以外のフォント: %s' % (pat, bad_f))
    elif small:
        ng('T8', '%-9s 14pt未満: %s' % (pat, sorted(set(small))))
    else:
        ok('T8', '%-9s Meiryo のみ／最小 %.0fpt' % (pat, min(sizes) if sizes else 0))

# ---------------------------------------------------------------- T9
print('\n=== T9｜cycle の均等配分（段階減衰にしない）===')
pal = PAL('Blue')
prs = P(); sl = BLANK(prs)
DRAW(sl, 'cycle', pal, mk('cycle', n=6, score_base=90))
xml = shapes_xml(sl).decode('utf-8', 'ignore').upper()
cyc = NS.get('V17_CYCLE_TIERS') or NS.get('V17_CYCLE_COLORS')
fills = set()
for sh in sl.shapes:
    try:
        if sh.fill.type is not None and sh.fill.fore_color.rgb is not None:
            fills.add(str(sh.fill.fore_color.rgb).upper())
    except Exception:
        pass
used = [k for k in cyc if pal[k].lstrip('#').upper() in fills]
if len(used) >= 3:
    ok('T9', 'n=6 で循環色 %d 種を使用（%s）' % (len(used), used))
else:
    ng('T9', '循環色が %d 種しか使われていない: %s' % (len(used), used))
if pal.get('lightest', '').lstrip('#').upper() not in xml:
    ok('T9', '段階減衰色（lightest）を使っていない＝均等配分')
else:
    ng('T9', 'lightest が使われている（段階減衰の疑い）')

# ---------------------------------------------------------------- T10
print('\n=== T10｜contrast の2極化（中間色を排する）===')
pal = PAL('Blue')
prs = P(); sl = BLANK(prs)
DRAW(sl, 'contrast', pal, mk('contrast', score_base=90))
xml = shapes_xml(sl).decode('utf-8', 'ignore').upper()
pole = NS.get('V17_CONTRAST_TIERS') or ['lightest', 'primary']
hit = [k for k in pole if pal[k].lstrip('#').upper() in xml]
if len(hit) == 2:
    ok('T10', '2極色 %s の両方を使用' % pole)
else:
    ng('T10', '2極色が揃っていない: %s' % hit)
if SPEC['contrast']['min_elements'] == SPEC['contrast']['max_elements'] == 2:
    ok('T10', 'min=max=2 固定（原本 use「対照的な2要素」準拠）')
else:
    ng('T10', 'contrast の要素数が2固定でない')

# ---------------------------------------------------------------- T11
print('\n=== T11｜draw_pattern ディスパッチ（11パターン＋退避）===')
for pat in ALL11:
    prs = P(); sl = BLANK(prs)
    try:
        d = mk(pat) if pat in P3 else {
            'title': 't',
            'categories': [{'label': 'a', 'score': 70, 'description': 'x'}] * 3,
            'whole': {'label': '全体', 'value': 100},
            'components': [{'label': 'a', 'value': 50}, {'label': 'b', 'value': 50}],
            'items': [{'label': 'a', 'score': 70}, {'label': 'b', 'score': 60}],
            'levels': [{'label': 'a', 'score': 70}] * 3,
            'steps': [{'label': 'a', 'score': 70}] * 3,
            'cells': [{'row': 0, 'col': 0, 'label': 'a', 'score': 70}] * 4,
            'axis_x_label': 'x', 'axis_y_label': 'y',
        }
        rep = DRAW(sl, pat, PAL('Blue'), d)
        if rep.get('pattern'):
            ok('T11', '%-11s → %s' % (pat, rep.get('pattern')))
        else:
            ng('T11', '%s レポートが空' % pat)
    except Exception as e:
        ng('T11', '%s 例外 %s: %s' % (pat, type(e).__name__, e))
# 未実装キー
prs = P(); sl = BLANK(prs)
rep = DRAW(sl, 'integration', PAL('Blue'),
           {'title': 't', 'categories': [{'label': 'a', 'score': 70}] * 3})
if rep.get('fallback_from') == 'integration':
    ok('T11', '未実装キー integration → category へ退避')
else:
    ng('T11', '未実装キーの退避が機能しない: %r' % rep)

# ---------------------------------------------------------------- T12
print('\n=== T12｜後方互換（P1・P2 の描画が変化しない）===')
LEGACY = {
    'category':   {'title': 't', 'categories': [
        {'label': 'a', 'score': 82, 'description': 'x'},
        {'label': 'b', 'score': 64, 'description': 'y'},
        {'label': 'c', 'score': 38, 'description': 'z'}]},
    # ⚠️ breakdown の正式キーは 'value'（'ratio' ではない）
    'breakdown':  {'title': 't', 'whole': {'label': '全体', 'value': 100},
                   'components': [
                       {'label': 'a', 'value': 50}, {'label': 'b', 'value': 30},
                       {'label': 'c', 'value': 20}]},
    'comparison': {'title': 't', 'items': [
        {'label': 'a', 'score': 70}, {'label': 'b', 'score': 45}]},
    'pyramid':    {'title': 't', 'levels': [
        {'label': 'a', 'score': 38}, {'label': 'b', 'score': 64},
        {'label': 'c', 'score': 82}]},
    'sequence':   {'title': 't', 'steps': [
        {'label': 'a', 'score': 78}, {'label': 'b', 'score': 52},
        {'label': 'c', 'score': 34}]},
    'framework':  {'title': 't', 'axis_x_label': 'x', 'axis_y_label': 'y',
                   'cells': [
                       {'row': 0, 'col': 0, 'label': 'a', 'score': 88},
                       {'row': 0, 'col': 1, 'label': 'b', 'score': 62},
                       {'row': 1, 'col': 0, 'label': 'c', 'score': 55},
                       {'row': 1, 'col': 1, 'label': 'd', 'score': 31}]},
}
for pat, d in LEGACY.items():
    prs = P(); sl = BLANK(prs)
    try:
        rep = DRAW(sl, pat, PAL('Blue'), copy.deepcopy(d))
        if rep.get('fallback_from'):
            ng('T12', '%s が想定外にフォールバック: %s' % (pat, rep.get('notes')))
        elif len(sl.shapes) < 2:
            ng('T12', '%s 図形数不足 %d' % (pat, len(sl.shapes)))
        else:
            ok('T12', '%-11s shapes=%d（P3 追記後も正常）' % (pat, len(sl.shapes)))
    except Exception as e:
        ng('T12', '%s 例外 %s: %s' % (pat, type(e).__name__, e))
# 原本 DIAGRAM_PATTERNS の無改変
if len(NS['DIAGRAM_PATTERNS']) == 12:
    ok('T12', '原本 DIAGRAM_PATTERNS は12件（無改変）')
else:
    ng('T12', '原本 DIAGRAM_PATTERNS が %d 件（改変の疑い）' % len(NS['DIAGRAM_PATTERNS']))
if len(SPEC) == 11:
    ok('T12', '拡張 DIAGRAM_PATTERN_SPEC は11件（P1 3＋P2 3＋P3 5）')
else:
    ng('T12', '拡張 SPEC が %d 件' % len(SPEC))

# ---------------------------------------------------------------- 集計
print('\n' + '=' * 72)
print(' 結果: PASS %d / FAIL %d  （合計 %d 項目）'
      % (len(PASS), len(FAIL), len(PASS) + len(FAIL)))
print('=' * 72)
if FAIL:
    print('\n⚠️ FAIL 一覧:')
    for tid, msg in FAIL:
        print('  [%s] %s' % (tid, msg))
    sys.exit(1)
print('\n⭐ 全項目 PASS')
