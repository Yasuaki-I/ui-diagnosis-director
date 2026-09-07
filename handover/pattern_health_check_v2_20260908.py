#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
パターン健全性チェック v2（キュー③｜fallback_from 判定の追加）
==============================================================
発行: 2026-09-08（火）｜AIスライド（実装領域）
根拠: 2026-09-07 19:00 統括判定 §3（諮問C 承認・キュー③として登録）
対象: 03_pptx_builder.py v17.2.2（258,551B / md5 818c840c5c64）

■ v1（9/4）からの変更点
  1. report['fallback_from'] が None であることを判定条件に追加
  2. 全12パターンで再測定（正しいデータキーを使用）
  3. 意図的に誤キーを渡し、fallback が異常として検出されることを検証

■ 判断原理19の適用｜本検査が「検出できないこと」（先に宣言する）
  - D1: PowerPoint レンダラ固有の字形差による実描画幅の差（python-pptx は
        テキストの実測幅を持たないため、枠高不足は行数×行高の計算値で判定する）
  - D2: 図形の重なりのうち「意図的な重ね」（軸線とドット等）の妥当性
  - D3: 色のコントラスト比・可読性（テーマ仕様としての灰色は検出対象外）
  - D4: notes に記録された警告文の妥当性（文言自体の正しさは検査しない）
  - D5: integration（12番目）の構造的対象外という判断の妥当性
        （fallback_from='integration' は「異常」ではなく「仕様」として扱う）
"""

import sys, os, importlib.util

CANVAS_W, CANVAS_H = 1280, 720
EMU_PER_PX = 9525  # 1280px = 13.333in = 12192000 EMU


def load_builder(path):
    spec = importlib.util.spec_from_file_location('builder', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def px(emu):
    return emu / EMU_PER_PX


# ---------------------------------------------------------------
# 正しいデータキー（9/7・9/8 に実装から確定｜判断原理18）
# ---------------------------------------------------------------
def make_data(pattern, n):
    label = lambda i: '項目%d' % (i + 1)
    desc = 'これは説明文であり2行に折返す長さの日本語テキストである'
    el = lambda i: {'label': label(i), 'score': 62 + i * 3, 'description': desc}

    base = {'title': '%s パターン検査（n=%d）' % (pattern, n)}
    KEY = {
        'category': 'categories', 'breakdown': 'components',
        'comparison': 'items', 'pyramid': 'levels', 'sequence': 'steps',
        'framework': 'cells', 'funnel': 'stages', 'cycle': 'phases',
        'contrast': 'sides', 'timeline': 'milestones', 'network': 'nodes',
        'integration': 'categories',
    }
    items = [el(i) for i in range(n)]

    if pattern == 'breakdown':
        # breakdown は value（数値）必須｜合計>0 でなければ前提不成立
        # （9/8 実装確認｜L4141-4155 の数値健全性チェック）
        for i, it in enumerate(items):
            it['value'] = 10 + i * 5
            it['note'] = desc
        base['whole'] = {'label': '全体', 'value': sum(10 + i * 5
                                                       for i in range(n))}

    if pattern == 'timeline':
        # requires_axes=True｜axis_label と各要素の axis が必須
        for i, it in enumerate(items):
            it['axis'] = '第%d期' % (i + 1)
        base['axis_label'] = '時間軸（四半期）'
    elif pattern == 'framework':
        # requires_axes=True｜axis_x_label / axis_y_label が必須
        base['axis_x_label'] = '実現容易性'
        base['axis_y_label'] = '効果の大きさ'
        base['axis_x_low'], base['axis_x_high'] = '低', '高'
        base['axis_y_low'], base['axis_y_high'] = '小', '大'
    elif pattern == 'comparison':
        base['attribute_labels'] = ['指標A', '指標B']
        base['comparison_axis'] = '比較軸'
    elif pattern == 'cycle':
        base['cycle_name'] = '改善サイクル'
    elif pattern == 'network':
        base['edges'] = [{'from': label(i), 'to': label(i + 1)}
                         for i in range(n - 1)]

    base[KEY[pattern]] = items
    return base


# 12パターン × 検査要素数
#   各 spec の max_elements を採用（9/8 実装から確定｜判断原理18）
#   ⚠️ v1（9/4）は network に n=9 を渡していたが実装 max は 7 であった
TARGETS = [
    ('category', 6), ('breakdown', 7), ('comparison', 3),
    ('pyramid', 5), ('sequence', 6), ('framework', 9),
    ('funnel', 6), ('cycle', 6), ('contrast', 2),
    ('timeline', 7), ('network', 7), ('integration', 6),
]


def inspect(B, pattern, n, data=None):
    """1パターンを描画し、fallback_from・枠外・重なりを測る"""
    prs = B.create_presentation()
    palette = B.get_theme_palette('Blue')
    d = data if data is not None else make_data(pattern, n)
    slide, report = B.add_diagram_slide(prs, pattern, palette, d,
                                        page_num=1, total=1)

    shapes = []
    for sh in slide.shapes:
        try:
            l, t = px(sh.left), px(sh.top)
            w, h = px(sh.width), px(sh.height)
        except (TypeError, AttributeError):
            continue
        shapes.append({'name': sh.shape_type, 'l': l, 't': t,
                       'w': w, 'h': h, 'r': l + w, 'b': t + h})

    oob = [s for s in shapes
           if s['l'] < -0.5 or s['t'] < -0.5
           or s['r'] > CANVAS_W + 0.5 or s['b'] > CANVAS_H + 0.5]

    return {
        'pattern': pattern, 'n': n,
        'fallback_from': report.get('fallback_from'),
        'elements_drawn': report.get('elements_drawn'),
        'notes': report.get('notes', []),
        'shapes': len(shapes), 'oob': len(oob),
        'oob_detail': [(round(s['l'], 1), round(s['t'], 1),
                        round(s['r'], 1), round(s['b'], 1)) for s in oob[:3]],
    }


def judge(r):
    """判定｜v2 で fallback_from を条件に追加"""
    ng = []
    # 【v2 追加】fallback_from が None でなければ異常
    #   ただし integration は構造的対象外のため仕様として扱う（D5）
    if r['fallback_from'] is not None and r['pattern'] != 'integration':
        ng.append('fallback_from=%s（フォールバック混入）' % r['fallback_from'])
    if r['oob'] > 0:
        ng.append('枠外 %d件 %s' % (r['oob'], r['oob_detail']))
    if r['elements_drawn'] == 0:
        ng.append('描画要素0件')
    return ng


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else '/tmp/b.py'
    B = load_builder(path)

    print('=' * 72)
    print('パターン健全性チェック v2｜fallback_from 判定を追加')
    print('builder:', path)
    print('=' * 72)

    print('\n[検査1] 全12パターン｜正しいデータキーで再測定')
    print('-' * 72)
    print('%-12s %3s %-14s %5s %5s %5s  %s'
          % ('pattern', 'n', 'fallback_from', 'drawn', 'shape', 'oob', '判定'))
    ng_total = []
    results = []
    for p, n in TARGETS:
        r = inspect(B, p, n)
        ng = judge(r)
        results.append((r, ng))
        if ng:
            ng_total.append((p, ng))
        print('%-12s %3d %-14s %5s %5d %5d  %s'
              % (p, n, str(r['fallback_from']), r['elements_drawn'],
                 r['shapes'], r['oob'], 'NG: ' + '; '.join(ng) if ng else 'OK'))

    print('\n[検査2] 誤キー注入｜fallback が異常として検出されるかの検証')
    print('-' * 72)
    print('%-12s %-24s %-14s  %s'
          % ('pattern', '注入内容', 'fallback_from', '検出'))
    inj = []
    # timeline: 9/4 に実際に起きた誤り（elements キー）
    d = make_data('timeline', 7)
    d2 = {'title': d['title'], 'elements': d['milestones']}
    r = inspect(B, 'timeline', 7, d2)
    inj.append(('timeline', 'elements キー（9/4の誤り）', r))
    # timeline: axis_label 欠落
    d3 = dict(make_data('timeline', 7)); d3.pop('axis_label')
    r = inspect(B, 'timeline', 7, d3)
    inj.append(('timeline', 'axis_label 欠落', r))
    # framework: 軸ラベル欠落
    d4 = dict(make_data('framework', 9))
    d4.pop('axis_x_label'); d4.pop('axis_y_label')
    r = inspect(B, 'framework', 9, d4)
    inj.append(('framework', '軸ラベル欠落', r))
    # funnel: 要素数レンジ外
    r = inspect(B, 'funnel', 9)
    inj.append(('funnel', '要素数9（max6超）', r))

    inj_ok = 0
    for p, what, r in inj:
        ng = judge(r)
        detected = any('fallback_from' in x for x in ng)
        inj_ok += 1 if detected else 0
        print('%-12s %-24s %-14s  %s'
              % (p, what, str(r['fallback_from']),
                 '検出' if detected else '⚠️ 未検出'))

    print('\n' + '=' * 72)
    print('検査1（12パターン）  : 異常 %d件' % len(ng_total))
    print('検査2（誤キー注入4件）: 検出 %d/4件' % inj_ok)
    print('=' * 72)
    return results, ng_total, inj


if __name__ == '__main__':
    main()
