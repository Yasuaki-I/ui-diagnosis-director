# -*- coding: utf-8 -*-
"""v3.7 固定入力セット（2026-10-04｜AIスライド）
11パターン × 要素数下限・上限（22件）＋ funnel の境界条件（4件）。
⚠️ funnel 新規入力（V37-1/2）は入力キーが未確定のため本版には含めない（各実装時に追補版を発行）。
説明文は推奨上限（全角30字）を基本とし、funnel 境界条件のみ推奨超過の長文（LONG）を用いる。"""
L = '広告接触とSNS経由で初回流入した推定ユーザー数の月間合計値'   # 推奨上限ちょうど（PLAIN J-7-5｜全角30字）
LONG = '広告接触およびSNS経由で初回流入した推定ユニークユーザー数の月間合計値である'   # 推奨超過（9/17 長文検証データと同一）
S = '初回流入の月間合計'
def items(n, key='label', desc=L, score=True, extra=None):
    out = []
    for i in range(n):
        d = {key: '要素%d' % (i + 1), 'score': (90 - i * 12) if score else None, 'description': desc}
        if extra: d.update(extra(i))
        out.append(d)
    return out
def comp(n):
    return {'title': '内訳（n=%d）' % n, 'whole': {'label': '全体', 'value': 100 * n},
            'components': [{'label': '構成%d' % (i+1), 'value': 100 - i*10, 'score': 80 - i*8, 'note': L[:24]} for i in range(n)]}
def cmpd(n):
    return {'title': '比較（n=%d）' % n, 'comparison_axis': '評価軸', 'attribute_labels': ['速度', '品質', '費用'],
            'items': [{'label': '案%d' % (i+1), 'score': 70 + i*5, 'attributes': {'速度': '高', '品質': '中', '費用': '低'}} for i in range(n)]}
def fw(n):
    rows, cols = {4:(2,2),5:(3,2),6:(3,2),7:(3,3),8:(3,3),9:(3,3)}[n]
    cells = []
    for k in range(n):
        cells.append({'row': k // rows, 'col': k % rows, 'label': 'セル%d' % (k+1), 'score': 85 - k*7, 'items': ['項目A', '項目B']})
    return {'title': '枠組み（n=%d）' % n, 'axis_x_label': '効果', 'axis_y_label': '難易度',
            'axis_x_low': '低', 'axis_x_high': '高', 'axis_y_low': '易', 'axis_y_high': '難', 'cells': cells}
def net(n):
    nodes = [{'id': 'n%d' % i, 'label': 'ノード%d' % (i+1), 'score': 80 - i*6} for i in range(n)]
    edges = [{'from': 'n0', 'to': 'n%d' % i} for i in range(1, min(n, 3))] + \
            [{'from': 'n%d' % ((i-3) % 2 + 1), 'to': 'n%d' % i} for i in range(3, n)]
    return {'title': 'ネットワーク（n=%d）' % n, 'nodes': nodes, 'edges': edges}
VALS = [120000, 42000, 12600, 8200, 5080, 1960]
def funnel(n, desc=L, dense=False, tn=False):
    st = items(n, desc=desc, extra=(lambda i: {'value': VALS[i], 'unit': '件'}) if dense else None)
    d = {'title': '絞り込み（n=%d%s）' % (n, '｜密度強化' if dense else ''), 'stages': st}
    if dense:
        st[1]['action'] = 'LP改善'
    if tn: d['total_note'] = '= FY25売上680億円'
    return d
CASES = []
def add(key, pattern, data): CASES.append((key, pattern, data))
for n in (3, 6): add('category_n%d' % n, 'category', {'title': '分類（n=%d）' % n, 'categories': items(n)})
for n in (3, 7): add('breakdown_n%d' % n, 'breakdown', comp(n))
for n in (2, 3): add('comparison_n%d' % n, 'comparison', cmpd(n))
for n in (3, 5): add('pyramid_n%d' % n, 'pyramid', {'title': '階層（n=%d）' % n, 'levels': items(n)})
for n in (3, 6): add('sequence_n%d' % n, 'sequence', {'title': '手順（n=%d）' % n, 'steps': items(n)})
for n in (4, 9): add('framework_n%d' % n, 'framework', fw(n))
for n in (3, 6): add('funnel_n%d' % n, 'funnel', funnel(n))
for n in (3, 6): add('cycle_n%d' % n, 'cycle', {'title': '循環（n=%d）' % n, 'cycle_name': '改善', 'phases': items(n)})
add('contrast_n2', 'contrast', {'title': '対比（n=2）', 'sides': [{'label': '現状', 'score': 40, 'items': ['課題A', '課題B', '課題C']}, {'label': '理想', 'score': 85, 'items': ['解決A', '解決B', '解決C']}]})
add('contrast_n2_long', 'contrast', {'title': '対比（n=2｜長文）', 'sides': [{'label': '現状', 'score': 40, 'items': [LONG, LONG]}, {'label': '理想', 'score': 85, 'items': [LONG, LONG]}]})
for n in (3, 7): add('timeline_n%d' % n, 'timeline', {'title': '時系列（n=%d）' % n, 'axis_label': '2026年', 'milestones': items(n, extra=lambda i: {'axis': '%d月' % (i+4)})})
for n in (3, 7): add('network_n%d' % n, 'network', net(n))
# funnel 境界条件（V37-3 の対象分岐を必ず通す）
add('funnel_n6_short', 'funnel', funnel(6, desc=S))
add('funnel_n6_long', 'funnel', funnel(6, desc=LONG))
add('funnel_n6_dense_long', 'funnel', funnel(6, desc=LONG, dense=True, tn=True))
add('funnel_n5_dense_long', 'funnel', funnel(5, desc=LONG, dense=True, tn=True))
add('funnel_n6_noscore', 'funnel', {'title': '絞り込み（n=6｜score無）', 'stages': items(6, score=False)})
