# -*- coding: utf-8 -*-
"""第16条 段階3（統合済）判定｜稼働ビルダー本体に当該ロジックが存在するかの実体確認"""
import re
src=open('builder_v17_iriye.py',encoding='utf-8').read()
g={'__name__':'b'}; exec(compile(src,'b','exec'),g)

rows=[]
def chk(label, cond, evidence):
    rows.append((label,'✅' if cond else '❌', evidence))

# 段階3の判定基準：ビルダー実ファイル内に当該ロジックが存在する
for name in ('draw_category','draw_breakdown','draw_comparison'):
    m=re.search(r'^def %s\(' % name, src, re.M)
    line=src[:m.start()].count('\n')+1 if m else None
    chk('段階3｜%s の描画ロジックが本体に存在' % name, m is not None, '%d行' % line if line else '未検出')

# 辞書を読んで図形を描くロジックの存在（8/22検出3への直接の答え）
chk('段階3｜DIAGRAM_PATTERN_SPEC を参照した描画',
    "DIAGRAM_PATTERN_SPEC['category']" in src and 'spec[' in src,
    'draw_category 内で spec 参照')
chk('段階3｜DIAGNOSIS_TO_PATTERN を読むロジック',
    'DIAGNOSIS_TO_PATTERN.get' in src, 'resolve_pattern 内')
chk('段階3｜DIGITAL_AGENCY_PALETTE を読むロジック',
    'DIGITAL_AGENCY_PALETTE[theme_id]' in src, 'get_theme_palette 内')

# 細則7｜段階判定の根拠は __version__
chk("細則7｜__version__ == '17.0.0'", g.get('__version__')=='17.0.0', repr(g.get('__version__')))
chk("細則7｜__version_date__ == '2026-08-23'", g.get('__version_date__')=='2026-08-23', repr(g.get('__version_date__')))

# 実際に呼び出して描画が成立するか（動作の実証）
from lxml import etree
prs=g['create_presentation']()
sl=g['_blank_slide'](prs)
pal=g['get_theme_palette']('Blue')
rep=g['draw_pattern'](sl, g['resolve_pattern']('proposal_categorization'), pal,
    {'title':'統合確認','categories':[{'label':'A','score':70,'description':'x'},
     {'label':'B','score':50,'description':'y'},{'label':'C','score':30,'description':'z'}]})
chk('段階3｜resolve_pattern→draw_pattern の一気通貫',
    rep['pattern']=='category' and rep['elements_drawn']==3, str(rep))
chk('段階3｜図形が実際に生成された', len(sl.shapes)>0, '%d shapes' % len(sl.shapes))

# 21件の既存 add_* が全て健在
n_add=len(re.findall(r'^def (add_\w+)\(', src, re.M))
chk('後方互換｜既存 add_* 関数が21件健在', n_add==21, '%d件' % n_add)

print('\n===== 第16条 段階3（統合済）実体確認 =====')
for l,s,e in rows:
    print('%s %-48s %s' % (s,l,e))
ng=[r for r in rows if r[1]=='❌']
print('-----')
print('合計 %d 項目｜✅ %d｜❌ %d' % (len(rows),len(rows)-len(ng),len(ng)))
