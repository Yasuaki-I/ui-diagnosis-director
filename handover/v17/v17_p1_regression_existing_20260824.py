# -*- coding: utf-8 -*-
"""手順書 §5 段1｜既存C-1〜C-3の後方互換｜統合版 vs 旧版の出力XML一致検証"""
import os, sys
from lxml import etree

def load(path):
    g={'__name__':'b'}
    exec(compile(open(path,encoding='utf-8').read(),path,'exec'), g)
    return g

OLD = load('builder_v16_5.py')
NEW = load('builder_v17_1.py')

DIAG = {
    'service_name':'サンプルEC','project_type':'ec',
    'scores':[{'label':'ファーストビュー','score':82,'comment':'訴求は明快'},
              {'label':'情報設計','score':64,'comment':'階層に飛びあり'},
              {'label':'導線設計','score':38,'comment':'CTAが画面外'}],
    'strengths':['訴求が明快','配色が統一','写真の質が高い'],
    'issues':['CTAが画面外','見出し階層の飛び','入力項目が多い'],
    'conclusion':'導線設計の是正が最優先。訴求力は既に十分に高い。',
    'total_score':61,
}
PROPOSALS = {
    'service_name':'サンプルEC',
    'proposals':[{'title':'CTAを画面内に固定','priority':'高','category':'導線',
                  'point':'ファーストビュー内にCTAを常時表示し、離脱前の接触機会を確保する。'},
                 {'title':'見出し階層を整理','priority':'中','category':'情報設計',
                  'point':'h2/h3の飛びを解消し、走り読みでも構造が伝わる状態にする。'}],
    'summary':'導線と情報設計の2点に絞れば短期で改善効果が出る。',
}

def sig(mod, fn, *a, **k):
    prs = mod['create_presentation']()
    fn(prs, *a, **k)
    return etree.tostring(prs.slides._sldIdLst.getparent()), \
           [etree.tostring(s._element) for s in prs.slides]

cases = []

def run(label, builder):
    try:
        o = builder(OLD); n = builder(NEW)
        ok = (o == n)
        cases.append((label, 'PASS' if ok else 'FAIL', '' if ok else 'XML差分あり'))
    except Exception as e:
        cases.append((label, 'ERROR', '%s: %s' % (type(e).__name__, e)))

def b_cover(m):
    prs=m['create_presentation']()
    m['add_cover'](prs, title='UI診断レポート', date='2026年8月', author='UI診断ディレクター', page_total=3)
    return [etree.tostring(s._element) for s in prs.slides]

def b_agenda(m):
    prs=m['create_presentation']()
    m['add_agenda'](prs, items=['現状分析','課題整理','改善提案','スケジュール'], page_num=2, total=5)
    return [etree.tostring(s._element) for s in prs.slides]

def b_issue(m):
    prs=m['create_presentation']()
    m['add_issue_summary'](prs, title='課題サマリ',
        cards=[{'title':'導線','body':'CTAが画面外に落ちている'},
               {'title':'情報設計','body':'見出し階層に飛びがある'},
               {'title':'可読性','body':'字間に改善の余地'}], page_num=3, total=5)
    return [etree.tostring(s._element) for s in prs.slides]

def b_matrix(m):
    prs=m['create_presentation']()
    m['add_priority_matrix'](prs, title='優先度マトリクス',
        items=[{'label':'CTA固定','x':0.85,'y':0.80,'priority':'S'},
               {'label':'見出し整理','x':0.70,'y':0.55,'priority':'A'},
               {'label':'写真差替','x':0.30,'y':0.25,'priority':'C'}], page_num=4, total=5)
    return [etree.tostring(s._element) for s in prs.slides]

def b_closing(m):
    prs=m['create_presentation']()
    m['add_closing'](prs, message='Thank you.', next_step='次回：改善後の再診断', page_num=5, total=5)
    return [etree.tostring(s._element) for s in prs.slides]

def b_scorecard(m):
    prs=m['create_presentation']()
    m['add_scorecard_onepager'](prs, DIAG, page_num=1, total=2)
    return [etree.tostring(s._element) for s in prs.slides]

def b_proposal(m):
    prs=m['create_presentation']()
    m['add_proposal_onepager'](prs, PROPOSALS, page_num=2)
    return [etree.tostring(s._element) for s in prs.slides]

def b_action(m):
    prs=m['create_presentation']()
    m['add_action_table'](prs, title='アクション', columns=['施策','担当','期限'],
        rows=[['CTA固定','制作','9/5'],['見出し整理','編集','9/12']], page_num=3, total=5)
    return [etree.tostring(s._element) for s in prs.slides]

for label, b in [
    ('段1-1｜add_cover', b_cover),
    ('段1-2｜add_agenda', b_agenda),
    ('段1-3｜add_issue_summary', b_issue),
    ('段1-4｜add_priority_matrix', b_matrix),
    ('段1-5｜add_action_table', b_action),
    ('段1-6｜add_closing', b_closing),
    ('段1-7｜add_scorecard_onepager (C-1)', b_scorecard),
    ('段1-8｜add_proposal_onepager (C-2)', b_proposal),
]:
    run(label, b)

# 辞書・定数の無改変
def cmp_obj(label, key, conv=lambda x:x):
    try:
        ok = conv(OLD[key]) == conv(NEW[key])
        cases.append((label,'PASS' if ok else 'FAIL','' if ok else '内容相違'))
    except Exception as e:
        cases.append((label,'ERROR',str(e)))

cmp_obj('段1-9｜DIAGRAM_PATTERNS 無改変','DIAGRAM_PATTERNS')
cmp_obj('段1-10｜DIAGNOSIS_TO_PATTERN 無改変','DIAGNOSIS_TO_PATTERN')
cmp_obj('段1-11｜DIGITAL_AGENCY_PALETTE 無改変','DIGITAL_AGENCY_PALETTE')
cmp_obj('段1-12｜LIMITS 無改変','LIMITS')
cmp_obj('段1-13｜DIGITAL_AGENCY_THRESHOLD 無改変','DIGITAL_AGENCY_THRESHOLD')

# validate_length の挙動同一
def b_validate():
    try:
        for key,txt in [('c1_comment','あ'*40),('c2_point','い'*150),('c3_summary','う'*80)]:
            assert OLD['validate_length'](txt,key)==NEW['validate_length'](txt,key)
        for key,txt in [('c1_comment','あ'*41)]:
            eo=en=None
            try: OLD['validate_length'](txt,key)
            except Exception as e: eo=str(e)
            try: NEW['validate_length'](txt,key)
            except Exception as e: en=str(e)
            assert eo==en and eo is not None
        cases.append(('段1-14｜validate_length 挙動同一','PASS',''))
    except Exception as e:
        cases.append(('段1-14｜validate_length 挙動同一','FAIL',str(e)))
b_validate()

npass=sum(1 for _,s,_ in cases if s=='PASS')
print('\n===== 手順書 §5 段1｜既存C-1〜C-3 後方互換検証 =====')
for l,s,m in cases:
    print('%-42s %-5s %s' % (l,s,m))
print('-----')
print('合計 %d 項目｜PASS %d｜FAIL/ERROR %d' % (len(cases),npass,len(cases)-npass))
sys.exit(0 if npass==len(cases) else 1)
