# r5 の崩し入力 t1〜t11・u1・u3・追加4種（定義から再構成。r5 当時のファイルそのものではない）
import json, copy, os, shutil, subprocess
from pptx import Presentation
R6='/tmp/r6work/r6/ui-diagnosis-director'; W='/tmp/r6work/bt3'; shutil.rmtree(W,ignore_errors=True); os.makedirs(W)
BASE=json.load(open(R6+'/assets/sample_input_rankB.json',encoding='utf-8'))
def gen(data,d):
    jp=d+'_gen.json'; json.dump(data,open(jp,'w',encoding='utf-8'),ensure_ascii=False)
    return subprocess.run(['python3','/tmp/r6work/harness_r6.py',R6,jp,d],capture_output=True,text=True)
def chk(data,d):
    jp=d+'_chk.json'; json.dump(data,open(jp,'w',encoding='utf-8'),ensure_ascii=False)
    c=subprocess.run(['python3',R6+'/scripts/check_output.py',jp,d],capture_output=True,text=True); return c.returncode,c.stdout
rows=[]
def T(name, gen_data, chk_data, want, extra=None):
    d=f'{W}/{name}'; g=gen(gen_data,d)
    if g.returncode: rows.append((name,'builder停止',g.stderr.strip().splitlines()[-1][:100],'—')); return
    if extra: extra(d)
    rc,out=chk(chk_data,d)
    hit=[l.strip() for l in out.splitlines() if want in l]
    rows.append((name,rc,hit[0][:130] if hit else '(期待の理由なし)','⭕' if rc==1 and hit else '🚨'))
B=lambda: copy.deepcopy(BASE)
V='visual_data'; D='diagnosis'
def m(f): d=B(); f(d); return d
T('t1 総評の末尾…', m(lambda d:d[V].update(summary='FV内CTA追加で誘導…')), m(lambda d:d[V].update(summary='FV内CTA追加で誘導…')), '省略記号')
T('t2 総評が空', m(lambda d:d[V].update(summary='')), m(lambda d:d[V].update(summary='')), '総評が空')
T('t3 総評の項目なし', m(lambda d:d[V].pop('summary')), m(lambda d:d[V].pop('summary')), '項目が無い')
T('t4 JSONの総評だけ書換', B(), m(lambda d:d[V].update(summary='別の総評に差し替え')), '総評「別の総評に差し替え」が見当たらない')
T('t5 rank_label 長い', m(lambda d:[d[k].update(rank_label='標準的（改善余地あり）') for k in (D,V)]), m(lambda d:[d[k].update(rank_label='標準的（改善余地あり）') for k in (D,V)]), 'ランクの説明が')
T('t6 rank_label 優良（B）', m(lambda d:[d[k].update(rank_label='優良') for k in (D,V)]), m(lambda d:[d[k].update(rank_label='優良') for k in (D,V)]), 'ランクの説明が')
T('t7 rank_label なし', m(lambda d:[d[k].pop('rank_label') for k in (D,V)]), m(lambda d:[d[k].pop('rank_label') for k in (D,V)]), 'rank_label）が無い')
T('t8 総評18字', m(lambda d:d[V].update(summary='あ'*18)), m(lambda d:d[V].update(summary='あ'*18)), '総評が18字')
T('t9 rank_label を直したが C-1 を作り直さない', m(lambda d:[d[k].update(rank_label='優良') for k in (D,V)]), B(), 'ランク欄が入力')
T('t10 課題の文だけ書換', B(), m(lambda d:[d[k].__setitem__('priority_issues',['別の課題A','別の課題B','別の課題C']) for k in (D,V)] and None or d[V].__setitem__('top_issues',['別の課題A','別の課題B','別の課題C'])), '課題『別の課題A』が PPTX に見当たらない')
def two(d):
    for k in (D,V):
        d[k]['strengths']=d[k]['strengths'][:2]; d[k]['priority_issues']=d[k]['priority_issues'][:2]
    d[V]['top_issues']=d[V]['top_issues'][:2]
T('t11 2件で生成・JSONは3件', m(two), B(), 'が PPTX に見当たらない')
def r3(d):
    for h in d[V]['highlights']:
        h['area']=h.pop('target_area'); h.pop('priority'); h.pop('effort')
    d[V].pop('flow_summary')
T('u1 area のまま・flow_summary なし', m(r3), m(r3), 'area ではなく target_area')
T('u3 PPTX は r3 型・JSON だけ足す', m(r3), B(), '箇所：')
T('追加 priority＝最優先', m(lambda d:d[V]['highlights'][0].update(priority='最優先')), m(lambda d:d[V]['highlights'][0].update(priority='最優先')), '優先度（priority）')
T('追加 flow_summary なし', m(lambda d:d[V].pop('flow_summary')), m(lambda d:d[V].pop('flow_summary')), 'flow_summary）が無い')
T('追加 flow_summary 41字', m(lambda d:d[V].update(flow_summary='流'*41)), m(lambda d:d[V].update(flow_summary='流'*41)), '41字')
T('追加 target_area なし', m(lambda d:d[V]['highlights'][0].pop('target_area')), m(lambda d:d[V]['highlights'][0].pop('target_area')), '箇所（target_area）が無い')
def add_test(d):
    shutil.copy(d+'/01_UI_Scorecard_sample_lp.pptx', d+'/test_demo.pptx')
T('test_*.pptx の混入', B(), B(), 'test_*.pptx がある', extra=add_test)
for r in rows: print(r[3] if len(r)>3 else '', r[0],'| exit',r[1],'|',r[2])
print('---', sum(1 for r in rows if r[-1]=='⭕'),'/',len(rows))
