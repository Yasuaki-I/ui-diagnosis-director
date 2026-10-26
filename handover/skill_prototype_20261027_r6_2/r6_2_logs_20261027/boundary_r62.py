# 境目の試験（設計 r2 §6-3 ＋ 判定118 カード内容の境界）｜/tmp だけで実行・test_*.pptx は作らない
import json, copy, os, sys, shutil, subprocess, importlib.util
R6='/tmp/r6work/r62/ui-diagnosis-director'
spec=importlib.util.spec_from_file_location('co',R6+'/scripts/check_output.py'); co=importlib.util.module_from_spec(spec); spec.loader.exec_module(co)
BASE=json.load(open(R6+'/assets/sample_input_rankB.json',encoding='utf-8'))
WORK='/tmp/r6work/bt'; shutil.rmtree(WORK,ignore_errors=True); os.makedirs(WORK)
results=[]
def gen_check(data,tag):
    d=f'{WORK}/{tag}'; os.makedirs(d,exist_ok=True); jp=f'{d}.json'
    json.dump(data,open(jp,'w',encoding='utf-8'),ensure_ascii=False)
    g=subprocess.run(['python3','/tmp/r6work/harness_r6.py',R6,jp,d],capture_output=True,text=True)
    if g.returncode!=0:
        return None,'builder停止：'+(g.stderr.strip().splitlines() or ['?'])[-1][:120]
    c=subprocess.run(['python3',R6+'/scripts/check_output.py',jp,d,'--report',d+'_rep.json'],capture_output=True,text=True)
    return c.returncode,c.stdout
def setv(data,path,val):
    o=data
    for k in path[:-1]: o=o[k]
    o[path[-1]]=val
def case(name, mut, expect, want=None, mode='input', gen=False):
    data=copy.deepcopy(BASE); mut(data)
    if mode=='input' and not gen:
        ng,notes=co.check_input_r6(data)
        got='NG' if ng else '合格'
        txt=' / '.join(ng) if ng else ' / '.join(notes)
    else:
        rc,out=gen_check(data,name.replace(' ','_')[:40])
        if rc is None: got='NG'; txt=out
        else:
            got='合格' if rc==0 else 'NG'
            txt=' / '.join(l.strip() for l in out.splitlines() if ('[NG]' in l or '[未検査]' in l or '[要実機確認]' in l and '固定カード' not in l or l.strip().startswith('- ') or '表示名' in l))[:400]
    ok=(got==expect) and (want is None or want in txt)
    results.append((name,expect,got,'⭕' if ok else '🚨',txt[:300]))
F='visual_data'; P='proposals'; D='diagnosis'
def both_area(data,i,val):
    data[P]['proposals'][i]['target_area']=val; data[F]['highlights'][i]['target_area']=val
def both(data,i,key,val):
    data[P]['proposals'][i][key]=val; data[F]['highlights'][i][key]=val
def all_service(data,val):
    for k in (D,P,F): data[k]['service_name']=val
# ---- 正常：上限ちょうど（生成して点検）----
def at_limit(d):
    all_service(d,'あ'*11)
    for k in (D,F): d[k]['input_type']='URL / '+'エ'*7  # 3+7=10換算
    d[F]['summary']='あ'*17
    for i,s in enumerate(d[F]['sections']): s['label']='見'*9 if i==0 else s['label']; 
    d[F]['sections'][0]['desc']='説'*7   # 9+7=16
    d[F]['sections'][1]['label']='見'*6; d[F]['sections'][1]['desc']='説'*10  # 6+10=16
    d[F]['flow_summary']='流'*31
    for i in range(3):
        both_area(d,i,'箇'*9+'一二三'[i]); both(d,i,'before','前'*24); both(d,i,'after','後'*24)  # r6.1：提案ごとに異なる10換算（対応を一意に）
        d[F]['highlights'][i]['title']='題'*27
    d[F]['purpose']='目'*14; d[F]['target']='対'*14
    for k in (D,F): d[k]['strengths'][0]='強'*28
case('上限ちょうど（全欄）',at_limit,'合格',gen=True)
# ---- C より1多い ----
case('サービス名 12換算',lambda d: all_service(d,'あ'*12),'NG','表示上限超過')
case('箇所 11換算（C-2/C-3 同じ）',lambda d: both_area(d,0,'箇'*11),'NG','表示上限超過')
case('現状 25換算',lambda d: both(d,0,'before','前'*25),'NG','表示上限超過')
case('構造マップ 見出し10換算',lambda d: d[F]['sections'][0].update(label='見'*10,desc='説'),'NG','表示上限超過')
case('構造マップ 説明11換算',lambda d: d[F]['sections'][0].update(label='見',desc='説'*11),'NG')
case('構造マップ 合計17換算',lambda d: d[F]['sections'][0].update(label='見'*8,desc='説'*9),'NG','合計' if False else '16')
case('フロー見出し 7換算',lambda d: d[F]['flow_steps'][0].update(label='見'*7),'NG')
case('フロー説明 10換算',lambda d: d[F]['flow_steps'][0].update(note='説'*10),'NG')
case('まとめ 32換算',lambda d: d[F].update(flow_summary='流'*32),'NG')
case('ハイライト title 28換算',lambda d: d[F]['highlights'][0].update(title='題'*28),'NG')
case('目的 15換算',lambda d: d[F].update(purpose='目'*15),'NG')
# ---- len・builder 換算 ----
case('総評 len18（換算17以下）',lambda d: d[F].update(summary='ab'*9),'NG','len 18')
case('フロー ASCII12＋ASCII18（len25超過）',lambda d: d[F]['flow_steps'][0].update(label='a'*12,note='b'*18),'NG','len')
case('構造マップ 見出しASCII17＋説明7換算',lambda d: d[F]['sections'][0].update(label='a'*17,desc='説'*7),'合格',gen=True)
case('強み 全角29字',lambda d: [d[k]['strengths'].__setitem__(0,'強'*29) for k in (D,F)],'NG','28')
case('強み ASCII29字（14.5換算）',lambda d: [d[k]['strengths'].__setitem__(0,'a'*29) for k in (D,F)],'合格',gen=True)
case('強み 全角20＋ASCII16（28換算・len36）',lambda d: [d[k]['strengths'].__setitem__(0,'強'*20+'a'*16) for k in (D,F)],'合格',gen=True)
case('半角記号の混在「CTA/FAQを追加」（箇所 4.5＋3＝7.5換算）',lambda d: both_area(d,0,'CTA/FAQを追加'),'合格')
case('半角記号の混在 箇所「CTA/FAQ導線を追加する」（12.5換算）',lambda d: both_area(d,0,'CTA/FAQ導線を追加する'),'NG')
# ---- 文字種 ----
case('明示の改行',lambda d: d[D].update(conclusion='一行目\n二行目'),'NG','改行')
case('半角カナ',lambda d: d[F].update(summary='ｻﾝﾌﾟﾙ'),'NG','半角カナ')
case('絵文字',lambda d: d[F].update(summary='良い😀'),'NG','絵文字')
case('異体字セレクタ',lambda d: d[F].update(summary='葛\U000E0100'),'NG')
case('異体字セレクタ VS16',lambda d: d[F].update(summary='注意\ufe0f'),'NG','異体字')
case('未対応の文字種 ✓',lambda d: d[F].update(summary='対応✓'),'NG','未対応')
# ---- 複数行・長い単語 ----
case('改善方向に長い ASCII 単語（生成）',lambda d: d[F].update(direction='改善には'+'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmn'+'を導入して導線を整えるべきだ。'),'合格',gen=True)
case('フロー6段すべて上限（見出し6・説明9）',lambda d: [s.update(label='見'*6,note='説'*9) for s in d[F]['flow_steps']],'合格',gen=True)
case('結論 2行の枠を超える（len60のまま）は不可→ len 61',lambda d: d[D].update(conclusion='結'*61),'NG','len 61')
# ---- record ----
case('正式名称30字＋表示名11字',lambda d: (d['record'].update(service_name_official='正'*30), all_service(d,'表'*11)),'合格','表示名')
case('service_name_official 欠落',lambda d: d['record'].pop('service_name_official'),'NG','service_name_official')
case('URL 起点で target_url が null',lambda d: d['record'].update(target_url=None),'NG','target_url')
case('input_origin 不明値',lambda d: d['record'].update(input_origin='web'),'NG','input_origin')
case('record 欠落',lambda d: d.pop('record'),'NG','record')
case('input_type と origin の食い違い',lambda d: d['record'].update(input_origin='screenshot',target_url=None),'NG')
case('service_name 3か所不一致',lambda d: d[P].update(service_name='別の名前'),'NG','一致しない')
for n,e,g,ok,t in results: print(ok,n,'| 期待',e,'| 結果',g,'|',t)
print('---',sum(1 for r in results if r[3]=='⭕'),'/',len(results))
