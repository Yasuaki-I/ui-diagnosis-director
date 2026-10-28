# 読み取り専用で manifest 記載のファイルを /tmp/ro/mirror に取得する（ドライブへの書込み操作は一切使わない）
import subprocess, json, os, sys, urllib.request, concurrent.futures as cf
BASE='/ui-diagnosis-director/'; OUT='/tmp/ro/mirror'
rows=[l.rstrip('\n').split('\t') for l in open(sys.argv[1],encoding='utf-8') if l.strip() and not l.startswith('#')]
def url_of(p):
    r=subprocess.run(['gsk','drive','get_readable_url','--path',BASE+p],capture_output=True,text=True,timeout=90)
    d=json.loads(r.stdout); t=d['data']['result']
    return t.split(': ',1)[1].split('\n',1)[0].strip()
def job(row):
    p=row[0]; dst=os.path.join(OUT,p)
    for a in range(3):
        try:
            u=url_of(p); os.makedirs(os.path.dirname(dst),exist_ok=True)
            open(dst,'wb').write(urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':'Mozilla/5.0'}),timeout=120).read()); return p,'ok'
        except Exception as e: err=str(e)[:120]
    return p,'取得失敗:'+err
with cf.ThreadPoolExecutor(8) as ex:
    res=list(ex.map(job,rows))
bad=[r for r in res if r[1]!='ok']
print('取得', len(res)-len(bad), '/', len(res)); [print(b) for b in bad]
