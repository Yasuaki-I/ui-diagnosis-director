# V1 ハーネス（スキルフォルダの外・試作の検証用）：SKILL.md の読込み3段を再現して PPTX 3点を生成
import json, os, sys, glob
SKILL = sys.argv[1]; INP = sys.argv[2]; OUT = sys.argv[3]
os.makedirs(OUT, exist_ok=True); os.chdir(OUT)
# ① 探す
cands = [os.path.join(SKILL, 'scripts', '03_pptx_builder.py'), os.path.join(os.getcwd(), '03_pptx_builder.py')]
path = next((p for p in cands if os.path.isfile(p)), None) or next(iter(glob.glob('/tmp/**/03_pptx_builder.py', recursive=True)), None)
if not path: print('builder が見つからない → 生成しない'); sys.exit(3)
# ② 2行方式で読込み → 場所と版を報告
_ns = {'__name__': 'pptx_builder'}; exec(open(path, encoding='utf-8').read(), _ns)
globals().update({k: v for k, v in _ns.items() if not k.startswith('__')})
print(f'builder：{path}｜版 {_ns.get('__version__')}')
d = json.load(open(INP, encoding='utf-8')); slug = d['site_slug']
for data, name, fn in [(d['diagnosis'], f'01_UI_Scorecard_{slug}', add_scorecard_onepager),
                       (d['proposals'], f'02_Proposals_{slug}', add_proposal_onepager),
                       (d['visual_data'], f'03_Visual_Board_{slug}', add_visual_board)]:
    prs = create_presentation(); fn(prs, data, page_num=1, slide_no=name[1]); prs.save(os.path.join(OUT, name + '.pptx'))
print('生成：', sorted(os.listdir(OUT)))
