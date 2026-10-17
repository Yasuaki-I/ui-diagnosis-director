# -*- coding: utf-8 -*-
# 【購入者限定】本ファイルの再配布・転載・転売は固く禁じます。｜UI診断ディレクター スキル版 prototype（2026-10-18）
"""check_output.py｜生成後の点検（不合格なら AI が直して作り直す）

使い方:
    python3 check_output.py <入力JSON> <出力フォルダ>
    入力JSON = {"site_slug":..., "diagnosis":{...}, "proposals":{...}, "visual_data":{...}}

点検項目:
    1. 総評（visual_data.summary）が17字以内
    2. 件数：強み・課題がランクどおり（S/A/B=3・C=4・D=5）。C-1 と C-3・top_issues の件数一致
    3. ファイル名が英数字（01_UI_Scorecard_/02_Proposals_/03_Visual_Board_ + 英数字 + .pptx）で3点そろう
    4. 出力フォルダに test_*.pptx が無い
    5. PPTX にハイパーリンクが無い
    6. フォントがメイリオ・14pt 以上（参考：builder 側でも止まる）

終了コード: 0＝合格／1＝不合格（直す箇所を表示）／2＝引数誤り
"""
import json, os, re, sys, glob, zipfile

RANK_COUNT = {'S': 3, 'A': 3, 'B': 3, 'C': 4, 'D': 5}
SUMMARY_MAX = 17
NAME_RE = {
    'C-1': re.compile(r'^01_UI_Scorecard_[A-Za-z0-9_]+\.pptx$'),
    'C-2': re.compile(r'^02_Proposals_[A-Za-z0-9_]+\.pptx$'),
    'C-3': re.compile(r'^03_Visual_Board_[A-Za-z0-9_]+\.pptx$'),
}


def check(data, out_dir):
    ng = []
    d, v = data.get('diagnosis', {}), data.get('visual_data', {})
    s = v.get('summary') or ''
    if len(s) > SUMMARY_MAX:
        ng.append(f'総評が{len(s)}字（上限{SUMMARY_MAX}字）：「{s}」→ 意味を保って{SUMMARY_MAX}字以内に言い換える（末尾「…」禁止）')
    rank = str(d.get('rank', '')).strip()
    need = RANK_COUNT.get(rank)
    if need is None:
        ng.append(f'ランク「{rank}」が S/A/B/C/D のどれでもない')
    else:
        for label, arr in [('C-1 強み', d.get('strengths')), ('C-1 課題', d.get('priority_issues')),
                           ('C-3 強み', v.get('strengths')), ('C-3 課題', v.get('priority_issues')),
                           ('C-3 top_issues', v.get('top_issues'))]:
            n = len(arr or [])
            if n != need:
                ng.append(f'{label}が{n}件（ランク{rank}は{need}件）')
    files = sorted(os.path.basename(p) for p in glob.glob(os.path.join(out_dir, '*.pptx')))
    for key, rx in NAME_RE.items():
        hit = [f for f in files if rx.match(f)]
        if len(hit) != 1:
            ng.append(f'{key} のファイルが{len(hit)}件（1件であるべき・英数字のファイル名）')
    for f in files:
        if not re.match(r'^[A-Za-z0-9_.-]+$', f):
            ng.append(f'英数字以外のファイル名：{f}')
    tests = [f for f in files if f.startswith('test_')]
    if tests:
        ng.append(f'test_*.pptx がある：{tests}（builder の読み込み方を確認。デモが動いている）')
    for f in files:
        with zipfile.ZipFile(os.path.join(out_dir, f)) as z:
            for n in z.namelist():
                if n.endswith('.rels') and 'hyperlink' in z.read(n).decode('utf-8', 'ignore'):
                    ng.append(f'ハイパーリンクがある：{f}（{n}）')
                if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
                    x = z.read(n).decode('utf-8', 'ignore')
                    for sz in re.findall(r'<a:rPr[^>]*\ssz="(\d+)"', x):
                        if int(sz) < 1400:
                            ng.append(f'14pt 未満の文字：{f}（{n}・{int(sz)/100}pt）'); break
                    faces = set(re.findall(r'typeface="([^"]+)"', x)) - {'+mn-lt', '+mj-lt', '+mn-ea', '+mj-ea', '+mn-cs', '+mj-cs'}
                    other = sorted(t for t in faces if t not in ('Meiryo', 'メイリオ', 'Meiryo UI'))
                    if other:
                        ng.append(f'メイリオ以外のフォント：{f}（{n}・{other}）')
    return files, ng


def main(argv):
    if len(argv) != 2:
        print(__doc__); return 2
    data = json.load(open(argv[0], encoding='utf-8'))
    files, ng = check(data, argv[1])
    print('点検対象：', ', '.join(files) or '（なし）')
    if ng:
        print(f'🚨 不合格 {len(ng)}件'); [print('  -', m) for m in ng]; return 1
    print('⭕ 合格（総評17字・件数・ファイル名・test_*.pptx 無し・リンク無し・メイリオ14pt以上）'); return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
