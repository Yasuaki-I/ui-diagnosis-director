# -*- coding: utf-8 -*-
# 【購入者限定】本ファイルの再配布・転載・転売は固く禁じます。｜UI診断ディレクター スキル版 prototype r5 案A（2026-10-23｜統括判定待ち）
"""check_output.py｜生成後の点検（不合格なら AI が直して作り直す）

使い方:
    python3 check_output.py <入力JSON> <出力フォルダ>
    入力JSON = {"site_slug":..., "diagnosis":{...}, "proposals":{...}, "visual_data":{...}}

点検項目:
    1. 総評（visual_data.summary）：項目があり・空でなく・17字以内・末尾が「…」「...」でない
    2. ランクの説明（rank_label）：ランクに対応する1語（S＝卓越／A＝優良／B＝標準的／C＝要改善／D＝致命的）
    3. 件数：強み・課題とも各3件（ランクに関わらず・PPTX に載る件数）。C-1・C-3・top_issues で同じ
    4. ファイル名が英数字で3点そろう／test_*.pptx が無い
    5. PPTX にハイパーリンクが無い／メイリオ・14pt 以上
    6. PPTX の中の文字が入力と一致（C-3 の総評・C-1 のランク欄・強み・課題・top_issues の各文言）
       ＝JSON だけ直して PPTX を作り直していない場合（文言の差し替え・件数の水増し）を止める
    7. C-3 の Before/After：各件に target_area（箇所）・priority（高中低）・effort（小中大）があり、
       PPTX に「箇所：{target_area}」がある／行動の流れのまとめ（flow_summary）が40字以内で入り、PPTX にある

終了コード: 0＝合格／1＝不合格（直す箇所を表示）／2＝引数誤り
"""
import json, os, re, sys, glob, zipfile

ITEM_COUNT = 3  # 強み・課題の件数（builder の描画上限｜r5 案A）
RANK_SET = ('S', 'A', 'B', 'C', 'D')
RANK_LABEL = {'S': '卓越', 'A': '優良', 'B': '標準的', 'C': '要改善', 'D': '致命的'}
SUMMARY_MAX = 17
NAME_RE = {
    'C-1': re.compile(r'^01_UI_Scorecard_[A-Za-z0-9_]+\.pptx$'),
    'C-2': re.compile(r'^02_Proposals_[A-Za-z0-9_]+\.pptx$'),
    'C-3': re.compile(r'^03_Visual_Board_[A-Za-z0-9_]+\.pptx$'),
}


def slide_texts(path):
    """PPTX の各スライドの文字（段落ごと）を返す"""
    out = []
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)),
                       key=lambda n: int(re.findall(r'\d+', n)[0]))
        for n in names:
            x = z.read(n).decode('utf-8', 'ignore')
            paras = []
            for p in re.findall(r'<a:p>.*?</a:p>|<a:p .*?</a:p>', x, flags=re.S):
                t = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', p))
                if t:
                    paras.append(t.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>'))
            out.append(paras)
    return out


def check(data, out_dir):
    ng = []
    d, v = data.get('diagnosis', {}), data.get('visual_data', {})
    # 1. 総評
    if 'summary' not in v:
        ng.append('総評（visual_data.summary）の項目が無い → 17字以内の総評を入れる')
        s = None
    else:
        s = v.get('summary')
        if not isinstance(s, str) or not s.strip():
            ng.append('総評が空 → 17字以内の総評を入れる')
        else:
            if len(s) > SUMMARY_MAX:
                ng.append(f'総評が{len(s)}字（上限{SUMMARY_MAX}字）：「{s}」→ 意味を保って{SUMMARY_MAX}字以内に言い換える')
            if s.rstrip().endswith(('…', '...', '‥')):
                ng.append(f'総評の末尾が省略記号：「{s}」→ 省略せず言い切る形に言い換える')
    # 2・3. ランクの説明と件数
    rank = str(d.get('rank', '')).strip()
    need = ITEM_COUNT if rank in RANK_SET else None
    if need is None:
        ng.append(f'ランク「{rank}」が S/A/B/C/D のどれでもない')
    else:
        for label, src in [('C-1', d), ('C-3', v)]:
            if 'rank_label' in src and src.get('rank_label') != RANK_LABEL[rank]:
                ng.append(f'{label} のランクの説明が「{src.get("rank_label")}」→ ランク{rank}は「{RANK_LABEL[rank]}」をそのまま使う')
        if d.get('rank_label') is None:
            ng.append(f'C-1 のランクの説明（rank_label）が無い → 「{RANK_LABEL[rank]}」を入れる')
        for label, arr in [('C-1 強み', d.get('strengths')), ('C-1 課題', d.get('priority_issues')),
                           ('C-3 強み', v.get('strengths')), ('C-3 課題', v.get('priority_issues')),
                           ('C-3 top_issues', v.get('top_issues'))]:
            n = len(arr or [])
            if n != need:
                ng.append(f'{label}が{n}件（ランクに関わらず{need}件｜4件目以降は改善提案・結論に回す）')
    # 4. ファイル
    files = sorted(os.path.basename(p) for p in glob.glob(os.path.join(out_dir, '*.pptx')))
    found = {}
    for key, rx in NAME_RE.items():
        hit = [f for f in files if rx.match(f)]
        if len(hit) != 1:
            ng.append(f'{key} のファイルが{len(hit)}件（1件であるべき・英数字のファイル名）')
        else:
            found[key] = os.path.join(out_dir, hit[0])
    for f in files:
        if not re.match(r'^[A-Za-z0-9_.-]+$', f):
            ng.append(f'英数字以外のファイル名：{f}')
    tests = [f for f in files if f.startswith('test_')]
    if tests:
        ng.append(f'test_*.pptx がある：{tests}（builder の読み込み方を確認。デモが動いている）')
    # 5. リンク・フォント
    for f in files:
        with zipfile.ZipFile(os.path.join(out_dir, f)) as z:
            for n in z.namelist():
                if n.endswith('.rels') and 'hyperlink' in z.read(n).decode('utf-8', 'ignore'):
                    ng.append(f'ハイパーリンクがある：{f}（{n}）')
                if re.match(r'ppt/slides/slide\d+\.xml$', n):
                    x = z.read(n).decode('utf-8', 'ignore')
                    small = [int(sz) for sz in re.findall(r'<a:rPr[^>]*\ssz="(\d+)"', x) if int(sz) < 1400]
                    if small:
                        ng.append(f'14pt 未満の文字：{f}（{n}・{min(small)/100}pt）')
                    faces = set(re.findall(r'typeface="([^"]+)"', x)) - {'+mn-lt', '+mj-lt', '+mn-ea', '+mj-ea', '+mn-cs', '+mj-cs'}
                    other = sorted(t for t in faces if t not in ('Meiryo', 'メイリオ', 'Meiryo UI'))
                    if other:
                        ng.append(f'メイリオ以外のフォント：{f}（{n}・{other}）')
    # 6. PPTX の中身と入力の照合
    if 'C-3' in found and isinstance(s, str) and s.strip():
        paras = [p for sl in slide_texts(found['C-3']) for p in sl]
        if s not in paras:
            ng.append(f'C-3 の PPTX に総評「{s}」が見当たらない → 入力を直したら PPTX も作り直す')
    if 'C-1' in found and need is not None:
        want = f'{rank}  {d.get("rank_label") or ""}'.rstrip()
        sl = slide_texts(found['C-1'])
        hit_all = all(want in paras for paras in sl[:2]) if len(sl) >= 2 else (want in sl[0] if sl else False)
        if not hit_all:
            ng.append(f'C-1 の PPTX のランク欄が入力（「{want}」）と一致しない → 入力を直したら PPTX も作り直す')
    # 6-2. 強み・課題・top_issues の各文言が PPTX の中にあるか（判定68）
    def _norm(t):
        return re.sub(r'\s+', '', t or '')
    for key, src, items in [('C-1', d, [('強み', 'strengths'), ('課題', 'priority_issues')]),
                            ('C-3', v, [('強み', 'strengths'), ('課題', 'priority_issues'), ('top_issues', 'top_issues')])]:
        if key not in found:
            continue
        whole = _norm(''.join(p for sl in slide_texts(found[key]) for p in sl))
        for jp, field in items:
            for t in (src.get(field) or []):
                if isinstance(t, str) and t.strip() and _norm(t) not in whole:
                    ng.append(f'{key} {jp}『{t}』が PPTX に見当たらない → 入力を直したら PPTX も作り直す')
    # 7. C-3 の Before/After と行動の流れのまとめ（判定78）
    hls = v.get('highlights') or []
    c3_paras = [p for sl in slide_texts(found['C-3']) for p in sl] if 'C-3' in found else []
    c3_whole = _norm(''.join(c3_paras))
    for i, h in enumerate(hls[:3], start=1):
        if not isinstance(h, dict):
            ng.append(f'C-3 Before/After #{i} の形が違う'); continue
        ta = h.get('target_area')
        if not (isinstance(ta, str) and ta.strip()):
            if h.get('area'):
                ng.append(f'C-3 Before/After #{i}：箇所が area になっている → area ではなく target_area と書く（C-2 の対応する提案と同じ値）')
            else:
                ng.append(f'C-3 Before/After #{i}：箇所（target_area）が無い → C-2 の対応する提案の target_area を入れる')
        elif 'C-3' in found and _norm('箇所：' + ta[:20]) not in c3_whole:
            ng.append(f'C-3 Before/After #{i}：PPTX に「箇所：{ta}」が見当たらない → 入力を直したら PPTX も作り直す')
        if h.get('priority') not in ('高', '中', '低'):
            ng.append(f'C-3 Before/After #{i}：優先度（priority）が「{h.get("priority")}」→ 高・中・低のどれか（C-2 と同じ値）')
        if h.get('effort') not in ('小', '中', '大'):
            ng.append(f'C-3 Before/After #{i}：工数（effort）が「{h.get("effort")}」→ 小・中・大のどれか（C-2 と同じ値）')
    fs = v.get('flow_summary')
    if not (isinstance(fs, str) and fs.strip()):
        ng.append('C-3 行動の流れのまとめ（flow_summary）が無い・空 → 40字以内の1文を入れる')
    else:
        if len(fs) > 40:
            ng.append(f'C-3 行動の流れのまとめが{len(fs)}字（上限40字）→ 言い換える')
        elif 'C-3' in found and _norm('→ ' + fs) not in c3_whole:
            ng.append(f'C-3 の PPTX に行動の流れのまとめ「{fs}」が見当たらない → 入力を直したら PPTX も作り直す')
    return files, ng


def main(argv):
    if len(argv) != 2:
        print(__doc__); return 2
    data = json.load(open(argv[0], encoding='utf-8'))
    files, ng = check(data, argv[1])
    print('点検対象：', ', '.join(files) or '（なし）')
    if ng:
        print(f'🚨 不合格 {len(ng)}件'); [print('  -', m) for m in ng]; return 1
    print('⭕ 点検合格（総評17字・ランクの説明・件数・ファイル名・test_*.pptx 無し・リンク無し・メイリオ14pt以上・PPTX と入力の一致・Before/After の箇所と優先度・行動の流れのまとめ）'); return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
