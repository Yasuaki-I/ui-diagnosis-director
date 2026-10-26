# -*- coding: utf-8 -*-
# 【購入者限定】本ファイルの再配布・転載・転売は固く禁じます。｜UI診断ディレクター スキル版 prototype r6.2（2026-10-27｜設計 r2・判定113・117〜128）
"""check_output.py｜生成後の点検（r6）｜不合格なら AI が入力を直して作り直す

使い方:
    python3 check_output.py <入力JSON> <出力フォルダ>
        入力JSON = {"site_slug":..., "record":{...}, "diagnosis":{...}, "proposals":{...}, "visual_data":{...}}
    python3 check_output.py --pptx-only <出力フォルダ>
        入力の無い PPTX（旧版など）の表示だけを調べる（試験用）。内容・構成の点検は実施しない
    どちらも末尾に --report <JSONの保存先> を付けると、欄ごとの結果を JSON で保存する

結果は3段で表示する:
    1. 内容・構成点検（r5 の点検をすべて維持＋入力の A・B・C・文字種・record）
    2. 表示の簡易推定（幅・高さ・枠の対応・例外表・フッター・切り詰め）
       ※「基準内」は実際の描画の合格を意味しない
    3. Windows PowerPoint 目視：未実施（最終の判定は Windows の PowerPoint での目視）

終了コード: 0＝内容・構成が合格かつ 表示の推定で NG 0・未検査 0／1＝不合格（直す箇所を表示）
            2＝引数誤り／3＝--pptx-only で NG 0・未検査 0（内容は未実施のため合格ではない）

r5 の点検項目（一つも外さない）:
    1. 総評（visual_data.summary）：項目があり・空でなく・17字以内・末尾が「…」「...」でない
    2. ランクの説明（rank_label）：ランクに対応する1語（S＝卓越／A＝優良／B＝標準的／C＝要改善／D＝致命的）
    3. 件数：強み・課題とも各3件（ランクに関わらず・PPTX に載る件数）。C-1・C-3・top_issues で同じ
    4. ファイル名が英数字で3点そろう／test_*.pptx が無い
    5. PPTX にハイパーリンクが無い／メイリオ・14pt 以上
    6. PPTX の中の文字が入力と一致（C-3 の総評・C-1 のランク欄・強み・課題・top_issues の各文言）
    7. C-3 の Before/After：target_area・priority・effort と PPTX の「箇所：」／flow_summary 40字以内で PPTX にある
"""
import json, os, re, sys, glob, zipfile

ITEM_COUNT = 3  # 強み・課題の件数（builder の描画上限｜r5 から維持）
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




# =====================================================================
# ここから r6 の追加（設計 r2｜3891789c7f87・判定113）。r5 の点検は上の check() のまま
# =====================================================================
import math, unicodedata

AUTHOR = 'UI診断ディレクター'
EMU_PX = 9525.0            # 1280×720 座標（builder の px()）
TOL = 2                    # 枠の位置の許容差（px）

# ---- §2-2 固定記号の幅（固定記号の一覧を先に当てる）----
FIXED_W = {'✓': 1.0, '✕': 1.0, '⚠': 1.0, '→': 1.0, '◎': 1.0, '◤': 1.0, '●': 1.0,
           '？': 1.0, '\u2060': 0.0}


def char_class(ch):
    """§2-1 の文字種。戻り値 (幅 or None, NG理由 or None)"""
    o = ord(ch)
    if ch in '\n\r\t':
        return None, '改行・タブは使わない'
    if 0x20 <= o <= 0x7E:
        return 0.5, None
    if 0xFF61 <= o <= 0xFF9F:
        return None, '半角カナは全角に直す'
    if (0x0300 <= o <= 0x036F) or (0xFE00 <= o <= 0xFE0F) or (0xE0100 <= o <= 0xE01EF) or (0x200B <= o <= 0x200D):
        return None, '結合文字・異体字セレクタ・ゼロ幅文字は使わない'
    if (0xE000 <= o <= 0xF8FF) or o >= 0xF0000:
        return None, '私用領域の文字は使わない'
    if o >= 0x1F000:
        return None, '絵文字は使わない'
    if unicodedata.category(ch) in ('Cc', 'Cf', 'Cs', 'Co', 'Cn'):
        return None, '制御文字・書式文字は使わない'
    if ch == '\u3000':
        return 1.0, None
    eaw = unicodedata.east_asian_width(ch)
    if eaw in ('W', 'F', 'A'):
        return 1.0, None
    return None, f'未対応の文字種（U+{o:04X}）'


def disp_w(s):
    """入力文字列の表示幅（C の数え方）。(幅, NG理由のリスト)"""
    w, bad = 0.0, []
    for ch in s:
        cw, why = char_class(ch)
        if why:
            bad.append(f'「{ch}」(U+{ord(ch):04X})：{why}')
        else:
            w += cw
    return w, bad


def builder_w(s):
    """builder の _truncate_full の数え方（ord<128 を0.5・それ以外 1.0）"""
    return sum(0.5 if ord(c) < 128 else 1.0 for c in s)


def pptx_char_w(ch):
    """PPTX 上の文字の幅（固定記号を先に当て、次に文字種）。定義外は None"""
    if ch in FIXED_W:
        return FIXED_W[ch]
    cw, why = char_class(ch)
    return None if why else cw


# ---- §1-2 上限（A＝len／B＝builder 換算／C＝表示換算）----
LIM = {
    'service_name':  dict(A=30, C=11),
    'input_type':    dict(C=10),
    'comment':       dict(A=40, C=40),
    'strength':      dict(A=40, B=28),
    'issue_top':     dict(A=20, B=20, C=20),     # priority_issues・top_issues
    'conclusion':    dict(A=60, C=114),
    'p_title':       dict(A=30, B=36, C=32),
    'target_area':   dict(A=25, B=20, C=10),
    'p_issue':       dict(A=25, C=28),
    'before_after':  dict(A=40, B=38, C=24),
    'p_summary':     dict(A=120, B=120, C=114),
    'purpose':       dict(A=20, C=14),
    'diag_date':     dict(C=14),
    'v_summary':     dict(A=17, B=17, C=17),
    'sec_label':     dict(A=20, B=12, C=9),
    'sec_desc':      dict(B=16, C=10),
    'flow_label':    dict(B=6, C=6),
    'flow_note':     dict(B=14, C=9),
    'flow_summary':  dict(A=40, B=40, C=31),
    'hl_title':      dict(A=30, B=30, C=27),
    'direction':     dict(A=120, B=120, C=112),
    'top_issue':     dict(A=40, C=20),
}
LIM_NAME = {'A': 'len', 'B': 'builder換算', 'C': '表示換算'}


def check_input_r6(data):
    """入力の段階の点検（§1・§2・§5）。NG のリストを返す"""
    ng, notes = [], []
    d = data.get('diagnosis') or {}
    p = data.get('proposals') or {}
    v = data.get('visual_data') or {}

    def chk(label, s, key):
        if s is None:
            return
        if not isinstance(s, str):
            ng.append(f'{label}：文字列でない'); return
        w, bad = disp_w(s)
        for b in bad:
            ng.append(f'{label}：{b}')
        lim = LIM[key]
        if 'A' in lim and len(s) > lim['A']:
            ng.append(f'{label}：len {len(s)}（上限 {lim["A"]}）「{s}」→ 言い換える')
        if 'B' in lim and builder_w(s) > lim['B']:
            ng.append(f'{label}：builder 換算 {builder_w(s):g}（上限 {lim["B"]}）「{s}」→ 言い換える（builder が「…」で切るため）')
        if 'C' in lim and not bad and w > lim['C']:
            ng.append(f'{label}：表示上限超過 {w:g}（上限 {lim["C"]}）「{s}」→ 言い換える')
        return w

    for lbl, src in (('C-1', d), ('C-2', p), ('C-3', v)):
        chk(f'{lbl} service_name', src.get('service_name'), 'service_name')
    names = [src.get('service_name') for src in (d, p, v)]
    if len(set(names)) != 1:
        ng.append(f'service_name が3か所で一致しない：{names}')
    for lbl, src in (('C-1', d), ('C-3', v)):
        chk(f'{lbl} input_type', src.get('input_type'), 'input_type')
    for i, sc in enumerate(d.get('scores') or [], 1):
        if isinstance(sc, dict):
            chk(f'C-1 scores#{i} comment', sc.get('comment'), 'comment')
    for lbl, src in (('C-1', d), ('C-3', v)):
        for i, s in enumerate(src.get('strengths') or [], 1):
            chk(f'{lbl} 強み#{i}', s, 'strength')
        for i, s in enumerate(src.get('priority_issues') or [], 1):
            chk(f'{lbl} 課題#{i}', s, 'issue_top')
    for i, s in enumerate(v.get('top_issues') or [], 1):
        chk(f'C-3 top_issues#{i}', s, 'issue_top')
    chk('C-1 結論（conclusion）', d.get('conclusion'), 'conclusion')
    for i, pr in enumerate(p.get('proposals') or [], 1):
        if not isinstance(pr, dict):
            continue
        chk(f'C-2 提案#{i} title', pr.get('title'), 'p_title')
        chk(f'C-2 提案#{i} target_area', pr.get('target_area'), 'target_area')
        chk(f'C-2 提案#{i} issue', pr.get('issue'), 'p_issue')
        chk(f'C-2 提案#{i} before', pr.get('before'), 'before_after')
        chk(f'C-2 提案#{i} after', pr.get('after'), 'before_after')
    chk('C-2 summary（POINT）', p.get('summary'), 'p_summary')
    chk('C-3 purpose', v.get('purpose'), 'purpose')
    chk('C-3 target', v.get('target'), 'purpose')
    chk('C-3 diagnosis_date', v.get('diagnosis_date'), 'diag_date')
    chk('C-3 総評（summary）', v.get('summary'), 'v_summary')
    for i, sec in enumerate(v.get('sections') or [], 1):
        if not isinstance(sec, dict):
            continue
        a = chk(f'C-3 構造マップ#{i} label', sec.get('label'), 'sec_label')
        b = chk(f'C-3 構造マップ#{i} desc', sec.get('desc'), 'sec_desc')
        if a is not None and b is not None and a + b > 16:
            ng.append(f'C-3 構造マップ#{i}：見出し＋説明の表示換算 {a + b:g}（上限 16）→ 言い換える')
    for i, st in enumerate((v.get('flow_steps') or [])[:6], 1):
        if not isinstance(st, dict):
            continue
        chk(f'C-3 フロー#{i} label', st.get('label'), 'flow_label')
        chk(f'C-3 フロー#{i} note', st.get('note'), 'flow_note')
        comb = ((st.get('label') or '') + ' ' + (st.get('note') or '')).strip()
        if len(comb) > 25:
            ng.append(f'C-3 フロー#{i}：label＋note の len {len(comb)} 超過（上限 25）→ 言い換える')
    chk('C-3 flow_summary', v.get('flow_summary'), 'flow_summary')
    for i, h in enumerate((v.get('highlights') or [])[:3], 1):
        if not isinstance(h, dict):
            continue
        chk(f'C-3 ハイライト#{i} title', h.get('title'), 'hl_title')
        chk(f'C-3 ハイライト#{i} target_area', h.get('target_area'), 'target_area')
        chk(f'C-3 ハイライト#{i} before', h.get('before'), 'before_after')
        chk(f'C-3 ハイライト#{i} after', h.get('after'), 'before_after')
    chk('C-3 改善方向（direction）', v.get('direction'), 'direction')
    if v.get('top_issue') is not None:
        chk('C-3 top_issue', v.get('top_issue'), 'top_issue')

    # ---- 判定122-3：C-2 の提案と C-3 のハイライトの対応（箇所で一意に特定）----
    props = [x for x in (p.get('proposals') or []) if isinstance(x, dict)]
    used = {}
    for i, h in enumerate((v.get('highlights') or [])[:3], 1):
        if not isinstance(h, dict):
            continue
        cand = [q for q in props if q.get('target_area') == h.get('target_area')]
        if not cand:
            ng.append(f'C-3 ハイライト#{i}：対応する C-2 の提案が無い（target_area「{h.get("target_area")}」が一致する提案が無い）')
            continue
        if len(cand) > 1:
            ng.append(f'C-3 ハイライト#{i}：対応する C-2 の提案が複数ある（target_area「{h.get("target_area")}」）→ 箇所を提案ごとに分ける')
            continue
        q = cand[0]
        qn = q.get('no')
        if qn in used:
            ng.append(f'C-3 ハイライト#{i}：ハイライト#{used[qn]} と同じ提案（no {qn}）に対応している')
        used[qn] = i
        for k, jp in (('priority', '優先度'), ('effort', '工数'), ('before', '現状（before）'), ('after', '改善後（after）')):
            if h.get(k) != q.get(k):
                ng.append(f'C-3 ハイライト#{i} と C-2 提案 no {qn}：{jp}が一致しない（C-2「{q.get(k)}」／C-3「{h.get(k)}」）→ 同じ文字列にする')

    # ---- §5 record ----
    r = data.get('record')
    if not isinstance(r, dict):
        ng.append('record が無い・型が違う → input.json のいちばん上に record を置く')
    else:
        so, tu, io = r.get('service_name_official'), r.get('target_url', 'MISSING'), r.get('input_origin')
        if not (isinstance(so, str) and so.strip()):
            ng.append('record.service_name_official が無い・空 → 正式名称を入れる')
        if tu == 'MISSING' or not (tu is None or isinstance(tu, str)):
            ng.append('record.target_url が無い・型が違う（文字列か null）')
        if io not in ('url', 'screenshot'):
            ng.append(f'record.input_origin が「{io}」→ url か screenshot のどちらか')
        it = str(d.get('input_type') or '')
        if io == 'url':
            if not (isinstance(tu, str) and re.match(r'^https?://\S+$', tu)):
                ng.append('input_origin が url なのに target_url が空・http(s):// で始まらない')
        if io == 'screenshot' and 'スクショ' not in it:
            ng.append(f'input_origin が screenshot なのに input_type「{it}」に「スクショ」が無い')
        if ('URL' in it) != (io == 'url'):
            ng.append(f'input_type「{it}」と input_origin「{io}」が食い違う（URL を含む ⇔ url）')
        if isinstance(so, str) and so.strip() and names[0] and so != names[0]:
            notes.append(f'表示名「{names[0]}」（正式名称：{so}）')
    return ng, notes


# ---- PPTX の読み出し ----
NSA = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _emu_px(v):
    return int(v) / EMU_PX


def read_frames(path):
    """スライドごとに、文字枠（表のセルを含む）を図形の順に返す"""
    from pptx import Presentation
    prs = Presentation(path)
    slides = []
    for sl in prs.slides:
        frames = []
        for sh in sl.shapes:
            box = (sh.left / EMU_PX, sh.top / EMU_PX, sh.width / EMU_PX, sh.height / EMU_PX) \
                if sh.width is not None else None
            if sh.has_text_frame:
                frames.append(dict(kind='text', box=box, body=sh.text_frame._txBody,
                                   text=sh.text_frame.text))
            if getattr(sh, 'has_table', False) and sh.has_table:
                tbl = sh.table
                x = box[0]
                colw = [c.width / EMU_PX for c in tbl.columns]
                y = box[1]
                for ri, row in enumerate(tbl.rows):
                    rh = row.height / EMU_PX
                    cx = x
                    for ci, cell in enumerate(row.cells):
                        tc = cell._tc
                        pr = tc.find(NSA + 'tcPr')
                        mar = {k: (int(pr.get(k)) / EMU_PX if pr is not None and pr.get(k) else d)
                               for k, d in (('marL', 91440 / EMU_PX), ('marR', 91440 / EMU_PX),
                                            ('marT', 45720 / EMU_PX), ('marB', 45720 / EMU_PX))}
                        frames.append(dict(kind='cell', box=(cx, y, colw[ci], rh), body=cell.text_frame._txBody,
                                           text=cell.text_frame.text, row=ri, col=ci, mar=mar,
                                           table_box=box))
                        cx += colw[ci]
                    y += rh
        slides.append(frames)
    return slides


def paragraphs(body):
    """段落ごとに runs[(text, size_pt or None, bold, spc)]・改行で区切った segs・行間・前後間隔
    spc＝文字間隔（1/100 pt・各文字の後ろに加える）。継承元に文字間隔があれば spc_unknown"""
    out = []
    lst = body.find(NSA + 'lstStyle')
    inh = lst is not None and any(el.get('spc') not in (None, '0') for el in lst.iter())
    for p in body.findall(NSA + 'p'):
        runs, segs, has_br, spc_unknown = [], [[]], False, inh
        ppr = p.find(NSA + 'pPr')
        if ppr is not None:
            dr = ppr.find(NSA + 'defRPr')
            if dr is not None and dr.get('spc') not in (None, '0'):
                spc_unknown = True
        for el in p:
            if el.tag == NSA + 'r':
                rpr = el.find(NSA + 'rPr')
                sz = rpr.get('sz') if rpr is not None else None
                sp = rpr.get('spc') if rpr is not None else None
                t = el.find(NSA + 't')
                run = ((t.text or '') if t is not None else '', int(sz) / 100 if sz else None,
                       rpr is not None and rpr.get('b') == '1', int(sp) if sp else 0)
                runs.append(run); segs[-1].append(run)
            elif el.tag == NSA + 'br':
                has_br = True
                segs.append([])
        ln = sa = sb = None
        if ppr is not None:
            x = ppr.find(NSA + 'lnSpc')
            if x is not None and len(x):
                c = x[0]
                ln = ('pct', int(c.get('val')) / 100000) if c.tag.endswith('spcPct') else ('pts', int(c.get('val')) / 100)
            for tag, key in (('spcAft', 'sa'), ('spcBef', 'sb')):
                x = ppr.find(NSA + tag)
                if x is not None and len(x):
                    c = x[0]
                    val = ('pct', int(c.get('val')) / 100000) if c.tag.endswith('spcPct') else ('pts', int(c.get('val')) / 100)
                    if key == 'sa':
                        sa = val
                    else:
                        sb = val
        endp = p.find(NSA + 'endParaRPr')
        end_sz = int(endp.get('sz')) / 100 if endp is not None and endp.get('sz') else None
        out.append(dict(runs=runs, segs=segs, ln=ln, sa=sa, sb=sb, br=has_br, end_sz=end_sz,
                        spc_unknown=spc_unknown, text=''.join(r[0] for r in runs)))
    return out


def insets(fr):
    if fr['kind'] == 'cell':
        m = fr['mar']; return m['marL'], m['marR'], m['marT'], m['marB']
    bp = fr['body'].find(NSA + 'bodyPr')
    g = lambda k, d: (int(bp.get(k)) if bp is not None and bp.get(k) is not None else d) / EMU_PX
    return g('lIns', 91440), g('rIns', 91440), g('tIns', 45720), g('bIns', 45720)


def _runs_width(runs):
    """文字区分の並びの幅（px）。文字間隔 spc（1/100 pt）を各文字の後ろに加える"""
    w, undefined, nosize, sizes = 0.0, [], False, set()
    for t, sz, b, spc in runs:
        if not t:
            continue
        if sz is None:
            nosize = True; continue
        sizes.add(sz)
        for ch in t:
            cw = pptx_char_w(ch)
            if cw is None:
                undefined.append(ch)
            else:
                w += sz * 4 / 3 * cw + spc / 100 * 4 / 3
    return w, undefined, nosize, sizes


def para_width(par):
    """段落の幅（px・改行を無視した合計）。(幅, 定義外の文字, 大きさ不明, 最大の大きさ pt, 大きさ混在)"""
    w, undefined, nosize, sizes = _runs_width(par['runs'])
    mx = max(sizes) if sizes else par.get('end_sz')
    return w, undefined, nosize, mx, len(sizes) > 1


def seg_widths(par, default_pt):
    """明示改行で区切った行ごとの (幅px, 行の最大の大きさ pt)。空の行は幅0・段落の大きさ"""
    out = []
    for seg in par['segs']:
        w, _, _, sizes = _runs_width(seg)
        out.append((w, max(sizes) if sizes else default_pt))
    return out


def _spacing_px(val, size_pt):
    if not val:
        return 0.0
    kind, x = val
    return x * 4 / 3 if kind == 'pts' else size_pt * 4 / 3 * x


def long_ascii_word(text, cap_chars):
    for m in re.finditer(r'[!-~]+', text):
        if len(m.group(0)) * 0.5 >= cap_chars / 3:
            return m.group(0)
    return None


# ---- §3・§4 表示の簡易推定（r6.1〜r6.2｜判定122-1・122-2・122-4）----
# 期待する枠の表（builder 17.2.17 の座標と描画規則から作る）と、実際の枠を照合する。
# 欠落・重複・多重一致・想定外の枠・固定文言の変更を、それぞれ理由を示して止める
FILE_SLIDES = {'C-1': (2,), 'C-2': (1, 2), 'C-3': (3,)}
TEN_ITEMS = ('ファーストビュー', 'キャッチコピー', 'CTA設計', '信頼性・権威性', 'フォーム設計',
             'レスポンシブ', '読みやすさ', '情報設計', 'ブランド一貫性', '表示速度・技術')
RANK_WORDS = '卓越|優良|標準的|要改善|致命的'


def _tf(s, w):
    """builder の _truncate_full と同じ切り方（期待する文字を作るため）"""
    t = 0.0
    for i, ch in enumerate(s):
        t += 0.5 if ord(ch) < 128 else 1.0
        if t > w:
            return s[:i].rstrip() + '…'
    return s


def E(fid, slide, box, cls, kind, exp, **kw):
    """期待する枠1つ。exp＝('eq', 文字) か ('re', 正規表現)"""
    return dict(fid=fid, slide=slide, box=box, cls=cls, kind=kind, exp=exp, **kw)


def _eq(s):
    return ('eq', s)


def _re(p):
    return ('re', p)


def _common(key, slide, total, title, sub, title_w):
    t = key[-1]
    return [E(f'{t}-{slide}-題名', slide, (60, 20, title_w, 34), 'fixed', 'line', _eq(title)),
            E(f'{t}-{slide}-副題', slide, (60, 58, title_w, 22), 'fixed', 'line', sub),
            E(f'{t}-{slide}-フッター', slide, (40, 682, 400, 24), 'footer', 'line', _eq(AUTHOR)),
            E(f'{t}-{slide}-ページ番号', slide, (1140, 685, 120, 24), 'fixed', 'line', _eq(f'{slide} / {total}'))]


def expect_c1(d):
    """C-1：2枚。d＝diagnosis（None なら PPTX だけの点検用に、値は形で照合）"""
    X = []
    subs = {1: 'サマリ：診断対象・スコア・強み・最優先課題・結論', 2: '10項目評価：各項目のスコアと一言所見（全文表示）'}
    for s in (1, 2):
        X += _common('C-1', s, 2, '【1】UI診断スコアカード', _eq(subs[s]), 1140)
        if d:
            c = [_eq('診断対象\n' + str(d.get('service_name', ''))), _eq('入力種別\n' + str(d.get('input_type', ''))),
                 _eq(f'総合スコア\n{d.get("total_score")}  / 50'), _eq(f'ランク\n{d.get("rank")}  {d.get("rank_label") or ""}')]
        else:
            c = [_re(r'診断対象\n[^\n]+'), _re(r'入力種別\n[^\n]+'), _re(r'総合スコア\n\d{1,2}  / 50'),
                 _re(rf'ランク\n[SABCD]  ({RANK_WORDS})')]
        X += [E(f'1-{s}-診断対象', s, (58, 100, 265, 66), 'input', 'card', c[0]),
              E(f'1-{s}-入力種別', s, (361, 100, 265, 66), 'input', 'card', c[1]),
              E(f'1-{s}-総合スコア', s, (664, 100, 265, 66), 'fixed', 'card', c[2], card=('C-1', '総合スコア')),
              E(f'1-{s}-ランク', s, (967, 100, 265, 66), 'fixed', 'card', c[3], card=('C-1', 'ランク'))]
    if d:
        st = _eq('\n'.join('● ' + x for x in (d.get('strengths') or [])[:3]))
        iss = _eq('\n'.join(f'{i + 1}. ' + x for i, x in enumerate((d.get('priority_issues') or [])[:3])))
        con = _eq(_tf(str(d.get('conclusion', '')).replace('\n', ' '), 120))
    else:
        st, iss, con = _re(r'(● [^\n]+)(\n● [^\n]+){0,2}'), _re(r'(1\. [^\n]+)(\n\d\. [^\n]+){0,2}'), _re(r'[^\n]+')
    X += [E('1-1-強み見出し', 1, (58, 190, 556, 36), 'fixed', 'line', _eq('◎ 強み（活かすべき点）')),
          E('1-1-強み', 1, (60, 244, 552, 328), 'input', 'multi', st),
          E('1-1-課題見出し', 1, (666, 190, 556, 36), 'fixed', 'line', _eq('⚠ 最優先で直すべき点')),
          E('1-1-課題', 1, (668, 244, 552, 328), 'input', 'multi_line', iss),
          E('1-1-結論ラベル', 1, (40, 594, 88, 62), 'fixed', 'line', _eq('結論')),
          E('1-1-結論', 1, (168, 594, 1072, 62), 'input', 'multi', con)]
    # 2枚目の表（11行×4列）。表そのものの位置も照合する
    rows = (d.get('scores') or [])[:10] if d else None
    n = len(rows) if rows is not None else 10
    X.append(E('1-2-表', 2, (40, 180, 1200, None), 'table', 'table', None, rows=n + 1, cols=4))
    for c, h in enumerate(('No', '項目', 'スコア', '一言所見')):
        X.append(E(f'1-2-表-見出し{c + 1}', 2, None, 'fixed', 'line', _eq(h), cell=(0, c)))
    for r in range(1, n + 1):
        it = rows[r - 1] if rows else None
        X += [E(f'1-2-表-{r}-No', 2, None, 'fixed', 'line', _eq(str(r)), cell=(r, 0)),
              E(f'1-2-表-{r}-項目', 2, None, 'fixed', 'line',
                _eq(_tf(it.get('name') or it.get('category', ''), 10)) if it else _re('|'.join(map(re.escape, TEN_ITEMS))), cell=(r, 1)),
              E(f'1-2-表-{r}-スコア', 2, None, 'fixed', 'line',
                _eq(f'{it.get("score", 0)}/{it.get("max", 5)}') if it else _re(r'(\d|－)/5'), cell=(r, 2)),
              E(f'1-2-表-{r}-所見', 2, None, 'input', 'multi',
                _eq(str(it.get('comment', ''))) if it else _re(r'[^\n]*'), cell=(r, 3))]
    return X


PL = {'S': '高', 'A': '高', 'B': '中', 'C': '低', '高': '高', '中': '中', '低': '低'}


def expect_c2(p, n):
    """C-2：提案3件以下なら1枚、4〜5件なら2枚（builder の add_proposal_onepager の規則）"""
    props = (p.get('proposals') or []) if p else None
    split = n >= 4
    total = 2 if split else 1
    pages = [(1, list(range(3))), (2, list(range(3, n)))] if split else [(1, list(range(n)))]
    card_h = (558 - 20) // 3 if split else (486 - 10 * (n - 1)) // max(1, n)
    X = []
    for s, idx in pages:
        if split:
            sub = '優先度の高い改善アクションを整理（1/2 前半 3件）' if s == 1 else f'優先度の高い改善アクションを整理（2/2 後半 {n - 3}件・POINT）'
        else:
            sub = f'優先度の高い改善アクションを整理（全{n}件）'
        X += _common('C-2', s, total, '【2】改善提案リスト', _eq(sub), 820)
        X += [E(f'2-{s}-凡例', s, (880, 30, 80, 22), 'fixed', 'line', _eq('優先度')),
              E(f'2-{s}-凡例高', s, (970, 33, 70, 22), 'fixed', 'line', _eq('高')),
              E(f'2-{s}-凡例中', s, (1050, 33, 70, 22), 'fixed', 'line', _eq('中')),
              E(f'2-{s}-凡例低', s, (1130, 33, 70, 22), 'fixed', 'line', _eq('低'))]
        for j, k in enumerate(idx):
            y = 100 + j * (card_h + 10); ht = y + 10; bt = ht + 44; bb = bt + 52
            q = props[k] if props else None
            g = lambda f, pre='': _eq(pre + str(q.get(f, ''))) if q else _re(re.escape(pre) + r'[^\n]*')
            X += [E(f'2-提案{k + 1}-番号', s, (58, ht + 4, 34, 26), 'fixed', 'line', _eq(str(q.get('no', ''))) if q else _re(r'\d')),
                  E(f'2-提案{k + 1}-title', s, (102, ht + 2, 770, 32), 'input', 'line',
                    _eq(_tf(q.get('title', ''), 36)) if q else _re(r'[^\n]+')),
                  E(f'2-提案{k + 1}-優先度', s, (980, ht + 1, 70, 30), 'fixed', 'line',
                    _eq(PL.get(q.get('priority', 'B'), '中')) if q else _re(r'[高中低]'), pid=k),
                  E(f'2-提案{k + 1}-工数', s, (1060, ht + 5, 170, 22), 'fixed', 'line',
                    _eq('工数：' + str(q.get('effort', '中'))) if q else _re(r'工数：[小中大]'), pid=k),
                  E(f'2-提案{k + 1}-箇所', s, (56, bt, 568, 48), 'input', 'line', g('target_area', '改善箇所\u3000'), pid=k),
                  E(f'2-提案{k + 1}-課題', s, (640, bt, 584, 48), 'input', 'line', g('issue', '課題\u3000')),
                  E(f'2-提案{k + 1}-現状', s, (64, bb + 4, 552, 40), 'input', 'line', g('before', '現状\u3000'), pid=k),
                  E(f'2-提案{k + 1}-改善後', s, (648, bb + 4, 568, 40), 'input', 'line', g('after', '改善後\u3000'), pid=k)]
    if p:
        sm = _tf(str(p.get('summary', '')).replace('\n', ' '), 120)
        if sm.endswith('。'):
            sm = sm[:-1] + '\u2060。'
        es = _eq(sm)
    else:
        es = _re(r'[^\n]+')
    X += [E('2-POINTラベル', total, (40, 594, 90, 62), 'fixed', 'line', _eq('POINT')),
          E('2-POINT', total, (170, 594, 1070, 62), 'input', 'multi', es)]
    return X, total


def _hl_no(h, index, top_issues):
    """判定126-1：builder 17.2.17 の _c3_resolve_issue_no と同じ規則で期待番号を作る
    ①no の明示を優先 ②未指定なら 1/3 の最重要課題の文と照合した連番 ③照合できなければ出現順"""
    explicit = h.get('no')
    if explicit not in (None, ''):
        return str(explicit)
    def _norm(v):
        return ''.join(ch for ch in str(v)
                       if not ch.isspace() and ch not in '。、．，,.:：;；・/／-ー―（）()「」【】')
    key = _norm(h.get('issue_ref') or h.get('title') or '')
    if key:
        for n, issue in enumerate(top_issues or [], start=1):
            ikey = _norm(issue)
            if not ikey:
                continue
            if key == ikey or key in ikey or ikey in key:
                return str(n)
    return str(index + 1)


def expect_c3(v, n_sec, n_flow, n_issues, n_h, n_sc):
    X = []
    subs = {1: '構造マップ・総評・最重要課題・ユーザー行動フロー', 2: 'スコア視覚化：10項目評価と強み・最優先課題の俯瞰',
            3: 'Before/After ハイライト：優先度の高い改善Top3'}
    for s in (1, 2, 3):
        X += _common('C-3', s, 3, '【3】ビジュアル診断ボード', _eq(subs[s]), 1140)
        if s in (1, 2):
            if v:
                m = [_eq('目的\n' + str(v.get('purpose', '-'))), _eq('対象\n' + str(v.get('target', '-'))),
                     _eq('診断日\n' + str(v.get('diagnosis_date', '-'))),
                     _eq(f'総合スコア\n{v.get("total_score", 0)}  / 50    {v.get("rank", "-")}')]
            else:
                m = [_re(r'目的\n[^\n]+'), _re(r'対象\n[^\n]+'), _re(r'診断日\n[^\n]+'), _re(r'総合スコア\n\d{1,2}  / 50    [SABCD]')]
            X += [E(f'3-{s}-目的', s, (58, 98, 265, 50), 'input', 'card', m[0]),
                  E(f'3-{s}-対象', s, (361, 98, 265, 50), 'input', 'card', m[1]),
                  E(f'3-{s}-診断日', s, (664, 98, 265, 50), 'input', 'card', m[2]),
                  E(f'3-{s}-総合スコア', s, (967, 98, 265, 50), 'fixed', 'card', m[3], card=('C-3', '総合スコア'))]
        dr = _eq(_tf(str(v.get('direction', '')).replace('\n', ' '), 120)) if v else _re(r'[^\n]+')
        X += [E(f'3-{s}-改善方向ラベル', s, (40, 594, 110, 62), 'fixed', 'line', _eq('改善方向')),
              E(f'3-{s}-改善方向', s, (190, 594, 1050, 62), 'input', 'multi', dr)]
    # 1枚目
    X += [E('3-1-構造マップ見出し', 1, (56, 158, 220, 28), 'fixed', 'line', _eq('◤ LP構造マップ')),
          E('3-1-凡例良好', 1, (288, 164, 40, 16), 'fixed', 'line', _eq('良好')),
          E('3-1-凡例要改善', 1, (358, 164, 60, 16), 'fixed', 'line', _eq('要改善'))]
    row_h = max(26, 248 // max(1, n_sec))
    secs = (v.get('sections') or []) if v else None
    for i in range(n_sec):
        y = 190 + i * row_h
        sec = secs[i] if secs else None
        if sec is not None:
            nm = _tf(sec.get('label') or sec.get('name', ''), 12); ds = _tf(sec.get('desc', ''), 16)
            e, en = _eq(nm + (f'  {ds}' if ds else '')), _eq(str(sec.get('no', i + 1)))
        else:
            e, en = _re(r'[^\n]+'), _re(r'\d')
        X += [E(f'3-1-構造マップ{i + 1}-番号', 1, (48, y, 22, 22), 'fixed', 'line', en),
              E(f'3-1-構造マップ{i + 1}', 1, (78, y, 332, row_h - 2), 'input', 'map', e)]
    if v:
        ti = (v.get('top_issues') or ([v.get('top_issue')] if v.get('top_issue') else []) or v.get('priority_issues', []))[:4]
        ei = _eq('\n'.join(f'{i + 1}. ' + _tf(x, 20) for i, x in enumerate(ti)))
        es = _eq('\n'.join('● ' + _tf(x, 28) for x in (v.get('strengths') or [])[:3]))
        es2 = _eq('\n'.join('● ' + _tf(x, 40) for x in (v.get('strengths') or [])[:3]))
        ei2 = _eq('\n'.join(f'{i + 1}. ' + _tf(x, 50) for i, x in enumerate((v.get('priority_issues') or [])[:3])))
        ef = _eq('→ ' + _tf(v.get('flow_summary', ''), 40))
        sm = _eq(_tf(v.get('summary', ''), 17))
    else:
        ei = ei2 = _re(r'(1\. [^\n]+)(\n\d\. [^\n]+){0,3}'); es = es2 = _re(r'(● [^\n]+)(\n● [^\n]+){0,2}')
        ef, sm = _re(r'→ [^\n]+'), _re(r'[^\n]+')
    X += [E('3-1-総評見出し', 1, (448, 158, 428, 28), 'fixed', 'line', _eq('◤ 総評')),
          E('3-1-総評', 1, (446, 186, 432, 40), 'input', 'line', sm),
          E('3-1-課題見出し', 1, (448, 238, 428, 28), 'fixed', 'line', _eq(f'⚠ 最重要課題（Top{n_issues}）')),
          E('3-1-最重要課題', 1, (446, 272, 432, 158), 'input', 'multi_line', ei),
          E('3-1-強み見出し', 1, (920, 158, 304, 28), 'fixed', 'line', _eq('◎ 強み（活かすべき点）')),
          E('3-1-強み', 1, (920, 194, 304, 232), 'input', 'multi', es),
          E('3-1-フロー見出し', 1, (56, 444, 500, 28), 'fixed', 'line', _eq('◤ ユーザー行動フローと課題')),
          E('3-1-フローまとめ', 1, (560, 444, 624, 28), 'input', 'line', ef)]
    if n_flow:
        step_w = (1200 - 20 - 8 * (n_flow - 1) - 60) // n_flow
        arrow_w = 60 // max(1, n_flow - 1) if n_flow > 1 else 0
        steps = (v.get('flow_steps') or [])[:6] if v else None
        for i in range(n_flow):
            x = 50 + i * (step_w + 8 + arrow_w)
            stp = steps[i] if steps else None
            if stp:
                st = stp.get('status', '✓')
                st = '✓' if st in ('✓', 'OK', True, 'success', '成功') else ('✕' if st in ('✕', '×', 'NG', False, 'fail', '離脱') else '？')
                e, est = _eq(_tf(stp.get('label', ''), 6) + '\n' + _tf(stp.get('note', ''), 14)), _eq(st)
            else:
                e, est = _re(r'[^\n]+\n[^\n]+'), _re(r'[✓✕？]')
            X += [E(f'3-1-フロー{i + 1}-記号', 1, (x + step_w // 2 - 22, 480, 44, 44), 'fixed', 'line', est),
                  E(f'3-1-フロー{i + 1}', 1, (x, 528, step_w, 56), 'input', 'flow', e)]
            if i < n_flow - 1:
                X.append(E(f'3-1-矢印{i + 1}', 1, (x + step_w + 8, 488, arrow_w, 44), 'exception', 'arrow', _eq('→')))
    # 2枚目
    X.append(E('3-2-スコア見出し', 2, (56, 158, 548, 28), 'fixed', 'line', _eq('◤ 10項目スコア（満点5・ギャップ可視化）')))
    scs = (v.get('scores') or [])[:10] if v else None
    rh = min(42, 400 // max(1, n_sc))
    for i in range(n_sc):
        y = 192 + i * rh
        it = scs[i] if scs else None
        X += [E(f'3-2-項目{i + 1}', 2, (56, y + 8, 180, 22), 'fixed', 'line',
                _eq(it.get('name') or it.get('category', '')) if it else _re('|'.join(map(re.escape, TEN_ITEMS)))),
              E(f'3-2-点{i + 1}', 2, (408, y + 8, 62, 22), 'fixed', 'line',
                _eq('－/5' if it.get('score', '－') in ('－', '－/5') else f'{it.get("score")}/5') if it else _re(r'(\d|－)/5'))]
    X += [E('3-2-強み見出し', 2, (656, 158, 568, 28), 'fixed', 'line', _eq('◎ 強み（活かすべき点）')),
          E('3-2-強み', 2, (656, 194, 568, 174), 'input', 'multi', es2),
          E('3-2-課題見出し', 2, (656, 388, 568, 28), 'fixed', 'line', _eq('⚠ 最優先で直すべき点')),
          E('3-2-課題', 2, (656, 424, 568, 154), 'input', 'multi_line', ei2)]
    # 3枚目
    hls = (v.get('highlights') or [])[:3] if v else None
    hl_h = (494 - 12 * max(0, n_h - 1)) // max(1, n_h)
    for i in range(n_h):
        y = 92 + i * (hl_h + 12); ht = y + 10; bt = ht + 50
        h = hls[i] if hls else None
        if h is not None:
            _ti = (v.get('top_issues') or ([v.get('top_issue')] if v.get('top_issue') else []) or v.get('priority_issues', []))
            hn = _eq(_hl_no(h, i, list(_ti)))
        else:
            hn = _re(r'\d{1,2}')     # 旧PPTX（入力なし）：形だけ確認。入力由来の番号との一致は未実施
        X += [E(f'3-3-HL{i + 1}-番号', 3, (60, ht + 4, 32, 24), 'fixed', 'line', hn),
              E(f'3-3-HL{i + 1}-title', 3, (102, ht + 4, 660, 28), 'input', 'line', _eq(_tf(h.get('title', ''), 30)) if h else _re(r'[^\n]+')),
              E(f'3-3-HL{i + 1}-優先度', 3, (780, ht + 5, 70, 22), 'fixed', 'line', _eq(str(h.get('priority', '中'))) if h else _re(r'[高中低]'), hid=i),
              E(f'3-3-HL{i + 1}-工数', 3, (860, ht + 5, 90, 22), 'fixed', 'line', _eq('工数：' + str(h.get('effort', '中'))) if h else _re(r'工数：[小中大]'), hid=i),
              E(f'3-3-HL{i + 1}-箇所', 3, (960, ht + 5, 260, 22), 'input', 'line', _eq('箇所：' + _tf(h.get('target_area', ''), 20)) if h else _re(r'箇所：[^\n]+'), hid=i),
              E(f'3-3-HL{i + 1}-BEFOREラベル', 3, (60, bt, 90, 28), 'fixed', 'line', _eq('BEFORE')),
              E(f'3-3-HL{i + 1}-現状', 3, (158, bt, 462, 28), 'input', 'line', _eq(_tf(h.get('before', ''), 38)) if h else _re(r'[^\n]+'), hid=i),
              E(f'3-3-HL{i + 1}-AFTERラベル', 3, (644, bt, 90, 28), 'fixed', 'line', _eq('AFTER')),
              E(f'3-3-HL{i + 1}-改善後', 3, (742, bt, 462, 28), 'input', 'line', _eq(_tf(h.get('after', ''), 38)) if h else _re(r'[^\n]+'), hid=i)]
        if h is not None and str(h.get('next_action', '') or '').strip():
            X.append(E(f'3-3-HL{i + 1}-次の一手ラベル', 3, (60, bt + 38, 90, 28), 'fixed', 'line', _eq('次の一手')))
            X.append(E(f'3-3-HL{i + 1}-次の一手', 3, (158, bt + 38, 1072, 28), 'input', 'line', _eq(_tf(str(h.get('next_action')).strip(), 38))))
    return X


# ---- 判定118：固定カードの高さの限定例外（3種類・6出現）----
# 採用 builder 17.2.17（47629772eaff）の _scorecard_info_cards／_visual_meta_header の構造と一致するときだけ、
# 高さ推定の超過を「要実機確認」へ移す。幅・値・改行・大きさ・枠・段落構造は免除しない
BUILDER_MD5_12 = '47629772eaff'
FIXED_CARD = {
    # (ファイル, 欄名): (枠 x,y,w,h, 見出し, 2段落目の区分[(大きさ,太字,文字の型)], 行間[段落1,段落2])
    ('C-1', '総合スコア'): ((664, 100, 265, 66), '総合スコア',
                         [(32.0, True, r'\d{1,2}'), (14.0, False, r'  / 50')], [1.0, 1.0]),
    ('C-1', 'ランク'): ((967, 100, 265, 66), 'ランク',
                      [(32.0, True, r'[SABCD]'), (14.0, False, r'  (卓越|優良|標準的|要改善|致命的)')], [1.0, 1.0]),
    ('C-3', '総合スコア'): ((967, 98, 265, 50), '総合スコア',
                         [(26.0, True, r'\d{1,2}'), (14.0, False, r'  / 50    '), (26.0, True, r'[SABCD]')], [1.0, 1.0]),
}


def fixed_card_match(key, field, sno, fr, data=None):
    """判定118 の適用条件を満たすか。(True/False, 理由)"""
    spec = FIXED_CARD.get((key, field))
    if spec is None:
        return False, '対象外'
    allowed = {'C-1': (1, 2), 'C-3': (1, 2)}[key]
    if sno not in allowed:
        return False, 'スライドが対象外'
    box, label, segs, lns = spec
    if not all(abs(round(a) - b) <= TOL for a, b in zip(fr['box'], box)):
        return False, '枠の位置・寸法が builder と違う'
    if tuple(round(v, 3) for v in insets(fr)) != (0.0, 0.0, 0.0, 0.0):
        return False, '枠の内側の余白が builder（0）と違う'
    pars = paragraphs(fr['body'])
    if len(pars) != 2:
        return False, f'段落数が {len(pars)}（2であるべき）'
    p1, p2 = pars
    if [(r[0], r[1], r[2], r[3]) for r in p1['runs']] != [(label, 14.0, True, 0)] or p1['spc_unknown'] or p2['spc_unknown']:
        return False, '見出しの文字・大きさ・太字が builder と違う'
    if len(p2['runs']) != len(segs):
        return False, '値の文字区分の数が builder と違う'
    for (t, sz, b, spc), (esz, eb, rx) in zip(p2['runs'], segs):
        if sz != esz or bool(b) != eb or spc or not re.fullmatch(rx, t):
            return False, f'値の区分「{t}」の大きさ・太字・形が builder と違う'
    for par, ln in zip(pars, lns):
        if par['ln'] != ('pct', ln) or par['br'] or par['sa'] or par['sb']:
            return False, '行間・段落の間隔・改行が builder と違う'
    if data:
        src = (data.get('diagnosis') if key == 'C-1' else data.get('visual_data')) or {}
        if field == '総合スコア' and p2['runs'][0][0] != str(src.get('total_score')):
            return False, '総合点が入力と一致しない'
        if key == 'C-1' and field == 'ランク':
            if p2['runs'][0][0] != str(src.get('rank')) or p2['runs'][1][0] != '  ' + str(src.get('rank_label')):
                return False, 'ランク・説明が入力と一致しない'
        if key == 'C-3' and p2['runs'][2][0] != str(src.get('rank')):
            return False, 'ランクが入力と一致しない'
    return True, ''


def _line_h(par, size_pt):
    ln = par['ln']
    if ln is None:
        return size_pt * 4 / 3 * 1.2, True
    kind, x = ln
    return (size_pt * 4 / 3 * 1.2 * x if kind == 'pct' else x * 4 / 3), False


def estimate(fr, kind):
    """1枠の推定。戻り値 dict(status, why, detail)"""
    l, r, t, b = insets(fr)
    x, y, w, h = fr['box']
    if not w or not h:
        return dict(status='未検査', why='枠の幅・高さが取れない')
    cw = w - l - r
    ch = h - t - b
    pars = paragraphs(fr['body'])
    total, total_aux, assumed, notes = 0.0, 0.0, False, []
    width_ng = []
    for i, par in enumerate(pars):
        pw, undefined, nosize, mx, mixed = para_width(par)
        if nosize:
            return dict(status='未検査', why=f'段落{i+1}の文字の大きさが分からない')
        if undefined:
            return dict(status='未検査', why=f'定義されていない文字種：{"".join(sorted(set(undefined)))}')
        if mx is None:
            mx = 14.0
        if mixed:
            notes.append('大きさ混在')
        lh, asm = _line_h(par, mx)
        assumed |= asm
        one_line = kind in ('line', 'map', 'multi_line', 'multi_cell_line') or \
            (kind == 'card' and i == 1) or kind == 'flow'
        if par['br']:
            if one_line:
                width_ng.append(f'段落{i+1}に明示の改行'); 
            notes.append('明示改行あり')
        lines = max(1, -(-pw // cw)) if pw > 0 else 1
        aux = max(1, -(-pw // max(1.0, cw - mx * 4 / 3))) if pw > 0 else 1
        if one_line and pw > cw + 1e-6:
            width_ng.append(f'段落{i+1}の幅 {pw:.1f}px > 枠 {cw:.1f}px「{par["text"]}」')
        lw = long_ascii_word(par['text'], cw / (mx * 4 / 3))
        if lw and not one_line:
            notes.append(f'長い ASCII 単語「{lw}」')
        sp = _spacing_px(par['sb'], mx) + (_spacing_px(par['sa'], mx) if i < len(pars) - 1 else 0)
        total += lines * lh + sp
        total_aux += aux * lh + sp
    res = dict(text=fr['text'], height=round(total, 2), cap=round(ch, 2), notes=notes)
    if width_ng:
        res.update(status='NG', why='幅 NG：' + '／'.join(width_ng))
        if total > ch + 1e-6:
            res['why'] += f'／同じ枠の高さ超過 {total:.1f}px > {ch:.1f}px'
            res['height_ng'] = True
        return res
    # §3-2 の対象＝複数行を許す欄と、複数の段落がある枠（1段落・1行の枠は幅で点検）
    height_scope = kind in ('multi', 'multi_line', 'card', 'flow') or len(pars) >= 2
    if height_scope and total > ch + 1e-6:
        res.update(status='NG', why=f'高さ超過 {total:.1f}px > 枠 {ch:.1f}px', height_ng=True)
        return res
    if not height_scope:
        total_aux = 0.0      # 1段落・1行の枠は幅で点検する（高さ・補助の判定の対象外）
    reasons = []
    if total_aux > ch + 1e-6:
        reasons.append(f'折り返しの余裕不足（補助の判定 {total_aux:.1f}px > {ch:.1f}px）')
    if assumed and height_scope:
        notes.append('行間：仮定1.0')
        if (ch - total) / ch < 0.10:
            reasons.append('行間を仮定し高さの余裕10%未満')
    if any(n.startswith('長い ASCII') for n in notes):
        reasons.append('長い ASCII 単語')
    if reasons:
        res.update(status='要実機確認', why='／'.join(reasons))
    else:
        res.update(status='基準内', why='')
    return res


def _box_ok(fbox, ebox):
    return all(b is None or abs(round(a) - b) <= TOL for a, b in zip(fbox, ebox))


def _text_ok(text, exp):
    kind, val = exp
    return text == val if kind == 'eq' else re.fullmatch(val, text) is not None


# ---- 判定122-2：矢印の例外は条件をすべて満たすときだけ ----
def arrow_match(fr, e):
    """例外表 #1 の全条件：対象スライド・期待の位置（フローの段の間）・幅12・高さ44（±TOL）・
    1段落・1文字区分・文字「→」・20pt・太字・改行なし・文字間隔なし"""
    if not _box_ok(fr['box'], e['box']):
        return False, '位置・寸法が期待（フローの段の間・幅12・高さ44）と違う'
    pars = paragraphs(fr['body'])
    if len(pars) != 1:
        return False, f'段落数が {len(pars)}（1であるべき）'
    p0 = pars[0]
    if len(p0['runs']) != 1 or p0['br']:
        return False, '文字区分が1つでない・改行がある'
    t, sz, b, spc = p0['runs'][0]
    if t != '→':
        return False, f'文字が「{t}」（「→」1字であるべき）'
    if sz != 20.0 or not b:
        return False, f'文字の大きさ・太字が違う（{sz}pt・太字{b}）'
    if spc or p0['spc_unknown']:
        return False, '文字間隔が指定されている'
    return True, ''


def estimate(fr, kind):
    """1枠の推定（設計 r2 §3）。明示改行は行ごとに数え、文字間隔は幅に加える"""
    l, r, t, b = insets(fr)
    x, y, w, h = fr['box']
    if not w or not h:
        return dict(status='未検査', why='枠の幅・高さが取れない')
    cw = w - l - r
    ch = h - t - b
    if cw <= 0 or ch <= 0:
        return dict(status='未検査', why='枠の有効な幅・高さが0以下')
    pars = paragraphs(fr['body'])
    total, total_aux, assumed, notes, width_ng = 0.0, 0.0, False, [], []
    for i, par in enumerate(pars):
        pw, undefined, nosize, mx, mixed = para_width(par)
        if nosize:
            return dict(status='未検査', why=f'段落{i + 1}の文字の大きさが分からない')
        if undefined:
            return dict(status='未検査', why=f'定義されていない文字種：{"".join(sorted(set(undefined)))}')
        if par['spc_unknown']:
            return dict(status='未検査', why=f'段落{i + 1}：継承された文字間隔があり、幅を計算できない')
        if mx is None:
            mx = 14.0
        if mixed:
            notes.append('大きさ混在')
        one_line = kind in ('line', 'map', 'multi_line') or (kind == 'card' and i == 1) or kind == 'flow'
        if par['br']:
            if one_line:
                width_ng.append(f'段落{i + 1}に明示の改行（1行必須の欄）')
            notes.append(f'明示改行 {len(par["segs"]) - 1}回')
        sp = _spacing_px(par['sb'], mx) + (_spacing_px(par['sa'], mx) if i < len(pars) - 1 else 0)
        segs = seg_widths(par, mx)
        for j, (sw, smx) in enumerate(segs):
            lh, asm = _line_h(par, smx)
            assumed |= asm
            lines = max(1, -(-sw // cw)) if sw > 0 else 1
            aux = max(1, -(-sw // max(1.0, cw - smx * 4 / 3))) if sw > 0 else 1
            if one_line and sw > cw + 1e-6:
                seg_txt = ''.join(rr[0] for rr in par['segs'][j])
                width_ng.append(f'段落{i + 1}の幅 {sw:.1f}px > 枠 {cw:.1f}px「{seg_txt}」')
            total += lines * lh
            total_aux += aux * lh
        total += sp
        total_aux += sp
        lw = long_ascii_word(par['text'], cw / (mx * 4 / 3))
        if lw and not one_line:
            notes.append(f'長い ASCII 単語「{lw}」')
    res = dict(height=round(total, 2), cap=round(ch, 2), notes=notes)
    height_scope = kind in ('multi', 'multi_line', 'card', 'flow') or len(pars) >= 2 or any(p['br'] for p in pars)
    if width_ng:
        res.update(status='NG', why='幅 NG：' + '／'.join(width_ng))
        if total > ch + 1e-6:
            res['why'] += f'／同じ枠の高さ超過 {total:.1f}px > {ch:.1f}px'
            res['height_ng'] = True
        return res
    if height_scope and total > ch + 1e-6:
        res.update(status='NG', why=f'高さ超過 {total:.1f}px > 枠 {ch:.1f}px', height_ng=True)
        return res
    if not height_scope:
        total_aux = 0.0
    reasons = []
    if total_aux > ch + 1e-6:
        reasons.append(f'折り返しの余裕不足（補助の判定 {total_aux:.1f}px > {ch:.1f}px）')
    if assumed and height_scope:
        notes.append('行間：仮定1.0')
        if (ch - total) / ch < 0.10:
            reasons.append('行間を仮定し高さの余裕10%未満')
    if any(n.startswith('長い ASCII') for n in notes):
        reasons.append('長い ASCII 単語')
    res.update(status='要実機確認' if reasons else '基準内', why='／'.join(reasons))
    return res


def _counts(data, slides, key):
    """値や件数で変わる構造の数。入力があれば入力から、無ければ PPTX から数える（旧PPTX用）"""
    if key == 'C-2':
        if data:
            return len((data.get('proposals') or {}).get('proposals') or [])
        return sum(1 for frs in slides for fr in frs if fr['kind'] == 'text' and fr['text'].startswith('改善箇所\u3000'))
    v = (data or {}).get('visual_data') if data else None
    s1 = slides[0] if slides else []
    if v is not None:
        ti = (v.get('top_issues') or ([v.get('top_issue')] if v.get('top_issue') else []) or v.get('priority_issues', []))
        return dict(n_sec=len(v.get('sections') or []) or 1, n_flow=len((v.get('flow_steps') or [])[:6]),
                    n_issues=min(len(ti), 4), n_h=len((v.get('highlights') or [])[:3]), n_sc=len((v.get('scores') or [])[:10]))
    n_sec = sum(1 for fr in s1 if fr['kind'] == 'text' and abs(round(fr['box'][0]) - 78) <= TOL
                and abs(round(fr['box'][2]) - 332) <= TOL and fr['text'].strip())
    n_flow = sum(1 for fr in s1 if fr['kind'] == 'text' and abs(round(fr['box'][1]) - 528) <= TOL and fr['text'].strip())
    mm = [re.search(r'Top(\d)', fr['text']) for fr in s1 if fr['text'].startswith('⚠ 最重要課題')]
    n_issues = int(mm[0].group(1)) if mm and mm[0] else 3
    s3 = slides[2] if len(slides) > 2 else []
    n_h = sum(1 for fr in s3 if fr['text'] == 'BEFORE')
    s2 = slides[1] if len(slides) > 1 else []
    n_sc = sum(1 for fr in s2 if fr['kind'] == 'text' and abs(round(fr['box'][0]) - 408) <= TOL and abs(round(fr['box'][2]) - 62) <= TOL)
    return dict(n_sec=n_sec, n_flow=n_flow, n_issues=n_issues, n_h=n_h, n_sc=n_sc)


def check_display(found, data=None):
    """期待する枠の表と実際の枠を照合して推定する。戻り値：欄ごとの結果・追加の NG 理由・対応の記録"""
    rows, ng, match = [], [], {}
    for key in ('C-1', 'C-2', 'C-3'):
        if key not in found:
            continue
        fn = os.path.basename(found[key])
        try:
            slides = read_frames(found[key])
        except Exception as e:  # noqa
            rows.append(dict(file=fn, slide=0, field='（ファイル）', status='未検査', why=f'PPTX が開けない：{e}'))
            continue
        if key == 'C-1':
            X = expect_c1((data or {}).get('diagnosis') if data else None); want = (2,)
        elif key == 'C-2':
            n = _counts(data, slides, key)
            X, tot = expect_c2((data or {}).get('proposals') if data else None, n); want = (tot,)
        else:
            c = _counts(data, slides, key)
            X = expect_c3((data or {}).get('visual_data') if data else None, **c); want = (3,)
        if len(slides) not in want:
            rows.append(dict(file=fn, slide=0, field='（枚数）', status='未検査',
                             why=f'{len(slides)}枚（想定 {"・".join(map(str, want))}枚）'))
            continue
        found_ct = {e['fid']: 0 for e in X}
        for sno, frames in enumerate(slides, 1):
            tables = [fr for fr in frames if fr['kind'] == 'cell']
            for fr in frames:
                if fr['kind'] == 'text' and not fr['text'].strip():
                    continue            # 文字の無い枠（図形の地）は対象外
                if 'by GPTs' in fr['text']:
                    ng.append(f'{fn} {sno}枚目：「by GPTs」を含む枠がある「{fr["text"]}」')
                base = dict(file=fn, slide=sno, box=[round(v) for v in fr['box']], text=fr['text'])
                if fr['kind'] == 'cell':
                    tb = [e for e in X if e['kind'] == 'table' and e['slide'] == sno]
                    if len(tb) != 1 or not _box_ok(fr['table_box'], tb[0]['box']):
                        rows.append(dict(base, field='（想定外の表）', status='未検査', why='期待する表の位置・スライドに無い'))
                        continue
                    cand = [e for e in X if e.get('cell') == (fr['row'], fr['col']) and e['slide'] == sno]
                else:
                    cand = [e for e in X if e['slide'] == sno and e['box'] is not None and e['kind'] != 'table'
                            and _box_ok(fr['box'], e['box'])]
                if not cand:
                    rows.append(dict(base, field='（想定外の枠）', status='未検査',
                                     why='期待する枠の表のどれにも位置が合わない'))
                    continue
                if len(cand) > 1:
                    rows.append(dict(base, field='（多重一致）', status='未検査',
                                     why='2つ以上の期待枠に位置が合う：' + '・'.join(e['fid'] for e in cand)))
                    continue
                e = cand[0]
                found_ct[e['fid']] += 1
                if found_ct[e['fid']] > 1:
                    rows.append(dict(base, field=e['fid'], status='NG', why='同じ期待枠が重複している（2つ目以降）'))
                    continue
                match[(key, e['fid'])] = fr
                if e['cls'] == 'footer':
                    # 判定126-2：文言一致に加え、通常の1行必須欄として幅・改行・文字間隔・文字情報を検査する
                    r = estimate(fr, 'line')
                    whys = [r['why']] if r.get('why') else []
                    st = r['status']
                    if fr['text'] != AUTHOR:
                        whys.insert(0, f'フッター：表記の不一致「{fr["text"]}」（正：{AUTHOR}）')
                        if st != '未検査':
                            st = 'NG'
                    rows.append(dict(base, field=e['fid'], kind='footer', cls='footer',
                                     **dict({k: v for k, v in r.items() if k not in ('status', 'why')}, status=st, why='／'.join(whys))))
                    continue
                if e['cls'] == 'exception':
                    ok, why = arrow_match(fr, e)
                    if ok:
                        pw = para_width(paragraphs(fr['body'])[0])[0]
                        rows.append(dict(base, field=e['fid'], status='参考', kind='arrow',
                                         why=f'例外表 #1（幅 {pw:.1f}px／枠 {fr["box"][2]:.0f}px）・実機確認'))
                    else:
                        r = estimate(fr, 'line')
                        rows.append(dict(base, field=e['fid'], status='NG' if r['status'] != '未検査' else '未検査',
                                         why=f'矢印の例外の条件を満たさない：{why}' + (f'／{r["why"]}' if r.get('why') else '')))
                    continue
                r = estimate(fr, e['kind'])
                if not _text_ok(fr['text'], e['exp']):
                    lab = '固定文言' if e['cls'] == 'fixed' else '入力の欄'
                    r = dict(r, status='NG', why=(r.get('why') + '／' if r.get('why') else '')
                             + f'{lab}の文字が期待と一致しない「{fr["text"][:40]}」')
                if e['cls'] == 'input' and '…' in fr['text'] and r['status'] != 'NG':
                    r = dict(r, status='NG', why='末尾が「…」＝builder の切り詰め（B の制約に抵触）')
                if e.get('card'):
                    okc, whyc = fixed_card_match(*e['card'], sno, fr, data)
                    if not okc:
                        r = dict(r, status='NG' if r.get('status') != '未検査' else '未検査',
                                 why=(r.get('why') + '／' if r.get('why') else '') + f'固定カードの構造が builder と違う：{whyc}')
                    elif r.get('status') == 'NG' and r.get('height_ng') and not str(r.get('why', '')).startswith(('幅 NG', '固定文言', '入力の欄')) \
                            and '一致しない' not in str(r.get('why', '')):
                        r = dict(r, status='要実機確認', exception='固定カード',
                                 why=f'固定カード／高さ推定{r["height"]}px・枠{r["cap"]:g}px／高さ判定は限定例外（判定118）／Windows PowerPoint未確認')
                rows.append(dict(base, field=e['fid'], kind=e['kind'], cls=e['cls'], **r))
        for e in X:
            if e['kind'] == 'table':
                continue
            if found_ct[e['fid']] == 0:
                rows.append(dict(file=fn, slide=e['slide'], field=e['fid'], status='NG', why='期待する枠が見つからない（欠落・移動）'))
    return rows, ng, match


def check_cross_pptx(data, found, match):
    """判定122-3：C-2 の提案と C-3 のハイライトの対応を、生成後の枠単位で完全一致まで確かめる"""
    ng = []
    p = (data.get('proposals') or {}).get('proposals') or []
    hls = ((data.get('visual_data') or {}).get('highlights') or [])[:3]
    for i, h in enumerate(hls):
        cand = [k for k, q in enumerate(p) if isinstance(q, dict) and isinstance(h, dict) and q.get('target_area') == h.get('target_area')]
        if len(cand) != 1:
            continue                    # 入力の点検で止まる
        k = cand[0]
        pairs = [('箇所', f'2-提案{k + 1}-箇所', '改善箇所\u3000', f'3-3-HL{i + 1}-箇所', '箇所：'),
                 ('優先度', f'2-提案{k + 1}-優先度', '', f'3-3-HL{i + 1}-優先度', ''),
                 ('工数', f'2-提案{k + 1}-工数', '工数：', f'3-3-HL{i + 1}-工数', '工数：'),
                 ('現状', f'2-提案{k + 1}-現状', '現状\u3000', f'3-3-HL{i + 1}-現状', ''),
                 ('改善後', f'2-提案{k + 1}-改善後', '改善後\u3000', f'3-3-HL{i + 1}-改善後', '')]
        for jp, f2, pre2, f3, pre3 in pairs:
            a, b = match.get(('C-2', f2)), match.get(('C-3', f3))
            if a is None or b is None:
                ng.append(f'C-2 提案{k + 1}／C-3 ハイライト{i + 1}：{jp}の枠が見つからない'); continue
            ta, tb = a['text'], b['text']
            if not (ta.startswith(pre2) and tb.startswith(pre3)) or ta[len(pre2):] != tb[len(pre3):]:
                ng.append(f'C-2 提案{k + 1}／C-3 ハイライト{i + 1}：PPTX の{jp}が一致しない（C-2「{ta}」／C-3「{tb}」）')
    return ng


def summarize(rows):
    return {k: [r for r in rows if r['status'] == k] for k in ('基準内', '要実機確認', 'NG', '未検査', '参考')}


def print_display(rows, extra_ng):
    cnt = summarize(rows)
    bad_frames = len(cnt['NG'])
    print(f'表示の簡易推定：基準内（{len(cnt["基準内"])}欄）／要実機確認（{len(cnt["要実機確認"])}欄）'
          f'／NG（{bad_frames}枠＋追加理由{len(extra_ng)}件）／未検査（{len(cnt["未検査"])}件）／参考（{len(cnt["参考"])}件）')
    clean = lambda t: re.sub(r'[\n\r\x0b\x0c]', '⏎', str(t))   # 改行・垂直タブを1行の表示に（報告の行が割れないように）
    for k in ('NG', '未検査', '要実機確認', '参考'):
        for r in cnt[k]:
            print(f'  [{k}] {r["file"]} {r["slide"]}枚目 {r["field"]}：{clean(r.get("why", ""))}')
    for m in extra_ng:
        print(f'  [NG] {clean(m)}')
    return bad_frames + len(extra_ng), len(cnt['未検査'])


def find_files(out_dir):
    files = sorted(os.path.basename(x) for x in glob.glob(os.path.join(out_dir, '*.pptx')))
    found = {}
    for key, rx in NAME_RE.items():
        hit = [f for f in files if rx.match(f)]
        if len(hit) == 1:
            found[key] = os.path.join(out_dir, hit[0])
    return found


def main(argv):
    report = None
    if '--report' in argv:
        i = argv.index('--report'); report = argv[i + 1]; argv = argv[:i] + argv[i + 2:]
    if len(argv) == 2 and argv[0] == '--pptx-only':
        found = find_files(argv[1])
        rows, extra, _ = check_display(found)
        print('点検対象：', ', '.join(os.path.basename(v) for v in found.values()) or '（なし）')
        print('内容・構成点検：未実施（入力なし・--pptx-only。入力との照合・record・C-2/C-3 の対応は未実施）')
        n_ng, n_un = print_display(rows, extra)
        print('Windows PowerPoint 目視：未実施')
        print('※「基準内」は実際の描画の合格を意味しない。最終の判定は Windows の PowerPoint での目視。')
        if report:
            json.dump(dict(rows=rows, extra_ng=extra), open(report, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        return 3 if (n_ng == 0 and n_un == 0) else 1
    if len(argv) != 2:
        print(__doc__); return 2
    data = json.load(open(argv[0], encoding='utf-8'))
    files, ng = check(data, argv[1])                  # r5 の点検（そのまま）
    ng_in, notes = check_input_r6(data)               # 入力の A・B・C・文字種・record・C-2/C-3 の対応
    found = find_files(argv[1])
    rows, extra, match = check_display(found, data)   # 表示の簡易推定（入力の結果に関わらず実行）
    ng_px = check_cross_pptx(data, found, match)      # 生成後の C-2/C-3 の対応（枠単位の完全一致）
    content_ng = ng + ng_in + ng_px
    print('点検対象：', ', '.join(files) or '（なし）')
    for n in notes:
        print('  表示名の報告：', n)
    if content_ng:
        print(f'内容・構成点検：🚨 不合格 {len(content_ng)}件')
        for m in content_ng:
            print('  -', m)
    else:
        print('内容・構成点検：⭕ 合格（総評17字・ランクの説明・件数・ファイル名・test_*.pptx 無し・リンク無し・'
              'メイリオ14pt以上・PPTX と入力の一致・Before/After の箇所と優先度・行動の流れのまとめ・'
              '字数 A/B/C・文字種・record・C-2/C-3 の対応）')
    n_ng, n_un = print_display(rows, extra)
    print('Windows PowerPoint 目視：未実施')
    print('※「基準内」は実際の描画の合格を意味しない。最終の判定は Windows の PowerPoint での目視。')
    if report:
        json.dump(dict(content_ng=content_ng, rows=rows, extra_ng=extra), open(report, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1)
    ok = (not content_ng) and n_ng == 0 and n_un == 0
    print('⭕ 点検合格（表示は簡易推定・実機の目視は別途）' if ok else '🚨 不合格 → 入力を直して PPTX を作り直す')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
