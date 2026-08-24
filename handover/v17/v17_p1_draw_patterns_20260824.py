# -*- coding: utf-8 -*-
"""
v17 P1｜図解パターン描画ロジック（category / breakdown / comparison）
=====================================================================

位置づけ
--------
第16条「段階2（実装済）」までを対象とするAIスライド成果物。
段階3（統合済）は入江さんが本ファイルの内容を
`03_pptx_builder_v16_5_20260728.py` へ挿入し、動作確認をもって成立する。

挿入位置（実装手順書 §2 参照）
--------------------------------
`03_pptx_builder_v16_5_20260728.py` の
「▲▲▲ Phase A 追加ブロック ここまで ▲▲▲」（213行）の直後、
「# キャンバス（1280×720 想定）」（215行）の直前に、
本ファイルの「▼▼▼ v17 追加ブロック」以降を丸ごと貼り付ける。
ただし本ブロックはビルダー後半で定義される
`px` / `add_text` / `add_shape` / `add_paragraph_box` / `set_run` /
`CANVAS_W_PX` / `TEXT` / `WHITE` / `NAVY` を実行時に参照するため、
**関数内参照のみ**で構成してある（import時評価はしない）。
モジュール末尾に置いても動作する。

仕様の出所（第9条・第17条準拠｜参照行を明記）
------------------------------------------------
1. `DIAGRAM_PATTERNS`      : ビルダー 181〜194行（原本辞書｜12種×3属性）
2. `DIAGNOSIS_TO_PATTERN`  : ビルダー 197〜209行（11マッピング）
3. `DIGITAL_AGENCY_PALETTE`: ビルダー 95〜165行（7テーマ×8色階調）
4. 拡張プロパティ（min/max/direction/color_gradation）:
   `v35_core_extended_pattern_definitions.md`（13,473B）
   「拡張定義｜集約表」P1-2 / P1-3 / P1-4 行
5. 横断的原則4件（フォールバック／警告オーバーライド／決定論性／基本図形優先）:
   同ファイル「横断的な設計原則｜全パターン共通｜4件」

作成: 2026-08-23（日）15:00連結②／AIスライド（実装領域）
"""

# =====================================================================
# ▼▼▼ v17 追加ブロック（P1｜category / breakdown / comparison） ▼▼▼
# 既存の C-1〜C-3 描画ロジックには一切触れない（後方互換完全維持）
# =====================================================================

# ---------------------------------------------------------------------
# v17-0｜拡張パターン仕様（原本3属性に対する拡張層）
#   出典: v35_core_extended_pattern_definitions.md 拡張定義集約表
#   ※ 原本 DIAGRAM_PATTERNS は書き換えない（2層構造を維持）
# ---------------------------------------------------------------------
DIAGRAM_PATTERN_SPEC = {
    'category': {
        'min_elements': 3, 'max_elements': 6, 'requires_axes': False,
        'direction': 'grid', 'color_gradation': 'uniform_parallel',
        'grid_map': {3: (3, 1), 4: (2, 2), 5: (3, 2), 6: (3, 2)},
    },
    'breakdown': {
        'min_elements': 3, 'max_elements': 7, 'requires_axes': False,
        'direction': 'vertical', 'color_gradation': 'proportional',
    },
    'comparison': {
        'min_elements': 2, 'max_elements': 3, 'requires_axes': False,
        'direction': 'horizontal', 'color_gradation': 'discrete_contrast',
    },
}

# 描画領域（ヘッダ60px／フッター境界660pxの内側）
V17_AREA = {
    'left': 40, 'right': 1240, 'width': 1200,
    'title_top': 90, 'body_top': 138, 'body_bottom': 646,
    'gap': 16,
}

# 警告オーバーライドの閾値（原則②｜B-6 4.4節）
V17_WARNING_SCORE = 40

# カード高さの上限（px）
# v17-fix1（8/24 実機検証）: 上限230pxではラベル+スコア+説明1行=108pxに対し
# 充填率47%となり「カード内が空白だらけ」に見える事象を検出。
# 説明文2行（44px）を収容しつつ充填率60%以上を確保する値として186pxを採用。
V17_CARD_H_MAX = 186


# ---------------------------------------------------------------------
# v17-1｜色ユーティリティ
# ---------------------------------------------------------------------
def hex_to_rgb(hex_str):
    """'#0017C1' / '0017C1' -> RGBColor"""
    from pptx.dml.color import RGBColor
    s = str(hex_str).lstrip('#')
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def get_theme_palette(theme_id):
    """テーマIDから8色階調辞書（hex文字列）を返す。未知IDは SolidGray。"""
    if theme_id not in DIGITAL_AGENCY_PALETTE:
        theme_id = 'SolidGray'
    return DIGITAL_AGENCY_PALETTE[theme_id]


def select_theme_by_project_type(project_type, warning_flag=False):
    """プロジェクトタイプ＋警告フラグからテーマIDを決定（決定論的）。"""
    if warning_flag:
        return 'Red'
    return {
        'corporate': 'Blue',
        'ec':        'Orange',
        'lp':        'Orange',
        'webapp':    'Cyan',
        'media':     'Green',
    }.get(project_type, 'SolidGray')


def _relative_luminance(hex_str):
    """WCAG 相対輝度（0.0〜1.0）"""
    s = str(hex_str).lstrip('#')
    out = []
    for i in (0, 2, 4):
        v = int(s[i:i + 2], 16) / 255.0
        out.append(v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4)
    return 0.2126 * out[0] + 0.7152 * out[1] + 0.0722 * out[2]


def _text_color_on(bg_hex):
    """背景色に対して可読な文字色を返す（WHITE または TEXT）。

    v17 拡張｜拡張定義集約表には規定がないため本実装で追加した。
    根拠: テーマ Green/Cyan/Red/Orange の warning 色は '#CCCCCC'（明色）で、
    白文字を固定すると score<40 の警告セルが判読不能になる。
    輝度0.5を境に自動反転する（決定論的｜同一入力→同一出力）。
    """
    return WHITE if _relative_luminance(bg_hex) < 0.5 else TEXT


def _tier_fill(palette, score, base_key):
    """基準色 + 警告オーバーライド（原則②）。

    score が None の場合は警告判定をスキップする。
    """
    if score is not None and score < V17_WARNING_SCORE:
        return palette['warning']
    return palette[base_key]


# ---------------------------------------------------------------------
# v17-2｜共通描画ヘルパー
# ---------------------------------------------------------------------
def _v17_title(slide, title, palette):
    """パターン共通のタイトル行（ヘッダ帯とは別のスライド内見出し）"""
    add_text(slide, V17_AREA['left'], V17_AREA['title_top'], V17_AREA['width'],
             str(title), 22, bold=True, color=hex_to_rgb(palette['primary']),
             height_px=34)


def _v17_axis_label(slide, text, palette):
    """任意の軸ラベル（比較軸・時間軸等）。斜体は使わない（条項13）。"""
    add_text(slide, V17_AREA['left'], V17_AREA['title_top'] + 36,
             V17_AREA['width'], str(text), 14, bold=True,
             color=hex_to_rgb(palette['secondary']), height_px=22)


def _v17_card(slide, x, y, w, h, fill_hex, palette, radius_px=8):
    """カード図形（塗り＋枠線のみ｜条項7：図形に文字を入れない）"""
    return _add_bg_frame(slide, x, y, w, h,
                         fill=hex_to_rgb(fill_hex),
                         line=hex_to_rgb(palette['midtone']),
                         line_width_pt=1, radius_px=radius_px)


def _v17_rule(slide, x, y, w, h, hex_color):
    """罫線・コネクタ代替の矩形（原則④：基本図形のみ）"""
    from pptx.enum.shapes import MSO_SHAPE
    return add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, max(w, 1), max(h, 1),
                     fill=hex_to_rgb(hex_color))


def _v17_normalize(items, key_label='label'):
    """要素を dict 化して正規化（str 入力も許容）"""
    out = []
    for it in items or []:
        if isinstance(it, dict):
            d = dict(it)
        else:
            d = {key_label: str(it)}
        d.setdefault(key_label, '')
        out.append(d)
    return out


def _v17_report(pattern, drawn, fallback_from=None, notes=None):
    """描画結果レポート（回帰確認・ログ記録用）"""
    return {
        'pattern': pattern,
        'elements_drawn': drawn,
        'fallback_from': fallback_from,
        'notes': list(notes or []),
    }


# ---------------------------------------------------------------------
# v17-3｜category（分類）｜grid / uniform_parallel / 3〜6要素
# ---------------------------------------------------------------------
def draw_category(slide, palette, data, _fallback_from=None):
    """診断結果を並列カテゴリ（グリッド）で描画する。

    Args:
        slide   : python-pptx Slide
        palette : 8色階調辞書（get_theme_palette の戻り値）
        data    : {'title': str,
                   'categories': [{'label': str, 'score': int|None,
                                   'description': str}, ...]}
    Returns:
        dict : 描画レポート

    仕様: 拡張定義集約表 P1-2 行（min3 / max6 / grid / uniform_parallel）
    フォールバック本体のため例外は送出しない（原則①）。
    """
    spec = DIAGRAM_PATTERN_SPEC['category']
    notes = []
    items = _v17_normalize(data.get('categories', []))

    # 要素数の丸め込み（例外を投げない｜原則①）
    if len(items) > spec['max_elements']:
        notes.append('要素数 %d > max %d｜先頭 %d 件で描画'
                     % (len(items), spec['max_elements'], spec['max_elements']))
        items = items[:spec['max_elements']]
    if len(items) < spec['min_elements']:
        notes.append('要素数 %d < min %d｜そのまま描画（劣化描画で通す）'
                     % (len(items), spec['min_elements']))
    if not items:
        _v17_title(slide, data.get('title', ''), palette)
        notes.append('要素0件｜タイトルのみ描画')
        return _v17_report('category', 0, _fallback_from, notes)

    _v17_title(slide, data.get('title', ''), palette)

    n = len(items)
    cols, rows = spec['grid_map'].get(n, (min(n, 3), (n + 2) // 3))
    gap = V17_AREA['gap']
    avail_h = V17_AREA['body_bottom'] - V17_AREA['body_top']
    cell_w = (V17_AREA['width'] - gap * (cols - 1)) / float(cols)
    cell_h = min((avail_h - gap * (rows - 1)) / float(rows), V17_CARD_H_MAX)
    block_h = cell_h * rows + gap * (rows - 1)
    y0 = V17_AREA['body_top'] + (avail_h - block_h) / 2.0

    for i, it in enumerate(items):
        r, c = i // cols, i % cols
        in_row = min(cols, n - r * cols)
        row_w = cell_w * in_row + gap * (in_row - 1)
        x = V17_AREA['left'] + (V17_AREA['width'] - row_w) / 2.0 + c * (cell_w + gap)
        y = y0 + r * (cell_h + gap)

        # uniform_parallel: 全セル同一の基準色（並列性の担保）＋警告オーバーライド
        score = it.get('score')
        fill_hex = _tier_fill(palette, score, 'secondary')
        fg = _text_color_on(fill_hex)
        _v17_card(slide, x, y, cell_w, cell_h, fill_hex, palette)

        pad = 16
        add_text(slide, x + pad, y + 12, cell_w - pad * 2,
                 str(it.get('label', '')), 16, bold=True, color=fg, height_px=26)
        if score is not None:
            add_text(slide, x + pad, y + 42, cell_w - pad * 2,
                     '%s%%' % score, 26, bold=True, color=fg, height_px=40)
        desc_top = y + (86 if score is not None else 46)
        desc = str(it.get('description', ''))
        if desc:
            add_text(slide, x + pad, desc_top, cell_w - pad * 2, desc, 14,
                     color=fg, height_px=max(int(cell_h - (desc_top - y) - 12), 22),
                     line_height=1.4)

    return _v17_report('category', n, _fallback_from, notes)


# ---------------------------------------------------------------------
# v17-4｜breakdown（分解）｜vertical / proportional / 3〜7要素
# ---------------------------------------------------------------------
def _breakdown_tier(ratio):
    """構成比から色階調キーを決定（proportional｜決定論的）"""
    if ratio >= 0.30:
        return 'primary'
    if ratio >= 0.20:
        return 'secondary'
    if ratio >= 0.10:
        return 'midtone'
    return 'light'


def draw_breakdown(slide, palette, data):
    """全体を構成要素に分解して縦積みで描画する（構成比連動の色階調）。

    Args:
        data : {'title': str,
                'whole': {'label': str, 'value': int|float|None},
                'components': [{'label': str, 'value': int|float,
                                'score': int|None, 'note': str}, ...]}
    仕様: 拡張定義集約表 P1-3 行（min3 / max7 / vertical / proportional）
    要素数逸脱・データ不正時は draw_category へフォールバック（原則①）。
    """
    spec = DIAGRAM_PATTERN_SPEC['breakdown']
    notes = []
    comps = _v17_normalize(data.get('components', []))

    # 数値健全性チェック
    values = []
    bad = False
    for c in comps:
        try:
            v = float(c.get('value'))
        except (TypeError, ValueError):
            bad = True
            break
        if v < 0:
            bad = True
            break
        values.append(v)
    total_in = sum(values) if values else 0.0

    if (not (spec['min_elements'] <= len(comps) <= spec['max_elements'])
            or bad or total_in <= 0):
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': c.get('label', ''),
                            'score': c.get('score'),
                            'description': str(c.get('note', ''))} for c in comps],
        }, _fallback_from='breakdown')
        rep['notes'].insert(0, 'breakdown 前提不成立（要素数=%d／数値不正=%s／合計=%s）｜'
                            'category へフォールバック' % (len(comps), bad, total_in))
        return rep

    _v17_title(slide, data.get('title', ''), palette)

    whole = data.get('whole') or {}
    whole_label = str(whole.get('label', '全体'))
    whole_value = whole.get('value')
    if whole_value is None:
        whole_value = total_in

    # 上段｜全体ボックス
    wh_h = 56
    wh_y = V17_AREA['body_top']
    _v17_card(slide, V17_AREA['left'], wh_y, V17_AREA['width'], wh_h,
              palette['primary'], palette)
    fg_w = _text_color_on(palette['primary'])
    add_text(slide, V17_AREA['left'] + 20, wh_y + 15, 700, whole_label, 18,
             bold=True, color=fg_w, height_px=26)
    add_text(slide, V17_AREA['left'] + V17_AREA['width'] - 320, wh_y + 13, 300,
             _fmt_num(whole_value), 20, bold=True, color=fg_w, height_px=28,
             align=_pp_right())

    # 下段｜構成要素（縦積み・幅を構成比に比例）
    n = len(comps)
    gap = 10
    top = wh_y + wh_h + 22
    avail = V17_AREA['body_bottom'] - top
    row_h = min((avail - gap * (n - 1)) / float(n), 62)
    label_w = 300
    bar_x = V17_AREA['left'] + label_w + 12
    bar_max_w = V17_AREA['width'] - label_w - 12 - 150

    # 左スパイン（親→子の接続｜基本図形のみ）
    spine_bottom = top + row_h * n + gap * (n - 1) - row_h / 2.0
    _v17_rule(slide, V17_AREA['left'] + 24, wh_y + wh_h,
              3, max(spine_bottom - (wh_y + wh_h), 1), palette['midtone'])

    for i, c in enumerate(comps):
        v = values[i]
        ratio = v / total_in
        y = top + i * (row_h + gap)
        tier = _breakdown_tier(ratio)
        fill_hex = _tier_fill(palette, c.get('score'), tier)
        fg = _text_color_on(fill_hex)

        # 横枝（スパイン→ラベル）
        _v17_rule(slide, V17_AREA['left'] + 24, y + row_h / 2.0 - 1,
                  28, 3, palette['midtone'])

        add_text(slide, V17_AREA['left'] + 58, y + (row_h - 24) / 2.0,
                 label_w - 58, str(c.get('label', '')), 16, bold=True,
                 color=hex_to_rgb(palette['primary']), height_px=26)

        bar_w = max(bar_max_w * ratio, 40)
        _v17_card(slide, bar_x, y, bar_w, row_h, fill_hex, palette, radius_px=6)

        # 値ラベル｜バー内に収まる場合は内側（白抜き/濃字）、
        # 収まらない場合はバー右外に出す（文字切れ防止｜条項5）
        val_text = '%s（%.1f%%）' % (_fmt_num(v), ratio * 100)
        need_w = _est_text_w(val_text, 14) * 1.2 + 28
        if bar_w >= need_w:
            add_text(slide, bar_x + 14, y + (row_h - 24) / 2.0, bar_w - 28,
                     val_text, 14, bold=True, color=fg, height_px=24)
        else:
            add_text(slide, bar_x + bar_w + 10, y + (row_h - 24) / 2.0,
                     _est_text_w(val_text, 14) * 1.2 + 20, val_text, 14,
                     bold=True, color=hex_to_rgb(palette['primary']),
                     height_px=24)

        note = str(c.get('note', ''))
        if note:
            add_text(slide, bar_x + bar_max_w + 16, y + (row_h - 22) / 2.0, 134,
                     note, 14, color=hex_to_rgb(palette['secondary']),
                     height_px=24)

    notes.append('構成比合計=%.1f%%（入力合計 %s を100%%として正規化）'
                 % (100.0, _fmt_num(total_in)))
    return _v17_report('breakdown', n, None, notes)


# ---------------------------------------------------------------------
# v17-5｜comparison（比較）｜horizontal / discrete_contrast / 2〜3要素
# ---------------------------------------------------------------------
def draw_comparison(slide, palette, data):
    """2〜3要素を並列カラムで対比描画する。

    Args:
        data : {'title': str,
                'comparison_axis': str|None,      # 任意（原本 requires_axes=False）
                'attribute_labels': [str, ...],
                'items': [{'label': str, 'score': int|None,
                           'attributes': {key: value}}, ...]}
    仕様: 拡張定義集約表 P1-4 行（min2 / max3 / horizontal / discrete_contrast）
    要素数逸脱時は draw_category へフォールバック（原則①）。
    """
    spec = DIAGRAM_PATTERN_SPEC['comparison']
    notes = []
    items = _v17_normalize(data.get('items', []))

    if not (spec['min_elements'] <= len(items) <= spec['max_elements']):
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': it.get('label', ''),
                            'score': it.get('score'),
                            'description': ''} for it in items],
        }, _fallback_from='comparison')
        rep['notes'].insert(0, 'comparison 要素数 %d が %d〜%d の範囲外｜'
                            'category へフォールバック'
                            % (len(items), spec['min_elements'], spec['max_elements']))
        return rep

    _v17_title(slide, data.get('title', ''), palette)
    axis = data.get('comparison_axis')
    body_top = V17_AREA['body_top']
    if axis:
        _v17_axis_label(slide, '比較軸：%s' % axis, palette)
        body_top += 26

    # discrete_contrast: 中間色を挟まず index で離散的に割り当てる
    contrast_keys = {2: ['midtone', 'primary'],
                     3: ['midtone', 'secondary', 'primary']}[len(items)]

    n = len(items)
    gap = 24
    col_w = (V17_AREA['width'] - gap * (n - 1)) / float(n)
    col_h = V17_AREA['body_bottom'] - body_top
    head_h = 108
    attr_labels = list(data.get('attribute_labels') or [])

    for i, it in enumerate(items):
        x = V17_AREA['left'] + i * (col_w + gap)
        score = it.get('score')
        fill_hex = _tier_fill(palette, score, contrast_keys[i])
        fg = _text_color_on(fill_hex)

        # 本体カード（薄色）＋ヘッダ帯（濃色）
        _v17_card(slide, x, body_top, col_w, col_h, palette['lightest'], palette)
        _v17_card(slide, x, body_top, col_w, head_h, fill_hex, palette)

        add_text(slide, x + 18, body_top + 14, col_w - 36,
                 str(it.get('label', '')), 18, bold=True, color=fg, height_px=28)
        if score is not None:
            add_text(slide, x + 18, body_top + 48, col_w - 36, '%s%%' % score,
                     32, bold=True, color=fg, height_px=48)

        # 属性行
        attrs = it.get('attributes') or {}
        keys = attr_labels or list(attrs.keys())
        ay = body_top + head_h + 18
        row_h = 46
        for k in keys:
            if ay + row_h > body_top + col_h - 8:
                notes.append('カラム%d｜属性が領域を超過したため以降を省略' % (i + 1))
                break
            add_text(slide, x + 18, ay, col_w - 36, str(k), 14, bold=True,
                     color=hex_to_rgb(palette['secondary']), height_px=22)
            add_text(slide, x + 18, ay + 20, col_w - 36,
                     str(attrs.get(k, '－')), 16,
                     color=hex_to_rgb(palette['primary']), height_px=24)
            _v17_rule(slide, x + 18, ay + row_h - 2, col_w - 36, 1,
                      palette['light'])
            ay += row_h

    return _v17_report('comparison', n, None, notes)


# ---------------------------------------------------------------------
# v17-6｜共通ディスパッチャ
# ---------------------------------------------------------------------
def resolve_pattern(diagnosis_key):
    """診断カテゴリ → パターンキー（原本 DIAGNOSIS_TO_PATTERN を参照）。

    未定義キーは 'category' にフォールバック（原則①）。
    """
    return DIAGNOSIS_TO_PATTERN.get(diagnosis_key, 'category')


def draw_pattern(slide, pattern_key, palette, data):
    """パターンキーで描画関数を振り分ける。未実装キーは category に退避。"""
    table = {
        'category':   draw_category,
        'breakdown':  draw_breakdown,
        'comparison': draw_comparison,
    }
    fn = table.get(pattern_key)
    if fn is None:
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': data.get('categories', []) or [
                {'label': str(x.get('label', '')), 'score': x.get('score'),
                 'description': ''} for x in (data.get('items')
                                              or data.get('components') or [])],
        }, _fallback_from=pattern_key)
        rep['notes'].append('パターン "%s" は v17 P1 の対象外（P2/P3 で実装）'
                            % pattern_key)
        return rep
    return fn(slide, palette, data)


def _est_text_w(text, size_pt):
    """メイリオ想定の概算テキスト幅（px）。

    全角=size_pt*1.34px／半角=size_pt*0.70px で見積る。
    バー内に値ラベルが収まるかの判定に用いる（文字切れ防止）。
    """
    w = 0.0
    for ch in str(text):
        w += size_pt * (0.70 if ord(ch) < 0x2000 else 1.34)
    return w


def add_diagram_slide(prs, pattern_key, palette, data, page_num=1, total=1,
                      header_label=None, author='紺＆クリーン スライド作成'):
    """図解パターン1枚を「ヘッダ帯＋フッター＋パターン描画」で1スライド出力する。

    既存の add_* 関数群と同じ呼び出し規約（prs を受け取り slide を返す）に揃えた
    ラッパ。v17-fix2（8/24 実機検証）で検出した「_add_header と draw_* の
    両方がタイトルを描き、同一文言が2箇所に出る」事象を構造的に防ぐ。

    ヘッダ帯には data['title'] を出さず、原本 DIAGRAM_PATTERNS の `ja`
    （例: 'category' → '分類'）を既定ラベルとして表示する。
    スライド内見出しは draw_* 側が data['title'] を1箇所だけ描く。

    Args:
        prs         : python-pptx Presentation
        pattern_key : 'category' / 'breakdown' / 'comparison' 等
        palette     : get_theme_palette の戻り値
        data        : 各 draw_* のデータ辞書
        page_num    : ページ番号
        total       : 総ページ数
        header_label: ヘッダ帯の左側文言（None なら原本 `ja` を使用）
    Returns:
        (slide, report)
    """
    slide = _blank_slide(prs)
    if header_label is None:
        meta = DIAGRAM_PATTERNS.get(pattern_key) or {}
        header_label = meta.get('ja', '図解')
    _add_header(slide, header_label, str(pattern_key).upper())
    _add_footer(slide, page_num, total, author=author)
    report = draw_pattern(slide, pattern_key, palette, data)
    return slide, report


def _fmt_num(v):
    """数値の表示整形（整数はそのまま、小数は1桁）"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if abs(f - int(f)) < 1e-9 else '%.1f' % f


def _pp_right():
    from pptx.enum.text import PP_ALIGN
    return PP_ALIGN.RIGHT

# =====================================================================
# ▲▲▲ v17 追加ブロック ここまで ▲▲▲
# =====================================================================
