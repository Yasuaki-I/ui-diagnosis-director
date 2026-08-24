# -*- coding: utf-8 -*-
"""
v17 P2｜図解パターン描画ロジック（pyramid / sequence / framework）
=====================================================================

位置づけ
--------
第16条「段階2（実装済）」までを対象とするAIスライド成果物。
段階3（統合済）は入江さんが本ファイルの内容を
`03_pptx_builder_v17_1_20260824.py` へ追記し、動作確認をもって成立する。

挿入位置
--------
`03_pptx_builder_v17_1_20260824.py` の
「▲▲▲ v17 追加ブロック ここまで ▲▲▲」の直後（＝ファイル末尾）に、
本ファイルの「▼▼▼ v17 P2 追加ブロック」以降を貼り付ける。
P1ブロックのヘルパー（`_v17_card` / `_tier_fill` / `_text_color_on` /
`_est_text_w` / `_v17_normalize` / `_v17_report` / `draw_category`）を
実行時に参照するため、必ず P1ブロックより後に置く。

仕様の出所（第9条・第17条 細則7｜参照行を明記）
------------------------------------------------
1. `DIAGRAM_PATTERNS`      : ビルダー 181〜194行（原本辞書｜`ja`/`use`/`shape`）
   - `pyramid`   : 182行｜'ピラミッド'／'階層・優先順位を上下で表現'／'triangle'
   - `sequence`  : 184行｜'順序'／'ステップ・時系列を左→右'／'arrow_chain'
   - `framework` : 192行｜'フレームワーク'／'4象限マトリクス等'／'quadrant'
2. `DIAGNOSIS_TO_PATTERN`  : ビルダー 197〜209行
   - `priority_ranking` → pyramid｜`user_flow` → sequence｜
     `impact_cost_matrix` → framework
3. 拡張プロパティ : `v35_core_extended_pattern_definitions.md`
   「拡張定義｜集約表」P2-1 / P2-2 / P2-3 行
   - pyramid   : min3／max5／axes False／vertical／hierarchical
   - sequence  : min3／max6／axes False／horizontal／progressive
   - framework : min4／max9／axes True／grid／positional_quadrant
4. 横断的原則4件 : 同ファイル「横断的な設計原則｜全パターン共通｜4件」

⚠️ 8/24 統括指示により、要素数・軸要否は「集約表（8/15）を正」とした。
   実装記録（8/9〜8/11）との相違2件は事例015に登録する。

作成: 2026-08-24（月）15:00連結②／AIスライド（実装領域）
"""

# =====================================================================
# ▼▼▼ v17 P2 追加ブロック（pyramid / sequence / framework） ▼▼▼
# P1ブロックのヘルパーを再利用する。既存 C-1〜C-3 には一切触れない。
# =====================================================================

# ---------------------------------------------------------------------
# v17-P2-0｜拡張パターン仕様（集約表 P2-1／P2-2／P2-3 行を正とする）
# ---------------------------------------------------------------------
DIAGRAM_PATTERN_SPEC.update({
    'pyramid': {
        'min_elements': 3, 'max_elements': 5, 'requires_axes': False,
        'direction': 'vertical', 'color_gradation': 'hierarchical',
        # 頂点30%→基層90%（原本 use「階層・優先順位を上下で表現」）
        'width_ratio_top': 0.30, 'width_ratio_bottom': 0.90,
    },
    'sequence': {
        'min_elements': 3, 'max_elements': 6, 'requires_axes': False,
        'direction': 'horizontal', 'color_gradation': 'progressive',
        'arrow_gap_px': 44,
    },
    'framework': {
        'min_elements': 4, 'max_elements': 9, 'requires_axes': True,
        'direction': 'grid', 'color_gradation': 'positional_quadrant',
        # 要素数→グリッド形状（決定論的｜原則③）
        'grid_map': {4: (2, 2), 5: (3, 2), 6: (3, 2),
                     7: (3, 3), 8: (3, 3), 9: (3, 3)},
        'axis_y_w': 72, 'axis_x_h': 46,
    },
})

# 階層色（hierarchical）｜段数別の色キー列（決定論的）
V17_PYRAMID_TIERS = {
    3: ['primary', 'midtone', 'light'],
    4: ['primary', 'secondary', 'midtone', 'light'],
    5: ['primary', 'secondary', 'midtone', 'light', 'lightest'],
}

# 進行色（progressive）｜段数別の色キー列
V17_SEQUENCE_TIERS = {
    3: ['primary', 'secondary', 'midtone'],
    4: ['primary', 'secondary', 'midtone', 'light'],
    5: ['primary', 'secondary', 'midtone', 'light', 'lightest'],
    6: ['primary', 'primary', 'secondary', 'midtone', 'light', 'lightest'],
}


# ---------------------------------------------------------------------
# v17-P2-1｜pyramid（ピラミッド）｜vertical / hierarchical / 3〜5段
# ---------------------------------------------------------------------
def draw_pyramid(slide, palette, data):
    """階層・優先順位を上下で表現する（頂点＝最重要）。

    Args:
        data : {'title': str,
                'levels': [{'label': str, 'score': int|None,
                            'description': str}, ...]}   # 3〜5・index0が頂点
    Returns:
        dict : 描画レポート

    仕様: 集約表 P2-1 行（min3／max5／vertical／hierarchical／原本shape=triangle）
    ⚠️ 台形は `add_freeform` を使わず、幅可変の矩形段で近似する（原則④）。
       `add_freeform` は python-pptx のバージョン依存があり、PowerPoint実機での
       再現性を担保できないため。段ごとの幅差で三角形の輪郭を表現する。
    範囲外は draw_category へフォールバック（原則①｜例外を投げない）。
    """
    spec = DIAGRAM_PATTERN_SPEC['pyramid']
    notes = []
    levels = _v17_normalize(data.get('levels', []))

    if not (spec['min_elements'] <= len(levels) <= spec['max_elements']):
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': lv.get('label', ''), 'score': lv.get('score'),
                            'description': str(lv.get('description', ''))}
                           for lv in levels],
        }, _fallback_from='pyramid')
        rep['notes'].insert(0, 'pyramid 段数 %d が %d〜%d の範囲外｜category へフォールバック'
                            % (len(levels), spec['min_elements'], spec['max_elements']))
        return rep

    _v17_title(slide, data.get('title', ''), palette)

    n = len(levels)
    tiers = V17_PYRAMID_TIERS[n]
    gap = 8
    avail = V17_AREA['body_bottom'] - V17_AREA['body_top']
    band_h = min((avail - gap * (n - 1)) / float(n), 104)
    block_h = band_h * n + gap * (n - 1)
    y0 = V17_AREA['body_top'] + (avail - block_h) / 2.0
    cx = V17_AREA['left'] + V17_AREA['width'] / 2.0

    r_top, r_bot = spec['width_ratio_top'], spec['width_ratio_bottom']
    for i, lv in enumerate(levels):
        # 段の幅：頂点 r_top → 基層 r_bot を線形に配分
        ratio = r_top + (r_bot - r_top) * (i / float(n - 1))
        w = V17_AREA['width'] * ratio
        x = cx - w / 2.0
        y = y0 + i * (band_h + gap)

        score = lv.get('score')
        fill_hex = _tier_fill(palette, score, tiers[i])
        fg = _text_color_on(fill_hex)
        _v17_card(slide, x, y, w, band_h, fill_hex, palette, radius_px=6)

        # 段ラベル｜頂点ほど大きく（優先度の視覚的強調）
        label_size = 20 if i == 0 else (18 if i == 1 else 16)
        pad = 20
        add_text(slide, x + pad, y + 8, w - pad * 2, str(lv.get('label', '')),
                 label_size, bold=True, color=fg, height_px=int(label_size * 1.5),
                 align=_pp_center())
        sub = []
        if score is not None:
            sub.append('%s%%' % score)
        desc = str(lv.get('description', ''))
        if desc:
            sub.append(desc)
        if sub:
            add_text(slide, x + pad, y + 8 + int(label_size * 1.5), w - pad * 2,
                     '　'.join(sub), 14, color=fg,
                     height_px=max(int(band_h - int(label_size * 1.5) - 16), 22),
                     align=_pp_center())

        # 優先度の序列を左外に添える（1が最上位）
        add_text(slide, V17_AREA['left'], y + band_h / 2.0 - 13, 40,
                 str(i + 1), 16, bold=True,
                 color=hex_to_rgb(palette['secondary']), height_px=26,
                 align=_pp_center())

    return _v17_report('pyramid', n, None, notes)


# ---------------------------------------------------------------------
# v17-P2-2｜sequence（順序）｜horizontal / progressive / 3〜6ステップ
# ---------------------------------------------------------------------
def draw_sequence(slide, palette, data):
    """ステップ・時系列を左→右で表現する。

    Args:
        data : {'title': str,
                'steps': [{'label': str, 'score': int|None,
                           'description': str}, ...]}    # 3〜6
    仕様: 集約表 P2-2 行（min3／max6／horizontal 固定／progressive）
    ⚠️ 原本 use「ステップ・時系列を**左→右**」により direction は horizontal 固定。
       集約表も direction=horizontal と規定するため、縦方向は実装しない。
    範囲外は draw_category へフォールバック（原則①）。
    """
    spec = DIAGRAM_PATTERN_SPEC['sequence']
    notes = []
    steps = _v17_normalize(data.get('steps', []))

    if not (spec['min_elements'] <= len(steps) <= spec['max_elements']):
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': st.get('label', ''), 'score': st.get('score'),
                            'description': str(st.get('description', ''))}
                           for st in steps],
        }, _fallback_from='sequence')
        rep['notes'].insert(0, 'sequence ステップ数 %d が %d〜%d の範囲外｜'
                            'category へフォールバック'
                            % (len(steps), spec['min_elements'], spec['max_elements']))
        return rep

    _v17_title(slide, data.get('title', ''), palette)

    n = len(steps)
    tiers = V17_SEQUENCE_TIERS[n]
    agap = spec['arrow_gap_px']
    card_w = (V17_AREA['width'] - agap * (n - 1)) / float(n)
    card_h = min(V17_AREA['body_bottom'] - V17_AREA['body_top'], 232)
    y = V17_AREA['body_top'] + \
        (V17_AREA['body_bottom'] - V17_AREA['body_top'] - card_h) / 2.0

    for i, st in enumerate(steps):
        x = V17_AREA['left'] + i * (card_w + agap)
        score = st.get('score')
        fill_hex = _tier_fill(palette, score, tiers[i])
        fg = _text_color_on(fill_hex)
        _v17_card(slide, x, y, card_w, card_h, fill_hex, palette)

        pad = 14
        add_text(slide, x + pad, y + 12, card_w - pad * 2, 'STEP %d' % (i + 1),
                 14, bold=True, color=fg, height_px=22)
        add_text(slide, x + pad, y + 40, card_w - pad * 2,
                 str(st.get('label', '')), 16, bold=True, color=fg,
                 height_px=52, line_height=1.3)
        if score is not None:
            add_text(slide, x + pad, y + 98, card_w - pad * 2, '%s%%' % score,
                     24, bold=True, color=fg, height_px=38)
        desc = str(st.get('description', ''))
        if desc:
            dtop = y + (142 if score is not None else 98)
            add_text(slide, x + pad, dtop, card_w - pad * 2, desc, 14, color=fg,
                     height_px=max(int(y + card_h - dtop - 10), 22),
                     line_height=1.4)

        # ステップ間の矢印（原則④｜python-pptx標準図形・adjustments 未使用）
        if i < n - 1:
            ax = x + card_w + 6
            aw = agap - 12
            ah = 26
            add_shape(slide, _mso_right_arrow(), ax, y + card_h / 2.0 - ah / 2.0,
                      aw, ah, fill=hex_to_rgb(palette['accent']))

    return _v17_report('sequence', n, None, notes)


# ---------------------------------------------------------------------
# v17-P2-3｜framework（フレームワーク）｜grid / positional_quadrant / 4〜9セル
#   ⚠️ 実装済11パターン中で唯一 requires_axes=True
# ---------------------------------------------------------------------
def _framework_tier(row, col, rows, cols):
    """セル位置から色キーを決定（positional_quadrant｜決定論的）。

    2x2 は原本 use「4象限マトリクス等」に従い象限別の意味付けを行う
    （右上＝最重要／左下＝最軽微）。それ以外は行位置ベース。
    """
    if (rows, cols) == (2, 2):
        return {(0, 1): 'primary', (0, 0): 'secondary',
                (1, 1): 'midtone', (1, 0): 'light'}.get((row, col), 'midtone')
    if row == 0:
        return 'primary'
    if row == rows - 1:
        return 'light'
    return 'secondary'


def draw_framework(slide, palette, data):
    """2軸マトリクス（4象限等）で構造を提示する。

    Args:
        data : {'title': str,
                'axis_x_label': str, 'axis_y_label': str,     # ⚠️ 必須
                'axis_x_low': str, 'axis_x_high': str,
                'axis_y_low': str, 'axis_y_high': str,
                'cells': [{'row': int, 'col': int, 'label': str,
                           'score': int|None, 'items': [str]}, ...]}  # 4〜9
    仕様: 集約表 P2-3 行（min4／max9／requires_axes=True／grid／positional_quadrant）
    ⚠️ 集約表は「4〜9の範囲」を規定するため、5・7・8セルも受け付け、
       要素数からグリッド形状を決定論的に自動選定する（実装記録の固定値縛りは採らない）。
    ⚠️ 軸ラベル欠落時も例外を投げず draw_category へフォールバックする（原則①）。
    """
    spec = DIAGRAM_PATTERN_SPEC['framework']
    notes = []
    cells = _v17_normalize(data.get('cells', []))
    n = len(cells)

    def _fallback(reason):
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': c.get('label', ''), 'score': c.get('score'),
                            'description': ''} for c in cells],
        }, _fallback_from='framework')
        # ⚠️ v17.1.0 修正：ローカル notes ではなく戻り値の notes に記録する。
        # 旧実装はローカル notes に append した後 draw_category の新しい report を
        # 返していたため、フォールバック理由が呼び出し側に届かなかった。
        rep['notes'].insert(0, reason)
        return rep

    if not (spec['min_elements'] <= n <= spec['max_elements']):
        return _fallback('framework セル数 %d が %d〜%d の範囲外｜category へフォールバック'
                         % (n, spec['min_elements'], spec['max_elements']))

    ax_label = str(data.get('axis_x_label') or '').strip()
    ay_label = str(data.get('axis_y_label') or '').strip()
    if not ax_label or not ay_label:
        return _fallback('framework は requires_axes=True｜軸ラベル欠落のため '
                         'category へフォールバック（例外は送出しない）')

    _v17_title(slide, data.get('title', ''), palette)

    cols, rows = spec['grid_map'][n]
    notes.append('セル数 %d → グリッド %dx%d を自動選定' % (n, cols, rows))

    ay_w, ax_h = spec['axis_y_w'], spec['axis_x_h']
    gx = V17_AREA['left'] + ay_w
    gy = V17_AREA['body_top']
    gw = V17_AREA['width'] - ay_w
    gh = V17_AREA['body_bottom'] - gy - ax_h
    gap = 10
    cw = (gw - gap * (cols - 1)) / float(cols)
    ch = (gh - gap * (rows - 1)) / float(rows)

    # 軸ラベル（framework固有｜requires_axes=True）
    add_text(slide, gx, V17_AREA['body_bottom'] - ax_h + 10, gw,
             '%s ← %s → %s' % (data.get('axis_x_low', '低'), ax_label,
                               data.get('axis_x_high', '高')),
             14, bold=True, color=hex_to_rgb(palette['secondary']),
             height_px=24, align=_pp_center())
    add_paragraph_box(slide, V17_AREA['left'], gy + gh / 2.0 - 56, ay_w,
                      [{'text': str(data.get('axis_y_high', '高')), 'size': 14,
                        'bold': True, 'align': _pp_center()},
                       {'text': '↑', 'size': 14, 'align': _pp_center()},
                       {'text': ay_label, 'size': 14, 'bold': True,
                        'align': _pp_center()},
                       {'text': '↓', 'size': 14, 'align': _pp_center()},
                       {'text': str(data.get('axis_y_low', '低')), 'size': 14,
                        'bold': True, 'align': _pp_center()}],
                      height_px=112, default_color=hex_to_rgb(palette['secondary']),
                      line_height=1.15)

    # セル（row/col 指定がなければ左上から順に充填）
    used = set()
    for idx, c in enumerate(cells):
        r = c.get('row')
        col = c.get('col')
        if not isinstance(r, int) or not isinstance(col, int) \
                or not (0 <= r < rows and 0 <= col < cols) or (r, col) in used:
            r, col = idx // cols, idx % cols
        used.add((r, col))

        x = gx + col * (cw + gap)
        y = gy + r * (ch + gap)
        score = c.get('score')
        fill_hex = _tier_fill(palette, score, _framework_tier(r, col, rows, cols))
        fg = _text_color_on(fill_hex)
        _v17_card(slide, x, y, cw, ch, fill_hex, palette)

        pad = 14
        add_text(slide, x + pad, y + 10, cw - pad * 2, str(c.get('label', '')),
                 16, bold=True, color=fg, height_px=26)
        top = y + 38
        if score is not None:
            add_text(slide, x + pad, top, cw - pad * 2, '%s%%' % score,
                     20, bold=True, color=fg, height_px=32)
            top += 34
        items = [str(i) for i in (c.get('items') or [])][:3]
        if items and top + 22 <= y + ch - 8:
            add_paragraph_box(slide, x + pad, top, cw - pad * 2,
                              [{'text': '・' + t, 'size': 14} for t in items],
                              height_px=max(int(y + ch - top - 8), 22),
                              default_color=fg, line_height=1.35,
                              space_after_pt=2)

    return _v17_report('framework', n, None, notes)


# ---------------------------------------------------------------------
# v17-P2-4｜ディスパッチ表への登録（P1の draw_pattern を置き換える）
# ---------------------------------------------------------------------
def draw_pattern(slide, pattern_key, palette, data):
    """パターンキーで描画関数を振り分ける（P1 3種＋P2 3種＝6種に対応）。

    未実装キー（funnel / timeline / contrast / cycle / network / integration）は
    category に退避し、notes に理由を記録する（例外は送出しない）。
    """
    table = {
        'category':   draw_category,
        'breakdown':  draw_breakdown,
        'comparison': draw_comparison,
        'pyramid':    draw_pyramid,
        'sequence':   draw_sequence,
        'framework':  draw_framework,
    }
    fn = table.get(pattern_key)
    if fn is None:
        src = (data.get('categories') or data.get('items')
               or data.get('components') or data.get('levels')
               or data.get('steps') or data.get('cells') or [])
        rep = draw_category(slide, palette, {
            'title': data.get('title', ''),
            'categories': [{'label': str(x.get('label', '')) if isinstance(x, dict) else str(x),
                            'score': x.get('score') if isinstance(x, dict) else None,
                            'description': ''} for x in src],
        }, _fallback_from=pattern_key)
        rep['notes'].append('パターン "%s" は v17 P3 で実装予定（P1/P2 の対象外）'
                            % pattern_key)
        return rep
    return fn(slide, palette, data)


def _pp_center():
    from pptx.enum.text import PP_ALIGN
    return PP_ALIGN.CENTER


def _mso_right_arrow():
    from pptx.enum.shapes import MSO_SHAPE
    return MSO_SHAPE.RIGHT_ARROW

# =====================================================================
# ▲▲▲ v17 P2 追加ブロック ここまで ▲▲▲
# =====================================================================
