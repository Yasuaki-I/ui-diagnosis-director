#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v17 P3｜視覚的実機確認サンプル生成（第16条 細則9 対応）

================================================================================
■ 目的
================================================================================
運用ルール第16条 細則9（2026-08-24 統括判定⑤により新設）は、

  「新規追加された描画パターンは、少なくとも1回は意図的に呼び出した実機出力での
    目視確認を段階4の必須条件とする。テストのPASSおよび細則8の主要ブロック存在
    確認は、これを代替しない。」

と規定する。本スクリプトは v17 P3 の5パターン
（funnel / cycle / contrast / timeline / network）を「意図的に呼び出した」
実機出力を1ファイルで生成する。

================================================================================
■ 入江さんへの依頼（所要3分）
================================================================================
  1. GPTs のチャットに本ファイルを添付する（Knowledge 登録は不要）
  2. 「添付したスクリプトをそのまま実行し、標準出力を全文貼り付け、
      生成された .pptx をダウンロードできる形で提示してください」と伝える
  3. 出力された `v17_p3_visual_verify.pptx`（5枚）を PowerPoint で開く
  4. 5枚それぞれに図形が描かれているかを目視確認する
     ・1枚目 funnel  ：上が広く下が狭い5段（絞り込み）
     ・2枚目 cycle   ：4つの角丸ボックスが円環状に並び、間を矢印が結ぶ
                       ＋円環中央に「PDCA」
     ・3枚目 contrast：左右2分割（左が淡色／右が濃色）＋中央に区切り線
     ・4枚目 timeline：上部に時間軸ラベル（2026/8〜）＋水平バー＋丸ドット
                       ＋下部にカード5枚
     ・5枚目 network ：上下3段のノード＋ノード間を結ぶ直線
  5. 結果（正常／異常）をご連絡ください

⚠️ 白紙・図形なし・文字化け・図形の重なり・軸ラベル欠落があれば、
   それが検出事項です。そのままご報告ください。

================================================================================
■ 版数
================================================================================
  対象ビルダー : 03_pptx_builder.py（v17.2.0）
  作成         : 2026-08-25（火）18:00連結③｜AIスライド（実装領域）
  根拠         : 第16条 細則9（8/24 統括判定⑤）／事例017
"""

import os
import sys

BUILDER_CANDIDATES = [
    '03_pptx_builder.py',                          # 配布環境の正式名
    '/mnt/data/03_pptx_builder.py',                # GPTs 実行環境
    '03_pptx_builder_v17_20_20260825.py',          # AIドライブ原本名
]


def _load_builder():
    """ビルダーを exec で読み込み、名前空間を返す。"""
    for path in BUILDER_CANDIDATES:
        if os.path.exists(path):
            ns = {'__name__': 'builder_v17'}
            with open(path, encoding='utf-8') as f:
                code = f.read()
            exec(compile(code, path, 'exec'), ns)
            print('[OK] ビルダー読込: %s' % path)
            print('     __version__      = %r' % ns.get('__version__'))
            print('     __version_date__ = %r' % ns.get('__version_date__'))
            return ns
    print('[NG] ビルダーが見つかりません。以下のいずれかを同じ場所に置いてください:')
    for p in BUILDER_CANDIDATES:
        print('     - %s' % p)
    sys.exit(1)


# ------------------------------------------------------------------------------
# サンプルデータ｜⚠️ 集約表の要素数規定に収まる値のみを使う
#   funnel   : min3 / max6  → 5段
#   cycle    : min3 / max6  → 4工程（キーは 'phases'）
#   contrast : min2 / max2  → 2側（固定）
#   timeline : min3 / max7  → 5マイルストーン（⚠️ requires_axes=True）
#   network  : min3 / max7  → 6ノード（3階層）
# ------------------------------------------------------------------------------
FUNNEL_DATA = {
    'title': '申込までの絞り込み｜各段階の到達率',
    'stages': [
        {'label': 'サイト訪問',   'score': 100, 'description': '全流入の母数'},
        {'label': '商品一覧閲覧', 'score': 68,  'description': 'カテゴリから商品を探す'},
        {'label': '詳細ページ',   'score': 44,  'description': '仕様・価格を比較する'},
        {'label': 'カート投入',   'score': 21,  'description': '⚠️ ここで最も落ちる'},
        {'label': '申込完了',     'score': 12,  'description': 'フォーム入力を完了'},
    ],
}

CYCLE_DATA = {
    'title': 'UI改善の運用サイクル｜四半期ごとに反復',
    'cycle_name': 'PDCA',
    # ⚠️ cycle の要素キーは 'phases'（funnel の 'stages' とは別）
    'phases': [
        {'label': '計画',   'score': 78, 'description': '改善仮説を立てる'},
        {'label': '実装',   'score': 65, 'description': '施策をリリースする'},
        {'label': '計測',   'score': 52, 'description': '指標の変化を見る'},
        {'label': '見直し', 'score': 34, 'description': '⚠️ ここが形骸化しやすい'},
    ],
}

CONTRAST_DATA = {
    'title': '改善前後の対比｜ファーストビュー',
    'sides': [
        {'label': '改善前（Before）', 'score': 38,
         'items': ['CTAが画面外にある', '価値提案が3秒で伝わらない',
                   '主要導線が2階層下']},
        {'label': '改善後（After）',  'score': 82,
         'items': ['CTAをファーストビュー内へ', '価値提案を1行で明示',
                   '主要導線を最上部へ集約']},
    ],
}

TIMELINE_DATA = {
    'title': '改善ロードマップ｜四半期別マイルストーン',
    # ⚠️ requires_axes=True｜axis_label と各 axis は必須
    'axis_label': '実施時期',
    'milestones': [
        {'label': '現状診断',     'axis': '2026/8',  'score': 100,
         'description': '10項目のUI診断を完了'},
        {'label': 'CTA改善',      'axis': '2026/9',  'score': 72,
         'description': 'ファーストビュー内へ移設'},
        {'label': '情報設計刷新', 'axis': '2026/10', 'score': 55,
         'description': 'カテゴリ階層を4→2に'},
        {'label': 'フォーム改善', 'axis': '2026/11', 'score': 41,
         'description': '必須項目の明示と入力補助'},
        {'label': '効果検証',     'axis': '2026/12', 'score': 30,
         'description': '⚠️ 予算未確定のため要調整'},
    ],
}

NETWORK_DATA = {
    'title': '離脱要因の関係構造｜上位課題から派生要因へ',
    'nodes': [
        {'id': 'root', 'label': '申込完了率の低下', 'depth': 0, 'score': 31},
        {'id': 'fv',   'label': 'ファーストビュー', 'depth': 1, 'score': 38},
        {'id': 'ia',   'label': '情報設計',         'depth': 1, 'score': 64},
        {'id': 'fm',   'label': '入力フォーム',     'depth': 1, 'score': 71},
        {'id': 'cta',  'label': 'CTAが画面外',      'depth': 2, 'score': 28},
        {'id': 'deep', 'label': '階層が深い',       'depth': 2, 'score': 55},
    ],
    'edges': [
        {'from': 'root', 'to': 'fv'},
        {'from': 'root', 'to': 'ia'},
        {'from': 'root', 'to': 'fm'},
        {'from': 'fv',   'to': 'cta'},
        {'from': 'ia',   'to': 'deep'},
    ],
}

SAMPLES = [
    ('funnel',   FUNNEL_DATA,   '上が広く下が狭い5段（絞り込み）'),
    ('cycle',    CYCLE_DATA,    '角丸ボックス4つが円環状＋矢印＋中央にPDCA'),
    ('contrast', CONTRAST_DATA, '左右2分割（左が淡色／右が濃色）＋区切り線'),
    ('timeline', TIMELINE_DATA, '時間軸ラベル＋水平バー＋丸ドット＋カード5枚'),
    ('network',  NETWORK_DATA,  '上下3段のノード＋ノード間を結ぶ直線'),
]

OUT_PATH = 'v17_p3_visual_verify.pptx'
THEME_ID = 'Blue'


def _resolve_palette(ns):
    """テーマパレットを取得する（テーマIDの表記差に耐える）。"""
    get_theme_palette = ns['get_theme_palette']
    palette_dict = ns.get('DIGITAL_AGENCY_PALETTE') or {}
    for key in [THEME_ID] + list(palette_dict.keys()):
        try:
            pal = get_theme_palette(key)
            if pal:
                print('[OK] テーマパレット: %r' % key)
                return pal
        except Exception:
            continue
    raise RuntimeError('テーマパレットを取得できませんでした')


def main():
    ns = _load_builder()

    # 第16条 細則8｜主要ブロックの存在確認
    print('\n--- 第16条 細則8｜主要ブロックの存在確認 ---')
    required = ('DIAGRAM_PATTERNS', 'DIGITAL_AGENCY_PALETTE',
                'DIAGRAM_PATTERN_SPEC', 'draw_pattern', 'add_diagram_slide',
                'draw_funnel', 'draw_cycle', 'draw_contrast',
                'draw_timeline', 'draw_network')
    for name in required:
        mark = 'OK' if name in ns else 'NG'
        print('  [%s] %s' % (mark, name))
        if name not in ns:
            print('\n[NG] 主要ブロックが欠落しています。'
                  'ビルダーの差し替えが完了していない可能性があります。')
            sys.exit(1)

    spec = ns['DIAGRAM_PATTERN_SPEC']
    print('  [OK] DIAGRAM_PATTERN_SPEC = %d 件（11パターン想定）' % len(spec))
    print('  [OK] DIAGRAM_PATTERNS     = %d 件（原本12件・無改変想定）'
          % len(ns['DIAGRAM_PATTERNS']))

    prs = ns['create_presentation']()
    palette = _resolve_palette(ns)
    add_diagram_slide = ns['add_diagram_slide']

    print('\n--- 第16条 細則9｜意図的呼び出しによる描画 ---')
    total = len(SAMPLES)
    ng_count = 0
    for i, (pattern_key, data, expect) in enumerate(SAMPLES, start=1):
        slide, report = add_diagram_slide(
            prs, pattern_key, palette, data,
            page_num=i, total=total,
        )
        shapes = len(slide.shapes)
        fb = (report or {}).get('fallback_from')
        notes = (report or {}).get('notes') or []
        status = 'OK' if shapes > 1 and not fb else 'CHECK'
        if status != 'OK':
            ng_count += 1
        print('  [%s] p%d  %-9s shapes=%-3d  期待=%s'
              % (status, i, pattern_key, shapes, expect))
        if fb:
            print('        ⚠️ フォールバック発生: %s' % fb)
        for n in notes:
            print('        note: %s' % n)

    prs.save(OUT_PATH)
    size = os.path.getsize(OUT_PATH)
    print('\n[OK] 出力: %s（%s バイト｜%d枚）'
          % (OUT_PATH, format(size, ','), total))
    print('\n================================================================')
    print(' 次の手順｜PowerPoint で開いて5枚を目視確認してください')
    for i, (k, _d, expect) in enumerate(SAMPLES, start=1):
        print('  %d枚目 %-9s: %s' % (i, k, expect))
    print(' ⚠️ 白紙・図形なし・文字化け・重なり・軸ラベル欠落があれば')
    print('    それが検出事項です')
    print('================================================================')
    if ng_count:
        print('\n⚠️ CHECK が %d 件あります。上記 note をご報告ください。' % ng_count)


if __name__ == '__main__':
    main()
