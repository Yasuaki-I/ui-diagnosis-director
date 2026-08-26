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

と規定する。

本スクリプトは v17 P3 の5パターン（funnel / cycle / contrast / timeline / network）を
「意図的に呼び出した」実機出力を1ファイルで生成する。

■ 背景（判断原理14）
テスト169項目 全PASS・主要ブロック存在確認（細則8）を満たしていても、
「実際に描画されるか」は別の検証である。P3 5種は本サンプルの目視確認を
もって初めて段階4（反映済）の要件を満たす。

なお AIスライド側では 8/25 に自己目視を実施済（5種すべて正常描画）。
本サンプルは配布環境（GPTs）での再現を確認するためのものである。

================================================================================
■ 入江さんへの依頼（所要3分）
================================================================================
  1. 本ファイルを `03_pptx_builder.py` と同じ場所に置く
  2. `python v17_p2_visual_verify_sample_20260824.py` を実行する
  3. 出力された `v17_p3_visual_verify.pptx`（5枚）を PowerPoint で開く
  4. 5枚それぞれに図形が描かれているかを目視確認する
     ・1枚目：ファネル（5段／上が広く下が狭い＝絞り込み）
     ・2枚目：サイクル（4枚のカードが円環状＋矢印＋中央に「PDCA」）
     ・3枚目：対比（左右2枚のカード／左が赤・右が濃灰）
     ・4枚目：タイムライン（横一本の軸＋5個の丸＋上下交互のカード＋月ラベル）
     ・5枚目：ネットワーク（上下3階層のノード＋それを結ぶ線）
  5. 結果（正常／異常）をご連絡ください

⚠️ 白紙・図形なし・文字化け・図形の重なりがあれば、それが検出事項です。
⚠️ 判定はAIスライドの自己申告ではなく、この目視結果を根拠とします。

================================================================================
■ 版数
================================================================================
  対象ビルダー : 03_pptx_builder.py（v17.2.0 / 257,928B）
  作成         : 2026-08-25（火）｜AIスライド（実装領域）
  根拠         : 第16条 細則9（8/24 統括判定⑤）／判断原理14
"""

import os
import sys

# ------------------------------------------------------------------------------
# ビルダー読み込み
#   配布環境では `03_pptx_builder.py`（数字始まりのため import 文が使えない）
# ------------------------------------------------------------------------------
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
#   funnel    : min3 / max6  → 5段
#   cycle     : min3 / max6  → 4段
#   contrast  : min2 / max2  → 2固定
#   timeline  : min3 / max7  → 5マイルストーン（requires_axes=True）
#   network   : min3 / max7  → 5ノード / 4エッジ
# ------------------------------------------------------------------------------
FUNNEL_DATA = {
    'title': '流入から申込までの絞り込み',
    'stages': [
        {'label': 'サイト訪問',     'score': 100, 'description': '月間 24,800 セッション'},
        {'label': '商品一覧の閲覧', 'score': 62,  'description': 'カテゴリから商品を探す'},
        {'label': '詳細ページ到達', 'score': 41,  'description': '仕様・価格を比較する'},
        {'label': 'カート投入',     'score': 18,  'description': 'ここで最も絞られる'},
        {'label': '申込完了',       'score': 7,   'description': 'フォーム入力を完了する'},
    ],
}

CYCLE_DATA = {
    'title': 'UI改善の運用サイクル',
    'cycle_name': 'PDCA',
    'phases': [
        {'label': '計測', 'score': 82, 'description': 'GA4で行動を把握'},
        {'label': '診断', 'score': 64, 'description': '12観点でスコア化'},
        {'label': '改修', 'score': 38, 'description': '着手が遅れがち'},
        {'label': '検証', 'score': 71, 'description': 'A/Bで効果を確認'},
    ],
}

# ⚠️ contrast の各サイドは 'description' ではなく 'items': [str]（最大5件）
CONTRAST_DATA = {
    'title': 'ファーストビュー改修の前後',
    'sides': [
        {'label': '改修前', 'score': 32,
         'items': ['CTAが画面外', '価値提案が3秒で伝わらない',
                   '離脱率 68%', '直帰後の再訪なし']},
        {'label': '改修後', 'score': 81,
         'items': ['CTAを画面内に配置', '見出しで価値を明示',
                   '離脱率 41%へ改善', '再訪率が2.3倍']},
    ],
}

# ⚠️ timeline は requires_axes=True｜axis_label と各 axis が必須
TIMELINE_DATA = {
    'title': 'UI改善の実施計画',
    'axis_label': '2026年度 下期',
    'milestones': [
        {'label': '現状診断',   'axis': '8月',  'score': 82, 'description': '12観点のスコア化'},
        {'label': '優先度確定', 'axis': '9月',  'score': 66, 'description': '効果×工数で選定'},
        {'label': 'FV改修',     'axis': '10月', 'score': 38, 'description': '工数が最大'},
        {'label': '導線整理',   'axis': '11月', 'score': 54, 'description': '階層を4→2クリック'},
        {'label': '効果検証',   'axis': '12月', 'score': 73, 'description': 'A/Bテストで確認'},
    ],
}

NETWORK_DATA = {
    'title': '離脱要因の関係構造',
    'nodes': [
        {'id': 'r', 'label': '申込率の低下',   'score': 31},
        {'id': 'a', 'label': 'FVの訴求不足',   'score': 38},
        {'id': 'b', 'label': '導線の複雑さ',   'score': 52},
        {'id': 'c', 'label': 'CTAが画面外',    'score': 29},
        {'id': 'd', 'label': '階層が4クリック', 'score': 61},
    ],
    'edges': [
        {'from': 'r', 'to': 'a'}, {'from': 'r', 'to': 'b'},
        {'from': 'a', 'to': 'c'}, {'from': 'b', 'to': 'd'},
    ],
}

SAMPLES = [
    ('funnel',   FUNNEL_DATA,   'ファネル（5段／上が広く下が狭い）'),
    ('cycle',    CYCLE_DATA,    'サイクル（4カードが円環＋矢印＋中央PDCA）'),
    ('contrast', CONTRAST_DATA, '対比（左右2枚／左が赤・右が濃灰）'),
    ('timeline', TIMELINE_DATA, 'タイムライン（軸＋5個の丸＋上下交互カード）'),
    ('network',  NETWORK_DATA,  'ネットワーク（3階層のノード＋結ぶ線）'),
]

OUT_PATH = 'v17_p3_visual_verify.pptx'
THEME_ID = 'v17_blue'


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
        print('  [%s] %s' % ('OK' if name in ns else 'NG', name))
        if name not in ns:
            print('\n[NG] 主要ブロックが欠落しています。'
                  'ビルダーの差し替えが完了していない可能性があります。')
            sys.exit(1)

    prs = ns['create_presentation']()
    palette = _resolve_palette(ns)
    add_diagram_slide = ns['add_diagram_slide']

    print('\n--- 第16条 細則9｜意図的呼び出しによる描画 ---')
    total = len(SAMPLES)
    for i, (pattern_key, data, expect) in enumerate(SAMPLES, start=1):
        slide, report = add_diagram_slide(
            prs, pattern_key, palette, data, page_num=i, total=total)
        shapes = len(slide.shapes)
        fb = (report or {}).get('fallback_from')
        status = 'OK' if shapes > 3 and not fb else 'CHECK'
        print('  [%s] p%d  %-9s shapes=%-3d  期待=%s'
              % (status, i, pattern_key, shapes, expect))
        if fb:
            print('        ⚠️ フォールバック発生: %s' % fb)
        for n in (report or {}).get('notes') or []:
            print('        note: %s' % n)

    prs.save(OUT_PATH)
    print('\n[OK] 出力: %s（%s バイト｜%d枚）'
          % (OUT_PATH, format(os.path.getsize(OUT_PATH), ','), total))
    print('\n================================================================')
    print(' 次の手順｜PowerPoint で開いて5枚を目視確認してください')
    print('  1枚目: ファネル（5段／上が広く下が狭い＝絞り込み）')
    print('  2枚目: サイクル（4カードが円環状＋矢印＋中央に PDCA）')
    print('  3枚目: 対比（左右2枚のカード／左が赤・右が濃灰）')
    print('  4枚目: タイムライン（横一本の軸＋5個の丸＋上下交互のカード＋月ラベル）')
    print('  5枚目: ネットワーク（上下3階層のノード＋それを結ぶ線）')
    print(' ⚠️ 白紙・図形なし・文字化け・重なりがあれば、それが検出事項です')
    print('================================================================')


if __name__ == '__main__':
    main()
