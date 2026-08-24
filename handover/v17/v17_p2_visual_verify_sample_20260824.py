#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v17 P2｜視覚的実機確認サンプル生成（第16条 細則9 対応）

================================================================================
■ 目的
================================================================================
運用ルール第16条 細則9（2026-08-24 統括判定⑤により新設）は、

  「新規追加された描画パターンは、少なくとも1回は意図的に呼び出した実機出力での
    目視確認を段階4の必須条件とする。テストのPASSおよび細則8の主要ブロック存在
    確認は、これを代替しない。」

と規定する。

本スクリプトは v17 P2 の3パターン（pyramid / sequence / framework）を
「意図的に呼び出した」実機出力を1ファイルで生成する。

■ 背景（事例017）
v17.1.0 の実機出力7枚は C-1（スコアカード）・C-2（改善提案リスト）・
C-3（ビジュアル診断ボード）で構成され、P2の3パターン自体を描画したスライドが
一枚も含まれていなかった。テスト118項目 全PASS・XML同一性確認・主要ブロック
存在確認（細則8）を満たしていても、「実際に描画されるか」は未検証である。

================================================================================
■ 入江さんへの依頼（所要3分）
================================================================================
  1. 本ファイルを `03_pptx_builder.py` と同じ場所に置く
  2. `python v17_p2_visual_verify_sample_20260824.py` を実行する
  3. 出力された `v17_p2_visual_verify.pptx`（3枚）を PowerPoint で開く
  4. 3枚それぞれに図形が描かれているかを目視確認する
     ・1枚目：ピラミッド（4段の階層／上が狭く下が広い）
     ・2枚目：ステップ（5個の矢印が左→右に並ぶ）
     ・3枚目：フレームワーク（2x2の4象限＋縦横の軸ラベル＋各象限の箇条項目）
  5. 結果（正常／異常）をご連絡ください

⚠️ 白紙・図形なし・文字化け・図形の重なりがあれば、それが検出事項です。
⚠️ 判定はAIスライドの自己申告ではなく、この目視結果を根拠とします。

================================================================================
■ 版数
================================================================================
  対象ビルダー : 03_pptx_builder.py（v17.1.0 / 223,802B）
  作成         : 2026-08-24（月）｜AIスライド（実装領域）
  根拠         : 第16条 細則9（8/24 統括判定⑤）／事例017
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
    '03_pptx_builder_v17_10_20260824.py',          # AIドライブ原本名
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
#   pyramid   : min3 / max5  → 4段
#   sequence  : min3 / max7  → 5ステップ
#   framework : min4 / max9  → 4セル（2x2 自動選定）
# ------------------------------------------------------------------------------
PYRAMID_DATA = {
    'title': 'UI改善の優先順位｜上位が最重要',
    'levels': [
        {'label': 'ファーストビュー', 'score': 38,
         'description': 'CTAが画面外にあり最初の3秒で価値が伝わらない'},
        {'label': '情報設計',         'score': 64,
         'description': 'カテゴリ階層が深く目的ページまで4クリック必要'},
        {'label': '入力フォーム',     'score': 71,
         'description': '必須項目の表示が弱く離脱が発生している'},
        {'label': '表示速度',         'score': 82,
         'description': '画像最適化済み。現状は許容範囲'},
    ],
}

SEQUENCE_DATA = {
    'title': 'ユーザー行動フロー｜流入から申込まで',
    'steps': [
        {'label': '流入',       'score': 78, 'description': '検索・広告からの着地'},
        {'label': '一覧閲覧',   'score': 66, 'description': 'カテゴリから商品を探す'},
        {'label': '詳細確認',   'score': 52, 'description': '仕様・価格を比較する'},
        {'label': 'カート投入', 'score': 34, 'description': '⚠️ ここで最も離脱する'},
        {'label': '申込完了',   'score': 41, 'description': 'フォーム入力を完了する'},
    ],
}

FRAMEWORK_DATA = {
    'title': '改善施策マトリクス｜効果 × 工数',
    'axis_x_label': '実装工数',
    'axis_x_low': '小',
    'axis_x_high': '大',
    'axis_y_label': '改善効果',
    'axis_y_low': '低',
    'axis_y_high': '高',
    # ⚠️ framework の各セルは 'description' ではなく 'items': [str] を使う
    #    （draw_framework の docstring 記載の仕様。最大3件まで箇条描画される）
    'cells': [
        {'row': 0, 'col': 0, 'label': 'CTA位置の変更',   'score': 88,
         'items': ['即着手', '効果大・工数小']},
        {'row': 0, 'col': 1, 'label': '検索機能の追加',   'score': 62,
         'items': ['計画的に実施', '効果大・工数大']},
        {'row': 1, 'col': 0, 'label': 'ラベル文言の調整', 'score': 55,
         'items': ['余力で実施', '効果小・工数小']},
        {'row': 1, 'col': 1, 'label': '会員基盤の刷新',   'score': 31,
         'items': ['見送り', '効果小・工数大']},
    ],
}

# 3パターンの出力定義（描画順＝ページ順）
SAMPLES = [
    ('pyramid',   PYRAMID_DATA,   'ピラミッド（4段｜上が狭く下が広い）'),
    ('sequence',  SEQUENCE_DATA,  'ステップ（5個の矢印が左→右）'),
    ('framework', FRAMEWORK_DATA, 'フレームワーク（2x2の4象限＋軸ラベル＋箇条項目）'),
]

OUT_PATH = 'v17_p2_visual_verify.pptx'
THEME_ID = 'v17_blue'   # 既定テーマ。get_theme_palette が受け付ける最初の候補で代替する


def _resolve_palette(ns):
    """テーマパレットを取得する（テーマIDの表記差に耐える）。"""
    get_theme_palette = ns['get_theme_palette']
    palette_dict = ns.get('DIGITAL_AGENCY_PALETTE') or {}
    candidates = [THEME_ID] + list(palette_dict.keys())
    for key in candidates:
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
    for name in ('DIAGRAM_PATTERNS', 'DIGITAL_AGENCY_PALETTE',
                 'DIAGRAM_PATTERN_SPEC', 'draw_pattern', 'add_diagram_slide',
                 'draw_pyramid', 'draw_sequence', 'draw_framework'):
        mark = 'OK' if name in ns else 'NG'
        print('  [%s] %s' % (mark, name))
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
            prs, pattern_key, palette, data,
            page_num=i, total=total,
        )
        shapes = len(slide.shapes)
        fb = (report or {}).get('fallback_from')
        notes = (report or {}).get('notes') or []
        status = 'OK' if shapes > 1 and not fb else 'CHECK'
        print('  [%s] p%d  %-10s shapes=%-3d  期待=%s'
              % (status, i, pattern_key, shapes, expect))
        if fb:
            print('        ⚠️ フォールバック発生: %s' % fb)
        for n in notes:
            print('        note: %s' % n)

    prs.save(OUT_PATH)
    size = os.path.getsize(OUT_PATH)
    print('\n[OK] 出力: %s（%s バイト｜%d枚）' % (OUT_PATH, format(size, ','), total))
    print('\n================================================================')
    print(' 次の手順｜PowerPoint で開いて3枚を目視確認してください')
    print('  1枚目: ピラミッド（4段の階層／上が狭く下が広い）')
    print('  2枚目: ステップ（5個の矢印が左→右に並ぶ）')
    print('  3枚目: フレームワーク（2x2の4象限＋縦横の軸ラベル＋各象限の箇条項目）')
    print(' ⚠️ 白紙・図形なし・文字化け・重なりがあれば、それが検出事項です')
    print('================================================================')


if __name__ == '__main__':
    main()
