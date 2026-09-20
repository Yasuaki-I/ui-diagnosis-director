#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_labels.py｜⚠️ 廃止（2026-09-22 統括判定 第2便 判定②）

本スクリプトは `verify_render.py` に統合され、廃止された。

■ 廃止理由

builder のソースから「固定文字列 × 固定枠幅」の add_text を静的に抽出する方式は、
f-string・変数・関数戻り値で組み立てられる文字列を原理的に検出できない。

    add_text(slide, x, y, 50, f'{sv}/5', 14, bold=score_bold)   ← 検出不能
    add_text(slide, 56, y, 180, cat_disp, 14, bold=True)        ← 検出不能

2026-09-22 の実機目視で発見された2件（太字スコア値 55.5px>50px／
項目名 186.7px>180px）はいずれもこの死角にあり、本スクリプトは
「NG 0件」を返していた。

■ 移行先

    python3 verify_render.py 生成物.pptx                        # 全テキスト段落を実測
    python3 verify_render.py --check '次の一手' 90 14 --bold     # 単発（旧 --check 相当）

`verify_render.py` は生成された .pptx を開いて実際に描かれた文字列を測るため、
組み立て方によらず漏れない。

■ 関連

PLAIN J-7-6-2 B / B-1（検査対象を生成物実測へ転換）
SCRIPTS_REGISTRY_20260922.md（常設スクリプト一覧）
"""
import sys

MSG = """\
⚠️ verify_labels.py は廃止されました（2026-09-22 統括判定 第2便 判定②）。

  静的解析では f-string・変数・関数戻り値の文字列を検出できないため、
  生成済 .pptx を実測する verify_render.py に統合されました。

  移行先:
    python3 verify_render.py 生成物.pptx
    python3 verify_render.py --check '次の一手' 90 14 --bold
"""

if __name__ == '__main__':
    sys.stderr.write(MSG)
    sys.exit(2)
