#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ヘッドメッセージ 文型検査 v1（順2｜うちた式 工程1・工程5 の機械判定）
====================================================================
発行: 2026-09-11（金）｜AIスライド（実装領域）
根拠: 2026-09-10 19:00 統括判定 §1（順2 着手承認｜「検査のみ」であり
      細則3のキュー登録を経ず即日着手してよい）
分析: uchita_gap_analysis_20260910.md §2-2（判断語率 55% 対 規定なし）

■ 本検査の目的
  Instructions は従来「文字数上限」（所見40字／総評17字 等）のみを規定し、
  「文の形」を規定していなかった。記事の指摘（「原因は指示の側で、長さでなく
  文の形を規定した」）に対応し、ヘッドメッセージの文型を機械判定する。

■ 判定7項目（うちた式 工程5 の検査項目に対応）
  H1: 判断語を持つか（提言・要請・限定・評価・因果・可能・帰結のいずれか）
  H2: 二段構え「A。B」でないか（文末以外の句点）
  H3: 体言止め／題名でないか（「〜の分析」「〜について」等）
  H4: ダッシュ（——／—）を使っていないか
  H5: 自己言及（本ページ／この1枚 等）がないか
  H6: 要素の数え上げ（課題は3つ 等）がないか
  H7: 実質数値が2つ以下か／字幅が上限内か（1行36・2行56）

■ 判断語辞書の確定根拠（判断原理18｜推定でなく実測）
  うちた式カタログ61枚から機械抽出したヘッド60本の語尾を集計し、
  出現した語尾パターンを正規表現化した。最頻は「〜べきだ」（20件超）。
  検証結果: うちた式 59/60 = 98% 合格（残1件は ▲18億円 の ▲ を
  数値と数えた偽陽性）／対照6本（我々の現状・AI典型）は 0/6 = 偽陰性なし。

■ 判断原理19の適用｜本検査が「検出できないこと」（先に宣言する）
  - E1: 判断語があっても「主張として妥当か」は検査しない（文の形のみ）
  - E2: ヘッドと本体の整合（ヘッドの数字＝本体の合計か）は検査対象外
  - E3: 隣接3枚での文型の重複（うちた式「文型を散らす」）は本検査に含まない
        （複数枚の台帳が前提となるため、単票検査では判定できない）
  - E4: 判断語辞書は実測60本に基づくため、未出現の語尾は偽陽性となりうる
        （新たな語尾を検出した場合は辞書へ追加し、本宣言も更新する）
        9/11 追記: 自製の書き換え9本の検査で「引き上げられる」「足りる」の
        2件が未登録であることを検出し辞書へ追加した。可能形・受身形など
        活用形は原形と別に登録する必要がある。
  - E5: 「▲」「△」等の記号付き数値は数値として数える（H7 の既知の偽陽性）
"""

import re
import unicodedata

# ---------------------------------------------------------------
# 判断語辞書｜うちた式カタログ60本の語尾実測から確定（判断原理18）
# ---------------------------------------------------------------
JUDGE_RE = re.compile(
    r'(べきだ|べきである|べきではない'                 # 提言
    r'|必要だ|必要である|必要がある|必要ない'           # 要請・否定
    r'|に限られる|に限るべきだ|だけで足りる|に絞られる'   # 限定
    r'|に値する|に値しない'                          # 評価
    r'|左右しない|左右される|に依存する'                # 因果
    r'|ではなく|にすぎず|にとどまる'                   # 否定的限定
    r'|見込める|できる|前倒しできる'                   # 可能
    r'|確定させる|引き上げる|引き上げられる|引き下げる'    # 帰結
    r'|引き下げられる|縮む|招く|足りる'                 # 帰結（活用形）
    r'|先決になる|が先決だ|とすべきだ'                 # 優先
    r'|てある|にある'                                # 状態・所在の断定
    r'|である$|だ$|になる$|となる$'                   # 断定（文末）
    r')')

TITLE_RE = re.compile(
    r'(の分析|について|の整理|の検討|の状況|の一覧'
    r'|の推移|の比較|の概要|の結果|の課題|の傾向)$')

COUNT_RE = re.compile(
    r'(課題は\d+つ|論点は\d+つ|施策は\d+つ'
    r'|\d+点ある|\d+項目のうち|\d+工程のうち)')

SELF_RE = re.compile(r'(本ページ|この1枚|本スライド|本図|本資料|以下に示す)')

NUM_RE = re.compile(r'\d+(?:[,.]\d+)?%?')

W_1LINE = 36.0   # 1行の字幅上限
W_2LINE = 56.0   # 2行の字幅上限


def text_width(s):
    """全角=1.0／半角=0.5 の字幅を返す。"""
    return sum(1.0 if unicodedata.east_asian_width(c) in 'FWA' else 0.5
               for c in s)


def check_head(text):
    """ヘッドメッセージ1本を検査し判定結果を返す。

    Returns:
        dict : {'ok': bool, 'ng': [str], 'width': float, 'lines': int,
                'detail': {H1..H7: bool}}
    """
    t = (text or '').strip()
    ng = []
    detail = {}

    # H1 判断語
    detail['H1_judge'] = bool(JUDGE_RE.search(t))
    if not detail['H1_judge']:
        ng.append('H1 判断語なし')

    # H2 二段構え
    body = t[:-1] if t.endswith('。') else t
    detail['H2_single'] = ('。' not in body)
    if not detail['H2_single']:
        ng.append('H2 二段構え（句点で分割）')

    # H3 体言止め／題名
    detail['H3_not_title'] = not bool(TITLE_RE.search(body))
    if not detail['H3_not_title']:
        ng.append('H3 体言止め／題名')

    # H4 ダッシュ
    detail['H4_no_dash'] = ('——' not in t and '—' not in t)
    if not detail['H4_no_dash']:
        ng.append('H4 ダッシュ使用')

    # H5 自己言及
    detail['H5_no_self'] = not bool(SELF_RE.search(t))
    if not detail['H5_no_self']:
        ng.append('H5 自己言及')

    # H6 数え上げ
    detail['H6_no_count'] = not bool(COUNT_RE.search(t))
    if not detail['H6_no_count']:
        ng.append('H6 数え上げ')

    # H7 数値個数・字幅
    nums = NUM_RE.findall(t)
    w = text_width(t)
    detail['H7_nums'] = len(nums)
    detail['H7_width'] = w
    if len(nums) > 2:
        ng.append('H7 数値%d個（上限2）' % len(nums))
    if w > W_2LINE:
        ng.append('H7 字幅%.1f（上限%.0f）' % (w, W_2LINE))

    return {'ok': not ng, 'ng': ng, 'width': w,
            'lines': 1 if w <= W_1LINE else 2, 'detail': detail}


def check_heads(texts):
    """複数のヘッドを検査し集計を返す。"""
    rows = [(t, check_head(t)) for t in texts]
    ok = sum(1 for _, r in rows if r['ok'])
    reasons = {}
    for _, r in rows:
        for x in r['ng']:
            reasons[x] = reasons.get(x, 0) + 1
    return {'total': len(rows), 'ok': ok,
            'rate': (ok / len(rows) * 100) if rows else 0.0,
            'reasons': reasons, 'rows': rows}


def _report(title, texts):
    res = check_heads(texts)
    print('■ %s' % title)
    print('  合格 %d/%d = %.0f%%' % (res['ok'], res['total'], res['rate']))
    for t, r in res['rows']:
        mark = 'OK  ' if r['ok'] else 'NG  '
        print('  [%s] w=%5.1f L%d %-34s %s'
              % (mark, r['width'], r['lines'],
                 (', '.join(r['ng']) if r['ng'] else '-'), t[:44]))
    print()
    return res


if __name__ == '__main__':
    # 自己検証｜うちた式の実ヘッド（合格するべき）
    good = [
        '削減で生む原資が成長投資を支えるため、4件は一括での承認が必要だ',
        '引合の8割は見積後に消えるため、打ち手は見積の即断化に絞るべきだ',
        '同じ目盛で並べれば、投資に値するのはFAではなく成長する車載と医療だ',
        '4件を一括で承認いただければ、中計は初年度の削減施策から始動できる',
    ]
    # 対照｜我々の現状・AI典型（NGであるべき）
    bad = [
        '対照 n=6・説明文2行',
        'UI診断結果の分析',
        '課題は3つある。優先度をつけて対応すべきだ',
        'ナビゲーションの階層が深く回遊性が低い',
        '情報設計の課題について',
        'スマホ表示の改善——優先度A',
    ]
    g = _report('うちた式の実ヘッド（合格するべき）', good)
    b = _report('対照｜我々の現状・AI典型（NGであるべき）', bad)
    print('■ 自己検証の結論')
    print('  偽陽性（合格すべきものがNG）: %d件' % (g['total'] - g['ok']))
    print('  偽陰性（NGすべきものが合格）: %d件' % b['ok'])
    ok = (g['ok'] == g['total']) and (b['ok'] == 0)
    print('  判定: %s' % ('PASS' if ok else 'FAIL'))
