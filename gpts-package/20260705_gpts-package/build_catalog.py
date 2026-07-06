# -*- coding: utf-8 -*-
"""
04_layout_catalog.pdf 生成スクリプト

10種類のレイアウト見本を1ファイルにまとめた紹介PPTXを生成する。
GPTがKnowledge参照時に「どのレイアウト関数がどんな見た目になるか」を確認するための資料。
"""

import os, sys

# pptx_builder を読み込み
builder_path = os.path.join(os.path.dirname(__file__), '03_pptx_builder.py')
with open(builder_path, encoding='utf-8') as f:
    exec(f.read(), globals())


def build_catalog():
    """10種類のレイアウトを1ファイルにまとめる"""
    prs = create_presentation()
    total = 10

    # =========================================================
    # Layout 1: 表紙
    # =========================================================
    add_cover(prs,
        title='紺＆クリーン スライド作成\nレイアウト・カタログ',
        date='2026年版',
        author='Layout Catalog',
        subtitle='10種類のテンプレート関数の見本集')

    # =========================================================
    # Layout 2: アジェンダ
    # =========================================================
    add_agenda(prs,
        items=[
            {'title': '表紙（add_cover）', 'desc': 'タイトル・日付・著者'},
            {'title': 'アジェンダ（add_agenda）', 'desc': '全体目次'},
            {'title': '課題整理（add_issue_summary）', 'desc': '3カード並列'},
            {'title': '優先度マトリクス', 'desc': '4象限プロット'},
            {'title': 'OK・NG例', 'desc': '左右対比'},
            {'title': '施策一覧表', 'desc': 'PPTX表組'},
            {'title': 'KPIカード', 'desc': '数値強調'},
            {'title': 'スケジュール', 'desc': 'ガント風'},
            {'title': '運用フロー比較', 'desc': 'Before/After'},
            {'title': 'クロージング', 'desc': '最終ページ'},
        ],
        page_num=2, total=total,
        title='アジェンダ',
        lead='本カタログでは10種類のレイアウトを順に紹介します。')

    # =========================================================
    # Layout 3: 課題整理
    # =========================================================
    add_issue_summary(prs,
        title='【見本】課題整理レイアウト',
        cards=[
            {'no': '01', 'heading': 'カード見出し1',
             'body': 'カード本文の例。\n複数行に対応。\n3〜4行が読みやすい。'},
            {'no': '02', 'heading': 'カード見出し2',
             'body': '紺の左帯がアクセント。\n背景はライトグレー。\nメイリオ16pt。'},
            {'no': '03', 'heading': 'カード見出し3',
             'body': '番号は44ptで大きく。\n結論帯は任意。\nPOINTラベル付き。'},
        ],
        page_num=3, total=total,
        lead='3つの並列要素を整理するときに使う。リード文→3カード→結論帯の構成。',
        conclusion='【関数呼び出し】 add_issue_summary(prs, title, cards, conclusion=...)')

    # =========================================================
    # Layout 4: 優先度マトリクス
    # =========================================================
    add_priority_matrix(prs,
        title='【見本】優先度マトリクス',
        items=[
            {'label': '最優先項目', 'x': 0.78, 'y': 0.85, 'priority': 'S'},
            {'label': '重要項目', 'x': 0.62, 'y': 0.70, 'priority': 'A'},
            {'label': '通常項目1', 'x': 0.35, 'y': 0.55, 'priority': 'B'},
            {'label': '通常項目2', 'x': 0.55, 'y': 0.30, 'priority': 'B'},
            {'label': '低優先項目', 'x': 0.20, 'y': 0.20, 'priority': 'C'},
        ],
        page_num=4, total=total,
        x_axis_name='実行難易度', y_axis_name='インパクト',
        lead='4象限で施策を整理。赤円=最優先（S/A）、紺円=通常（B/C）。')

    # =========================================================
    # Layout 5: OK/NG
    # =========================================================
    add_ok_ng_pair(prs,
        title='【見本】OK・NG例の対比',
        ng={'heading': 'NGの典型例',
            'body': '避けるべきパターンを示す。\n本文は3〜4行で簡潔に。\nキャプションで補足。',
            'caption': '※ なぜダメか1行で説明'},
        ok={'heading': 'OKの推奨例',
            'body': '推奨パターンを示す。\nNGと同じ行数で対比性を出す。\n緑バッジ＋紺色見出し。',
            'caption': '※ どこが優れているか1行で'},
        page_num=5, total=total,
        lead='OK/NG対比は、品質基準の周知や改善ポイントの説明に有効。',
        conclusion='【関数呼び出し】 add_ok_ng_pair(prs, title, ng=..., ok=...)')

    # =========================================================
    # Layout 6: 施策一覧表
    # =========================================================
    add_action_table(prs,
        title='【見本】施策一覧表（add_action_table）',
        columns=['No', '施策名', '担当', '効果', '工数'],
        rows=[
            ['01', '施策Aの内容', '営業部', '+15%', '高'],
            ['02', '施策Bの内容', 'CS部', '-3pt', '中'],
            ['03', '施策Cの内容', 'IT部', '+8pt', '中'],
            ['04', '施策Dの内容', '人事部', '-20%', '低'],
            ['05', '施策Eの内容', 'マーケ', '+12%', '低'],
        ],
        page_num=6, total=total,
        lead='PPTXネイティブの表組。emphasize引数でセル内の数字を赤強調できる。',
        emphasize=[(0, 3, '+15%', '+15%'), (1, 3, '-3pt', '-3pt')],
        col_widths=[8, 40, 18, 22, 12],
        source='【出所】サンプルデータ（実数値ではありません）')

    # =========================================================
    # Layout 7: KPIカード
    # =========================================================
    add_kpi_card(prs,
        title='【見本】KPIカード（add_kpi_card）',
        kpis=[
            {'label': '指標A', 'value': '108', 'unit': '%',
             'desc': '前年比＋8%で目標達成', 'color': 'navy'},
            {'label': '指標B', 'value': '24', 'unit': '%',
             'desc': '前年比-10pt、要対策', 'color': 'red'},
            {'label': '指標C', 'value': '8.2', 'unit': '%',
             'desc': '悪化傾向、緊急対応', 'color': 'red'},
            {'label': '指標D', 'value': '42', 'unit': '社',
             'desc': '目標40社を超過達成', 'color': 'navy'},
        ],
        page_num=7, total=total,
        lead='核心数値を大きく表示（条項13）。赤=危機、紺=成果。1スライド最大4個。',
        conclusion='【関数呼び出し】 add_kpi_card(prs, title, kpis=[{label,value,unit,desc,color},...])')

    # =========================================================
    # Layout 8: スケジュール
    # =========================================================
    add_schedule_gantt(prs,
        title='【見本】スケジュール（add_schedule_gantt）',
        months=['7月', '8月', '9月', '10月', '11月', '12月'],
        tasks=[
            {'name': 'タスクA', 'start': 0, 'end': 2},
            {'name': 'タスクB', 'start': 1, 'end': 3},
            {'name': 'タスクC', 'start': 2, 'end': 4},
            {'name': 'マイルストーン①', 'start': 3, 'end': 3, 'milestone': True},
            {'name': 'マイルストーン②', 'start': 5, 'end': 5, 'milestone': True},
        ],
        page_num=8, total=total,
        lead='時系列計画。バー（紺）と マイルストーン（赤◆）の組み合わせ。')

    # =========================================================
    # Layout 9: 運用フロー比較
    # =========================================================
    add_flow_compare(prs,
        title='【見本】フロー比較（add_flow_compare）',
        before_steps=['STEP1', 'STEP2', 'STEP3', 'STEP4', 'STEP5'],
        after_steps=['STEP1', 'STEP2', 'STEP3', 'STEP4'],
        page_num=9, total=total,
        lead='Before/After フローを上下に並べて対比。改善効果を視覚化。',
        conclusion='【関数呼び出し】 add_flow_compare(prs, title, before_steps, after_steps)')

    # =========================================================
    # Layout 10: クロージング
    # =========================================================
    add_closing(prs,
        message='Thank you.',
        next_step='以上が10種類のレイアウト見本でした。\n各関数の詳細仕様は 02_design_spec.md を参照してください。',
        contact='紺＆クリーン スライド作成 / Layout Catalog',
        page_num=10, total=total)

    return prs


if __name__ == '__main__':
    out_path = os.path.join(os.path.dirname(__file__), '04_layout_catalog.pptx')
    prs = build_catalog()
    prs.save(out_path)
    print(f'Generated: {out_path}')
