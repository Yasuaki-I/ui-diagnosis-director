# -*- coding: utf-8 -*-
"""
紺＆クリーン スライド作成 ─ サンプル事例を生成するスクリプト

GPTs ユーザーが「どんなアウトプットが期待できるか」を確認できるように、
複数テーマでサンプル PPTX を生成します。

1. 販売戦略レビュー（10枚）
2. 組織改革プラン（8枚）
3. 新規事業提案（10枚）
4. UI/UX 診断スコアカード 1ページ集約版（1枚）
5. UI/UX 改善提案リスト 1ページ集約版（1枚）
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# pptx_builder を読み込み
builder_path = os.path.join(os.path.dirname(__file__), '..', '03_pptx_builder.py')
with open(builder_path, encoding='utf-8') as f:
    exec(f.read(), globals())


# =====================================================================
# Sample 1: Q2 販売戦略レビュー
# =====================================================================
def build_sample_1_sales():
    prs = create_presentation()
    total = 10

    add_cover(prs,
        title='Q2 販売戦略レビュー',
        date='2026年7月15日',
        author='○○株式会社 営業企画部',
        subtitle='上期実績の振り返りと下期に向けた重点施策')

    add_agenda(prs,
        items=[
            {'title': '市場環境の整理', 'desc': 'マクロ・競合・顧客動向'},
            {'title': '現状の課題', 'desc': '上期で見えた3つの論点'},
            {'title': '優先度マトリクス', 'desc': '4象限で施策を整理'},
            {'title': '主要KPIの現状', 'desc': '達成率と前年比'},
            {'title': '施策一覧', 'desc': '下期に実行する10施策'},
            {'title': '実行スケジュール', 'desc': '6か月のロードマップ'},
            {'title': 'まとめ', 'desc': '結論と次のアクション'},
        ],
        page_num=2, total=total)

    add_issue_summary(prs,
        title='現状の3つの課題',
        cards=[
            {'no': '01', 'heading': 'リード獲得の伸び悩み',
             'body': '主要チャネルで前年比横ばい。\n新規流入経路の追加検討が必要。\nWebマーケの投資配分も要見直し。'},
            {'no': '02', 'heading': '商談化率の低下',
             'body': 'リードから商談への転換が\n前年比10pt低下。\nスコアリングロジック再設計が必要。'},
            {'no': '03', 'heading': '受注後の解約率増',
             'body': '導入後3か月の解約が増加。\nオンボーディングが手薄。\nCS体制の強化が急務。'},
        ],
        page_num=3, total=total,
        lead='上期実績の分析から、下期に向けた3つの課題を整理しました。',
        conclusion='3課題は独立ではなく「リード→商談→受注→定着」の流れで連動している。')

    add_priority_matrix(prs,
        title='施策の優先度マトリクス',
        items=[
            {'label': 'オンボーディング刷新', 'x': 0.78, 'y': 0.85, 'priority': 'S'},
            {'label': 'スコアリング再設計', 'x': 0.65, 'y': 0.72, 'priority': 'A'},
            {'label': '新規チャネル開拓', 'x': 0.85, 'y': 0.55, 'priority': 'A'},
            {'label': '営業ツール統合', 'x': 0.30, 'y': 0.60, 'priority': 'B'},
            {'label': 'ナレッジ整備', 'x': 0.25, 'y': 0.30, 'priority': 'C'},
            {'label': '名刺管理刷新', 'x': 0.60, 'y': 0.20, 'priority': 'C'},
        ],
        page_num=4, total=total,
        x_axis_name='実行難易度',
        y_axis_name='インパクト',
        x_label_low='易', x_label_high='難',
        y_label_low='低', y_label_high='高',
        lead='重要度×緊急度ではなく「インパクト×実行難易度」で施策を整理します。')

    add_kpi_card(prs,
        title='主要KPI（上期実績）',
        kpis=[
            {'label': '売上前年比', 'value': '108', 'unit': '%',
             'desc': '計画比103%。\n新規開拓が貢献。', 'color': 'navy'},
            {'label': '商談化率', 'value': '24', 'unit': '%',
             'desc': '前年34%から10pt低下。\n要対策。', 'color': 'red'},
            {'label': '受注後解約率', 'value': '8.2', 'unit': '%',
             'desc': '前年5.1%から悪化。', 'color': 'red'},
            {'label': '新規ロゴ獲得数', 'value': '42', 'unit': '社',
             'desc': '目標40社を超過達成。', 'color': 'navy'},
        ],
        page_num=5, total=total,
        lead='上期の主要KPIを4指標で整理します。',
        conclusion='売上は計画達成も、中間KPI（商談化・解約）に赤信号。')

    add_action_table(prs,
        title='下期の重点施策一覧',
        columns=['No', '施策名', '担当部署', '効果', '工数'],
        rows=[
            ['01', 'オンボーディング体制の再設計', 'CS部', '解約率 -3pt', '高'],
            ['02', 'リードスコアリング刷新', '営業企画', '商談化率 +8pt', '中'],
            ['03', 'パートナーチャネル拡張', '事業開発', 'リード +500件', '中'],
            ['04', '営業ツールの統合', '営業企画', '工数 -20%', '高'],
            ['05', 'ナレッジ整備とFAQ拡充', 'CS部', '問合せ -15%', '低'],
            ['06', '顧客成功事例の取材公開', 'マーケ', 'CV +12%', '低'],
        ],
        page_num=6, total=total,
        lead='優先度マトリクスから抽出した6施策を担当別に整理。',
        emphasize=[(0, 3, '解約率 -3pt', '-3pt'),
                    (1, 3, '商談化率 +8pt', '+8pt')],
        col_widths=[8, 42, 18, 22, 10],
        source='【出所】営業企画部 試算（2026年6月時点）')

    add_schedule_gantt(prs,
        title='下期スケジュール',
        months=['7月', '8月', '9月', '10月', '11月', '12月'],
        tasks=[
            {'name': 'オンボーディング刷新', 'start': 0, 'end': 2},
            {'name': 'スコアリング再設計', 'start': 0, 'end': 1},
            {'name': 'パートナー開拓', 'start': 1, 'end': 4},
            {'name': 'ツール統合', 'start': 2, 'end': 5},
            {'name': '中間レビュー', 'start': 2, 'end': 2, 'milestone': True},
            {'name': '通期レビュー', 'start': 5, 'end': 5, 'milestone': True},
        ],
        page_num=7, total=total,
        lead='6か月のロードマップ。中間レビューを9月末に設定。')

    add_flow_compare(prs,
        title='オンボーディングフローの刷新',
        before_steps=['契約', '初回連絡', '操作説明', '放置', '解約打診'],
        after_steps=['契約', 'キックオフ', '週次伴走', '導入完了', '定着支援'],
        page_num=8, total=total,
        lead='現状フローと刷新後フローを対比します。',
        conclusion='受動的な対応から、伴走型のオンボーディングへ転換。')

    add_ok_ng_pair(prs,
        title='提案資料のレビュー基準',
        ng={'heading': 'NG例: 数字の羅列',
            'body': '本文と同じサイズで数字を並べてしまい、\nどの数値を覚えて帰ればよいか伝わらない。\n読み手は全部見なければならず疲弊する。',
            'caption': '色も統一・サイズも均等 → 強調が分散'},
        ok={'heading': 'OK例: キー数値の強調',
            'body': 'スライドの結論を支えるキー数値を\n+60〜100%に拡大し、赤または紺で強調。\n1スライドに最大3つまでに絞る。',
            'caption': 'サイズ拡大＋色変更で訴求点が明確に'},
        page_num=9, total=total,
        lead='訴求数値は能動的に強調する（必達条項13）。',
        conclusion='読み手が「どの数字を覚えて帰るか」が明確になる。')

    add_closing(prs,
        message='Next Action',
        next_step='次回ミーティング：7月22日（火）14:00-15:00\n各施策の責任者から週次進捗を報告してください。\n中間レビューは9月末を予定しています。',
        contact='問い合わせ：営業企画部 山田\nyamada@example.com',
        page_num=10, total=total)

    return prs


# =====================================================================
# Sample 2: 組織改革プラン
# =====================================================================
def build_sample_2_org():
    prs = create_presentation()
    total = 8

    add_cover(prs,
        title='2026年度 組織改革プラン',
        date='2026年4月1日',
        author='△△株式会社 経営企画室',
        subtitle='事業ポートフォリオの再編に対応した組織再設計')

    add_agenda(prs,
        items=[
            {'title': '改革の背景', 'desc': '事業環境変化と現組織の課題'},
            {'title': '現状の3つの論点', 'desc': '組織構造の問題点'},
            {'title': '新組織のKPI', 'desc': '改革後の目標値'},
            {'title': 'Before / After フロー', 'desc': '意思決定プロセスの変化'},
            {'title': '移行スケジュール', 'desc': '6か月の段階的移行'},
            {'title': '次のアクション', 'desc': '部門長への説明計画'},
        ],
        page_num=2, total=total)

    add_issue_summary(prs,
        title='現組織の3つの論点',
        cards=[
            {'no': '01', 'heading': '意思決定の遅延',
             'body': '事業部間の調整に時間がかかり、\n新規施策の立ち上げが平均6か月遅延。\n権限委譲が不十分。'},
            {'no': '02', 'heading': '人材配置の硬直化',
             'body': '部署横断のアサインが困難。\nスキルマッチングがされず、\n機会損失が発生している。'},
            {'no': '03', 'heading': '評価制度の形骸化',
             'body': '年功型の評価が残り、\n若手の離職率が前年比1.5倍に。\n成果連動の再設計が必要。'},
        ],
        page_num=3, total=total,
        lead='経営層・現場ヒアリング（120名）から抽出した論点。',
        conclusion='3論点は「構造・人・制度」の3軸で連動して解決する必要がある。')

    add_kpi_card(prs,
        title='改革後の目標KPI',
        kpis=[
            {'label': '意思決定リードタイム', 'value': '-50', 'unit': '%',
             'desc': '6か月 → 3か月に短縮', 'color': 'red'},
            {'label': '部署横断アサイン率', 'value': '30', 'unit': '%',
             'desc': '現状8% → 30%に拡大', 'color': 'navy'},
            {'label': '若手離職率', 'value': '5', 'unit': '%以下',
             'desc': '現状12% → 半減目標', 'color': 'red'},
        ],
        page_num=4, total=total,
        lead='改革後12か月時点で達成すべきKPI。',
        conclusion='3指標が連動して改善することで「自律分散組織」へ移行する。')

    add_flow_compare(prs,
        title='意思決定プロセスの再設計',
        before_steps=['発議', '事業部長', '経営会議', '取締役会', '実行'],
        after_steps=['発議', '小委員会', '実行', '結果レビュー'],
        page_num=5, total=total,
        lead='5段階の階層型決裁から、4段階の自律型決裁へ。',
        conclusion='決裁段階を1段削減し、現場の小委員会に予算権限を委譲。')

    add_schedule_gantt(prs,
        title='移行スケジュール（6か月）',
        months=['4月', '5月', '6月', '7月', '8月', '9月'],
        tasks=[
            {'name': '部門長説明', 'start': 0, 'end': 0},
            {'name': '人事制度改定', 'start': 0, 'end': 2},
            {'name': '組織図確定', 'start': 1, 'end': 2},
            {'name': '評価制度移行', 'start': 2, 'end': 4},
            {'name': '新組織発足', 'start': 3, 'end': 3, 'milestone': True},
            {'name': '効果測定', 'start': 4, 'end': 5},
        ],
        page_num=6, total=total,
        lead='7月の新組織発足をマイルストーンに設定。')

    add_ok_ng_pair(prs,
        title='改革推進のレビュー観点',
        ng={'heading': 'NG: トップダウンの一斉移行',
            'body': '全社一斉で新組織を発足させると、\n現場の混乱と業績失速のリスクが高い。\n過去事例でも半数が失敗。',
            'caption': '組織変更は段階移行が原則'},
        ok={'heading': 'OK: パイロット部署から段階展開',
            'body': '効果検証可能な3部署で先行実施し、\n結果を踏まえて他部署へ展開する。\nリスクを限定しながら学習を進める。',
            'caption': '小さく始めて速く学ぶ'},
        page_num=7, total=total,
        lead='組織改革で失敗パターンを避けるための原則。',
        conclusion='アジャイル組織への移行は、組織自体がアジャイルに進めるべき。')

    add_closing(prs,
        message='まずは部門長への説明から',
        next_step='4月中に全部門長（35名）への個別説明を実施。\n5月の経営会議で正式決定を取りに行きます。',
        contact='問い合わせ：経営企画室 佐藤',
        page_num=8, total=total)

    return prs


# =====================================================================
# Sample 3: 新規事業提案
# =====================================================================
def build_sample_3_newbiz():
    prs = create_presentation()
    total = 10

    add_cover(prs,
        title='新規事業提案：法人向けAIアシスタント',
        date='2026年6月1日',
        author='□□株式会社 新規事業開発部',
        subtitle='3年で売上10億円規模を目指す新事業の立ち上げ計画')

    add_agenda(prs,
        items=[
            {'title': '市場機会', 'desc': '法人向けAI市場の規模と成長率'},
            {'title': '解決する課題', 'desc': 'ターゲット顧客の3つの困りごと'},
            {'title': 'プロダクト概要', 'desc': '提供価値とコア機能'},
            {'title': '事業計画KPI', 'desc': '3年間の売上・顧客数目標'},
            {'title': '優先度マトリクス', 'desc': '機能開発の優先順位'},
            {'title': '競合との差別化', 'desc': '主要競合4社との比較'},
            {'title': '実行スケジュール', 'desc': '12か月のロードマップ'},
            {'title': '次のアクション', 'desc': '承認後の動き方'},
        ],
        page_num=2, total=total)

    add_issue_summary(prs,
        title='ターゲット顧客の3つの課題',
        cards=[
            {'no': '01', 'heading': '情報検索の非効率',
             'body': '社内文書が分散し、\n必要な情報を見つけるのに\n1日平均1.5時間を費やしている。'},
            {'no': '02', 'heading': '定型業務の負担',
             'body': '議事録作成・報告書作成・\nメール対応に管理職時間の\n4割が消費されている。'},
            {'no': '03', 'heading': '専門知識の属人化',
             'body': 'ベテラン社員のノウハウが\n言語化されておらず、\n後継者への移転が困難。'},
        ],
        page_num=3, total=total,
        lead='ターゲット企業100社のヒアリング結果から課題を抽出。',
        conclusion='3課題はいずれも「社内ナレッジへのアクセス」が原因。')

    add_kpi_card(prs,
        title='事業計画KPI（3年目標）',
        kpis=[
            {'label': '年間売上（3年目）', 'value': '10', 'unit': '億円',
             'desc': '導入企業100社×平均1,000万円', 'color': 'navy'},
            {'label': '導入企業数', 'value': '100', 'unit': '社',
             'desc': '3年で累計100社の獲得', 'color': 'navy'},
            {'label': '営業利益率', 'value': '25', 'unit': '%',
             'desc': 'SaaS事業として安定収益化', 'color': 'red'},
            {'label': '初期投資回収', 'value': '24', 'unit': 'か月',
             'desc': '2年目末で投資回収完了', 'color': 'red'},
        ],
        page_num=4, total=total,
        lead='事業立ち上げから3年間で達成すべき指標。',
        conclusion='SaaS型課金で安定収益基盤を構築し、4年目以降の海外展開に繋げる。')

    add_priority_matrix(prs,
        title='機能開発の優先度',
        items=[
            {'label': '社内文書検索', 'x': 0.32, 'y': 0.92, 'priority': 'S'},
            {'label': '議事録自動生成', 'x': 0.55, 'y': 0.80, 'priority': 'A'},
            {'label': 'FAQ自動応答', 'x': 0.40, 'y': 0.65, 'priority': 'A'},
            {'label': 'メール下書き', 'x': 0.62, 'y': 0.45, 'priority': 'B'},
            {'label': '専門知識Q&A', 'x': 0.80, 'y': 0.55, 'priority': 'B'},
            {'label': '多言語対応', 'x': 0.75, 'y': 0.20, 'priority': 'C'},
        ],
        page_num=5, total=total,
        x_axis_name='実装難易度', y_axis_name='顧客インパクト',
        x_label_low='易', x_label_high='難',
        y_label_low='低', y_label_high='高',
        lead='実装難易度×顧客インパクトで開発優先度を整理。')

    add_action_table(prs,
        title='主要競合との比較',
        columns=['項目', '当社案', 'A社', 'B社', 'C社'],
        rows=[
            ['機能数', '5', '8', '4', '6'],
            ['日本語精度', '◎', '○', '△', '◎'],
            ['価格（月額/人）', '1,200円', '2,000円', '800円', '1,800円'],
            ['導入支援', '○', '△', '×', '○'],
            ['オンプレ対応', '○', '×', '×', '○'],
        ],
        page_num=6, total=total,
        lead='主要競合3社と当社案を5項目で比較しました。',
        emphasize=[(2, 1, '1,200円', '1,200円')],
        col_widths=[24, 20, 18, 19, 19],
        source='【出所】各社公式サイト・営業ヒアリング（2026年5月時点）')

    add_schedule_gantt(prs,
        title='12か月の実行ロードマップ',
        months=['7月', '9月', '11月', '1月', '3月', '5月'],
        tasks=[
            {'name': '要件定義', 'start': 0, 'end': 0},
            {'name': 'MVP開発', 'start': 0, 'end': 2},
            {'name': 'パイロット顧客10社', 'start': 1, 'end': 3},
            {'name': '正式版開発', 'start': 2, 'end': 4},
            {'name': '正式リリース', 'start': 3, 'end': 3, 'milestone': True},
            {'name': '営業活動拡大', 'start': 3, 'end': 5},
        ],
        page_num=7, total=total,
        lead='1月の正式リリースをマイルストーンに、12か月で立ち上げ。')

    add_flow_compare(prs,
        title='営業プロセスの設計',
        before_steps=['資料請求', '商談', '個別開発提案', '長期検討', '受注'],
        after_steps=['Web申込', 'デモ', 'PoC開始', '本契約'],
        page_num=8, total=total,
        lead='受託型の長期商談から、SaaS型の短期受注へ。',
        conclusion='プロセス短縮で営業生産性を2倍に。')

    add_ok_ng_pair(prs,
        title='事業承認のレビュー観点',
        ng={'heading': 'NG: バラ色の数字並べ',
            'body': '3年で売上100億・利益50%など、\n根拠の薄い数字を並べると、\n経営層から信頼を得られない。',
            'caption': '楽観的すぎる数字は逆効果'},
        ok={'heading': 'OK: 保守的試算＋根拠明示',
            'body': '同業他社実績・ヒアリング結果に基づき、\n3年目10億円という現実的な数字を提示。\n計算根拠も併記する。',
            'caption': '保守試算で承認後に超過達成を狙う'},
        page_num=9, total=total,
        lead='新規事業の承認を得るための数字の作り方。',
        conclusion='承認後の運営で信頼を積むことが、次の投資承認に繋がる。')

    add_closing(prs,
        message='ご承認のお願い',
        next_step='本提案について、6月中の取締役会で承認をお願いします。\n承認後、7月から要件定義フェーズに着手します。',
        contact='問い合わせ：新規事業開発部 鈴木',
        page_num=10, total=total)

    return prs


# =====================================================================
# Main
# =====================================================================
if __name__ == '__main__':
    out_dir = os.path.dirname(__file__)

    prs1 = build_sample_1_sales()
    prs1.save(os.path.join(out_dir, 'sample_1_sales_strategy.pptx'))
    print('Generated: sample_1_sales_strategy.pptx')

    prs2 = build_sample_2_org()
    prs2.save(os.path.join(out_dir, 'sample_2_org_reform.pptx'))
    print('Generated: sample_2_org_reform.pptx')

    prs3 = build_sample_3_newbiz()
    prs3.save(os.path.join(out_dir, 'sample_3_new_business.pptx'))
    print('Generated: sample_3_new_business.pptx')

    # =====================================================================
    # Sample 6: UI診断スコアカード 1ページ集約版
    # =====================================================================
    onepager_diagnosis = {
        'service_name': 'スピークアップ英会話',
        'input_type': 'スクリーンショット画像',
        'total_score': 30,
        'rank': 'B',
        'rank_label': '標準的（改善余地あり）',
        'scores': [
            {'category': 'ファーストビュー',   'score': 3, 'max': 5,
             'comment': '清潔感はあるが、体験価値と成果が一瞬で伝わりにくい'},
            {'category': 'キャッチコピー',     'score': 3, 'max': 5,
             'comment': '開始喚起は明快だが、選ばれる理由が弱い'},
            {'category': 'CTA設計',           'score': 1, 'max': 5,
             'comment': 'FV・中盤に主要CTAがなく行動導線が弱い'},
            {'category': '信頼性・権威性',     'score': 3, 'max': 5,
             'comment': '実績数値はあるが、根拠や証拠の厚みが不足'},
            {'category': 'フォーム設計',       'score': 1, 'max': 5,
             'comment': '問い合わせ導線はあるが入力前の不安解消が少ない'},
            {'category': 'レスポンシブ',       'score': 3, 'max': 5,
             'comment': 'PC表示は整っているが、SPでCTA固定が必要'},
            {'category': '読みやすさ',         'score': 5, 'max': 5,
             'comment': '余白・行間・色数が整理され読みやすい'},
            {'category': '情報設計',           'score': 3, 'max': 5,
             'comment': '特徴と料金は見やすいが、比較・成果・不安解消が不足'},
            {'category': 'ブランド一貫性',     'score': 5, 'max': 5,
             'comment': '青基調で教育サービスらしい安心感がある'},
            {'category': '表示速度・技術',     'score': 3, 'max': 5,
             'comment': '画像中心のため最適化余地がある'},
        ],
        'strengths': [
            '余白が広く、読みやすいレイアウトで安心感がある',
            '料金・特徴・講師情報が整理され、サービス概要を把握しやすい',
            '青基調の配色で教育サービスらしい信頼感が出ている',
        ],
        'priority_issues': [
            'ファーストビューに主要CTAがなく、次の行動が分かりにくい',
            '受講後の成果・具体的なベネフィットが弱く、比較検討時の決め手に欠ける',
            '受講生の声が短く、信頼材料としての説得力が不足している',
        ],
        'conclusion': 'FV内CTA・成果訴求・信頼材料の3点を最優先で改善し、無料体験への離脱を防ぐ。',
    }

    prs4 = create_presentation()
    # 2スライド構成（サマリ + 詳細スコア表）
    add_scorecard_onepager(prs4, onepager_diagnosis,
                            page_num=1, total=2, slide_no='1')
    prs4.save(os.path.join(out_dir, 'sample_6_onepager_scorecard.pptx'))
    print('Generated: sample_6_onepager_scorecard.pptx')

    # =====================================================================
    # Sample 7: 改善提案リスト 1ページ集約版
    # =====================================================================
    onepager_proposals = {
        'service_name': 'スピークアップ英会話',
        'proposals': [
            {'no': 1, 'title': 'ファーストビューに無料体験CTAを追加',
             'priority': '高', 'effort': '小',
             'target_area': 'ファーストビュー',
             'issue': '興味を持った直後に行動できない',
             'before': '見出しと説明文はあるが、主要ボタンが見当たらない',
             'after': '無料体験・料金確認の2ボタンをFV内に配置',
             'target_score_item': 'CTA設計'},
            {'no': 2, 'title': 'キャッチコピーを成果訴求型に変更',
             'priority': '高', 'effort': '小',
             'target_area': 'メインコピー',
             'issue': '学習後のメリットが弱く差別化しにくい',
             'before': '「英語、始めませんか？」で汎用的な印象',
             'after': '忙しい人でも続く、成果が見えるオンライン英会話として訴求',
             'target_score_item': 'キャッチコピー'},
            {'no': 3, 'title': '実績数値に根拠と補足を追加',
             'priority': '中', 'effort': '中',
             'target_area': '実績帯',
             'issue': '数値はあるが信頼の裏付けが不足',
             'before': '開校年・採用率・段階数・時間帯のみ表示',
             'after': '累計受講者数、継続率、満足度など検討材料を追加',
             'target_score_item': '信頼性・権威性'},
            {'no': 4, 'title': '受講生の声を具体化',
             'priority': '中', 'effort': '小',
             'target_area': '受講生の声',
             'issue': '短文のみで利用シーンや成果が伝わりにくい',
             'before': '「続けられています」など一言コメント中心',
             'after': '課題・受講理由・成果を含む3行レビューへ拡張',
             'target_score_item': '信頼性・権威性'},
            {'no': 5, 'title': '料金セクション下に比較とCTAを追加',
             'priority': '高', 'effort': '中',
             'target_area': '料金プラン',
             'issue': 'プラン選択後の次アクションが弱い',
             'before': '料金カードの提示で終わり、相談導線が離れている',
             'after': 'おすすめプラン表示と無料相談CTAを直下に配置',
             'target_score_item': 'CTA設計'},
        ],
        'summary': '最優先の一手：ファーストビューに無料体験CTAを追加し、直下に成果・安心材料を3点で提示する。',
    }

    prs5 = create_presentation()
    # 2スライド構成（前半3件 + 後半2件 + POINT帯）
    # 提案件数に応じて自動的に1or2スライド（5件→2スライド／3件以下→1スライド）
    add_proposal_onepager(prs5, onepager_proposals,
                           page_num=1, slide_no='2')
    prs5.save(os.path.join(out_dir, 'sample_7_onepager_proposals.pptx'))
    print('Generated: sample_7_onepager_proposals.pptx')

    # =====================================================================
    # Sample 8: C-3 ビジュアル診断ボード（3枚構成）
    # =====================================================================
    onepager_visual = {
        # ヘッダー
        'service_name': 'スピークアップ英会話',
        'purpose': '無料体験予約・問い合わせ誘導',
        'target': 'スクール紹介LP（英会話）',
        'diagnosis_date': '2026年6月4日',
        'total_score': 30,
        'rank': 'B',
        'rank_label': '標準的（改善余地あり）',

        # スライド1：LP構造マップ
        'sections': [
            {'no': 1, 'name': 'ヘッダー', 'desc': 'ロゴ・ナビゲーション',
             'has_issue': False},
            {'no': 2, 'name': 'ファーストビュー',
             'desc': 'メインコピー・説明文', 'has_issue': True},
            {'no': 3, 'name': '実績バー',
             'desc': '数値で信頼を訴求', 'has_issue': False},
            {'no': 4, 'name': '特徴セクション',
             'desc': '7つの特徴を訴求', 'has_issue': True},
            {'no': 5, 'name': '講師紹介',
             'desc': '安心感を伝える', 'has_issue': False},
            {'no': 6, 'name': '料金プラン',
             'desc': '5つのプラン比較', 'has_issue': True},
            {'no': 7, 'name': '受講生の声',
             'desc': '口コミ・レビュー', 'has_issue': False},
            {'no': 8, 'name': '最終CTA',
             'desc': '問い合わせボタン', 'has_issue': True},
            {'no': 9, 'name': 'フッター',
             'desc': '会社情報・規約', 'has_issue': False},
        ],
        'summary': '見やすく整っているが、行動動機と信頼根拠の訴求が弱い。',
        'top_issues': [
            'ファーストビューに主要CTAがなく行動できない',
            '受講後の成果・ベネフィットが弱く決め手に欠ける',
            '受講生の声が短く信頼材料として説得力不足',
        ],

        # スライド1：行動フロー
        'flow_steps': [
            {'label': '見る', 'status': '✓', 'note': ''},
            {'label': '興味', 'status': '✓', 'note': ''},
            {'label': '信頼', 'status': '✕', 'note': '実績根拠が不足'},
            {'label': '料金確認', 'status': '✕', 'note': 'CTAが見つからない'},
            {'label': '迷う', 'status': '✕', 'note': '次のアクション不明'},
            {'label': '離脱', 'status': '✕', 'note': '最終CTA文言が弱い'},
        ],
        'flow_summary': '行動の壁が多く、興味を持ったユーザーを逃している',

        # スライド2：スコア視覚化（diagnosis から複製）
        'scores': onepager_diagnosis['scores'],
        'strengths': onepager_diagnosis['strengths'],
        'priority_issues': onepager_diagnosis['priority_issues'],

        # スライド3：Before/After Top3（proposals 上位3件から複製）
        'highlights': [
            {'no': p['no'], 'title': p['title'],
             'target_area': p['target_area'],
             'before': p['before'], 'after': p['after'],
             'priority': p['priority'], 'effort': p['effort']}
            for p in onepager_proposals['proposals'][:3]
        ],

        # 全体共通
        'direction': '成果訴求と信頼根拠を足し、無料体験への離脱率を改善する',
    }

    prs6 = create_presentation()
    add_visual_board(prs6, onepager_visual,
                     page_num=1, total=3, slide_no='3')
    prs6.save(os.path.join(out_dir, 'sample_8_visual_board.pptx'))
    print('Generated: sample_8_visual_board.pptx')

    # =====================================================================
    # Sample 9: 統合版 UI診断統合レポート（C-1+C-2+C-3 = 計7枚）
    # =====================================================================
    prs7 = build_full_report(onepager_diagnosis, onepager_proposals,
                              onepager_visual)
    prs7.save(os.path.join(out_dir, 'sample_9_full_report.pptx'))
    print('Generated: sample_9_full_report.pptx')
