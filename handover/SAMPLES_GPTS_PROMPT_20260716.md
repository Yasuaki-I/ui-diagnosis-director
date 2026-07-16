# フォーマット見本3種 GPTs投入プロンプト

**作成日**: 2026-07-16（木）
**目的**: `SAMPLES_DATA_20260716.md` の完成診断データをGPTsに投入し、3業種×3ファイル=9ファイルを生成する
**投入先**: UI診断ディレクター GPTs（v15.2 + Instructions v3.2 反映済み）

---

## 📋 使い方（入江さん向け・所要15分）

各業種ごとに、以下のプロンプトをGPTsの**新規チャットセッション**に貼付して送信してください。

**重要**：業種ごとに **必ず新規チャット** で実行してください。同一チャット内で連続実行すると、前業種のデータが混ざるリスクがあります。

投入順序：EC → LP → SaaS（順不同でも可）

生成されたPPTX3ファイル/業種を、ローカルに保存 → AI-Driveの `/ui-diagnosis-director/samples/` へアップロード。

---

## 🛍️ プロンプト1：EC（RURI STORE）

以下をコピーしてGPTsに送信：

```
以下は業種別フォーマット見本の生成依頼です。
販売ページ用のサンプル制作のため、診断結果を完全指定します。
通常の診断フローではなく、指定データに基づくPPTX生成のみ行ってください。

【重要な指示】
1. 診断は既に完了しています。追加の分析・スコア変更は不要です。
2. 以下の3辞書（diagnosis / proposals / visual_data）をそのまま使用してください。
3. Instructions v3.2 の J-3 標準コードに従い、3ファイルを生成してください。
4. site_slug は「ruri_store」を使用してください。
5. ファイル名は英数字表記（01_UI_Scorecard_ruri_store.pptx 等）としてください。

diagnosis = {
    'service_name': 'RURI STORE',
    'input_type': 'URL / EC',
    'total_score': 32,
    'rank': 'B',
    'rank_label': '標準的（改善余地あり）',
    'scores': [
        {'category': 'ファーストビュー', 'score': 3, 'max': 5, 'comment': 'ブランド世界観は伝わるが購入導線の起点が弱い'},
        {'category': 'キャッチコピー', 'score': 3, 'max': 5, 'comment': '雰囲気優先で商品価値の一文が曖昧'},
        {'category': 'CTA設計', 'score': 3, 'max': 5, 'comment': 'カート導線はあるが行動誘発が弱い'},
        {'category': '信頼性・権威性', 'score': 2, 'max': 5, 'comment': 'レビュー掲載が商品詳細末尾のみ'},
        {'category': 'フォーム設計', 'score': 3, 'max': 5, 'comment': '購入フォームの項目が多く入力負荷が高い'},
        {'category': 'レスポンシブ', 'score': 2, 'max': 5, 'comment': '商品一覧のタップ領域が小さくモバイル離脱の温床'},
        {'category': '読みやすさ', 'score': 4, 'max': 5, 'comment': '余白と写真構成が丁寧で回遊しやすい'},
        {'category': '情報設計', 'score': 3, 'max': 5, 'comment': 'カテゴリ設計は良好だが検索補助が弱い'},
        {'category': 'ブランド一貫性', 'score': 5, 'max': 5, 'comment': '写真・配色・トーンが統一され上質'},
        {'category': '表示速度・技術', 'score': 4, 'max': 5, 'comment': '画像最適化は良好だがLCPに改善余地'},
    ],
    'strengths': [
        '写真と余白でブランド世界観を強く表現',
        'カテゴリ構造が明快で商品発見がしやすい',
        '画像最適化により表示速度が快適',
    ],
    'priority_issues': [
        'FVのCTA可視性が弱く購入導線が不明',
        '商品詳細レビューが末尾で不安残る',
        'モバイル商品一覧のタップ領域が狭い',
    ],
    'conclusion': 'ブランド世界観は高水準。購入導線と信頼材料の配置改善で購入完了率を高められる。',
}

proposals = {
    'service_name': 'RURI STORE',
    'proposals': [
        {'no': 1, 'title': 'FV右上に固定カートCTAを設置する', 'priority': '高', 'effort': '小',
         'target_area': 'ファーストビュー', 'issue': '購入導線の起点が不明',
         'before': 'ロゴとナビゲーションのみで即行動できない',
         'after': '固定カートアイコン+新着ボタンを常時表示',
         'target_score_item': 'CTA設計'},
        {'no': 2, 'title': 'レビューを商品詳細の上部に移動する', 'priority': '高', 'effort': '中',
         'target_area': '商品詳細ページ', 'issue': '購入直前の不安解消が弱い',
         'before': 'レビューが商品詳細ページ末尾のみに配置',
         'after': '価格帯直下に評価スコア+抜粋3件を表示',
         'target_score_item': '信頼性・権威性'},
        {'no': 3, 'title': 'モバイル商品一覧のタップ領域を拡張する', 'priority': '高', 'effort': '中',
         'target_area': '商品一覧ページ(SP)', 'issue': 'タップミスによる離脱発生',
         'before': '商品画像小型で余白なく誤タップが多い',
         'after': '画像を1.5倍化し縦2列に変更・領域を確保',
         'target_score_item': 'レスポンシブ'},
        {'no': 4, 'title': '購入フォームを2段階に分割する', 'priority': '中', 'effort': '中',
         'target_area': '購入フォーム', 'issue': '入力項目が多く離脱を招く',
         'before': '配送・支払い・確認を1画面で入力',
         'after': '配送→支払い→確認の3ステップUIに分割',
         'target_score_item': 'フォーム設計'},
        {'no': 5, 'title': '商品検索に絞込みフィルタを追加する', 'priority': '低', 'effort': '中',
         'target_area': '検索・絞込み機能', 'issue': '商品発見の効率が低い',
         'before': 'カテゴリ選択のみで細かい絞込みができない',
         'after': 'サイズ・価格・色の3軸フィルタを追加',
         'target_score_item': '情報設計'},
    ],
    'summary': '最優先の一手は、FV固定カートCTA・商品詳細レビュー上部移動・モバイルタップ領域拡張で購入導線を通し、購入完了率を高めること。',
}

visual_data = {
    'service_name': 'RURI STORE',
    'total_score': 32, 'rank': 'B',
    'rank_label': '標準的（改善余地あり）',
    'scores': diagnosis['scores'],
    'strengths': diagnosis['strengths'],
    'priority_issues': diagnosis['priority_issues'],
    'top_issues': diagnosis['priority_issues'],
    'summary': 'FV導線と信頼材料を強化',
    'purpose': '購入完了率を高める',
    'target': 'ECサイト全体',
    'diagnosis_date': '2026年7月16日',
    'sections': [
        {'no': 1, 'name': 'ヘッダー', 'desc': 'ロゴ・ナビ明快', 'has_issue': False},
        {'no': 2, 'name': 'FV', 'desc': 'CTA可視性が弱い', 'has_issue': True},
        {'no': 3, 'name': '新着商品', 'desc': '写真構成が良好', 'has_issue': False},
        {'no': 4, 'name': 'カテゴリ一覧', 'desc': '構造明快', 'has_issue': False},
        {'no': 5, 'name': '商品詳細', 'desc': 'レビュー位置に難', 'has_issue': True},
        {'no': 6, 'name': '購入フォーム', 'desc': '入力負荷が高い', 'has_issue': True},
        {'no': 7, 'name': 'モバイル一覧', 'desc': 'タップ領域小', 'has_issue': True},
        {'no': 8, 'name': 'フッター', 'desc': '情報網羅性良好', 'has_issue': False},
    ],
    'flow_steps': [
        {'label': '訪問', 'status': '✓', 'note': 'ブランド認知'},
        {'label': 'FV確認', 'status': '✕', 'note': 'CTA不明'},
        {'label': '商品閲覧', 'status': '✓', 'note': '写真訴求良好'},
        {'label': '詳細確認', 'status': '✕', 'note': 'レビュー未見'},
        {'label': 'カート追加', 'status': '✓', 'note': '意思決定'},
        {'label': '購入完了', 'status': '✕', 'note': 'フォーム離脱'},
    ],
    'flow_summary': 'FVと商品詳細で意思決定の壁が発生',
    'direction': 'ブランド世界観の強みは維持しつつ、FV導線・レビュー配置・モバイル最適化・フォーム分割で購入完了率を高める。',
    'highlights': [
        {'no': 1, 'title': 'FV右上に固定カートCTAを設置', 'target_area': 'ファーストビュー',
         'before': 'ロゴとナビのみで即行動できない',
         'after': '固定カート+新着ボタンを常時表示',
         'priority': '高', 'effort': '小'},
        {'no': 2, 'title': 'レビューを商品詳細上部に移動', 'target_area': '商品詳細ページ',
         'before': 'レビューが末尾のみで購入直前の不安が残る',
         'after': '価格帯直下に評価スコア+抜粋3件を表示',
         'priority': '高', 'effort': '中'},
        {'no': 3, 'title': 'モバイル一覧のタップ領域拡張', 'target_area': '商品一覧ページ(SP)',
         'before': '画像小型で誤タップ多発',
         'after': '画像1.5倍化・縦2列で領域確保',
         'priority': '高', 'effort': '中'},
    ],
}

site_slug = 'ruri_store'

上記データを使用してPPTX 3ファイル（01_UI_Scorecard_ruri_store.pptx / 02_Proposals_ruri_store.pptx / 03_Visual_Board_ruri_store.pptx）を生成してください。
```

---

## 📄 プロンプト2：LP（TAX ASSIST PRO）

投入方法はEC版と同じ。**必ず新規チャット**で。

```
以下は業種別フォーマット見本の生成依頼です。
販売ページ用のサンプル制作のため、診断結果を完全指定します。
通常の診断フローではなく、指定データに基づくPPTX生成のみ行ってください。

【重要な指示】
1. 診断は既に完了しています。追加の分析・スコア変更は不要です。
2. 以下の3辞書をそのまま使用してください。
3. Instructions v3.2 の J-3 標準コードに従い、3ファイルを生成してください。
4. site_slug は「tax_assist_pro」を使用してください。
5. ファイル名は英数字表記としてください。

diagnosis = {
    'service_name': 'TAX ASSIST PRO',
    'input_type': 'URL / LP',
    'total_score': 30,
    'rank': 'B',
    'rank_label': '標準的（改善余地あり）',
    'scores': [
        {'category': 'ファーストビュー', 'score': 3, 'max': 5, 'comment': '実績数値は表示されるが提供価値が抽象的'},
        {'category': 'キャッチコピー', 'score': 2, 'max': 5, 'comment': '誰のどんな確定申告課題を解決するかが曖昧'},
        {'category': 'CTA設計', 'score': 2, 'max': 5, 'comment': '無料相談CTAがLP末尾のみで途中受け皿なし'},
        {'category': '信頼性・権威性', 'score': 4, 'max': 5, 'comment': '税理士監修バッジと利用者数を明示'},
        {'category': 'フォーム設計', 'score': 3, 'max': 5, 'comment': '相談申込みは短いが業種選択が任意'},
        {'category': 'レスポンシブ', 'score': 4, 'max': 5, 'comment': 'モバイル最適化は良好'},
        {'category': '読みやすさ', 'score': 4, 'max': 5, 'comment': 'セクション区切りと図解が丁寧'},
        {'category': '情報設計', 'score': 2, 'max': 5, 'comment': '料金プラン3種の比較表がなく違いが不明'},
        {'category': 'ブランド一貫性', 'score': 3, 'max': 5, 'comment': '配色統一だが装飾要素にばらつき'},
        {'category': '表示速度・技術', 'score': 3, 'max': 5, 'comment': '動画埋込みで初回表示が遅い'},
    ],
    'strengths': [
        '税理士監修バッジで初見の信頼を獲得',
        'モバイル最適化により広告経由の閲覧が快適',
        'セクション区切りと図解でLPの読了率が高い',
    ],
    'priority_issues': [
        'FVで確定申告課題の価値が曖昧',
        '料金プランの比較表無で違い不明',
        'LP末尾CTAのみで途中離脱多い',
    ],
    'conclusion': '信頼要素は強い。FV価値言語化と料金比較表と途中CTA配置で無料相談転換率を高められる。',
}

proposals = {
    'service_name': 'TAX ASSIST PRO',
    'proposals': [
        {'no': 1, 'title': 'FVに対象と価値の一文コピーを追加する', 'priority': '高', 'effort': '小',
         'target_area': 'ファーストビュー', 'issue': '対象と提供価値が伝わらない',
         'before': '「確定申告をラクに」のみの抽象表現',
         'after': 'フリーランス向け・税理士監修で節税を明示',
         'target_score_item': 'キャッチコピー'},
        {'no': 2, 'title': '料金プラン3種の比較表を追加する', 'priority': '高', 'effort': '中',
         'target_area': '料金セクション', 'issue': 'プラン違いが読み取れない',
         'before': '各プランをカード形式で個別に説明',
         'after': '3プラン横並びで機能・対象・料金を比較',
         'target_score_item': '情報設計'},
        {'no': 3, 'title': 'LP中盤に無料相談CTAを追加する', 'priority': '高', 'effort': '小',
         'target_area': 'LP中盤・料金セクション後', 'issue': '途中離脱の受け皿がない',
         'before': 'CTAがLP末尾のみで熱量が下がってから提示',
         'after': '料金比較表直下に「まず無料相談」を配置',
         'target_score_item': 'CTA設計'},
        {'no': 4, 'title': '相談申込みに業種選択を必須化する', 'priority': '中', 'effort': '小',
         'target_area': '相談申込みフォーム', 'issue': '相談前の情報粒度が粗い',
         'before': '氏名・メール・希望日時のみで業種は任意',
         'after': '業種選択を必須化し初回相談の質を上げる',
         'target_score_item': 'フォーム設計'},
        {'no': 5, 'title': '動画埋込みを遅延読込みに変更する', 'priority': '低', 'effort': '小',
         'target_area': 'FV下・紹介動画', 'issue': '初回LCPが遅い',
         'before': '動画が即時読込みで初回表示が3秒超',
         'after': '動画をレイジーロード化しLCP2秒以内に短縮',
         'target_score_item': '表示速度・技術'},
    ],
    'summary': '最優先の一手は、FV価値の言語化・料金比較表追加・LP中盤CTA配置で「対象・違い・行動」の3つの疑問を解消し、無料相談転換率を高めること。',
}

visual_data = {
    'service_name': 'TAX ASSIST PRO',
    'total_score': 30, 'rank': 'B',
    'rank_label': '標準的（改善余地あり）',
    'scores': diagnosis['scores'],
    'strengths': diagnosis['strengths'],
    'priority_issues': diagnosis['priority_issues'],
    'top_issues': diagnosis['priority_issues'],
    'summary': '価値と料金を明快化',
    'purpose': '無料相談転換率を高める',
    'target': 'LP全体',
    'diagnosis_date': '2026年7月16日',
    'sections': [
        {'no': 1, 'name': 'FV', 'desc': '価値言語化が弱い', 'has_issue': True},
        {'no': 2, 'name': '紹介動画', 'desc': 'LCP改善余地', 'has_issue': True},
        {'no': 3, 'name': '課題共感', 'desc': 'コピー訴求良好', 'has_issue': False},
        {'no': 4, 'name': 'サービス説明', 'desc': '図解が丁寧', 'has_issue': False},
        {'no': 5, 'name': '税理士監修', 'desc': '信頼シグナル強', 'has_issue': False},
        {'no': 6, 'name': '料金プラン', 'desc': '比較表なし', 'has_issue': True},
        {'no': 7, 'name': '中盤CTA', 'desc': '未設置', 'has_issue': True},
        {'no': 8, 'name': 'FAQ', 'desc': '網羅性良好', 'has_issue': False},
        {'no': 9, 'name': '末尾CTA', 'desc': '相談申込みフォーム', 'has_issue': False},
    ],
    'flow_steps': [
        {'label': '広告流入', 'status': '✓', 'note': 'LP到達'},
        {'label': 'FV理解', 'status': '✕', 'note': '価値抽象'},
        {'label': '課題共感', 'status': '✓', 'note': 'コピー刺さる'},
        {'label': '料金確認', 'status': '✕', 'note': '違い不明'},
        {'label': '離脱判断', 'status': '✕', 'note': '途中CTAなし'},
        {'label': '相談申込', 'status': '✓', 'note': '末尾フォーム'},
    ],
    'flow_summary': 'FV価値と料金理解で判断保留が多発',
    'direction': '税理士監修の信頼強みを活かしつつ、FV価値言語化・料金比較表・中盤CTA配置で「対象・違い・行動」の疑問を1本の導線で解消する。',
    'highlights': [
        {'no': 1, 'title': 'FVに対象と価値の一文コピー追加', 'target_area': 'ファーストビュー',
         'before': '「確定申告をラクに」のみの抽象表現',
         'after': 'フリーランス向け・税理士監修で節税を明示',
         'priority': '高', 'effort': '小'},
        {'no': 2, 'title': '料金プラン3種の比較表を追加', 'target_area': '料金セクション',
         'before': '各プランを個別カードで説明',
         'after': '3プラン横並びで機能・対象・料金を比較',
         'priority': '高', 'effort': '中'},
        {'no': 3, 'title': 'LP中盤に無料相談CTA追加', 'target_area': 'LP中盤・料金後',
         'before': 'CTAが末尾のみで熱量低下後に提示',
         'after': '料金比較直下に「まず無料相談」配置',
         'priority': '高', 'effort': '小'},
    ],
}

site_slug = 'tax_assist_pro'

上記データを使用してPPTX 3ファイル（01_UI_Scorecard_tax_assist_pro.pptx / 02_Proposals_tax_assist_pro.pptx / 03_Visual_Board_tax_assist_pro.pptx）を生成してください。
```

---

## 💼 プロンプト3：SaaS（WORKSYNC）

投入方法はEC・LP版と同じ。**必ず新規チャット**で。ランクC・課題数4項目に留意。

```
以下は業種別フォーマット見本の生成依頼です。
販売ページ用のサンプル制作のため、診断結果を完全指定します。
通常の診断フローではなく、指定データに基づくPPTX生成のみ行ってください。

【重要な指示】
1. 診断は既に完了しています。追加の分析・スコア変更は不要です。
2. 以下の3辞書をそのまま使用してください。
3. Instructions v3.2 の J-3 標準コードに従い、3ファイルを生成してください。
4. site_slug は「worksync」を使用してください。
5. ファイル名は英数字表記としてください。
6. このサンプルはランクCで課題数4項目です。priority_issuesを4件全て表示してください。

diagnosis = {
    'service_name': 'WORKSYNC',
    'input_type': 'URL / SaaS',
    'total_score': 28,
    'rank': 'C',
    'rank_label': '要改善(成果阻害あり)',
    'scores': [
        {'category': 'ファーストビュー', 'score': 3, 'max': 5, 'comment': '製品スクショはあるが誰の何を楽にするか不明'},
        {'category': 'キャッチコピー', 'score': 2, 'max': 5, 'comment': '機能名の羅列で価値の翻訳が弱い'},
        {'category': 'CTA設計', 'score': 2, 'max': 5, 'comment': '無料トライアルCTAが上下端のみで中盤なし'},
        {'category': '信頼性・権威性', 'score': 3, 'max': 5, 'comment': '導入企業ロゴはあるが事例が抽象的'},
        {'category': 'フォーム設計', 'score': 2, 'max': 5, 'comment': 'トライアル申込項目が10個超で入力負荷高'},
        {'category': 'レスポンシブ', 'score': 3, 'max': 5, 'comment': 'モバイルは閲覧可だが機能図が縮小しすぎ'},
        {'category': '読みやすさ', 'score': 3, 'max': 5, 'comment': '情報密度が高くセクション区切りが弱い'},
        {'category': '情報設計', 'score': 2, 'max': 5, 'comment': '料金は問合せ中心で比較検討段階で離脱'},
        {'category': 'ブランド一貫性', 'score': 4, 'max': 5, 'comment': '配色・タイポは統一され信頼感がある'},
        {'category': '表示速度・技術', 'score': 4, 'max': 5, 'comment': 'モダン実装で表示速度は良好'},
    ],
    'strengths': [
        '配色とタイポの統一でBtoBらしい信頼感を演出',
        'モダン実装により表示速度は業界トップクラス',
        '導入企業ロゴの掲示で初見の権威性を獲得',
    ],
    'priority_issues': [
        'FVの情シス向け価値翻訳が弱い',
        '料金プランが問合せ中心で途中離脱',
        '導入事例が抽象的で自社類似不明',
        '中盤トライアルCTA無しで行動起点弱',
    ],
    'conclusion': '技術基盤は強いが「誰の何を楽にするか」の翻訳と料金透明性で比較検討層の離脱を防ぐ改善が急務。',
}

proposals = {
    'service_name': 'WORKSYNC',
    'proposals': [
        {'no': 1, 'title': 'FVコピーを課題→効果の一文に変更する', 'priority': '高', 'effort': '小',
         'target_area': 'ファーストビュー', 'issue': '機能羅列で価値が伝わらない',
         'before': '「タスク管理・進捗共有・チーム連携」機能列挙',
         'after': '進捗確認の時間を週3時間削減の効果訴求',
         'target_score_item': 'キャッチコピー'},
        {'no': 2, 'title': '料金プラン比較表を公開する', 'priority': '高', 'effort': '中',
         'target_area': '料金ページ', 'issue': '問合せ中心で検討段階の離脱発生',
         'before': '「詳細はお問い合わせください」のみ',
         'after': '3プランの機能・利用人数・月額を透明化',
         'target_score_item': '情報設計'},
        {'no': 3, 'title': '導入事例を業種×規模で具体化する', 'priority': '高', 'effort': '中',
         'target_area': '導入事例セクション', 'issue': '抽象的で自社類似性が判断不能',
         'before': '「大手企業でも導入」等の抽象表現',
         'after': '「従業員30名の製造業で稼働率25%向上」等',
         'target_score_item': '信頼性・権威性'},
        {'no': 4, 'title': 'LP中盤にトライアルCTAを追加する', 'priority': '中', 'effort': '小',
         'target_area': '料金セクション後・事例セクション後', 'issue': 'CTAが上下端のみで途中離脱受皿なし',
         'before': 'ヘッダー・フッターのみのCTA配置',
         'after': '料金と事例の直後に「14日間無料で試す」配置',
         'target_score_item': 'CTA設計'},
        {'no': 5, 'title': 'トライアル申込フォームを3項目に簡素化する', 'priority': '中', 'effort': '小',
         'target_area': 'トライアル申込フォーム', 'issue': '入力負荷高く申込直前で離脱',
         'before': '会社名・部署・氏名・電話等10項目超',
         'after': 'メール・会社名・従業員規模の3項目のみ',
         'target_score_item': 'フォーム設計'},
    ],
    'summary': '最優先の一手は、FV価値翻訳・料金比較表公開・導入事例具体化で「誰に効くか・いくらか・実績あるか」の3疑問を解消し、比較検討層の離脱を防ぐこと。',
}

visual_data = {
    'service_name': 'WORKSYNC',
    'total_score': 28, 'rank': 'C',
    'rank_label': '要改善(成果阻害あり)',
    'scores': diagnosis['scores'],
    'strengths': diagnosis['strengths'],
    'priority_issues': diagnosis['priority_issues'],
    'top_issues': diagnosis['priority_issues'],
    'summary': '価値翻訳と料金透明化',
    'purpose': '比較検討層の離脱を防ぐ',
    'target': 'SaaSトップページ',
    'diagnosis_date': '2026年7月16日',
    'sections': [
        {'no': 1, 'name': 'FV', 'desc': '価値翻訳が弱い', 'has_issue': True},
        {'no': 2, 'name': '導入ロゴ帯', 'desc': '権威性◎', 'has_issue': False},
        {'no': 3, 'name': '機能説明', 'desc': '情報密度高い', 'has_issue': False},
        {'no': 4, 'name': '製品スクショ', 'desc': 'UI直感的', 'has_issue': False},
        {'no': 5, 'name': '導入事例', 'desc': '抽象的で不明', 'has_issue': True},
        {'no': 6, 'name': '料金プラン', 'desc': '問合せ中心', 'has_issue': True},
        {'no': 7, 'name': '中盤CTA', 'desc': '未設置', 'has_issue': True},
        {'no': 8, 'name': 'FAQ', 'desc': '網羅性良好', 'has_issue': False},
        {'no': 9, 'name': '末尾CTA', 'desc': 'トライアル導線', 'has_issue': True},
    ],
    'flow_steps': [
        {'label': '流入', 'status': '✓', 'note': '比較検討'},
        {'label': 'FV理解', 'status': '✕', 'note': '価値抽象'},
        {'label': '機能把握', 'status': '✓', 'note': 'スクショ有効'},
        {'label': '事例確認', 'status': '✕', 'note': '類似性不明'},
        {'label': '料金確認', 'status': '✕', 'note': '透明性なし'},
        {'label': 'トライアル', 'status': '✕', 'note': 'フォーム負荷'},
    ],
    'flow_summary': '事例・料金・フォームで多段階離脱発生',
    'direction': '技術基盤とブランド信頼の強みを活かしつつ、FV価値翻訳・料金比較表公開・導入事例具体化・中盤CTA・フォーム簡素化で比較検討層を無料トライアルへ通す。',
    'highlights': [
        {'no': 1, 'title': 'FVを課題→効果の一文に変更', 'target_area': 'ファーストビュー',
         'before': '機能名の羅列で価値が伝わらない',
         'after': '進捗確認の時間を週3時間削減を訴求',
         'priority': '高', 'effort': '小'},
        {'no': 2, 'title': '料金プラン比較表を公開', 'target_area': '料金ページ',
         'before': '「詳細はお問い合わせ」のみ',
         'after': '3プランの機能・人数・月額を透明化',
         'priority': '高', 'effort': '中'},
        {'no': 3, 'title': '導入事例を業種×規模で具体化', 'target_area': '導入事例セクション',
         'before': '「大手企業でも導入」等の抽象表現',
         'after': '「製造業30名で稼働率25%向上」等',
         'priority': '高', 'effort': '中'},
    ],
}

site_slug = 'worksync'

上記データを使用してPPTX 3ファイル（01_UI_Scorecard_worksync.pptx / 02_Proposals_worksync.pptx / 03_Visual_Board_worksync.pptx）を生成してください。
```

---

## ✅ 生成後の確認項目

各業種のPPTX 3ファイル/9ファイル全てで、以下を確認：

- [ ] ファイル名が英数字（`01_UI_Scorecard_[slug].pptx` 形式）
- [ ] C-1: 入力種別が2行以内で収まる
- [ ] C-2: 2枚目のカードが1枚目と同じ高さで上詰め
- [ ] C-2: POINT帯末尾の句点が孤立していない
- [ ] C-3: メタカード（目的・対象・診断日）に具体値が入る
- [ ] C-3: 最重要課題(Top3/4)が1行以内で切れずに表示
- [ ] SaaSのみ：priority_issues が4項目全て表示（B・C共通表示ロジックで正しく動作するか）

問題があれば即お知らせください。

---

**GPTs投入プロンプト 以上**
