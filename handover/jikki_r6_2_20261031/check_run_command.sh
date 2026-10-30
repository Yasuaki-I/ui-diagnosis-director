mkdir -p /home/claude/uniqlo && cd /home/claude/uniqlo && cat > build_input.py <<'EOF'
import json
scores=[
 ("ファーストビュー",3,"季節の告知は明快だが次の行動先が見えない"),
 ("キャッチコピー",3,"FVの訴求は汎用的、商品単位の説明は具体的"),
 ("CTA設計",3,"特集バナーに行動ボタンがなく誘導が弱い"),
 ("信頼性・権威性",5,"レビュー評価とAI要約で判断材料が厚い"),
 ("フォーム設計",4,"サイズと色の選択が明快、決済画面は未確認"),
 ("レスポンシブ",4,"SP最適化は良好だがカートが初期表示外"),
 ("読みやすさ",3,"写真上の白文字が読みにくい箇所がある"),
 ("情報設計",3,"カテゴリ一覧は良いが特集が縦に長く続く"),
 ("ブランド一貫性",5,"赤と白の配色と簡潔な語調で統一されている"),
 ("表示速度・技術",3,"画像382点と自動再生動画3本で読込が重い"),
]
sc=[{"category":c,"score":s,"max":5,"comment":m} for c,s,m in scores]
tot=round(sum(s for _,s,_ in scores)/10*10)
strengths=["レビュー評価とAI要約で購入前の不安を減らしている","到着予定日と在庫を商品ページ内で明示している","赤と白の配色で全ページの統一感を保っている"]
issues=["FVバナーに行動ボタンがない","SPでカートボタンが初期表示外","画像と動画が多く表示が重い"]
name="ユニクロ公式オンライン"
props=[
 dict(no=1,title="FVバナーに行動ボタンを添える",priority="高",effort="小",target_area="トップFV",issue="見た後の行動先が分かりにくい",before="期間告知の画像のみで行動ボタンがない",after="特集を見るボタンを画像の下に置く",target_score_item="CTA設計"),
 dict(no=2,title="SPでカートボタンを画面下に固定する",priority="高",effort="中",target_area="商品詳細SP",issue="価格とカートがスクロール後に見える",before="カートボタンが初期表示の外にある",after="価格とカートを画面下部に常時表示",target_score_item="レスポンシブ"),
 dict(no=3,title="価格に重なる案内表示をずらす",priority="中",effort="小",target_area="価格欄",issue="吹き出しが価格を一部隠している",before="サイズ案内の吹き出しが価格に重なる",after="吹き出しをサイズ欄の横に移す",target_score_item="情報設計"),
 dict(no=4,title="写真上の文字に下地を敷く",priority="中",effort="中",target_area="特集バナー",issue="白文字が写真に溶けて読みにくい",before="明るい写真の上に細い白文字を置く",after="文字の背面に半透明の下地を敷く",target_score_item="読みやすさ"),
 dict(no=5,title="動画と画像の読込を後回しにする",priority="低",effort="中",target_area="トップ全体",issue="画像と自動再生動画で読込が重い",before="動画3本が自動再生で同時に読み込む",after="画面に入った時だけ読み込み再生する",target_score_item="表示速度・技術"),
]
hl=[dict(title=t,target_area=p["target_area"],priority=p["priority"],effort=p["effort"],before=p["before"],after=p["after"]) for t,p in zip(["FVに行動ボタンを置く","SPでカートを常時表示","価格の重なりを解消"],props[:3])]
d={"site_slug":"uniqlo_jp",
 "record":{"service_name_official":"ユニクロ公式オンラインストア","target_url":"https://www.uniqlo.com/jp/ja/","input_origin":"url"},
 "diagnosis":{"service_name":name,"input_type":"URL / EC","total_score":tot,"rank":"B","rank_label":"標準的","scores":sc,"strengths":strengths,"priority_issues":issues,
   "conclusion":"FVバナーへの行動ボタン追加とSPのカート常時表示を優先すれば、商品詳細から購入までの到達率向上が見込める。"},
 "proposals":{"service_name":name,"proposals":props,"summary":"最優先の一手：トップFVのバナーに行動ボタンを添え、SPの商品詳細ではカートを画面下に固定して、見つけてから買うまでの迷いを減らす。"},
 "visual_data":{"service_name":name,"input_type":"URL / EC","total_score":tot,"rank":"B","rank_label":"標準的",
   "purpose":"購入CVを増やす","target":"トップと商品詳細","diagnosis_date":"2026年10月30日",
   "summary":"FVとSPカートの導線強化で購入増","top_issue":"FVバナーに行動ボタンがない",
   "direction":"トップFVのバナーに行動ボタンを置いて特集への移動を促し、SPの商品詳細ではカートを常時表示して購入までの手間を減らすべきだ。",
   "sections":[dict(no=i+1,label=l,desc=ds,has_issue=bad,status="issue" if bad else "good") for i,(l,ds,bad) in enumerate([
     ("ヘッダー","検索とカートが明快",False),("トップFV","行動ボタンがない",True),("カテゴリ一覧","アイコンで探せる",False),
     ("特集バナー","白文字が読みにくい",True),("商品詳細","レビューが充実",False),("カート追加","SPで初期表示外",True),("フッター","返品案内がある",False)])],
   "flow_steps":[dict(label=l,status=s,note=n) for l,s,n in [("流入","✓","検索や広告から"),("トップFV","✕","次の行動が不明"),("カテゴリ","✓","一覧で選べる"),("商品詳細","✓","情報が充実"),("サイズ選択","✓","ガイドがある"),("カート追加","✕","SPで要スクロール")]],
   "flow_summary":"トップFVとSPのカート追加で行動が止まりやすい",
   "highlights":hl,"scores":sc,"strengths":strengths,"priority_issues":issues,"top_issues":issues}}
json.dump(d,open("input.json","w"),ensure_ascii=False,indent=1)
print(tot)
EOF
python3 build_input.py && cat > gen.py <<'EOF'
import json
BUILDER_PATH='/mnt/skills/plugins/ui-diagnosis-director/scripts/03_pptx_builder.py'
_ns = {'__name__': 'pptx_builder'}; exec(open(BUILDER_PATH, encoding='utf-8').read(), _ns)
globals().update({k: v for k, v in _ns.items() if not k.startswith('__')})
d=json.load(open('input.json')); slug=d['site_slug']
for data, name, fn in [(d['diagnosis'], f'01_UI_Scorecard_{slug}', add_scorecard_onepager),
                       (d['proposals'], f'02_Proposals_{slug}', add_proposal_onepager),
                       (d['visual_data'], f'03_Visual_Board_{slug}', add_visual_board)]:
    prs = create_presentation(); fn(prs, data, page_num=1, slide_no=name[1], author='UI診断ディレクター'); prs.save('out/'+name + '.pptx')
print(_ns.get('__version__'))
EOF
mkdir -p out && python3 gen.py && python3 /mnt/skills/plugins/ui-diagnosis-director/scripts/check_output.py input.json out