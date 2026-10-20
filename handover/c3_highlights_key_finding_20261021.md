# 発見｜C-3 の Before/After（3枚目）で「箇所」が空欄・優先度と工数が既定値｜運用日 2026-10-21（水）｜AIスライド 新環境 15往復目
> 🔴 判定ではない。統括の判定材料。zip（e82ba6f995ba｜判定74 で V2〜V6 用に確定）は**変えていない**
> きっかけ：入江さんが中継した、統括の環境で出力された PPTX 3点と画像（運用日 10/21）

## ■ 1. 画像で見えたこと（03_Visual_Board 3/3）
- 3件とも「箇所：」の後ろが**空欄**
- 3件とも優先度バッジが「**中**」・「**工数：中**」（C-2 の改善提案では 高・高・中／工数 小）

## ■ 2. 原因（実測で確認）
| 箇所 | 中身 |
|---|---|
| builder（3,601〜3,720行） | highlights の各件から **`target_area`**（箇所）・**`priority`**（既定「中」）・**`effort`**（既定「中」）を読む |
| PLAIN J-3 のコメント（468行） | `highlights: Before/After3ペア [{area, before, after}]` ＝ **`area` と書かれている**。priority・effort・title の記載なし |
| AIスライドの input_schema.md・見本 JSON | PLAIN をそのまま写し、`{area, title, before, after}` にしていた（**AIスライドの見落とし**） |
- 試験：見本の各件に `target_area`・`priority`・`effort` を足して生成 → 「箇所：ファーストビュー」「高・工数：小」と正しく出た。verify_render は NG 0（変わらず）
- 🔴 **GPTs 経路でも同じことが起きている**：過去の実機出力 `_reference/v3.4_dogfooding_20260720/output/case01_lp_diagnosis/03_Visual_Board_speakup_eikaiwa.pptx` の3枚目も「箇所：」が3件とも空欄（case02〜05 は入っていた＝GPT がその回だけ `target_area` を書いていた）
- 点検スクリプト（r3）はこの空欄を検出しない（点検項目に入っていない）

## ■ 3. 直し方の案（判定はしない｜直すなら新しい名前 r4｜判定74）
1. input_schema.md：highlights を `{title, target_area, priority, effort, before, after}` に直す（area → target_area・priority・effort を追加。C-2 の該当する提案と同じ値を使う、と1行）
2. 見本 JSON：同じく直す
3. check_output.py：C-3 の3枚目の「箇所：」が空でないこと（または highlights の各件に target_area があること）を点検に追加
4. builder 本体は変えない（判定58-2）。PLAIN J-3 のコメントの誤記は記録のみ（PLAIN は採用版のため、直すかどうかは統括の判定）

## ■ 4. V2 への影響
- r3 のままで V2 を行うと、AI が input_schema.md のとおり `area` を書けば、3枚目の「箇所」が空欄・優先度が「中」になる見込み（V6 の目視で「欠け」に見える）
- 統括に判定をお願いしたい点：①r4 で直してから V2 を行うか ②r3 のまま V2 を行い、この欄は既知の不具合として V6 の評価から外すか
