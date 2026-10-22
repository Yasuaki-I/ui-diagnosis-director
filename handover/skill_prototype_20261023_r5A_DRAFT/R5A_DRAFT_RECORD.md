# r5 案A（下書き・統括判定待ち）の記録｜運用日 2026-10-23（金）｜AIスライド 新環境 20往復目
> 🔴 **統括の判定前のため、入江さんには渡さない**（zip 名に DRAFT）。判定で案A が採られたら、この中身を r5 として確定し、判定73 の方式で配布する。案B・C になったら作り直す
> 根拠：`skill_rankCD_count_finding_20261023.md`（bba08f636f1a）。r4（e8677c8c772f）は無変更で保存。builder 無改変（47629772eaff）

## ■ 1. 変更（r4 との差）
| ファイル | 変更 |
|---|---|
| SKILL.md §2 | 強み・課題の件数を「**ランクに関わらず各3件**（PPTX に載る件数）。4件目以降の課題は改善提案と結論に反映」に |
| references/scoring_rules.md §5 | 同上＋「C・D で3件に収まらないときは、成果への影響が大きい順に3件を選び、残りは改善提案（5件）と結論に反映」 |
| references/input_schema.md | strengths・priority_issues を「3件」に |
| scripts/check_output.py | 件数の点検を「ランクに関わらず3件」に（文面「4件目以降は改善提案・結論に回す」） |
| 版表記 | r4 → 「r5 案A（統括判定待ち）」 |

## ■ 2. 試験
| 試験 | 結果 |
|---|---|
| ランク B／**C**／**D**／S（強み3・課題3） | **4つとも点検合格**。verify_render は3ファイルとも NG 0。C-1 のランク欄「C  要改善」など正しく表示 |
| ランクC・課題4件（今回の V2 と同じ状況） | 🚨 不合格（課題が4件→3件にする／4件目が PPTX に無い） |
| r4 の点検でランクC（強み3・課題3） | 🚨 不合格5件（＝V2 で止まった状況の再現） |
| t1〜t11・u1・u3・追加4種（回帰） | すべて不合格のまま ⭕ |

## ■ 3. 下書き zip
| 項目 | 値 |
|---|---|
| ファイル | `ui-diagnosis-director_skill_prototype_r5A_20261023_DRAFT.zip` |
| B／md5 | 102,073B・`42fe55e5cb7c` |

| ファイル | B | md5 |
|---|---:|---|
| `ui-diagnosis-director/README.md` | 1154 | `3ee94d42786a` |
| `ui-diagnosis-director/SKILL.md` | 5503 | `de1b968f1048` |
| `ui-diagnosis-director/assets/sample_input_rankB.json` | 8997 | `6867e42ff10f` |
| `ui-diagnosis-director/references/diagram_patterns.md` | 13837 | `a3275de5276b` |
| `ui-diagnosis-director/references/dialogue_templates.md` | 5419 | `65c5c8c10e0f` |
| `ui-diagnosis-director/references/input_schema.md` | 3401 | `0fddd067ffe9` |
| `ui-diagnosis-director/references/scoring_rules.md` | 4729 | `4aa1c1fc6fde` |
| `ui-diagnosis-director/scripts/03_pptx_builder.py` | 287389 | `47629772eaff` |
| `ui-diagnosis-director/scripts/check_output.py` | 10943 | `dcf1c1372f37` |
