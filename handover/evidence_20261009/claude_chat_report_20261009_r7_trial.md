# 🧪 報告｜R-7（GPTs 経路の exec でデモが動く事象）解消案の再現試験｜2026-10-09（金）｜AIスライド側 17往復目
> 🔴 根拠：10/9 統括判定 第8便 §4（統括の提案は未検証・AIスライドが再現試験で確認）。builder・PLAIN・lite には触れていない。
> 🔴 入力：builder v17_216（`332199ad7dc2`）／Agent's Computer（Python 3.13.14）／各試験は空のフォルダで実行し、生成された `.pptx` を数えた。⚠️ ChatGPT の Code Interpreter 実機では試していない

## ■ 1. 結果
| 試験 | 読込方法 | デモの生成物 | その後の関数呼び出し |
|:---:|---|:---:|---|
| A（現行） | `exec(open('03_pptx_builder.py').read())` | 🚨 **5件** | 直接呼べる |
| B（統括案） | `exec(open(...).read(), {'__name__': 'pptx_builder'})` | ⭕ **0件** | ⚠️ **`create_presentation()` を直接呼ぶと `NameError`**（関数は渡した辞書の中にだけ定義される）。辞書から取り出せば動く（`g['create_presentation']`）・生成も正常 |
| C | B の書き方のまま、PLAIN J-7-3 の標準コードどおり関数を直接呼ぶ | ⭕ 0件 | ❌ `NameError: name 'create_presentation' is not defined` |
| D（代替案） | `_g = globals(); _s = _g['__name__']; _g['__name__'] = 'pptx_builder'; exec(open(...).read(), _g); _g['__name__'] = _s` | ⭕ **0件** | ⭕ **直接呼べる**（`__name__` も元に戻る） |

## ■ 2. 結論（事実）
- **統括案（B）は「デモを止める」点では成立する。** ただし PLAIN（445行・J-7-3）と lite（239行）の標準コードは、`exec` の後に関数を**直接**呼ぶ書き方であり、B に置き換えると関数が見つからず動かなくなる（C）。**B 単独では成立しない。**
- **代替案（D）は、builder を変えずにデモを止め、かつ既存の呼び出し方をそのまま使える**（Agent's Computer で確認）。
- どちらも builder のコードは変更しない（第3便の v17_217 の範囲制約に抵触しない）。PLAIN・lite の改訂は別件として起案が要る（第8便 §4 の「成立する場合」の扱い）。
- ⚠️ lite は 8000字制限があり、D は A より長い（約100字増）。収まるかは改訂案の起案時に確認する。
- ⚠️ Code Interpreter 実機での確認は、PLAIN・lite の改訂後に入江さんの GPTs で行う必要がある。

## ■ 3. 判定をお願いする事項
1. D で PLAIN・lite の改訂案を起案してよいか（または builder 側の対処＝v17_218 を選ぶか）
2. Code Interpreter 実機での確認を、改訂後の検証の条件にするか
