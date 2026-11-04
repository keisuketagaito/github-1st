"""買手向け資料（概要書・ノンネーム）に社内限の情報が混ざっていないか検査する。

    python check_confidential.py 企業概要書.pptx [追加の禁止語 ...]

標準の禁止語（他社仲介・テール・NG・前回プロセス・報酬取決め・社内進行）に加え、
案件ごとの語（他社仲介の社名、NG先の社名、関係者名、前回の価格など）を引数で渡す。
案件マスター「その他」シートの「プロセス関連（社内確認事項）」に書いた固有名詞は必ず渡すこと。
「要確認」の件数も出す（多すぎる頁は資料依頼・ヒアリングで潰す）。
"""
import sys
from pptx import Presentation

BASE = ['仲介会社', 'テール', 'NG先', 'ＮＧ', 'ロングリスト', 'ネームクリア', '前回プロセス', '前回ディール',
        'DDブレイク', '紹介料', '着手金の', '社内確認', '温度感', 'ショートリスト']
p = Presentation(sys.argv[1])
words = BASE + sys.argv[2:]
hits, todo = [], 0
for i, s in enumerate(p.slides, 1):
    texts = []
    for sh in s.shapes:
        if sh.has_text_frame:
            texts.append(sh.text_frame.text)
        if sh.has_table:
            texts += [c.text for r in sh.table.rows for c in r.cells]
    t = '\n'.join(texts)
    todo += t.count('要確認')
    for w in words:
        if w in t:
            j = t.find(w)
            hits.append((i, w, t[max(0, j - 25):j + 25].replace('\n', ' ')))
for h in hits:
    print('要確認（禁止語）', *h)
print(f'禁止語 {len(hits)}件／「要確認」{todo}件')
sys.exit(1 if hits else 0)
