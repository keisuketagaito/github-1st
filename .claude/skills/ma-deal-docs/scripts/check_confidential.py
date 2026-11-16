"""買手向け資料（概要書・ノンネーム）に社内限の情報が混ざっていないか検査する。

    python check_confidential.py 企業概要書.pptx [追加の禁止語 ...]

標準の禁止語（他社仲介・テール・NG・前回プロセス・報酬取決め・社内進行）に加え、
案件ごとの語（他社仲介の社名、NG先の社名、関係者名、前回の価格など）を引数で渡す。
案件マスター「その他」シートの「プロセス関連（社内確認事項）」に書いた固有名詞は必ず渡すこと。

あわせて「買手向けに書かない表現」も検査する（依頼者修正 越ヶ谷製氷冷蔵 261116）：
買手が持っていない資料（株式価値評価書）・まだ無いもの（データルーム、受領後に掲載）・
作業中であること（確認中）への言及は、買手には意味がなく、資料の未完成感だけが残る。
未確認の事実は「要確認」と書き、資料の予定は書かない。

非表示スライド（要確認事項など社内用の頁）は検査対象から外して件数だけ出す。
非表示の頁はPDFには出ないが pptx を開けば見えるので、買手へはPDFで渡す。
「要確認」の件数も出す（多すぎる頁は資料依頼・ヒアリングで潰す）。
"""
import sys
from pptx import Presentation

BASE = ['仲介会社', 'テール', 'NG先', 'ＮＧ', 'ロングリスト', 'ネームクリア', '前回プロセス', '前回ディール',
        'DDブレイク', '紹介料', '着手金の', '社内確認', '温度感', 'ショートリスト']
# 買手向けに書かない表現（社内資料・将来の予定・作業中への言及）
PHRASES = ['株式価値評価書', 'Valuation', 'データルーム', '受領後', '開示予定', '確認中', '追って', '後日開示', '別途ご説明']
p = Presentation(sys.argv[1])
words = BASE + sys.argv[2:]
hits, phr, todo, hidden = [], [], 0, []
for i, s in enumerate(p.slides, 1):
    if s._element.get('show') == '0':
        hidden.append(i)
        continue
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
    for w in PHRASES:
        if w in t:
            j = t.find(w)
            phr.append((i, w, t[max(0, j - 25):j + 25].replace('\n', ' ')))
for h in hits:
    print('要確認（禁止語）', *h)
for h in phr:
    print('要修正（買手向けに書かない表現）', *h)
print(f'禁止語 {len(hits)}件／買手向けに書かない表現 {len(phr)}件／「要確認」{todo}件'
      + (f'／非表示スライド {hidden}（検査対象外。買手へはPDFで渡す）' if hidden else ''))
sys.exit(1 if hits or phr else 0)
