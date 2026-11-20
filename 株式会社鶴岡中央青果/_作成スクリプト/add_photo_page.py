"""依頼者が手元で編集した概要書（261120）に、拠点写真ページを1枚追加する。

依頼者の手修正を残すため build_im.py から作り直さず、受領版に頁を差し込む。
「拠点概要（本社・市場）」頁を複製し、タイトル・字幕・セクション名だけ残して写真を置く。
    python add_photo_page.py <受領版.pptx> <出力.pptx>
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '../../.claude/skills/ma-deal-docs/scripts'))
from im_builder import *                     # noqa
from im_paste import scrub_doc_props

PH = os.path.join(HERE, '../写真')
src, out = sys.argv[1], sys.argv[2]
prs = Presentation(src)

idx = next(i for i, s in enumerate(prs.slides)
           if any(sh.has_text_frame and sh.text_frame.text.startswith('拠点概要（本社・市場）')
                  for sh in s.shapes))
s = duplicate_slide(prs, idx, after=idx)

# 頁枠（タイトル・字幕・右上のセクション名）以外を消す。拠点タブは見出しに流用する
tab = None
for sh in list(s.shapes):
    t = sh.text_frame.text if sh.has_text_frame else ''
    if sh.is_placeholder or t == '組織、拠点、設備等':
        continue
    if t == '拠点' and tab is None:
        tab = sh
        continue
    sh._element.getparent().remove(sh._element)

for sh in s.shapes:
    if sh.is_placeholder and sh.placeholder_format.idx == 0:
        set_text(sh, '拠点写真（本社・市場）')
    elif sh.is_placeholder and sh.placeholder_format.idx == 1:
        set_text(sh, '事務所棟と市場棟（荷捌き場・冷蔵庫）を同一敷地に構え、自社トラックで得意先へ配送している。')

# 見出しタブ（「拠点」タブの書式をそのまま複製）
import copy
def tab_at(text, y):
    el = copy.deepcopy(tab._element)
    s.shapes._spTree.append(el)
    sh = s.shapes[-1]
    sh.top = Inches(y)
    sh.text_frame.paragraphs[0].runs[0].text = text
    for r in sh.text_frame.paragraphs[0].runs[1:]:
        r.text = ''
tab_at('場内・設備', 4.80)
tab.top = Inches(1.30)
tab.text_frame.paragraphs[0].runs[0].text = '外観'

X, W = 0.34, 11.0
photo_strip(s, [(f'{PH}/外観_本社事務所.jpg', '本社事務所棟と社名看板'),
                (f'{PH}/外観_市場棟.jpg', '市場棟・倉庫（配送トラックの積込口）')],
            X, 1.74, W, 2.70, gap=0.20)
photo_strip(s, [(f'{PH}/場内_荷捌き場.jpg', '荷捌き場（出荷待ちの青果物と台車）'),
                (f'{PH}/場内_冷蔵庫.jpg', '場内の冷蔵庫とフォークリフト'),
                (f'{PH}/車両_配送トラック.jpg', '配送用トラック・営業車両')],
            X, 5.24, W, 1.84, gap=0.18)

force_white_headers(prs)
prs.save(out)
scrub_doc_props(out, title='企業概要書')
print('写真頁を', idx + 2, '頁に追加 →', out)
