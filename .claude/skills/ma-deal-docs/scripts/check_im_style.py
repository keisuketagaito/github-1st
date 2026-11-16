"""概要書の書式チェック（依頼者修正 越ヶ谷製氷冷蔵 261116 で直された点を機械的に拾う）。

    python check_im_style.py 企業概要書.pptx [--min 10]

検出するもの（いずれも render_check.py の目視の前に潰す）：
  1. 小さすぎる文字：本文・Comment・注記・図形内の文字が --min（既定10pt）未満
     - 表は10pt未満を検出。ただし比率行（斜体）と、財務情報詳細のような25行以上の明細表は対象外
     - テンプレートの定型頁（はじめに・報酬体系）は対象外
  2. Comment の ≪見出し≫ と本文が同じ段落にある（見出しは独立した行に、本文は次の行に）
  3. 2つ目以降の ≪見出し≫ の前に空行がない（ブロックの切れ目が見えない）
     ※枠の高さが足りない頁（P/LハイライトのCommentなど）は空行を省いてよい。文字を小さくして空行を入れるのは逆
  4. 表の科目列以外のセルが全角スペースで始まる（字下げは科目列だけ）
  6. 図形・表の下端がフッター（Confidential・ページ番号の帯、7.70インチより下）にかかる
  5. テキストのあふれ（推定）：文字数と枠幅から高さを見積もり、下の図形・表に食い込むもの
     （推定なので、検出された頁は必ず描画して目で確かめる）
"""
import math
import sys
import unicodedata
from pptx import Presentation
from pptx.util import Emu

MIN = 10.0
MIN_TABLE = 9.5      # 表は見出し行などで9.6pt（12pt×0.8）まで許容
if '--min' in sys.argv:
    MIN = float(sys.argv[sys.argv.index('--min') + 1])
SKIP_PAGES = ('重要なご留意事項', '成功報酬')     # テンプレートの定型頁（改変しない）
EMU_PT = 12700


def runs_of(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.text.strip():
                yield p, r


def width_units(text):
    return sum(1.0 if unicodedata.east_asian_width(c) in 'WFA' else 0.55 for c in text)


def est_height_pt(sh):
    tf = sh.text_frame
    bp = tf._txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
    li = int(bp.get('lIns', 91440)) + int(bp.get('rIns', 91440)) if bp is not None else 182880
    ti = int(bp.get('tIns', 45720)) + int(bp.get('bIns', 45720)) if bp is not None else 91440
    w = (sh.width - li) / EMU_PT
    if w <= 0:
        return 0
    h = ti / EMU_PT
    for p in tf.paragraphs:
        sz = max([r.font.size.pt for r in p.runs if r.font.size] or [10.5])
        n = max(1, math.ceil(width_units(p.text) * sz / w)) if p.text else 1
        sa = p.space_after.pt if p.space_after is not None else 0
        h += n * sz * 1.2 + sa
    return h


def boxes(slide):
    out = []
    for sh in slide.shapes:
        if sh.width is None or sh.top is None:
            continue
        out.append((sh, sh.left, sh.top, sh.left + sh.width, sh.top + sh.height))
    return out


p = Presentation(sys.argv[1])
issues = []
for i, s in enumerate(p.slides, 1):
    if s._element.get('show') == '0':
        continue
    alltext = ' '.join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
    if any(k in alltext for k in SKIP_PAGES):
        continue
    bx = boxes(s)
    detail_page = '財務情報詳細' in alltext
    for sh in s.shapes:
        if sh.shape_type != 13 and sh.top is not None and sh.height and sh.top < Emu(int(7.0 * 914400)) and sh.top + sh.height > Emu(int(7.70 * 914400)):
            issues.append((i, '下端がフッターにかかる', f'{sh.name}（下端{(sh.top + sh.height) / 914400:.2f}in）'))
        if sh.has_text_frame and not sh.is_placeholder:
            small = sorted({r.font.size.pt for _, r in runs_of(sh.text_frame) if r.font.size and r.font.size.pt < MIN})
            if small:
                issues.append((i, '文字が小さい', f'{sh.text_frame.text[:18]!r} {small}pt'))
            paras = [pp.text for pp in sh.text_frame.paragraphs]
            for k, t in enumerate(paras):
                for o, c in (('≪', '≫'), ('【', '】')):
                    if t.startswith(o) and c in t and t.index(c) < len(t.rstrip()) - 1:
                        issues.append((i, '見出しと本文が同じ行', t[:30]))
                if t.startswith('≪') and k > 0 and paras[k - 1].strip() and not paras[k - 1].startswith('【'):
                    issues.append((i, '見出しの前に空行がない（枠に余裕があれば空行）', t[:20]))
            # あふれ（推定）：自分の下端の推定値が、横に重なる下の図形の上端を越える
            if sh.text_frame.text.strip():
                bottom = sh.top + Emu(int(est_height_pt(sh) * EMU_PT))
                for o, l, t, r, b in bx:
                    if o is sh or r <= sh.left or l >= sh.left + sh.width:
                        continue
                    if sh.top + sh.height - EMU_PT * 2 <= t < bottom - EMU_PT * 4:
                        issues.append((i, 'あふれ（推定・要目視）', f'{sh.text_frame.text[:18]!r} → {o.name}'))
                        break
        if sh.has_table:
            rows = list(sh.table.rows)
            detail = len(rows) >= 25 or detail_page
            small = set()
            cols = {}
            for r in rows:
                for ci, c in enumerate(r.cells):
                    txt = c.text
                    body = txt.lstrip('\u3000 ')
                    if ci > 0 and body and not (body[0].isdigit() or body[0] in '△▲-−(（.%－'):
                        cols.setdefault(ci, []).append(len(txt) - len(txt.lstrip('\u3000')))
                    if detail:
                        continue
                    for _, run in runs_of(c.text_frame):
                        if run.font.size and run.font.size.pt < MIN_TABLE and not run.font.italic:
                            small.add(round(run.font.size.pt, 1))
            for ci, lv in cols.items():
                # 階層のない文字列の列（所在地・業種・備考等）が字下げされている
                if len(lv) >= 2 and len(set(lv)) == 1 and lv[0] > 0:
                    issues.append((i, '科目列以外の字下げ', f'{len(lv)}セル（列{ci + 1}）'))
            if small:
                issues.append((i, '表の文字が小さい', f'{sorted(small)}pt（fscaleを上げる）'))
seen = set()
for it in issues:
    if it in seen:
        continue
    seen.add(it)
    print(*it)
print(f'書式の指摘 {len(seen)}件')
sys.exit(1 if seen else 0)
