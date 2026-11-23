"""成約事例 発表資料（売：株式会社グルメルシー × 買：株式会社日本アメニティライフ協会）のビルドスクリプト。

    python 株式会社グルメルシー/build_seiyaku_jirei.py

当社の成約事例テンプレート（成約事例_素材/成約事例_template.pptx）を複製して中身を差し込む。
白紙から組み立てない。数値の出所：
  - 企業概要書（2025/1/21版）：会社概要・P/L・B/S（表は概要書の表をそのまま貼付：im_pl_table.xml / im_bs_table.xml）
  - 事業譲渡契約書（2026/7/23締結）：スキーム・対価・日付・前提条件・別紙1〜3
未確認の事項（担当者、初回面談〜トップ面談の日付、面談社数、買手の会社概要など）は「要確認」と赤字で記載。
"""
import copy
import os
import re

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '成約事例_素材')
TEMPLATE = os.path.join(SRC, '成約事例_template.pptx')
OUT = os.path.join(HERE, '成約事例_グルメルシー×日本アメニティライフ協会_261123.pptx')

NAVY = RGBColor(0x0E, 0x4B, 0x66)      # テンプレートの見出し帯（accent2 lumMod25%）相当
BLUE = RGBColor(0x08, 0x89, 0xC9)      # accent4
PALE = RGBColor(0xE1, 0xF3, 0xFB)      # accent2
GRAY = RGBColor(0xE3, 0xE7, 0xE9)      # lt2
DARK = RGBColor(0x26, 0x26, 0x26)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x00, 0x00)
FONT = '游ゴシック'


# ---------------------------------------------------------------- 基本ヘルパー
def shape_by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)


def remove(shape):
    el = shape._element
    el.getparent().remove(el)


def _style_run(r, sz=None, b=None, color=None):
    f = r.font
    if sz is not None:
        f.size = Pt(sz)
    if b is not None:
        f.bold = b
    if color is not None:
        f.color.rgb = color


def set_text(shape, paras, sz=None, b=None, color=None, align=None, red_kakunin=True):
    """テキストを差し替える。先頭ランの書式（フォント・色）を引き継ぐ。
    paras: [str | [(text, {sz,b,color}), ...], ...]。'要確認' は自動で赤字にする。"""
    tf = shape.text_frame
    txBody = tf._txBody
    tmpl_rpr, tmpl_ppr = None, None
    for p in txBody.findall(qn('a:p')):
        if tmpl_ppr is None and p.find(qn('a:pPr')) is not None:
            tmpl_ppr = copy.deepcopy(p.find(qn('a:pPr')))
        r = p.find(qn('a:r'))
        if r is not None and r.find(qn('a:rPr')) is not None:
            tmpl_rpr = copy.deepcopy(r.find(qn('a:rPr')))
            break
    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)
    for para in paras:
        p_el = etree.SubElement(txBody, qn('a:p'))
        if tmpl_ppr is not None:
            p_el.append(copy.deepcopy(tmpl_ppr))
        segs = para if isinstance(para, list) else [(para, {})]
        for seg in segs:
            text, opt = (seg, {}) if isinstance(seg, str) else seg
            pieces = re.split(r'(要確認)', text) if red_kakunin else [text]
            for piece in pieces:
                if piece == '':
                    continue
                r_el = etree.SubElement(p_el, qn('a:r'))
                rpr = copy.deepcopy(tmpl_rpr) if tmpl_rpr is not None else etree.Element(qn('a:rPr'))
                rpr.attrib.pop('err', None)
                rpr.set('lang', 'ja-JP')
                r_el.append(rpr)
                t = etree.SubElement(r_el, qn('a:t'))
                t.text = piece
        from pptx.text.text import _Paragraph
        para_obj = _Paragraph(p_el, tf)
        if align is not None:
            para_obj.alignment = align
        for run in para_obj.runs:
            o = {}
            for seg in segs:
                if not isinstance(seg, str) and seg[0] and run.text in seg[0]:
                    o = seg[1]
                    break
            _style_run(run, o.get('sz', sz), o.get('b', b), o.get('color', color))
            if red_kakunin and run.text == '要確認':
                run.font.color.rgb = RED
                run.font.bold = True


def add_box(slide, x, y, w, h, paras, fill=None, line=None, sz=10, b=False, color=DARK,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.RECTANGLE, margin=0.06):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(0.75)
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ('left', 'right', 'top', 'bottom'):
        setattr(tf, f'margin_{side}', Inches(margin))
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        segs = para if isinstance(para, list) else [(para, {})]
        for seg in segs:
            text, opt = (seg, {}) if isinstance(seg, str) else seg
            for piece in re.split(r'(要確認)', text):
                if not piece:
                    continue
                r = p.add_run()
                r.text = piece
                r.font.name = FONT
                r.font._rPr.get_or_add_latin().set('typeface', FONT)
                ea = r.font._rPr.find(qn('a:ea'))
                if ea is None:
                    ea = etree.SubElement(r.font._rPr, qn('a:ea'))
                ea.set('typeface', FONT)
                _style_run(r, opt.get('sz', sz), opt.get('b', b), opt.get('color', color))
                if piece == '要確認' and color != WHITE:
                    r.font.color.rgb = RED
                    r.font.bold = True
    return sp


def add_picture_fill(slide, path, x, y, w, h):
    """枠いっぱいに写真を置く（縦横比が合わない分はトリミング）。"""
    iw, ih = Image.open(path).size
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    box, img = w / h, iw / ih
    if img > box:
        c = (1 - box / img) / 2
        pic.crop_left = pic.crop_right = c
    else:
        c = (1 - img / box) / 2
        pic.crop_top = pic.crop_bottom = c
    return pic


def add_picture_fit(slide, path, x, y, w, h):
    iw, ih = Image.open(path).size
    s = min(w / iw, h / ih)
    pw, ph = iw * s, ih * s
    return slide.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))


def fix_ids(slide):
    seen, mx = set(), 0
    els = list(slide.shapes._spTree.iter(qn('p:cNvPr')))
    for e in els:
        mx = max(mx, int(e.get('id')))
    for e in els:
        i = int(e.get('id'))
        if i in seen:
            mx += 1
            e.set('id', str(mx))
            i = mx
        seen.add(i)


def dup_slide(prs, src_index, insert_at):
    """同じレイアウト・体裁の既存スライドを複製（画像を含まないスライドに限る）。"""
    src = prs.slides[src_index]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        remove(shp)
    for el in src.shapes._spTree.iterchildren():
        if etree.QName(el).localname in ('sp', 'grpSp', 'graphicFrame', 'cxnSp'):
            new.shapes._spTree.insert_element_before(copy.deepcopy(el), 'p:extLst')
    lst = prs.slides._sldIdLst
    ids = list(lst)
    lst.remove(ids[-1])
    lst.insert(insert_at, ids[-1])
    return new


def paste_table_xml(slide, xml_path, x, y, scale, font_scale):
    """概要書の表（graphicFrame）を位置・倍率を変えてそのまま貼る。"""
    el = etree.fromstring(open(xml_path, 'rb').read())
    xfrm = el.find(qn('p:xfrm'))
    xfrm.find(qn('a:off')).set('x', str(int(Inches(x))))
    xfrm.find(qn('a:off')).set('y', str(int(Inches(y))))
    ext = xfrm.find(qn('a:ext'))
    ext.set('cx', str(int(int(ext.get('cx')) * scale)))
    ext.set('cy', str(int(int(ext.get('cy')) * scale)))
    for gc in el.iter(qn('a:gridCol')):
        gc.set('w', str(int(int(gc.get('w')) * scale)))
    for tr in el.iter(qn('a:tr')):
        tr.set('h', str(int(int(tr.get('h')) * scale)))
    for tag in ('a:rPr', 'a:endParaRPr', 'a:defRPr'):
        for r in el.iter(qn(tag)):
            if r.get('sz'):
                r.set('sz', str(int(int(r.get('sz')) * font_scale)))
    slide.shapes._spTree.insert_element_before(el, 'p:extLst')
    return el


def ungroup_children(group):
    """グループ内の子要素を、スライド座標に直した複製として返す。"""
    g = group._element
    xfrm = g.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
    off, ext = xfrm.find(qn('a:off')), xfrm.find(qn('a:ext'))
    choff, chext = xfrm.find(qn('a:chOff')), xfrm.find(qn('a:chExt'))
    ox, oy, ex, ey = (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')))
    cx0, cy0, cex, cey = (int(choff.get('x')), int(choff.get('y')), int(chext.get('cx')), int(chext.get('cy')))
    sx, sy = (ex / cex if cex else 1), (ey / cey if cey else 1)
    out = []
    for ch in g:
        if etree.QName(ch).localname not in ('sp', 'cxnSp'):
            continue
        c = copy.deepcopy(ch)
        x = c.find('.//' + qn('a:xfrm'))
        o, e = x.find(qn('a:off')), x.find(qn('a:ext'))
        o.set('x', str(int(ox + (int(o.get('x')) - cx0) * sx)))
        o.set('y', str(int(oy + (int(o.get('y')) - cy0) * sy)))
        e.set('cx', str(int(int(e.get('cx')) * sx)))
        e.set('cy', str(int(int(e.get('cy')) * sy)))
        out.append(c)
    return out


def set_xfrm(el, x=None, y=None, w=None, h=None):
    xf = el.find('.//' + qn('a:xfrm'))
    o, e = xf.find(qn('a:off')), xf.find(qn('a:ext'))
    if x is not None:
        o.set('x', str(int(Inches(x))))
    if y is not None:
        o.set('y', str(int(Inches(y))))
    if w is not None:
        e.set('cx', str(int(Inches(w))))
    if h is not None:
        e.set('cy', str(int(Inches(h))))


def set_sp_text(el, slide, text_lines, sz=None, color=None):
    from pptx.shapes.autoshape import Shape
    set_text(Shape(el, slide.shapes), text_lines, sz=sz, color=color)


# ---------------------------------------------------------------- 組み立て
prs = Presentation(TEMPLATE)

# 構造の変更を先に済ませる：
#   ・「着地したスキーム」頁を譲渡希望条件の直後に追加（譲受企業頁＝8枚目の体裁を複製）
#   ・「事業譲渡に伴う移転手続き」頁を論点と具体施策の直後に追加（論点頁＝12枚目を複製）
scheme = dup_slide(prs, 7, 5)        # → 6枚目
procs = dup_slide(prs, 12, 13)       # 論点頁は追加後13枚目 → その直後14枚目
S = list(prs.slides)
s_cover, s_toc, s_sec1, s_seller, s_terms, s_scheme, s_pl, s_bs, s_buyer = S[0:9]
s_sec2, s_sched, s_sec3, s_issues, s_procs, s_sec4, s_learn = S[9:16]
assert s_scheme is scheme and s_procs is procs

# ---- 表紙
title = shape_by_id(s_cover, 3)
set_text(title, ['成約事例', '<売 株式会社グルメルシー × 買 株式会社日本アメニティライフ協会>'])
pp = title.text_frame.paragraphs
pp[1].runs[0].font.size = Pt(18)
l, w, h = title.left, title.width, title.height   # 位置はレイアウト継承なので先に読んでから明示する
title.left, title.width, title.height, title.top = l, w, h, Inches(3.30)
set_text(shape_by_id(s_cover, 6), ['売り主担当：要確認／買い主担当：要確認'])

# ---- 譲渡企業
set_text(shape_by_id(s_seller, 14), [
    '千葉県銚子市（工場：香取郡東庄町）で病院・福祉施設向け給食受託を営む。売上の大半は同族の医療法人向けで、'
    '3期連続の減収・2期連続の営業赤字、債務超過の状態にあった。'])
vals = {
    20: '1993年3月18日',
    24: '20,000千円',
    12: '代表取締役　内田 征一郎 様',
    27: '病院・福祉施設向け給食受託（クックチル方式のセントラルキッチン）',
    18: 'マルイチ商事㈱ 50％／医療法人厚仁会 25％／内田 征一郎 様 25％',
    30: '120,648千円（2024年8月期）',
    33: '△3,109千円（2024年8月期・調整前のEBITDA）※調整後は要確認',
    36: '84,648千円（2025年8月末簿価：建物・設備・商品等）※土地は別途売買',
}
for sid, v in vals.items():
    set_text(shape_by_id(s_seller, sid), [v], sz=9.5 if len(v) > 34 else None)
set_text(shape_by_id(s_seller, 39), [
    '・後継者の不在／事業の整理（企業概要書記載）',
    '・役員4名はいずれも現業に関与しておらず、同族の医療法人が割高な取引単価で経営を支援して事業を維持していた'])
logo = shape_by_id(s_seller, 3)
add_picture_fit(s_seller, os.path.join(SRC, 'logo.gif'), logo.left / 914400, logo.top / 914400,
                logo.width / 914400, logo.height / 914400)
remove(logo)
for sid, img in [(4, 'exterior.jpg'), (6, 'kitchen.jpg'), (7, 'medical_meal.jpg'), (8, 'paste_food.jpg')]:
    box = shape_by_id(s_seller, sid)
    add_picture_fill(s_seller, os.path.join(SRC, img), box.left / 914400, box.top / 914400,
                     box.width / 914400, box.height / 914400)
    remove(box)

# ---- 譲渡希望条件（希望 → 着地）
set_text(shape_by_id(s_terms, 24), ['事業譲渡'])
add_box(s_terms, 6.53, 2.42, 4.67, 0.25, ['（土地は所有者から買手へ別途売買）'], sz=10.5)
set_text(shape_by_id(s_terms, 56), ['2,000万円'])
add_box(s_terms, 3.55, 3.88, 3.5, 0.25, ['（役員退職慰労金を支給する場合は込み）'], sz=10, align=PP_ALIGN.CENTER)
set_text(shape_by_id(s_terms, 26), ['22,430千円（税抜）'])
add_box(s_terms, 6.53, 4.02, 4.67, 0.25, ['＋土地代金（別途売買・金額は要確認）'], sz=10.5)
cond3 = shape_by_id(s_terms, 256)
set_text(cond3, ['医療法人厚仁会からの短期借入金97,704千円の返済', '従業員の継続雇用',
                 '連帯保証の解除', '取引継続は協議'], sz=11)
cond3.width = Inches(3.9)
cond3.top = Inches(4.70)
cond3.height = Inches(1.05)
tri = shape_by_id(s_terms, 58)
tri3 = copy.deepcopy(tri._element)
s_terms.shapes._spTree.insert_element_before(tri3, 'p:extLst')
set_xfrm(tri3, x=6.62, y=5.07)
lab = shape_by_id(s_terms, 32)
lab3 = copy.deepcopy(lab._element)
s_terms.shapes._spTree.insert_element_before(lab3, 'p:extLst')
set_xfrm(lab3, x=8.13, y=4.67)
set_sp_text(lab3, s_terms, ['着地した条件'])
add_box(s_terms, 7.35, 5.00, 3.85, 0.86, [
    '負債は買手に承継せず売手に残置（第3条）',
    '従業員は個別同意で転籍、工場長は必須（第8条）',
    '主要顧客と譲渡後1年間の取引継続に協力（第9条）',
    '借入返済・連帯保証の扱いは要確認'], sz=10, anchor=MSO_ANCHOR.MIDDLE)
set_text(shape_by_id(s_terms, 264), ['後継者の不在・事業の整理'])

# ---- 着地したスキーム（新規頁）
for sh in list(s_scheme.shapes):
    if sh.shape_id not in (13, 5, 14):
        remove(sh)
set_text(shape_by_id(s_scheme, 13), ['着地したスキーム（事業譲渡）'])
set_text(shape_by_id(s_scheme, 14), [
    '債務超過の会社ごとではなく、負債を切り離して事業（建物・設備・顧客との契約・従業員）を譲渡。'
    '土地は所有者から買手へ直接売買する形で着地した。'])
add_box(s_scheme, 0.45, 3.05, 3.0, 1.45, [
    [('譲渡企業（甲）', {'sz': 10})], [('株式会社グルメルシー', {'sz': 14, 'b': True})],
    [('千葉県銚子市（工場：香取郡東庄町）', {'sz': 10})]], fill=NAVY, color=WHITE, align=PP_ALIGN.CENTER)
add_box(s_scheme, 8.25, 3.05, 3.0, 1.45, [
    [('譲受企業（乙）', {'sz': 10})], [('株式会社日本アメニティ', {'sz': 14, 'b': True})],
    [('ライフ協会', {'sz': 14, 'b': True})], [('神奈川県横浜市青葉区', {'sz': 10})]],
    fill=NAVY, color=WHITE, align=PP_ALIGN.CENTER)
add_box(s_scheme, 3.6, 2.95, 4.5, 0.42, ['事業（クックポートの建物・附属設備・什器・商品、',
                                       '顧客との契約上の地位、ノウハウ）'], sz=10, align=PP_ALIGN.CENTER)
add_box(s_scheme, 3.6, 3.40, 4.5, 0.38, [], fill=BLUE, shape=MSO_SHAPE.RIGHT_ARROW)
add_box(s_scheme, 3.6, 3.85, 4.5, 0.38, [], fill=GRAY, shape=MSO_SHAPE.LEFT_ARROW)
add_box(s_scheme, 3.6, 4.25, 4.5, 0.30, [[('譲渡対価 22,430千円（税抜）', {'b': True, 'sz': 11})]],
        align=PP_ALIGN.CENTER)
# 土地
add_box(s_scheme, 8.25, 1.55, 3.0, 0.80, [
    [('土地所有者（親族関連会社）', {'b': True, 'sz': 11})], [('※概要書上の貸主はマルイチ商事㈱', {'sz': 9.5})]],
    fill=GRAY, align=PP_ALIGN.CENTER)
add_box(s_scheme, 9.55, 2.38, 0.40, 0.64, [], fill=BLUE, shape=MSO_SHAPE.DOWN_ARROW)
add_box(s_scheme, 5.6, 2.38, 3.9, 0.55, ['土地売買（別紙3：12筆・計7,677㎡）', '代金は事業譲渡対価とは別建て（金額要確認）'],
        sz=10, align=PP_ALIGN.RIGHT)
# 顧客
add_box(s_scheme, 8.25, 5.30, 3.0, 0.85, [
    [('対象顧客', {'b': True, 'sz': 11})], [('医療法人厚仁会、介護老人保健施設、', {'sz': 10})],
    [('特別養護老人ホーム', {'sz': 10})]], fill=GRAY, align=PP_ALIGN.CENTER)
add_box(s_scheme, 9.55, 4.55, 0.40, 0.72, [], fill=BLUE, shape=MSO_SHAPE.UP_DOWN_ARROW)
add_box(s_scheme, 5.6, 4.62, 3.9, 0.62, ['取引基本契約を新たに締結（決済の前提条件）', '譲渡後1年間は現状の単価・発注数量を基準'],
        sz=10, align=PP_ALIGN.RIGHT)
# 甲に残るもの
add_box(s_scheme, 1.75, 4.55, 0.40, 0.72, [], fill=GRAY, shape=MSO_SHAPE.DOWN_ARROW)
add_box(s_scheme, 0.45, 5.30, 3.0, 1.15, [
    [('甲に残るもの（乙は承継しない）', {'b': True})],
    '・現預金、譲渡日前日までの売掛債権',
    '・借入金（厚仁会97,704千円、金融機関27,790千円）・未払金等の負債'], fill=PALE, sz=10)
# 個別に移す手続き
add_box(s_scheme, 3.6, 5.30, 4.5, 1.15, [
    [('個別に移す手続き（自動承継されない）', {'b': True})],
    '・リース／レンタル 5社37件：承継・再契約・買取（第7条）',
    '・従業員：個別同意で転籍、工場長は必須（第8条）',
    '・食品営業許可：乙が新規取得（第11条）'], fill=PALE, sz=10)
# 主要条件
add_box(s_scheme, 0.45, 6.60, 10.8, 0.85, [
    [('主要条件　', {'b': True}), ('契約締結 2026/7/23 ｜ 決済日 2026/10/30 ｜ 譲渡日 2026/11/1 ｜ 精算日 2026/12/1（いずれも契約書記載の予定日）', {})],
    [('　　　　　', {}), ('競業避止：譲渡日から10年・千葉県内 ｜ 補償：譲渡日から6か月以内の請求、上限は譲渡対価の20％', {})]],
    fill=None, line=NAVY, sz=10)

# ---- 財務ハイライト（P/L）：概要書の表を貼付
set_text(shape_by_id(s_pl, 16), [
    '3期連続の減収で2023年8月期から営業赤字。売上の大半は同族の医療法人向けで、単価は経営支援のため割高。'])
remove(shape_by_id(s_pl, 18))
paste_table_xml(s_pl, os.path.join(SRC, 'im_pl_table.xml'), 0.40, 1.55, 1.2, 1.1)
add_box(s_pl, 0.40, 7.22, 6.5, 0.30, ['※税込会計。2024年8月時点の繰越欠損金は25,433千円。'], sz=9)
set_text(shape_by_id(s_pl, 9), [
    [('≪売上≫', {'b': True})], '・3期連続の減収（136,632→128,173→120,648千円）。受注件数の減少に合わせ人員を削減し稼働を調整（聴取）',
    [('≪原価≫', {'b': True})], '・仕入高の売上比が35.0%→42.9%に上昇し、売上総利益率は65.1%→57.1%に低下',
    [('≪営業利益≫', {'b': True})], '・2023年8月期以降は営業赤字（△9,880千円、△10,505千円）',
    [('≪見通し≫', {'b': True})], '・進行期の着地見込みは要確認'], sz=10.5)
set_text(shape_by_id(s_pl, 10), [
    [('【調整項目】', {'b': True})],
    [('≪役員報酬≫', {'b': True})], '・役員4名とも報酬0千円。譲受後は工場運営の管理者人件費が別途かかると想定',
    [('≪取引単価≫', {'b': True})], '・主要顧客とは経営支援のための割高な単価で取引（概要書）。市場単価では収益性がさらに下がる可能性',
    [('✓ 調整後EBITDAは要確認（EBITDA過去3期平均は325千円）', {'b': True})]], sz=10)

# ---- 財務ハイライト（B/S）：概要書の表を貼付
set_text(shape_by_id(s_bs, 11), [
    '建物・設備は自社保有。同族の医療法人からの借入97,704千円を抱え、純資産△72,846千円の債務超過。'])
remove(shape_by_id(s_bs, 17))
paste_table_xml(s_bs, os.path.join(SRC, 'im_bs_table.xml'), 0.40, 1.55, 1.19, 1.1)
add_box(s_bs, 0.40, 7.08, 5.6, 0.45, ['短期借入金97,704千円：医療法人厚仁会（親族関連）',
                                    '長期借入金27,790千円：銚子信用金庫　※2024年8月期・税込会計'], sz=9)
set_text(shape_by_id(s_bs, 7), [
    [('≪建物≫', {'b': True})], '・建物77,967千円・附属設備9,512千円。1996年新築のセントラルキッチン（取得原価は建物214,620千円、附属設備172,682千円）。土地は親族関連会社からの賃借',
    [('≪短期借入金≫', {'b': True})], '・97,704千円は株主の医療法人厚仁会から。3期で29,347千円増加'], sz=10)
set_text(shape_by_id(s_bs, 8), [
    [('【スキームへの影響】', {'b': True})],
    [('≪純資産≫', {'b': True})], '・△72,846千円の債務超過。繰越欠損金25,433千円',
    [('≪借入金≫', {'b': True})], '・株式譲渡では借入金計125,494千円ごと引き継ぐことになり、買手がつきにくい。事業譲渡で資産と事業だけを切り出す形にした',
    '',
    [('✓ 純資産（簿価）　　　 △72,846千円', {'b': True})],
    [('✓ ネットキャッシュ　　△118,772千円', {'b': True})],
    [('✓ 譲渡対象資産（簿価）84,648千円（2025年8月末）', {'b': True})]], sz=10.5)

# ---- 譲受企業
set_text(shape_by_id(s_buyer, 14), ['神奈川県横浜市青葉区に本店を置く。特徴の一文は要確認（事業内容・規模を確認のうえ記載）。'])
for sid, v in {83: '要確認', 89: '要確認', 77: '代表取締役　江頭 瑞穂 様', 93: '要確認', 81: '要確認'}.items():
    set_text(shape_by_id(s_buyer, sid), [v])
set_text(shape_by_id(s_buyer, 97), [
    '要確認',
    '（仮説）セントラルキッチン（現状の稼働は1日約250食、製造キャパシティは日産1,000食と聴取）を取得し、'
    '給食・配食の製造拠点として活用する狙いと推察'])
set_text(shape_by_id(s_buyer, 3), ['ロゴ（要差替）'], red_kakunin=False)
for sid in (4, 6, 7, 8):
    set_text(shape_by_id(s_buyer, sid), ['本社外観・事業の写真（要差替）'], red_kakunin=False)

# ---- スケジュール
set_text(shape_by_id(s_sched, 6), [
    '2026/7/23に事業譲渡契約を先行締結し、決済日（10/30）→譲渡日（11/1）→精算日（12/1）の順で実行（いずれも契約書記載の日付）。'
    '初回面談〜トップ面談の日付と面談社数は要確認。'])
groups = [sh for sh in s_sched.shapes if sh.shape_type == 6]
proto = {}
for ch in ungroup_children(groups[0]):
    nm = ch.find('.//' + qn('p:cNvPr')).get('name')
    proto[nm] = ch
for g in groups:
    remove(g)
milestones = [
    ('要確認', '初回面談'),
    ('要確認', 'アドバイザリー\n業務契約締結'),
    ('25/1/21', '企業概要書作成\n（版の日付）'),
    ('要確認', 'トップ面談\n（社数 要確認）'),
    ('26/7/23', '事業譲渡契約\n締結'),
    ('要確認', '取引基本契約\n締結'),
    ('26/10/30', '決済日'),
    ('26/11/1', '譲渡日\n（クロージング）'),
]
tree = s_sched.shapes._spTree
for k, (date, ev) in enumerate(milestones):
    cx = 0.95 + k * 1.40
    line = copy.deepcopy(proto['直線コネクタ 6'])
    for cx_ in list(line.iter(qn('a:stCxn'))) + list(line.iter(qn('a:endCxn'))):
        cx_.getparent().remove(cx_)  # 先頭の点に接続されたままだと全線がそこへ引き寄せられる
    set_xfrm(line, x=cx, y=2.24, w=0, h=4.20)
    dot = copy.deepcopy(proto['楕円 24'])
    set_xfrm(dot, x=cx - 0.125, y=1.99, w=0.25, h=0.25)
    dl = copy.deepcopy(proto['正方形/長方形 25'])
    set_xfrm(dl, x=cx - 0.65, y=1.58, w=1.30, h=0.28)
    evb = copy.deepcopy(proto['正方形/長方形 23'])
    set_xfrm(evb, x=cx - 0.66, y=2.45 + k * 0.52, w=1.32, h=0.50)
    for el in (line, dot, dl, evb):
        tree.insert_element_before(el, 'p:extLst')
    set_sp_text(dl, s_sched, [date], sz=11)
    set_sp_text(evb, s_sched, ev.split('\n'), sz=10, color=WHITE)
    if '要確認' in ev:  # 濃紺の上の赤は読めないので白のまま太字
        from pptx.shapes.autoshape import Shape
        for p in Shape(evb, s_sched.shapes).text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = WHITE
cx5, cx8 = 0.95 + 4 * 1.40, 0.95 + 7 * 1.40
add_box(s_sched, cx5, 6.72, cx8 - cx5, 0.40,
        ['契約締結から譲渡日まで約3か月（前提条件の充足期間）'], fill=PALE, sz=10.5, align=PP_ALIGN.CENTER)

# ---- 論点と具体施策（3段を縦に広げる）
set_text(shape_by_id(s_issues, 4), ['3つの論点はいずれも「事業譲渡」を選んだことと表裏一体。売れる形にした一方、実行段階の負荷が増した。'])
rows_y, row_h = [2.05, 3.90, 5.75], 1.70
issues = [
    ([('① 重要な条件を後回しにした最終契約の先行締結', {'b': True})],
     '取り急ぎ事業譲渡契約（7/23）を先に締結し、主要顧客との取引条件の詰めを後回しにした。'
     'その結果、最後に取引基本契約書の前提条件で認識の不一致が生じ、交渉が難航した。'),
    ([('② 事業譲渡ゆえの手続きの煩雑さ', {'b': True})],
     '株式譲渡と違い、契約や許認可が自動では移らない。リース（5社37件）、水道・電話の契約切替、'
     '従業員の雇用契約のまき直し、駐車場契約の締結、食品営業許可の取得を一つずつ処理する必要があった。'),
    ([('③ マッチングの難しさ', {'b': True})],
     '親族取引（売上の大半が同族の医療法人向け・割高単価）、債務超過（△72,846千円）、2期連続の営業赤字に加え、'
     '銚子エリアという立地。会社ごとの譲渡では買手がつきにくかった。'),
]
measures = [
    ([('契約で手当てした内容', {'b': True})],
     '第9条で譲渡後1年間、現状の単価・発注数量を基準に取引を継続するよう協力する義務を置き、'
     '第18条で取引基本契約の締結を決済の前提条件とした。不一致の中身と着地のさせ方は要確認。'),
    ([('手続きの主体・期限・費用負担を契約に明記', {'b': True})],
     'リースは承継／再契約／買取から選ぶ（第7条）、従業員は個別同意で転籍し工場長は必須（第8条）、'
     '駐車場は条件据え置き（第10条）、水道光熱費等は譲渡日で区切り12/1に精算（第22条）。'),
    ([('事業譲渡で負債を切り離し、資産と立地の価値で探す', {'b': True})],
     'セントラルキッチン（建物・設備の簿価84,648千円、日産1,000食の能力）と土地（別途売買）に価値を見出す'
     '買手を探索し、横浜の事業者と成約。打診社数・トップ面談社数は要確認。'),
]
left_ids, right_ids, tri_ids = [8, 10, 11], [17, 19, 21], [23, 25, 27]
for i in range(3):
    for sid, (head, body) in ((left_ids[i], issues[i]), (right_ids[i], measures[i])):
        sh = shape_by_id(s_issues, sid)
        sh.top, sh.height = Inches(rows_y[i]), Inches(row_h)
        set_text(sh, [head, body], sz=11.5)
        sh.text_frame.margin_left = sh.text_frame.margin_right = Inches(0.10)
        sh.text_frame.word_wrap = True
    t = shape_by_id(s_issues, tri_ids[i])
    t.top = Inches(rows_y[i] + row_h / 2 - t.height / 914400 / 2)

# ---- 事業譲渡に伴う移転手続き（新規頁）
for sh in list(s_procs.shapes):
    if sh.shape_id not in (3, 2, 4):
        remove(sh)
set_text(shape_by_id(s_procs, 3), ['事業譲渡に伴う移転手続き'])
set_text(shape_by_id(s_procs, 4), ['論点②の内訳。会社ごと移る株式譲渡と違い、下表の一つずつに相手方の同意・審査・新規契約が要る。'])
rows = [
    ('項目', '対象', '契約上の手当て', '備考'),
    ('リース・レンタル', '5社37件（新規9件・再リース28件）。満了日は2026/4〜2031/2', '第7条：地位承継／乙の新規・再契約／乙の買取から選択。残債務・精算金の負担は譲渡日までに協議', '各社の審査結果は要確認'),
    ('従業員', '正社員5名・パート9名（概要書時点）', '第8条：自動承継せず、本人の同意で転籍。工場長は必須。給与は譲渡日で区分', '転籍人数は要確認'),
    ('キーマン', '工場長', '第18条：乙との雇用契約書の取得を決済の前提条件に', ''),
    ('主要顧客', '医療法人厚仁会、介護老人保健施設、特別養護老人ホーム', '第9条・第18条：取引基本契約の締結を決済の前提条件に', '論点①'),
    ('土地', '12筆・計7,677㎡（親族関連会社所有）', '第4条・第5条：別途の土地売買契約。決済日に代金支払と所有権移転登記', '売買価格は要確認'),
    ('駐車場', '厚仁会運営の介護老人保健施設との使用貸借', '第10条：譲渡日以降も条件を変更させない', ''),
    ('許認可', '食品営業許可（飲食店営業）', '第11条：乙が譲渡日までに新規取得', ''),
    ('水道・電話等', '水道・電話の契約名義の切替', '第22条：水道光熱費・公租公課・リース料は譲渡日で区分し、12/1に精算', ''),
    ('負債', '借入金・未払金・譲渡日前日までの債務', '第3条：乙は一切承継しない', ''),
]
cw = [1.35, 3.25, 4.75, 1.65]
tbl_sh = s_procs.shapes.add_table(len(rows), 4, Inches(0.34), Inches(1.65), Inches(sum(cw)), Inches(0.4 * len(rows)))
tbl = tbl_sh.table
tblPr = tbl._tbl.tblPr
for a in ('bandRow', 'firstRow'):
    tblPr.set(a, '0')
sid_el = tblPr.find(qn('a:tableStyleId'))
if sid_el is not None:
    tblPr.remove(sid_el)
for j, w in enumerate(cw):
    tbl.columns[j].width = Inches(w)
for i, row in enumerate(rows):
    tbl.rows[i].height = Inches(0.32 if i == 0 else 0.56)
    for j, v in enumerate(row):
        c = tbl.cell(i, j)
        c.margin_left = c.margin_right = Inches(0.06)
        c.margin_top = c.margin_bottom = Inches(0.03)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.fill.solid()
        c.fill.fore_color.rgb = NAVY if i == 0 else (GRAY if j == 0 else WHITE)
        tf = c.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if (i == 0 or j == 0) else PP_ALIGN.LEFT
        for piece in re.split(r'(要確認)', v or '　'):
            if not piece:
                continue
            r = p.add_run()
            r.text = piece
            r.font.size = Pt(10 if i else 10.5)
            r.font.bold = (i == 0 or j == 0 or piece == '要確認')
            r.font.color.rgb = WHITE if i == 0 else (RED if piece == '要確認' else DARK)
            r.font.name = FONT
            rp = r.font._rPr
            ea = rp.find(qn('a:ea'))
            if ea is None:
                ea = etree.SubElement(rp, qn('a:ea'))
            ea.set('typeface', FONT)
        # 罫線（細いグレー）
        tcPr = c._tc.get_or_add_tcPr()
        for ln in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
            e = etree.SubElement(tcPr, qn(ln), w='6350')
            sf = etree.SubElement(e, qn('a:solidFill'))
            etree.SubElement(sf, qn('a:srgbClr'), val='BFBFBF')
        # 罫線は塗りより前に置く必要がある
        fill = tcPr.find(qn('a:solidFill'))
        if fill is not None:
            tcPr.remove(fill)
            tcPr.append(fill)

# ---- 学び（3段を縦に広げる）
set_text(shape_by_id(s_learn, 4), ['スキームの選択で「売れる形」をつくり、実行段階の段取りで「締め切る」。両方を初期から設計しておく。'])
learn = [
    ([('重要な取引条件は最終契約の前に握る', {'b': True})],
     '主要顧客との取引継続のように事業価値の前提となる条件は、最終契約の前に条件書で合意しておく。'
     'スピードを優先して先行締結する場合も、前提条件の文言・期限・満たせなかった場合の扱いを具体的に決めておく。'),
    ([('事業譲渡は初期に「移す契約の一覧」を作る', {'b': True})],
     'リース、水道・電話、雇用、駐車場、許認可を一覧にし、担当と期限を決めて進捗を管理する。'
     'リース会社の審査など相手方の都合で日数がかかる手続きほど早く着手する。'),
    ([('株価がつかない案件は、資産と立地から買手像を描く', {'b': True})],
     '債務超過・赤字・親族取引の案件でも、事業譲渡で負債を切り離せば、建物・設備と立地に価値を見出す買手はいる。'
     '地域を絞らず、その設備を必要とする遠方の事業者まで探索範囲を広げる。'),
]
label_ids, body_ids = [9, 14, 7], [11, 16, 18]
for i in range(3):
    y = rows_y[i]
    lb = shape_by_id(s_learn, label_ids[i])
    lb.top, lb.height = Inches(y), Inches(row_h)
    bd = shape_by_id(s_learn, body_ids[i])
    bd.top, bd.height = Inches(y), Inches(row_h)
    set_text(bd, [learn[i][0], learn[i][1]], sz=12.5)
    bd.text_frame.margin_left = bd.text_frame.margin_right = Inches(0.12)
    bd.text_frame.word_wrap = True

for s in prs.slides:
    fix_ids(s)
prs.save(OUT)
print('saved', OUT)
