# -*- coding: utf-8 -*-
"""概要書の表を安全に編集するためのヘルパー（python-pptx）。

    from pptx_table import set_cell, resize_cols, expand_table, check_tables

既存テンプレートの表を複製・拡張すると、次の3つの壊れ方をする。本モジュールはすべて回避済み。

1. 列を足すとき <a:tc> を <a:tr> の末尾に append すると、行末の <a:extLst>（行ID）の後ろに入り
   スキーマ違反になる。PowerPoint は extLst で読み取りを打ち切るため、足した列が「空白の列」として
   表示され、表の編集もできなくなる。→ 最後の tc の直後（addnext）に挿入する。
2. gridCol / tr を deepcopy すると a16:colId / a16:rowId（Microsoft拡張のGUID）が重複し、
   PowerPoint 上で行・列の選択・挿入・削除が効かなくなる。→ 複製したら extLst を除去する。
3. 複製元セルに残った <a:br>（改行）で全セルが2行分の高さを主張し、行高が倍になって表が
   下の注記に重なる。→ set_cell で除去する。改行は <a:t> 内の「\\n」ではなく段落を分けて表現する。
"""
import copy
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_ANCHOR

BLANK = '　'   # 空欄には全角スペース（空文字だと既定サイズで行高を主張する）
JP = '游ゴシック'


def set_cell(cell, text, size=None, bold=None, color=None, align=None, fill=None,
             margins=(0.04, 0.04, 0.01, 0.01)):
    """既存ランの書式を保ったまま文字を差し替える。'\\n' は段落分割で表現。"""
    tf = cell.text_frame
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for br in p0._p.findall(qn('a:br')):
        p0._p.remove(br)
    lines = str(text if text not in ('', None) else BLANK).split('\n')
    if p0.runs:
        p0.runs[0].text = lines[0]
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        r = p0.add_run(); r.text = lines[0]; r.font.name = JP

    def style(p):
        r = p.runs[0]
        if size is not None: r.font.size = Pt(size)
        if bold is not None: r.font.bold = bold
        if color is not None: r.font.color.rgb = color
        if align is not None: p.alignment = align
    style(p0)
    for ln in lines[1:]:
        tf._txBody.append(copy.deepcopy(p0._p))
        p = tf.paragraphs[-1]; p.runs[0].text = ln; style(p)
    if fill is not None:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    l, r_, t, b = margins
    cell.margin_left, cell.margin_right = Inches(l), Inches(r_)
    cell.margin_top, cell.margin_bottom = Inches(t), Inches(b)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def strip_ids(el):
    """a16:rowId / a16:colId を持つ extLst を除去（PowerPoint が開くときに再採番する）。"""
    for ext in el.findall(qn('a:extLst')):
        el.remove(ext)
    return el


def clean_tr(tr):
    """複製した行から結合・ハイライト・IDを除く。"""
    for tc in tr.findall(qn('a:tc')):
        for a in ('rowSpan', 'vMerge', 'gridSpan', 'hMerge'):
            tc.attrib.pop(a, None)
        strip_ids(tc)
    for h in list(tr.iter(qn('a:highlight'))):
        h.getparent().remove(h)
    strip_ids(tr)
    return tr


def resize_cols(table, widths_in):
    """列数と幅（インチ）を合わせる。必ず expand_table / 値の書き込みより先に呼ぶ。"""
    tbl = table._tbl
    grid = tbl.find(qn('a:tblGrid'))
    cols = grid.findall(qn('a:gridCol'))
    n = len(widths_in)
    while len(cols) < n:
        grid.append(strip_ids(copy.deepcopy(cols[-1])))
        for tr in tbl.findall(qn('a:tr')):
            tcs = tr.findall(qn('a:tc'))
            tcs[-1].addnext(strip_ids(copy.deepcopy(tcs[-1])))   # extLst より前に入れる
        cols = grid.findall(qn('a:gridCol'))
    while len(cols) > n:
        grid.remove(cols[-1])
        for tr in tbl.findall(qn('a:tr')):
            tr.remove(tr.findall(qn('a:tc'))[-1])
        cols = grid.findall(qn('a:gridCol'))
    for c, w in zip(cols, widths_in):
        c.set('w', str(Inches(w)))
        strip_ids(c)
    for tr in tbl.findall(qn('a:tr')):
        clean_tr(tr)


def expand_table(table, nrows, proto=1):
    """行数を nrows に合わせる。proto 行を雛形に末尾へ追加（合計行などは後で上書き）。"""
    tbl = table._tbl
    trs = tbl.findall(qn('a:tr'))
    while len(trs) < nrows:
        tbl.append(clean_tr(copy.deepcopy(trs[min(proto, len(trs) - 1)])))
        trs = tbl.findall(qn('a:tr'))
    while len(trs) > nrows:
        tbl.remove(trs[-1]); trs = tbl.findall(qn('a:tr'))


def insert_rows_after(table, after_idx, count, proto_idx):
    """after_idx 行の直後に proto_idx 行の複製を count 行挿入し、挿入行の index を返す。"""
    trs = table._tbl.findall(qn('a:tr'))
    anchor = trs[after_idx]
    for _ in range(count):
        anchor.addnext(clean_tr(copy.deepcopy(trs[proto_idx])))
    return list(range(after_idx + 1, after_idx + 1 + count))


def copy_table(src_slide, dst_slide, x, y, w):
    """既存スライドの表を書式ごと複製して配置（shape id は重複しないよう振り直す）。"""
    src = next(sh for sh in src_slide.shapes if sh.has_table)
    el = copy.deepcopy(src._element)
    dst_slide.shapes._spTree.append(el)
    used = [int(e.get('id')) for e in dst_slide.shapes._spTree.iter(qn('p:cNvPr')) if e is not el.find('.//' + qn('p:cNvPr'))]
    el.find('.//' + qn('p:cNvPr')).set('id', str(max(used) + 1))
    sh = next(s for s in dst_slide.shapes if s._element is el)
    sh.left, sh.top, sh.width = Inches(x), Inches(y), Inches(w)
    return sh


def next_id(slide):
    return max(int(e.get('id')) for e in slide.shapes._spTree.iter(qn('p:cNvPr'))) + 1


def check_tables(prs, slides=None):
    """納品前チェック：ID残存・改行残存・tc順序違反・列数不一致・shape id 重複を列挙する。"""
    from lxml import etree
    probs = []
    for i, s in enumerate(prs.slides, 1):
        if slides and i not in slides:
            continue
        ids = [int(e.get('id')) for e in s.shapes._spTree.iter(qn('p:cNvPr'))]
        dup = sorted({x for x in ids if ids.count(x) > 1})
        if dup: probs.append((i, 'shape id 重複', dup))
        for sh in s.shapes:
            if not sh.has_table: continue
            tbl = sh.table._tbl
            ncol = len(tbl.find(qn('a:tblGrid')).findall(qn('a:gridCol')))
            for ri, tr in enumerate(tbl.findall(qn('a:tr'))):
                names = [etree.QName(c).localname for c in tr]
                if names.count('tc') != ncol: probs.append((i, '列数不一致', ri))
                if 'extLst' in names and 'tc' in names[names.index('extLst'):]:
                    probs.append((i, 'tc が extLst の後ろ', ri))
            if tbl.findall('.//' + qn('a:br')): probs.append((i, 'セル内に a:br 残存', sh.shape_id))
    return probs
