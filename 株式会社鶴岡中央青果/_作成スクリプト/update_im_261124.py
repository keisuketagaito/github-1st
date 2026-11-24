"""概要書261120（依頼者の手修正＋写真頁）に、買手候補先の質問3点への会社回答（11/24）を反映する。

- 依頼者の手修正を残すため、文言は段落単位で差し替え、表は案件マスター261124の該当セルと同じ文言に揃える
  （数値の変更はなく、変わるのは備考・説明文のみ）。
    python update_im_261124.py <概要書261120.pptx> <案件マスター261124.xlsx> <出力.pptx>
"""
import copy, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '../../.claude/skills/ma-deal-docs/scripts'))
from im_builder import *                     # noqa
from im_paste import scrub_doc_props
from openpyxl import load_workbook

src, mpath, out = sys.argv[1:4]
prs = Presentation(src)
wb = load_workbook(mpath, data_only=True)
miss = []


def shp(pno, sid):
    for sh in prs.slides[pno - 1].shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError((pno, sid))


def para_set(p, new):
    """段落の書式（先頭runの書式）を保ったまま文言を差し替える。"""
    runs = p.runs
    runs[0].text = new
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def rep(pno, sid, old, new):
    """段落内の文字列を置換。1つのrunに収まっていればそのrunだけ直す。"""
    sh = shp(pno, sid)
    for p in sh.text_frame.paragraphs:
        if old in p.text:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new)
                    return
            para_set(p, p.text.replace(old, new))
            return
    miss.append((pno, sid, old[:20]))


def add_para_after(pno, sid, anchor, new):
    sh = shp(pno, sid)
    for p in sh.text_frame.paragraphs:
        if anchor in p.text:
            el = copy.deepcopy(p._p)
            p._p.addnext(el)
            from pptx.text.text import _Paragraph
            para_set(_Paragraph(el, p._parent), new)
            return
    miss.append((pno, sid, anchor[:20]))


def cell_sync(pno, old_part, value):
    """表の中で old_part を含むセルを、マスターの値（value）に揃える。"""
    for sh in prs.slides[pno - 1].shapes:
        if not sh.has_table:
            continue
        for row in sh.table.rows:
            for c in row.cells:
                if old_part in c.text_frame.text:
                    tf = c.text_frame
                    lines = str(value).split('\n')
                    ps = tf.paragraphs
                    for i, ln in enumerate(lines):
                        if i < len(ps):
                            para_set(ps[i], ln)
                        else:
                            el = copy.deepcopy(ps[0]._p); tf._txBody.append(el)
                            from pptx.text.text import _Paragraph
                            para_set(_Paragraph(el, tf), ln)
                    for p in list(tf.paragraphs)[len(lines):]:
                        p._p.getparent().remove(p._p)
                    return
    miss.append((pno, 'table', old_part[:20]))


def add_note(pno, y, lines, tmpl=(14, 3)):
    """P14 の注記テキストボックスを複製して注記を置く（書式を揃える）。"""
    src_ = shp(*tmpl)
    el = copy.deepcopy(src_._element)
    prs.slides[pno - 1].shapes._spTree.append(el)
    sh = prs.slides[pno - 1].shapes[-1]
    sh.left, sh.top, sh.width, sh.height = Inches(0.34), Inches(y), Inches(11.0), Inches(0.17 * len(lines))
    tf = sh.text_frame
    para_set(tf.paragraphs[0], lines[0])
    for ln in lines[1:]:
        el2 = copy.deepcopy(tf.paragraphs[0]._p); tf._txBody.append(el2)
        from pptx.text.text import _Paragraph
        para_set(_Paragraph(el2, tf), ln)


M = lambda sh, ref: wb[sh][ref].value

# ---- P6 事業概要
p16 = shp(6, 16).text_frame.paragraphs
for p, t in zip(p16, ['地物：持込3~4割', '地物：集荷6~7割', '系統物・県内外市場']):
    para_set(p, t)
rep(6, 36, '生産者の商品を預かり販売し、販売額の8.5%を手数料として収受',
    '生産者の商品を預かり販売し、手数料（持込8.5%・集荷18.0%）を収受')

# ---- P11 事業内容
rep(11, 36, '■ 近郷の農家（持込30〜40%・残りは当社集荷）', '■ 近郷の農家（地物：持込30〜40%・残りは当社集荷）')
add_para_after(11, 36, '近郷の農家', '■ JA等（系統物）')
rep(11, 39, '生産者の商品を預かり販売し、販売額の8.5%を手数料として収受（純額計上）',
    '生産者の商品（地物・系統物）を預かり販売し、手数料（持込8.5%・集荷18.0%）を収受（純額計上）')

# ---- P12 事業の仕組みと収益源
rep(12, 16, '受託販売（手数料8.5%）', '受託販売（手数料：持込8.5%・集荷18.0%）')
cell_sync(12, '受託販売の手数料', M('売上構成', 'E7'))
rep(12, 30, '※手数料率8.5%から逆算した受託販売の取扱高は6か月で約235百万円（推計）。センター費は月約3.1百万円。',
    '※受託販売は手数料のみを売上計上し、委託品の取扱高は売上高の外数（2026/3期292百万円：地物243百万円・系統物等50百万円。売上高と合わせた取扱高は1,488百万円）。センター費は月約3.1百万円。')
add_para_after(12, 30, '※受託販売は手数料のみ',
    '※近郷KK部門・パッケージ部門（仲卸業務）は、市場から地物等を買い受け、買参人へ仕入額の15%以上、県内外販売先へ10%以上＋運賃、スーパーへ17%以上（センター費込み）を上乗せして販売し、部門別に採算を管理。')

# ---- P13 仕入・販売の仕組み
rep(13, 15, '・入荷の30〜40%は農家の持込で、職員が早朝に農家を回って集荷する分の方が多い。持込は数量が多くても原則すべて受け入れる。',
    '・近郷農家の地物は30〜40%が農家の持込で、残りは職員が早朝に農家を回って集荷する。持込は数量が多くても原則すべて受け入れる。地物・系統物（JA等）は受託販売で扱い、品揃えの補完分は県内外の市場の卸売業者から買い付ける。')
rep(13, 17, '・生産者の商品を預かって販売し、販売後に8.5%の手数料を受領（市場で一般的な水準）。',
    '・生産者の商品を預かって販売し、販売後に手数料を受領。市場持込は8.5%（市場で一般的な水準）、当社集荷は18.0%。2026/3期の受託取扱高は292百万円（売上高の外数）。')

# ---- P14 品目別売上
rep(14, 11, '年間取扱高は約12.9億円で3期とも横ばい。', '品目別の販売額（販売管理システム集計）は約12.9億円で3期とも横ばい。')

# ---- P16 主要得意先・P17 主要仕入先
# P16 の備考欄は依頼者が空欄にしているため、注記のみ追加する
add_note(16, 6.93, ['※社内勘定は、市場から地物等を買い受けて買参人・県内外販売先・スーパーへ販売する仲卸部門（近郷KK・パッケージ）の部門別勘定（部門採算の管理用）。'])
cell_sync(17, '性質は要確認', '\u3000' + M('仕入先', 'L15'))   # 備考欄は全角スペース始まり
para_set(shp(17, 15).text_frame.paragraphs[0],
         '決算書の仕入高（買付分）では公設庄内青果物地方卸売市場（三川町）の卸売業者2社が41.5%を占める。近郷農家の地物・JA等の系統物は受託販売で扱い、仕入高とは別に2026/3期で292百万円を取り扱う。')
add_note(17, 6.93, ['※荷主別順位表に近郷農家の地物は含まれない（系統物は含む）。地物・系統物は受託販売で手数料のみを売上計上するため、仕入高の外数（2026/3期の受託取扱高292百万円：地物243百万円・系統物等50百万円）。',
                    '※社内勘定は、市場から地物等を買い受けて買参人・県内外販売先・スーパーへ販売する仲卸部門（近郷KK・パッケージ）の部門別勘定（部門採算の管理用）。'])

# ---- 退職金（P22・P30・P34）
rep(22, 15, '※退職金規程あり（中退共・オリックス生命で外部積立。会社の追加負担は基本的にないとの認識）。',
    '※退職金規程あり（基本給×支給基準率×勤続年数）。中退共・アクサ生命（特定退職金共済）で外部積立し、外部からの支給額は規程額から控除（差額は会社負担。要支給額は確認中）。')
cell_sync(30, '受託販売手数料は8.5%', M('その他補足', 'C8'))
cell_sync(30, '賞与規定はなく', M('その他補足', 'C9'))
rep(33, 6, '受託販売は2019年10月以降、手数料（8.5%）のみの純額表示。',
    '受託販売は2019年10月以降、手数料（持込8.5%・集荷18.0%）のみの純額表示（2026/3期の受託取扱高292百万円は売上高の外数）。')
rep(34, 3, '・退職金は中退共・オリックス生命で外部積立。', '・退職金は中退共・アクサ生命（特定退職金共済）で外部積立。規程額との差額は会社負担（要支給額は確認中）。')

assert not miss, miss
force_white_headers(prs)
prs.save(out)
scrub_doc_props(out, title='企業概要書')
print('saved', out)
