# -*- coding: utf-8 -*-
"""ティーザーのビルドスクリプト。

雛形：依頼者から受領したティーザー（A4横・5頁。表紙／はじめに／基本情報／想定スキーム／裏表紙）。
雛形には別案件の情報（社名・写真・スキーム図）が入っているため、本スクリプトで全て差し替え・削除する。
雛形のパスは第1引数（省略時は受領ファイルの保存先）。

数値はすべて案件マスター（案件マスター_ウェルカム・バスケット.xlsx）から読む。パワポに数字を直接打たない。

構成（雛形＋財務ハイライト1頁を追加）:
  1 表紙  2 はじめに  3 基本情報  4 事業の特徴・財務ハイライト  5 想定スキーム・今後の進め方  6 裏表紙
"""
import copy
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.path.join(HERE, '..', '.claude', 'skills', 'ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
from im_builder import (BLUE, BODY, GREY, NAVY, NOTE, PALE, WHITE, JP, NUM,  # noqa: E402
                        duplicate_slide, fix_broken_pictures, house_table, rect,
                        rich, set_lead, set_text, strip_red_banners, textbox)
from pptx_table import set_cell, check_tables  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

SRC = sys.argv[1] if len(sys.argv) > 1 else \
    '/root/.claude/uploads/d6622e8e-0979-56f9-b1d4-7d89babe4a83/dd688421-______________.pptx'
OUT = os.path.join(HERE, 'ティーザー_ウェルカム・バスケット_261104.pptx')

# ================================================================ マスターから読む
wb = load_workbook(os.path.join(HERE, '案件マスター_ウェルカム・バスケット.xlsx'), data_only=True)
pl, sga, cr = wb['PL'], wb['SGA'], wb['CR']
bsa, bsl = wb['BS(借方)'], wb['BS (貸方)']
PC = ['C', 'E', 'G', 'I']                 # PL系の帳簿価額列（2022〜2025）
BC = ['C', 'D', 'E', 'F']                 # BS列
YEARS = [bsa[c + '4'].value for c in BC]  # '2022年12月期' …
assert sga['B29'].value == '減価償却費' and cr['B34'].value == '減価償却費'


def row(ws, r, cols=PC):
    return [ws['%s%d' % (c, r)].value or 0 for c in cols]


sales, gp, sgaa, op = row(pl, 11), row(pl, 17), row(pl, 18), row(pl, 19)
ordinary, ni = row(pl, 36), row(pl, 56)
dep = [a + b for a, b in zip(row(cr, 34), row(sga, 29))]
ebitda = [a + b for a, b in zip(op, dep)]
temp_staff, other_out = row(cr, 18), row(cr, 27)
officer = row(sga, 6)
assets, net_assets = row(bsa, 61, BC), row(bsl, 50, BC)
cash = row(bsa, 6, BC)
debt = [a + b + c for a, b, c in zip(row(bsl, 6, BC), row(bsl, 7, BC), row(bsl, 19, BC))]


def k(v):                                 # 千円・△表記
    v = round(v / 1000)
    return ('△{:,}'.format(-v)) if v < 0 else '{:,}'.format(v)


def m(v):                                 # 百万円（概数）
    return '{:,}'.format(round(v / 1e6))


def pct(a, b, nd=1):
    if not b:
        return '－'
    v = a / b * 100
    s = ('{:.%df}%%' % nd).format(abs(v))
    return ('△' + s) if v < 0 else s


def yoy(i, s):
    return '－' if i == 0 else pct(s[i] - s[i - 1], s[i - 1])


# ================================================================ 雛形を開く
prs = Presentation(SRC)
strip_red_banners(prs)
fix_broken_pictures(prs)
S_COVER, S_INTRO, S_BASIC, S_SCHEME, S_END = [prs.slides[i] for i in range(5)]


def find(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError(name)


def remove(slide, *names):
    for sh in list(slide.shapes):
        if sh.name in names:
            sh._element.getparent().remove(sh._element)


def set_runs(shape, text):
    """1段落目・1ラン目の書式を保ったまま文字を差し替える（雛形の見出し帯・脚注用）。"""
    tf = shape.text_frame
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in p0.runs[1:]:
        r._r.getparent().remove(r._r)
    p0.runs[0].text = text


HEAD_SRC = copy.deepcopy(find(S_BASIC, '正方形/長方形 9')._element)   # グラデーションの見出し帯


def head_bar(slide, x, y, w, text, h=0.39):
    """雛形の見出し帯（紺→青グラデーション・白14pt太字）を複製して置く。"""
    el = copy.deepcopy(HEAD_SRC)
    slide.shapes._spTree.append(el)
    sh = slide.shapes[-1]
    sh.left, sh.top, sh.width, sh.height = Inches(x), Inches(y), Inches(w), Inches(h)
    set_runs(sh, text)
    ids = [int(s.shape_id) for s in slide.shapes]
    el.nvSpPr.cNvPr.set('id', str(max(ids) + 1))
    return sh


def footnote(slide, text, x=0.34, y=7.33, w=10.9):
    return textbox(slide, x, y, w, 0.22, text, size=8, color=NOTE, font=JP)


# ================================================================ 1 表紙
title = find(S_COVER, 'タイトル 2')                    # ラン構成：'ティーザー ' / '<' / 業種 / '/' / 地区 / '>' / 全角空白
runs = title.text_frame.paragraphs[0].runs
assert runs[0].text.startswith('ティーザー') and runs[1].text == '<', [r.text for r in runs]
runs[1].text = '<引越業 / 東海地区>'                    # 2ラン目（24pt）の書式で1ランにまとめる
for r in runs[2:]:
    r._r.getparent().remove(r._r)
assert title.text_frame.text == 'ティーザー <引越業 / 東海地区>', title.text_frame.text

# ================================================================ 2 はじめに
# 免責文は雛形のまま（改変しない）。照会先の担当者欄は本件担当者が未定のため●で置く（要差替え）。
tbl = find(S_INTRO, 'Google Shape;188;p2').table
set_cell(tbl.cell(1, 1), '●●　●●')
set_cell(tbl.cell(1, 2), '●●●-●●●●-●●●●')
set_cell(tbl.cell(1, 3), '●●●●@ma-sg.co.jp')

# ================================================================ 3 基本情報
set_lead(S_BASIC, '東海地方で業歴20年超の引越会社。個人向け引越を主力に、法人・医療機関の移転、'
                  '電気工事等の付帯サービスまでワンストップで提供。直近期は営業黒字へ回復。')
t = find(S_BASIC, '表 11').table
BASIC = [
    ('会社名', '株式会社ウェルカム・バスケット'),
    ('本社所在地', '愛知県名古屋市南区阿原町5（登記上の本店）'),
    ('その他拠点', '本部・名東/名南支店：愛知県日進市赤池町\n'
                  '名西支店：愛知県名古屋市中川区新家\n'
                  '事業部拠点：愛知県名古屋市中川区服部\n'
                  '岐阜支店：岐阜県羽島郡岐南町（拠点は確認中）'),
    ('設立', '2001年3月1日'),
    ('資本金', '50,000千円'),
    ('代表者', '代表取締役社長　天野 清美'),
    ('事業内容', '・個人向け引越（主力）、オフィス・医療機関等の法人移転\n'
                '・付帯サービス（電気工事、ハウスクリーニング、ピアノ輸送、\n'
                '　不用品回収・買取、家電販売）'),
    ('売上高', '約%s百万円（%s）' % (m(sales[-1]), YEARS[-1])),
    ('本件の概要', '株式譲渡（100%）を想定\n'          # このセルは雛形側に行頭記号「•」がある
                  '譲渡理由：●●●（確認中）\n'
                  '希望譲渡価額：応相談'),
]
for i, (lab, val) in enumerate(BASIC, start=1):
    set_cell(t.cell(i, 0), lab)
    set_cell(t.cell(i, 1), val, align=PP_ALIGN.LEFT)
HEIGHTS = [0.28, 0.30, 0.30, 0.98, 0.30, 0.30, 0.30, 0.66, 0.30, 0.70]
for r_, h in zip(t.rows, HEIGHTS):
    r_.height = Inches(h)
find(S_BASIC, '表 11').height = Inches(sum(HEIGHTS))

# 右側：雛形の写真（別案件）を削除し、事業の柱をカードで示す
remove(S_BASIC, '図 65')
X, W = 6.80, 4.45
head_bar(S_BASIC, X, 1.66, W, '事業の柱')
CARDS = [
    ('個人向け引越（主力）',
     '単身から家族まで。自社集客による直接受注に加え、不動産仲介・住宅関連企業等からの紹介受注チャネルを持つ'),
    ('法人・医療機関の移転',
     'オフィス移転のほか、医療機関の院内配置転換など専門性の高い作業にも対応'),
    ('付帯サービス',
     'エアコン等の電気工事、ハウスクリーニング、ピアノ輸送、不用品回収・買取、家電販売を一括で提供し、1件当たりの付加価値を高める'),
    ('エリア・拠点',
     '愛知・岐阜を中心に東海地方をカバー。複数の支店・事業部拠点で地域密着の運営'),
]
y = 2.14
for i, (ttl, txt) in enumerate(CARDS, 1):
    h = 1.12
    rect(S_BASIC, X, y, W, h, fill=WHITE, line=GREY, lw=0.75)
    num = rect(S_BASIC, X, y, 0.42, h, fill=PALE)
    set_text(num, str(i), size=16, bold=True, color=BLUE, font=NUM, align=PP_ALIGN.CENTER)
    num.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    rich(S_BASIC, X + 0.55, y + 0.10, W - 0.65, h - 0.16, [
        (ttl, {'size': 11, 'bold': True, 'color': NAVY, 'space': 3}),
        (txt, {'newpara': True, 'size': 9, 'color': BODY, 'line_sp': 1.12}),
    ])
    y += h + 0.10
remove(S_BASIC, 'テキスト ボックス 67')
footnote(S_BASIC, '※対象会社の決算書、HP記載の情報及び公開情報を基に作成。拠点・事業内容の詳細は企業概要書にて開示予定',
         x=0.34, y=7.62)

# ================================================================ 4 事業の特徴・財務ハイライト（基本情報を複製）
S_FIN = duplicate_slide(prs, 2, after=2)
for sh in list(S_FIN.shapes):
    if not (sh.is_placeholder or sh.name == 'タイトル 1'):
        sh._element.getparent().remove(sh._element)
for sh in S_FIN.shapes:
    if sh.is_placeholder and sh.placeholder_format.idx == 0:
        set_text(sh, '財務ハイライト', font=JP)
d_sales = sales[-1] - sales[-2]
set_lead(S_FIN, '直近期は減収ながら売上総利益を前期並みに維持し、販管費の圧縮で'
                '営業利益%s百万円・EBITDA%s百万円へ黒字転換。' % (m(op[-1]), m(ebitda[-1])))

head_bar(S_FIN, 0.34, 1.66, 6.55, '損益の推移')
rows = [['（単位：千円）'] + [y_.replace('年12月期', '/12期') for y_ in YEARS]]
rows += [
    ['売上高'] + [k(v) for v in sales],
    ['　前期比'] + [yoy(i, sales) for i in range(4)],
    ['売上総利益'] + [k(v) for v in gp],
    ['　売上総利益率'] + [pct(a, b) for a, b in zip(gp, sales)],
    ['販売費及び一般管理費'] + [k(v) for v in sgaa],
    ['営業利益'] + [k(v) for v in op],
    ['　営業利益率'] + [pct(a, b) for a, b in zip(op, sales)],
    ['減価償却費'] + [k(v) for v in dep],
    ['EBITDA（調整前）'] + [k(v) for v in ebitda],
    ['経常利益'] + [k(v) for v in ordinary],
    ['当期純利益'] + [k(v) for v in ni],
]
house_table(S_FIN, 0.34, 2.08, 6.55, [2.0, 1.1, 1.1, 1.1, 1.1], rows, row_h=0.24, head_h=0.26,
            size=9, bold_rows=(1, 6, 9), fills={('row', 9): PALE})

head_bar(S_FIN, 0.34, 5.08 + 0.10, 6.55, '財政状態（期末）')
rows2 = [['（単位：千円）', YEARS[-2].replace('年12月期', '/12期末'), YEARS[-1].replace('年12月期', '/12期末')],
         ['総資産', k(assets[-2]), k(assets[-1])],
         ['純資産', k(net_assets[-2]), k(net_assets[-1])],
         ['自己資本比率', pct(net_assets[-2], assets[-2]), pct(net_assets[-1], assets[-1])],
         ['現預金', k(cash[-2]), k(cash[-1])],
         ['有利子負債（金融機関借入）', k(debt[-2]), k(debt[-1])]]
house_table(S_FIN, 0.34, 5.60, 6.55, [2.35, 2.1, 2.1], rows2, row_h=0.24, head_h=0.26, size=9)

# 右：Comment（数字はすべてマスター由来の変数から組む）
CX, CW_ = 7.10, 4.25
head_bar(S_FIN, CX, 1.66, CW_, 'Comment')
rect(S_FIN, CX, 2.05, CW_, 5.00, fill=WHITE, line=GREY, lw=0.75)
H = {'size': 10, 'bold': True, 'color': NAVY, 'newpara': True, 'space': 2}
B = {'size': 9, 'color': BODY, 'newpara': True, 'line_sp': 1.15, 'space': 8}
rich(S_FIN, CX + 0.15, 2.17, CW_ - 0.30, 4.80, [
    ('≪減収下でも粗利を維持≫', H),
    ('・売上高は前期比%s減の%s百万円。一方、臨時傭員費（%s→%s百万円）・外注費（%s→%s百万円）の抑制により、'
     '売上総利益は%s百万円と前期並みを維持し、売上総利益率は%sへ改善'
     % (pct(sales[-2] - sales[-1], sales[-2]), m(sales[-1]), m(temp_staff[-2]), m(temp_staff[-1]), m(other_out[-2]), m(other_out[-1]),
        m(gp[-1]), pct(gp[-1], sales[-1])), B),
    ('≪販管費の圧縮で黒字転換≫', H),
    ('・販管費は前期比%s百万円減の%s百万円（役員報酬、広告宣伝費、人件費等）。営業利益%s百万円、EBITDA%s百万円と、'
     '2期連続の営業赤字から回復' % (m(sgaa[-2] - sgaa[-1]), m(sgaa[-1]), m(op[-1]), m(ebitda[-1])), B),
    ('≪財務基盤≫', H),
    ('・純資産%s百万円・自己資本比率%s。有利子負債は前期末比%s百万円減の%s百万円'
     % (m(net_assets[-1]), pct(net_assets[-1], assets[-1], 0), m(debt[-2] - debt[-1]), m(debt[-1])), B),
    ('≪正常収益力について≫', H),
    ('・役員報酬（直近期%s百万円）の水準、固定資産賃貸の収支、営業外収益の内容等を踏まえた調整後EBITDAは、'
     '秘密保持契約締結後の企業概要書にてお示しします' % m(officer[-1]), B),
    ('≪今後の伸び代（弊社の評価）≫', H),
    ('・外注している電気工事等の内製化、法人移転・紹介チャネルの拡大により、収益性をさらに高める余地がある', B),
])
footnote(S_FIN, '※対象会社の決算書（%s〜%s）を基に作成。EBITDA＝営業利益＋減価償却費（売上原価・販管費計上分の合計）。'
                '千円未満を四捨五入しているため合計が一致しない場合があります' % (YEARS[0], YEARS[-1]), y=7.62)

# ================================================================ 5 想定スキーム（雛形の第二会社方式の図は本件に当てはまらないため描き直す）
S_SCHEME = prs.slides[4]
for sh in list(S_SCHEME.shapes):
    if not (sh.is_placeholder or sh.name == 'タイトル 1'):
        sh._element.getparent().remove(sh._element)
for sh in S_SCHEME.shapes:
    if sh.is_placeholder and sh.placeholder_format.idx == 0:
        set_text(sh, '想定スキーム・今後の進め方', font=JP)
set_lead(S_SCHEME, '現株主が保有する対象会社株式の100%を譲渡する株式譲渡を想定。'
                   '従業員・取引先・許認可等をそのまま引き継ぎ、事業を継続いただく形を想定しています。')

head_bar(S_SCHEME, 0.99, 1.66, 9.98, '想定スキーム（株式譲渡）')


def box(slide, x, y, w, h, title, sub=None, fill=WHITE, line=NAVY, tcolor=NAVY):
    sp = rect(slide, x, y, w, h, fill=fill, line=line, lw=1.25, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    sp.adjustments[0] = 0.10
    tf = sp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(sp, title, size=12, bold=True, color=tcolor, font=JP, align=PP_ALIGN.CENTER)
    if sub:
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = sub
        r.font.size, r.font.name = Pt(9), JP
        r.font.color.rgb = NOTE if tcolor == NAVY else WHITE
    return sp


def arrow(slide, x, y, w, h, text=None, rot=0):
    a = rect(slide, x, y, w, h, fill=BLUE, shape=MSO_SHAPE.RIGHT_ARROW)
    a.rotation = rot
    if text:
        textbox(slide, x - 0.3, y - 0.30, w + 0.6, 0.26, text, size=9.5, bold=True, color=NAVY,
                align=PP_ALIGN.CENTER)
    return a


textbox(S_SCHEME, 1.40, 2.20, 3.6, 0.3, '現状', size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
textbox(S_SCHEME, 6.30, 2.20, 4.3, 0.3, '譲渡後', size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
# 現状
box(S_SCHEME, 1.80, 2.62, 2.80, 0.70, '現株主', '個人株主等')
rect(S_SCHEME, 3.17, 3.32, 0.06, 0.55, fill=NAVY)
textbox(S_SCHEME, 3.30, 3.45, 0.8, 0.25, '100%', size=9, bold=True, color=NAVY, font=NUM)
box(S_SCHEME, 1.80, 3.87, 2.80, 0.95, '対象会社', '引越運送業・付帯サービス\n従業員・取引先・許認可等')
# 譲渡の矢印
arrow(S_SCHEME, 4.95, 3.05, 1.10, 0.42, '株式100%譲渡')
arrow(S_SCHEME, 4.95, 4.05, 1.10, 0.42, None, rot=180)
textbox(S_SCHEME, 4.65, 4.52, 1.7, 0.26, '譲渡対価の支払', size=9.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
# 譲渡後
box(S_SCHEME, 6.85, 2.62, 3.20, 0.70, '譲受企業（貴社）', fill=NAVY, line=NAVY, tcolor=WHITE)
rect(S_SCHEME, 8.42, 3.32, 0.06, 0.55, fill=NAVY)
textbox(S_SCHEME, 8.55, 3.45, 0.8, 0.25, '100%', size=9, bold=True, color=NAVY, font=NUM)
box(S_SCHEME, 6.85, 3.87, 3.20, 0.95, '対象会社（100%子会社）', '事業・従業員・取引先を\nそのまま承継し事業を継続')

head_bar(S_SCHEME, 0.99, 5.15, 9.98, '今後の進め方（想定）')
STEPS = ['秘密保持契約\n概要書開示', 'トップ面談\n現地見学', '意向表明', '基本合意', 'デュー\nデリジェンス', '最終契約\nクロージング']
sx, sw, gap = 0.99, 1.58, 0.10
for i, s_ in enumerate(STEPS):
    sp = rect(S_SCHEME, sx + i * (sw + gap), 5.66, sw, 0.80,
              fill=(NAVY if i == 0 else PALE), shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    set_text(sp, s_, size=8.5, bold=True, color=(WHITE if i == 0 else NAVY), font=JP, align=PP_ALIGN.CENTER)
    sp.adjustments[0] = 0.25                                 # 矢羽根を浅くして文字幅を確保
    sp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    sp.text_frame.margin_left = sp.text_frame.margin_right = Inches(0.20)
    sp.text_frame.word_wrap = True
footnote(S_SCHEME, '※スキーム・譲渡対象・条件は現時点の想定であり、今後の協議により変更となる可能性があります。'
                   '関係会社との取引関係の取扱いを含め、詳細は秘密保持契約締結後にご説明いたします', x=0.99, y=6.70, w=9.98)

# ================================================================ 検査・保存
# 図形を消しても画像のリレーションは残り、別案件の写真がファイル内に埋まったままになる。参照の無い画像関係を切る
for s_ in prs.slides:
    xml = s_._element.xml
    for rid, rel in list(s_.part.rels.items()):
        if rel.reltype.endswith('/image') and ('"%s"' % rid) not in xml:
            s_.part.drop_rel(rid)
assert not check_tables(prs), check_tables(prs)
txt = '\n'.join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
txt += '\n'.join(c.text for s in prs.slides for sh in s.shapes if sh.has_table for r_ in sh.table.rows for c in r_.cells)
for w in ('東日本流通', '民事再生', '古橋', '千葉', '三井 尚輝', 'mitsui', '080-4342', '第二会社', '過剰債務', 'スポンサー'):
    assert w not in txt, '雛形（別案件）の情報が残っている: ' + w
prs.save(OUT)
print('saved', OUT)
