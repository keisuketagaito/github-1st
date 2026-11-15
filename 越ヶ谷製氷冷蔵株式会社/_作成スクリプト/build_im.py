# -*- coding: utf-8 -*-
"""越ヶ谷製氷冷蔵 企業概要書：当社テンプレート（assets/im_template.pptx）を複製し、
公式マスターの貼付範囲を paste_range で貼る。リード文・Comment の数字もマスターから読む。

    python3 build_im.py <案件マスター.xlsx> <出力.pptx> [写真フォルダ]
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.environ.get('MA_SKILL', '/root/.claude/skills/synced/88fd3bc4-f86f-4f70-8539-650c944222ea_4f59f5da-cf81-4443-8419-ba745686a7ff/ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
import im_builder as ib
from im_paste import *
from openpyxl import load_workbook
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

MASTER, OUT = sys.argv[1], sys.argv[2]
PH = sys.argv[3] if len(sys.argv) > 3 else os.path.join(HERE, '..', '写真')
wb = load_workbook(MASTER, data_only=True)
IDX = json.load(open(os.path.join(HERE, 'master_index.json')))
v = lambda sh, a: wb[sh][a].value
k = lambda x: (f'△{abs(x) / 1000:,.0f}' if x < 0 else f'{x / 1000:,.0f}')
pct = lambda x: f'{x * 100:.1f}%'
ES, BZ, ORG, FIN, APX = 'エグゼクティブサマリー', '事業概要', '組織、拠点、設備等', '財務情報詳細', 'Appendix'

prs = ib.open_template()
ib.strip_red_banners(prs)
ib.fix_broken_pictures(prs)
remove_yellow(prs); remove_null_pics(prs); remove_watermark(prs)
T = list(prs.slides)                       # テンプレートの36頁（1始まりで T[n-1]）
t = lambda n: T[n - 1]

# ---------------------------------------------------------------- 頁構成（複製してから並べ替え）
s_lic = ib.duplicate_slide(prs, 22, after=22)      # 保有設備一覧 → 許認可・取引金融機関
s_todo = ib.duplicate_slide(prs, prs.slides.index(t(34)), after=prs.slides.index(t(34)))   # その他補足 → 要確認事項
ORDER = [t(1), t(2), t(3), t(4), t(5), t(32), t(6), t(7), t(8),
         t(9), t(10), t(11), t(12), t(13), t(14), t(15),
         t(16), t(17), t(18), t(20), t(22), t(23), s_lic,
         t(24), t(25), t(26), t(27), t(28),
         t(29), t(30), t(34), s_todo, t(35), t(36)]
lst = prs.slides._sldIdLst
els = {int(e.get('id')): e for e in lst}
keep = [s.slide_id for s in ORDER]
for e in list(lst):
    if int(e.get('id')) not in keep:
        prs.part.drop_rel(e.get(qn('r:id')))
    lst.remove(e)
for sid in keep:
    lst.append(els[sid])


LEAD_PROTO = [sh for sh in t(13).shapes if sh.is_placeholder and sh.placeholder_format.idx == 1][0]._element


def ensure_lead(s):
    """字幕（リード文）のプレースホルダーが無い頁（テンプレ11・12）に、他頁の字幕を複製して置く。"""
    import copy as _c
    if not any(sh.is_placeholder and sh.placeholder_format.idx == 1 for sh in s.shapes):
        el = _c.deepcopy(LEAD_PROTO)
        s.shapes._spTree.append(el)
        ids = [int(e.get('id')) for e in s.shapes._spTree.iter(qn('p:cNvPr'))]
        el.find('.//' + qn('p:cNvPr')).set('id', str(max(ids) + 1))


def frame(s, title, section, lead=None):
    ensure_lead(s)
    ib.set_title(s, title)
    ib.set_section(s, section)
    if lead:
        set_lead(s, lead)


def drop(s, *ids):
    for sid in ids:
        try:
            remove(s, sid)
        except KeyError:
            pass


def fit(ref, target, hmax=1.0):
    return min(hmax, target / range_height(wb, ref))


def paste(s, ref, x, y, w, fs, target=None, hmax=1.0):
    hs = fit(ref, target, hmax) if target else hmax
    paste_range(s, wb, ref, x, y, w, hscale=hs, fscale=fs)
    return y + range_height(wb, ref, hscale=hs)


PL_HEAD, PL_BODY = shape_by_id(t(6), 8), shape_by_id(t(6), 9)     # Comment枠の原本（P/L頁）
ph_ = lambda a: v('PLハイライト', a)
avg3, avg2 = ph_('F26'), ph_('F27')
sales26 = ph_('F4')

# ================================================================ 1 表紙・はじめに・目次
ib.set_cover_title(t(1), '製氷業', '関東')
import pptx_table as ptab
ct = shape_by_id(t(2), 3).table
for j, val in enumerate(['株式会社M&A承継機構', '田垣内　啓介', '080-4737-4326', 'tagaito_keisuke@ma-sg.co.jp']):
    ptab.set_cell(ct.cell(1, j), val)

# ================================================================ 基本情報
s = t(5)
drop(s, 3, 4, 6, 7, 8, 15)
frame(s, '基本情報', ES, f'1960年創業の純氷メーカー。自社工場で純水を約48時間かけて凍らせる伝統製法の「純氷」を、首都圏の食品卸・氷卸へ供給。'
                         f'売上高{k(sales26)}千円（2026年3月期）、調整後EBITDAは直近3期平均{k(avg3)}千円。')
yb = paste(s, paste_ref(wb, '会社概要'), 0.34, 2.06, 6.41, 0.83, target=5.0)
place(shape_by_id(s, 2), 0.34, 1.62, 6.41, 0.39)
mix = [v('月別売上', f'C{r}') for r in (21, 22, 23)]
box, body = comment(s, PL_HEAD, PL_BODY, 6.95, 1.62, 4.40, yb - 1.62, [
    '≪65年超の業歴と伝統製法の純氷≫',
    '・純水装置（イオン交換・RO膜）でろ過した水をアイス缶で約48時間かけて凍結。透明で溶けにくい純氷を自社ブランド「キングアイス」で展開。',
    '≪首都圏の食品卸・氷卸との継続取引≫',
    '・(株)日本アクセス・三菱食品(株)等の食品卸、氷卸と取引。掛売中心（月末締め翌月末回収）。',
    '≪製氷・加工・冷蔵・配送の一貫体制≫',
    f'・製氷設備日産40トン、冷蔵設備F級1,000トン。製氷関係が売上の{pct(mix[0])}。',
    '≪工場の土地・建物を自社保有≫',
    '・越谷市平方工業団地内の工場（土地2,145㎡）を保有。賃借は松戸営業所と駐車場のみ。',
    '≪論点（要確認）≫',
    f'・借入金{k(v("BSハイライト", "F27"))}千円と代表者の保証・担保提供、工場建物の会計上の償却停止、発行済株式数の記載相違。',
], head='本件のポイント', body_dy=0.58)
for p in body.text_frame.paragraphs:
    for r in p.runs:
        r.font.size = Pt(8.5 if not r.text.startswith('≪') else 9.5)

# ================================================================ 事業概要（テンプレ32「事業計画」を転用）
s = t(32)
ib.clear_body(s)
frame(s, '事業概要', ES, f'製氷関係が売上の{pct(mix[0])}を占め、ドライアイス等の商品販売と冷蔵保管を組み合わせる。'
                        f'純水製造から凍結・加工・保管・配送までを越谷工場で一貫して担い、夏季の需要ピークに対応する。')
ib.section_bar(s, 0.34, 1.62, 5.2, '事業構成（2026年3月期）', size=12, h=0.34)
cd = CategoryChartData()
labs = [v('月別売上', f'B{r}') for r in (21, 22, 23)]
cd.categories = [f'{l.split("（")[0]} {pct(x)}' for l, x in zip(labs, mix)]
cd.add_series('兼業割合', mix)
gf = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.34), Inches(2.05), Inches(2.6), Inches(2.6), cd)
ch = gf.chart
ch.has_title = False
ch.has_legend = False
ch.plots[0].has_data_labels = False
for i, col in enumerate(['0B3041', '0889C9', '9DC3E6']):
    pt = ch.plots[0].series[0].points[i]
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(col)
for i, (l, x) in enumerate(zip(labs, mix)):
    col = [ib.NAVY, ib.BLUE, RGBColor(0x9D, 0xC3, 0xE6)][i]
    ib.rect(s, 3.05, 2.15 + i * 0.8, 0.10, 0.62, fill=col)
    ib.textbox(s, 3.25, 2.12 + i * 0.8, 2.3, 0.3, pct(x), size=14, bold=True, color=ib.NAVY, font=ib.NUM)
    ib.textbox(s, 3.25, 2.44 + i * 0.8, 2.3, 0.3, l, size=8.5)
ib.textbox(s, 0.34, 4.62, 5.2, 0.2, '出所：法人事業概況説明書（第67期）の兼業割合', size=7.5, color=ib.NOTE)
ib.section_bar(s, 5.9, 1.62, 5.45, '拠点別売上高（2026年3月期）', size=12, h=0.34)
rows = [['拠点', '売上高（千円）', '構成比', '内容']]
for r in (26, 27):
    rows.append([v('月別売上', f'B{r}'), k(v('月別売上', f'C{r}')), pct(v('月別売上', f'D{r}')), v('月別売上', f'E{r}')])
rows.append(['合計', k(v('月別売上', 'C28')), '100.0%', ''])
ib.house_table(s, 5.9, 2.05, 5.45, [1.15, 1.15, 0.8, 2.35], rows, row_h=0.32, head_h=0.3, size=9)
ib.textbox(s, 5.9, 3.42, 5.45, 0.2, '出所：売上高等の事業所別内訳書（第67期）。登記上の本店は管理のみで売上なし', size=7.5, color=ib.NOTE)
ib.section_bar(s, 5.9, 3.75, 5.45, '主要設備', size=12, h=0.34)
paste(s, IDX['EQ_MAIN'], 5.9, 4.15, 5.45, 0.78, target=0.8)
ib.section_bar(s, 0.34, 5.0, 11.01, '事業フロー', size=12, h=0.34)
steps = [('純水製造', 'イオン交換・RO膜で\n原水をろ過'), ('凍結', 'アイス缶で\n約48時間'), ('加工', 'カット・砕氷・袋詰\n手切りキューブ'),
         ('保管', '冷蔵設備で保管\n寄託物品も受託'), ('配送・販売', '自社便・委託便で\n食品卸・氷卸へ')]
for i, (h, b) in enumerate(steps):
    x = 0.34 + i * 2.22
    sp = ib.rect(s, x, 5.45, 2.12, 0.42, fill=ib.NAVY, shape=MSO_SHAPE.PENTAGON if i < 4 else MSO_SHAPE.RECTANGLE)
    ib.set_text(sp, h, size=11, bold=True, color=ib.WHITE, align=PP_ALIGN.CENTER)
    sp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    bx = ib.rect(s, x, 5.95, 1.95, 0.62, fill=ib.WHITE, line=ib.GREY)
    ib.set_text(bx, b, size=9, color=ib.BODY, align=PP_ALIGN.CENTER)
    bx.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
ib.textbox(s, 0.34, 6.72, 11.0, 0.4, '※製法・設備の記載は同社公式サイト等の公開情報および法人事業概況説明書による。原水の取水方法、製氷缶の数量・稼働率は要確認。', size=7.5, color=ib.NOTE)

# ================================================================ 財務ハイライト（P/L）
s = t(6)
drop(s, 18)
ref = paste_ref(wb, 'PLハイライト')
yb = paste(s, ref, 0.34, 1.55, 6.62, 0.875, target=5.0)
note(s, 0.34, yb + 0.1, 6.62, [f'※EBITDA＝営業利益＋減価償却費（製造原価・販管費）。調整後EBITDAの直近3期平均は{k(avg3)}千円、直近2期平均は{k(avg2)}千円。',
                               '※調整方針は株式価値評価書の正常収益力と同一。進行期の試算表は未受領。'], size=7.5)
frame(s, '財務ハイライト（P/L）', ES, f'売上高は4期連続で増加し2026年3月期に{k(sales26)}千円。電力料が高騰した2023年3月期の営業赤字から回復し、'
                                      f'臨時的な退職給与を調整した調整後EBITDAは直近3期平均{k(avg3)}千円で安定。')
power = [ph_(f'{c}30') for c in 'CDEF']
fre = [ph_(f'{c}31') for c in 'CDEF']
fill_comment(shape_by_id(s, 9), [
    '≪売上高≫',
    f'・{k(ph_("C4"))}→{k(ph_("D4"))}→{k(ph_("E4"))}→{k(ph_("F4"))}千円と4期連続で増加。7〜8月の盛夏期に年間売上の約3割が集中する。',
    '≪売上総利益≫',
    f'・23/3期は電力料（動力費）が{k(power[0])}千円に膨らみ粗利率{pct(ph_("C7"))}。24/3期以降は電力料が{k(power[1])}〜{k(power[3])}千円に落ち着き、粗利率は{pct(ph_("F7"))}〜{pct(ph_("D7"))}。',
    '≪営業利益≫',
    f'・配送運賃の増加（24/3期{k(fre[1])}千円→26/3期{k(fre[3])}千円）で25/3期以降は減益。26/3期は臨時的な退職給与{k(ph_("F33"))}千円を含む。',
    '≪EBITDA≫',
    f'・減価償却費は24/3期以降{k(ph_("D15"))}〜{k(ph_("F15"))}千円（23/3期はほぼ未計上）。調整後EBITDAは34〜37百万円で安定。',
], body=8.5, head=9.5, after=2)
fill_comment(shape_by_id(s, 10), [
    '【調整項目】',
    '≪退職給与≫',
    '・過去に発生のない臨時的な退職金支出と推察し加算（24/3期307千円・26/3期6,178千円。規程・対象者は要確認）',
    '≪役員報酬≫',
    '・現行の代表者報酬7,200千円／年は標準的な水準と想定し調整なし',
    f'✓ 調整後EBITDA（直近3期平均）は{k(avg3)}千円',
], body=8, head=9, after=1)

# ================================================================ 財務ハイライト（B/S）
s = t(7)
drop(s, 17, 9)
yb = paste(s, paste_ref(wb, 'BSハイライト'), 0.37, 1.55, 5.47, 0.875, target=5.2)
note(s, 0.37, yb + 0.1, 5.47, ['※2026年3月期末（評価基準日）。4期推移は「財務情報詳細｜貸借対照表」。'], size=7.5)
debt = v('BSハイライト', 'F27')
frame(s, '財務ハイライト（B/S）', ES, f'越谷工場の土地・建物と製氷設備を自社保有する装置型の資産構成。借入金{k(debt)}千円を抱え、'
                                      '工場建物は会計上の償却を停止しているため簿価が時価を大きく上回る点に留意を要する。')
land_ev = v('不動産(所有)', 'J4'); bld_ev = v('不動産(所有)', 'J5')
fill_comment(shape_by_id(s, 7), [
    '≪現預金・保険積立金≫',
    f'・現預金{k(v("BS(借方)", "F6"))}千円（定期預金を含む）。生命保険の積立金{k(v("BS(借方)", "F8"))}千円は決算書上「現預金（諸掛金）」に含めて表示。',
    '≪固定資産≫',
    f'・工場の土地{k(v("BS(借方)", "F27"))}千円（固定資産税評価額{k(land_ev)}千円）、建物{k(v("BS(借方)", "F22"))}千円（同{k(bld_ev)}千円）。機械{k(v("BS(借方)", "F24"))}千円は製氷設備・冷凍機等。',
    '≪借入金≫',
    f'・金融機関6行庫{k(debt - v("BS (貸方)", "F57"))}千円と代表者借入{k(v("BS (貸方)", "F57"))}千円。日本公庫には資本性ローン{k(v("BS (貸方)", "F58"))}千円を含む。代表者が連帯保証・担保提供。',
], body=8.5, head=9.5, after=2)
place(shape_by_id(s, 2), 6.08, 1.51, 5.24, 3.0); place(shape_by_id(s, 7), 6.08, 2.05, 5.24, 2.4)
place(shape_by_id(s, 6), 6.08, 4.62, 5.24, 2.95); place(shape_by_id(s, 8), 6.08, 4.70, 5.24, 2.8)
fill_comment(shape_by_id(s, 8), [
    '【時価修正】（詳細は株式価値評価書）',
    f'≪建物≫ 固定資産税評価額へ評価替え（{k(v("修正BS(借方)", "D22"))}千円）',
    f'≪土地≫ 固定資産税評価額÷0.7で時価推計（＋{k(v("修正BS(借方)", "D27"))}千円）',
    f'≪構築物・機械≫ 当期の償却不足額を減額（{k(v("修正BS(借方)", "D23") + v("修正BS(借方)", "D24"))}千円）',
    '≪保険積立金・売掛金・棚卸資産≫ 資料未受領のため簿価（要確認）',
    f'✓ 簿価純資産{k(v("修正BS(貸方) ", "C52"))}千円 → 時価純資産{k(v("修正BS(貸方) ", "E52"))}千円',
    f'✓ ネットデット{k(v("ネットキャッシュ", "F15"))}千円（代表者借入金を含む）',
], body=8, head=9, after=1)

# ================================================================ 譲渡希望条件
s = t(8)
set_text(shape_by_id(s, 19), '株式100％譲渡')
set_text(shape_by_id(s, 56), '応相談')
sh = shape_by_id(s, 256)
set_text(sh, ['従業員の継続雇用', '代表者は引継ぎ完了後に退任予定（引継ぎ期間・条件は譲受企業と協議）',
              '代表者の連帯保証・個人所有土地の担保提供の解除', '代表者借入金（1,400万円）・駐車場賃貸借の取扱いは要協議'], size=12)
place(sh, 3.45, 4.62, 7.74, 1.2)
set_text(shape_by_id(s, 264), '後継者不在')

# ================================================================ 事業内容（商流）
s = t(10)
drop(s, 2)
frame(s, '事業内容（商流）', BZ, '包装資材・ドライアイス等の商品を仕入れ、越谷工場で製造した純氷を加工・袋詰めして首都圏の食品卸・氷卸・業務用顧客へ供給。'
                               '配送は自社便と運送会社への委託を併用する。')
set_text(shape_by_id(s, 8), ['鈴木段ボール工業(株)', '大三紙業(株)', '(株)シモジマ', '日本液炭(株)', '越谷松伏水道企業団'], size=9)
place(shape_by_id(s, 8), 0.49, 2.2, 2.36, 2.1)
set_text(shape_by_id(s, 13), ['(株)日本アクセス', '三菱食品(株)', '(株)日本橋冷凍手島商店', '金木屋商事(株)', '(株)ニチレイ・アイス', '(株)クールテックサガワ', 'ほか業務用顧客'], size=9)
place(shape_by_id(s, 13), 8.84, 2.2, 2.36, 2.4)
set_text(shape_by_id(s, 39), '製氷事業（越谷工場）', size=11)
set_text(shape_by_id(s, 51), ['純水製造・凍結', '角氷（約135kg）'], size=9)
set_text(shape_by_id(s, 52), ['加工・袋詰', '加工氷・袋氷'], size=9)
set_text(shape_by_id(s, 46), '商品販売・冷蔵保管', size=11)
set_text(shape_by_id(s, 49), ['商品販売', 'ドライアイス等'], size=9)
set_text(shape_by_id(s, 50), ['冷蔵保管', '寄託物品'], size=9)
set_text(shape_by_id(s, 47), '販売拠点', size=11)
set_text(shape_by_id(s, 55), ['松戸営業所', '千葉県松戸市'], size=9)
set_text(shape_by_id(s, 56), ['配送', '自社便・委託便'], size=9)
set_text(shape_by_id(s, 12), '主要外注・委託先', size=11)
set_text(shape_by_id(s, 25), ['配送委託', '運送会社2社 等'], size=8.5)
set_text(shape_by_id(s, 26), ['人材派遣', '繁忙期の製造要員'], size=8.5)
ib.textbox(s, 0.49, 4.5, 2.36, 1.0, '※包装資材（段ボール・袋）、ドライアイス等の商品、原水（上水道）を仕入。電力は新電力と契約。', size=7.5, color=ib.NOTE)
ib.textbox(s, 8.84, 4.75, 2.36, 1.2, '※販売は月末締め翌月末回収（掛売99%）。(株)日本橋冷凍手島商店は資材の仕入先でもある。', size=7.5, color=ib.NOTE)

# ================================================================ 純氷の製造工程（テンプレ11を転用）
s = t(11)
ib.clear_body(s)
frame(s, '純氷の製造工程', BZ, 'ろ過した純水をアイス缶に注ぎ、低温のブラインで約48時間かけてゆっくり凍らせる伝統製法。'
                               '空気で撹拌しながら不純物を中心に集めて除去するため、透明度が高く硬く溶けにくい純氷ができる。')
procs = [('① 純水製造', '原水を自社の純水装置でろ過。1964年にイオン交換方式、1990年に逆浸透膜（RO）ろ過装置を導入'),
         ('② 注水', '純水をアイス缶（製氷缶）へ注ぎ、缶グリッドごとブライン槽へ'),
         ('③ 凍結', '約－12℃のブラインで約48時間かけて凍結。空気撹拌で不純物を中央に集めて吸い取る'),
         ('④ 揚氷・脱缶', '約135kgの角氷として揚氷。氷の基本となる製品'),
         ('⑤ 加工・袋詰', 'カット・砕氷・袋詰（キングアイス1.1kg〜）、手切りのダイヤアイス、板氷 等'),
         ('⑥ 保管・出荷', '冷蔵設備（F級1,000トン）で保管し、自社便・委託便で配送')]
for i, (h, b) in enumerate(procs):
    x = 0.34 + i * 1.84
    sp = ib.rect(s, x, 1.62, 1.78, 0.46, fill=ib.NAVY if i % 2 == 0 else ib.BLUE, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    ib.set_text(sp, h, size=10, bold=True, color=ib.WHITE, align=PP_ALIGN.CENTER)
    sp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    ph_ok = ib.photo_tile(s, os.path.join(PH, f'工程{i + 1}.jpg'), x + 0.05, 2.2, 1.66, 1.15, h[2:])
    by = 3.55 if os.path.exists(os.path.join(PH, '工程1.jpg')) else 2.22
    bx = ib.rect(s, x + 0.05, by, 1.66, 1.5 if by > 3 else 2.6, fill=ib.WHITE, line=ib.GREY)
    ib.set_text(bx, b, size=8.5 if by > 3 else 9.5, color=ib.BODY, line_sp=1.25)
    bx.text_frame.vertical_anchor = MSO_ANCHOR.TOP
ib.section_bar(s, 0.34, 5.25, 5.3, '純氷の特徴（公開情報）', size=12, h=0.34)
ib.textbox(s, 0.44, 5.68, 5.15, 1.4, '・結晶の並びが規則的で透明度が高く、硬く割れにくい\n・熱が伝わりにくく溶けにくい。溶け方が一定で飲料の味を損ねにくい\n'
           '・料亭・日本料理店などの飲食店や、ぶっかき氷・ロックアイス用途で使用\n・自社ブランド「キングアイス」（1.1kg／2kg／3kg袋、板氷、かき氷用キューブ 等）',
           size=9, line_sp=1.2)
ib.section_bar(s, 5.95, 5.25, 5.4, '主要設備', size=12, h=0.34)
paste(s, IDX['EQ_MAIN'], 5.95, 5.68, 5.4, 0.78)
ib.textbox(s, 0.34, 7.05, 11.0, 0.2, '※工程の記載は同社公式サイト等の公開情報による。工程写真は受領後に掲載。製氷缶の数量・1回あたり生産量・稼働率は要確認。', size=7.5, color=ib.NOTE)

# ================================================================ 特徴・強み・成長性
s = t(12)
for sid, (h, b) in zip((14, 20, 29), [
        ('65年超の業歴を持つ純氷メーカー・自社ブランド「キングアイス」',
         '1960年創業。自社の純水設備（イオン交換・RO膜）でろ過した水を、アイス缶で約48時間かけて凍結する伝統製法により、透明度が高く溶けにくい純氷を製造。角氷から袋氷・手切りキューブまで自社で加工する。'),
        ('大手食品卸を含む首都圏の安定した販路と一貫供給体制',
         '(株)日本アクセス・三菱食品(株)等の食品卸や氷卸と継続取引（掛売99%）。製氷・加工・冷蔵保管・配送を自社で一貫して担い、夏季の需要ピークにも製氷能力日産40トン・冷蔵1,000トンの設備で対応する。'),
        ('猛暑による需要拡大と設備・物流網を活かした拡張余地（仮説）',
         '夏季需要が売上の約3割を占め、猛暑の長期化は追い風。冬季は冷蔵保管で稼働を補完しており、食品・物流関連の買手の販路・配送網との連携により、閑散期の稼働向上や取扱商品拡大の余地があると想定。')]):
    sh = shape_by_id(s, sid)
    set_text(sh, [h, b])
    ps = sh.text_frame.paragraphs
    for r in ps[0].runs:
        r.font.size = Pt(16); r.font.bold = True
    for r in ps[1].runs:
        r.font.size = Pt(11); r.font.bold = False
frame(s, '特徴・強み・成長性', BZ, '65年超の業歴に裏付けられた純氷の品質と、首都圏の食品卸・氷卸への安定した販路、製氷から配送までの一貫体制が強み。'
                                  '猛暑の長期化と買手の販路・物流網との連携で拡張余地がある（仮説）。')
for sh in s.shapes:            # リード文の分だけ3バンドを下げる（1段目+0.42、2段目+0.21）
    if sh.is_placeholder:
        continue
    y_ = sh.top / 914400
    if 1.1 <= y_ < 3.3:
        sh.top = Inches(y_ + 0.42)
    elif 3.3 <= y_ < 5.5:
        sh.top = Inches(y_ + 0.21)

# ================================================================ 月別売上高（季節性）（テンプレ13を転用）
s = t(13)
drop(s, 8)
ref = paste_ref(wb, '月別売上')
months = [v('月別売上', f'{c}4') for c in 'CDEFGHIJKLMN']
sm = [v('月別売上', f'{c}5') for c in 'CDEFGHIJKLMN']
st = [v('月別売上', f'{c}6') for c in 'CDEFGHIJKLMN']
frame(s, '月別売上高（季節性）', BZ, f'売上は気温に連動する季節性が高く、7〜8月の2か月で年間の{pct(v("月別売上", "C12"))}を占める。'
                                     '冬季（11〜2月）は夏季の約1/4に落ち込む一方、冷蔵保管料が増加し閑散期の設備稼働を補っている。')
cd = CategoryChartData()
cd.categories = months
cd.add_series('売上（凍氷）', sm)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.34), Inches(1.55), Inches(6.9), Inches(3.55), cd)
ch = gf.chart
ch.has_title = False; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False
ch.font.size = Pt(8); ch.font.name = JP
pl_ = ch.plots[0]; pl_.gap_width = 60
ser = pl_.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = ib.NAVY
ser.data_labels.show_value = True; ser.data_labels.number_format = '#,##0'; ser.data_labels.number_format_is_linked = False
ser.data_labels.font.size = Pt(7.5); ser.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = ib.GREY
ch.value_axis.tick_labels.number_format = '#,##0'; ch.value_axis.tick_labels.number_format_is_linked = False
ch.value_axis.minimum_scale = 0
ib.textbox(s, 5.6, 1.5, 1.6, 0.2, '（単位：千円）', size=8, color=ib.NOTE, align=PP_ALIGN.RIGHT)
yb = paste(s, ref, 0.34, 5.3, 11.0, 0.8)
note(s, 0.34, yb + 0.08, 11.0, ['※出所：第67期（2026年3月期）法人事業概況説明書。千円単位の記載を集計しているため決算書の売上高とは端数が一致しない。'], size=7.5)
comment(s, PL_HEAD, PL_BODY, 7.45, 1.55, 3.9, 3.55, [
    '≪季節性≫',
    f'・7月{sm[3]:,}千円・8月{sm[4]:,}千円がピーク。6〜9月の4か月で年間の{pct(v("月別売上", "C13"))}。',
    f'・11〜2月は月平均{v("月別売上", "C14"):,.0f}千円。',
    '≪費用・人員≫',
    '・人件費は7月・12月に賞与支給で増加。従事員数は23〜26名でほぼ一定、繁忙期は人材派遣で補完。',
    '≪冷蔵保管≫',
    f'・保管料は11〜12月に増加（12月{st[8]:,}千円）。年末の食品保管需要と推察。',
], head='Comment', body_dy=0.55)

# ================================================================ 主要販売先（テンプレ14を転用）
s = t(14)
drop(s, 10)
ref = paste_ref(wb, '取引先')
yb = paste(s, ref, 0.34, 1.55, 11.0, 0.75, target=4.2)
note(s, 0.34, yb + 0.15, 11.0, ['※取引先別の年間売上高は未受領のため、期末売掛金残高（月末締め翌月末回収）で代替。構成比は期末売掛金合計に対する比率。取引先別売上高はデータルームにて開示予定。',
                                '※26/3期の勘定科目内訳書は上位先の合計が残高を1,724千円上回るため、その他は残高との差額で表示（(株)小野田商店の計上区分は要確認）。'], size=7.5)
frame(s, '主要販売先（期末売掛金残高）', BZ, f'大手食品卸の(株)日本アクセス・三菱食品(株)をはじめ、食品卸・氷卸との継続的な取引関係を有する。'
                                            f'上位7社で期末売掛金の{pct(v("取引先", "K21"))}を占め、特定先への過度な依存はない。')

# ================================================================ 主要仕入先・主要経費（テンプレ15を転用）
s = t(15)
drop(s, 3)
ref = paste_ref(wb, '仕入先')
yb = paste(s, ref, 0.34, 1.55, 7.2, 0.78, target=3.0)
note(s, 0.34, yb + 0.3, 7.2, ['※仕入先別の年間仕入高は未受領。第67期の勘定科目内訳書（買掛金）の期末残高で代替。'], size=7.5)
ref2 = paste_ref(wb, '主要費用')
y2 = yb + 0.72
ib.section_bar(s, 0.34, y2, 7.2, '主要費用の推移', size=11, h=0.3)
yb2 = paste(s, ref2, 0.34, y2 + 0.36, 7.2, 0.85)
frame(s, '主要仕入先・主要経費', BZ, '仕入は包装資材（段ボール・袋）とドライアイス等の商品が中心。原価の主要項目は人件費・電力料・包装資材で、'
                                     '販管費では配送運賃の比重が大きい。')
comment(s, PL_HEAD, PL_BODY, 7.8, 1.55, 3.55, 5.5, [
    '≪包装資材≫',
    f'・原材料（段ボール・ビニール袋等）の仕入高は{k(v("CR", "I9"))}千円（26/3期）。',
    '≪電力料≫',
    f'・製氷・冷蔵の電力料は23/3期{k(power[0])}千円をピークに{k(power[3])}千円（26/3期）。新電力と契約（条件は要確認）。',
    '≪配送運賃≫',
    f'・自社便に加え運送会社へ委託。売上拡大に伴い{k(fre[0])}→{k(fre[3])}千円に増加。',
    '≪外注・人材派遣≫',
    f'・外注作業費{k(v("CR", "I24"))}千円（26/3期）。繁忙期は人材派遣で製造要員を補完。外注先別の内訳は要確認。',
], head='Comment', body_dy=0.55)

# ================================================================ 組織図
s = t(17)
frame(s, '組織図', ORG, '代表取締役の下、工場長が越谷工場の現場を統括する体制。松戸営業所に販売拠点を置き、経理・総務の一部は会計事務所に委託している。')
set_box(shape_by_id(s, 72), ['株主総会'])
set_box(shape_by_id(s, 30), ['小野田 洋', '代表取締役'])
set_box(shape_by_id(s, 31), ['越谷工場', '工場長　S氏'])
set_box(shape_by_id(s, 32), ['製造', 'H氏・M氏'])
set_box(shape_by_id(s, 33), ['配送', 'K氏'])
set_box(shape_by_id(s, 34), ['松戸営業所', '責任者　M氏'])
set_box(shape_by_id(s, 35), ['販売', '従業員　5名'])
set_box(shape_by_id(s, 39), ['営業', 'M氏'])
set_box(shape_by_id(s, 40), ['営業', '1名'])
set_box(shape_by_id(s, 41), ['管理', 'C氏・S氏'])
set_box(shape_by_id(s, 42), ['財務・経理', '一部外部委託'])
set_box(shape_by_id(s, 54), ['品質検査', '代表者が担当'])
drop(s, 49, 50, 19, 20)
place(shape_by_id(s, 2), 3.86, 5.0, 0.0, 0.56)
set_box(shape_by_id(s, 82), ['【人員構成】', '2026年3月期末', '役員　：1名', '従業員：24名', '合計　：25名'])
place(shape_by_id(s, 82), 0.44, 1.75, 2.0, 1.4)
note(s, 0.34, 7.2, 11.0, ['※図中の氏名はイニシャル表記。役割分担は受領資料に基づく。従業員名簿（年齢・勤続・資格・雇用形態）は未受領のため、データルームにて開示予定。'], size=7.5)

# ================================================================ 株主及び役員
s = t(18)
drop(s, 6, 11)
frame(s, '株主及び役員', ORG, '代表取締役の小野田洋氏が発行済株式の100%を保有するオーナー企業。取締役会・監査役は設置しておらず、株式の譲渡には株主総会の承認を要する。')
place(shape_by_id(s, 2), 0.34, 1.62)
yb = paste(s, paste_ref(wb, '株主'), 0.85, 2.08, 9.6, 0.78)
place(shape_by_id(s, 3), 0.34, yb + 0.35)
yb2 = paste(s, paste_ref(wb, '取締役'), 0.85, yb + 0.8, 9.6, 0.78)
note(s, 0.34, yb2 + 0.25, 11.0, ['※発行済株式数は登記上22,000株。決算書の株主資本等変動計算書注記（200株）・法人税申告書別表二（220株）と記載が異なるため、株主名簿・定款で確認中（要確認）。',
                                 '※株券発行会社（定款の定め）のため、株式譲渡には株券の交付を要する。株券の現存状況は要確認。'], size=7.5)

# ================================================================ 拠点概要
s = t(20)
drop(s, 7, 10, 11, 16)
frame(s, '拠点概要', ORG, '越谷市平方工業団地内の自社工場（土地2,145㎡・建物約2,094㎡）を中核拠点とし、千葉県松戸市に販売拠点の営業所を賃借している。')
set_text(shape_by_id(s, 2), '拠点')
yb = paste(s, paste_ref(wb, '拠点(本社等)'), 0.34, 2.1, 11.0, 0.83)
n1 = ib.photo_tile(s, os.path.join(PH, '工場外観.jpg'), 0.34, yb + 0.3, 5.3, 3.3, '越谷工場（外観）')
n2 = ib.photo_tile(s, os.path.join(PH, '所在地図.jpg'), 6.02, yb + 0.3, 5.3, 3.3, '所在地（越谷工場・松戸営業所）')
if not (n1 or n2):
    comment(s, PL_HEAD, PL_BODY, 0.34, yb + 0.35, 11.0, 7.05 - yb - 0.35, [
        '≪越谷工場（自社所有）≫',
        f'・平方工業団地内の土地2,145㎡に1991年築の鉄筋コンクリート造の工場棟。製氷設備4台（日産40トン）・冷蔵設備（F級1,000トン）を備え、製造・加工・保管・配送の中核。売上高{k(v("月別売上", "C26"))}千円、従業員19名。',
        '≪松戸営業所（賃借）≫',
        f'・千葉県松戸市の店舗（延床400㎡）を年1,200千円で賃借。角氷・加工氷等を販売し、売上高{k(v("月別売上", "C27"))}千円、従業員5名。',
        '≪本店・駐車場≫',
        '・登記上の本店は代表者住所（越谷市蒲生寿町）。隣接の駐車場を代表者から年545千円で賃借（譲渡後の継続条件は要協議）。',
    ], head='拠点の状況', body_dy=0.55)
note(s, 0.34, 7.2, 11.0, ['※延床面積は売上高等の事業所別内訳書の記載。越谷工場建物は登記上2,093.87㎡で相違あり（要確認）。工場の外観・内観写真は受領後に掲載。'], size=7.5)

# ================================================================ 不動産詳細
s = t(22)
drop(s, 18, 22)
frame(s, '不動産詳細', ORG, f'越谷工場の土地・建物を会社所有（固定資産税評価額：土地{k(land_ev)}千円・建物{k(bld_ev)}千円）。'
                           '工場は日本政策金融公庫の担保に供されている。賃借は松戸営業所の店舗と代表者所有の駐車場のみ。')
place(shape_by_id(s, 2), 0.34, 1.62)
yb = paste(s, paste_ref(wb, '不動産(所有)'), 0.34, 2.08, 11.0, 0.72)
place(shape_by_id(s, 3), 0.34, yb + 0.35)
yb2 = paste(s, paste_ref(wb, '不動産(賃借) '), 0.34, yb + 0.8, 11.0, 0.72)
note(s, 0.34, yb2 + 0.2, 11.0, ['※固定資産税評価額は2025年3月期時点の資料による（最新年度の評価証明書で要確認）。建物の簿価は会計上の減価償却を停止しているため固定資産税評価額を大きく上回る。',
                                '※代表者個人所有の土地を栃木銀行・商工組合中央金庫の借入担保に提供している（会社資産ではないため上表に含まない。明細は要確認）。'], size=7.5)

# ================================================================ 保有設備一覧
s = t(23)
frame(s, '保有設備一覧', ORG, '製氷設備4台・袋詰砕氷設備・冷蔵設備を保有。機械は127件・取得価額375百万円に上り、近年も製氷設備の更新・修繕投資を継続している。')
ib.section_bar(s, 0.34, 1.62, 11.0, '償却資産の状況（第67期 減価償却計算書）', size=12, h=0.34)
yb = paste(s, paste_ref(wb, '設備'), 0.34, 2.05, 11.0, 0.85)
ib.section_bar(s, 0.34, yb + 0.3, 11.0, '主要設備', size=12, h=0.34)
yb2 = paste(s, IDX['EQ_MAIN'], 0.34, yb + 0.72, 8.0, 0.85)
note(s, 0.34, yb2 + 0.2, 11.0, ['※建物（越谷工場、1991年3月取得）は当期の償却を行っておらず、当期償却不足額21,745千円。償却資産の明細（185件）はデータルームにて開示予定。主要設備の更新時期・修繕計画は要確認。'], size=7.5)

# ================================================================ 許認可・取引金融機関（追加）
s = s_lic
for sh in list(s.shapes):
    if sh.has_table or sh.shape_type == 13:
        sh._element.getparent().remove(sh._element)
frame(s, '許認可・取引金融機関', ORG, '氷雪製造業・氷雪販売業の食品営業許可を取得。借入は日本政策金融公庫・地方銀行・信用金庫等の6行庫で、代表者が連帯保証・担保提供を行っている。')
ib.section_bar(s, 0.34, 1.62, 11.0, '許認可', size=12, h=0.34)
yb = paste(s, paste_ref(wb, '許認可'), 0.34, 2.05, 11.0, 0.85)
ib.section_bar(s, 0.34, yb + 0.3, 11.0, '借入金の状況（2026年3月期末）', size=12, h=0.34)
yb2 = paste(s, IDX['LOAN_RANGE'], 0.34, yb + 0.72, 11.0, 0.8)
note(s, 0.34, yb2 + 0.12, 11.0, ['※出所：第67期 借入金及び支払利息の内訳書。上記のほか埼玉県信用保証協会への保証料559,650円を支払利息に含む。借入条件（返済期限・金利・保証）は要確認。'], size=7.5)

# ================================================================ 財務情報詳細（財務入力シートをそのまま貼る）
s = t(25)
drop(s, 6, 7)
yb = paste(s, "'BS(借方)'!B3:F61", 0.34, 1.52, 5.42, 0.68, target=5.7)
paste(s, "'BS (貸方)'!B3:F25", 5.88, 1.52, 5.42, 0.68, target=2.9, hmax=0.62)
ref_na = "'BS (貸方)'!B34:F51"
hna = min(0.62, 2.3 / range_height(wb, ref_na))
paste_range(s, wb, ref_na, 5.88, yb - range_height(wb, ref_na, hscale=hna), 5.42, hscale=hna, fscale=0.68)
frame(s, '貸借対照表', FIN, f'借入金の約定返済と利益の積上げにより純資産は4期で{k(v("BS (貸方)", "C50"))}千円から{k(v("BS (貸方)", "F50"))}千円へ増加。'
                           '工場建物は会計上の償却を停止しており4期とも同額。')

for n_, ref, cid, lead, lines in [
    (26, 'PL!B2:J56', 6, f'売上総利益率は2023年3月期の{pct(ph_("C7"))}から2024年3月期以降は{pct(ph_("F7"))}〜{pct(ph_("D7"))}へ回復。営業外では年{k(v("PL", "I29"))}千円の支払利息を負担。',
     ['≪売上高≫', '・製氷・加工氷・商品の売上に、冷蔵保管料を加えたもの。決算書では支払手数料を一般業務費として販管費の外に表示しているが、本表では販管費に含めている。',
      '≪売上原価≫', '・当期製品製造原価（「製造原価報告書」）と商品（ドライアイス等）の売上原価。',
      '≪営業外損益≫', '・65期の雑収入8,482千円・雑損5,192千円の内容は要確認。67期の雑収入には保険満期返戻金2,296千円（非経常）を含む。',
      '≪法人税等≫', '・67期は繰越欠損金を使用。特別損益は4期ともなし。']),
    (27, 'CR!B4:J51', 3, '製造原価の中心は労務費・電力料（動力費）・包装資材。2024年3月期以降は製氷設備の更新に伴い減価償却費が増加している。',
     ['≪材料費≫', '・段ボール・ビニール袋等の包装資材。', '≪労務費≫', f'・26/3期は臨時的な退職給与{k(v("CR", "I17"))}千円を含む（要確認）。',
      '≪動力費≫', f'・製氷・冷蔵の電力料。23/3期{k(v("CR", "C39"))}千円→26/3期{k(v("CR", "I39"))}千円。',
      '≪減価償却費≫', f'・23/3期はほぼ未計上。26/3期{k(v("CR", "I42"))}千円。']),
    (28, 'SGA!B2:J40', 3, '販管費の中心は配送運賃を含む輸送関係費と人件費。運賃は売上拡大に伴い増加傾向。',
     ['≪人件費≫', '・役員報酬は代表者1名の7,200千円／年で4期同額。', '≪運賃・車輛費≫', f'・配送運賃{k(v("SGA", "I25"))}千円（26/3期）。自社便と運送会社への委託を併用。',
      '≪地代家賃≫', '・松戸営業所の店舗賃料1,200千円と代表者からの駐車場賃借545千円。', '≪支払手数料≫', '・66期から計上。内容は要確認。'])]:
    s = t(n_)
    paste(s, ref, 0.34, 1.5, 7.3, 0.6, target=5.75)
    place(shape_by_id(s, 2), 7.8, 1.45, 3.55, 6.0)
    place(shape_by_id(s, cid), 7.8, 2.05, 3.53, 5.3)
    fill_comment(shape_by_id(s, cid), lines, body=8.5, head=9.5, after=2)
    set_lead(s, lead)

# ================================================================ Appendix
s = t(30)
yb = paste(s, paste_ref(wb, '沿革'), 0.49, 1.6, 10.71, 0.95)
note(s, 0.49, yb + 0.12, 10.71, ['※1960〜1990年は同社公式サイト等の公開情報、1991・2008年は減価償却計算書、2012年以降は登記による。代表者の経歴は受領後に掲載。'], size=7.5)
set_lead(s, '1960年に越谷で創業。早くから純水設備を導入し、品質にこだわった純氷の製造を65年以上にわたり継続している。')
s = t(34)
yb = paste(s, paste_ref(wb, 'その他補足'), 0.49, 1.6, 10.71, 0.9)
set_lead(s, '取引条件・営業時間・取引金融機関等の補足情報。販売は月末締め翌月末回収、仕入・給与は翌月20日払い。')
s = s_todo
ib.set_title(s, '要確認事項・ネクストステップ')
yb = paste(s, paste_ref(wb, 'その他'), 0.49, 1.55, 10.71, 0.85, target=5.6)
set_lead(s, '以下は資料の追加受領・ヒアリングで確認を進めている事項。トップ面談前に確定させる事項を「高」としている。')

cp = prs.core_properties          # テンプレート作成者の個人名を残さない
cp.author = '株式会社M&A承継機構'; cp.last_modified_by = ''; cp.title = ''
prs.save(OUT)
print('saved', OUT, len(prs.slides))
