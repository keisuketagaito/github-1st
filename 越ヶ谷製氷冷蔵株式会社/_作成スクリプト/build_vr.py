# -*- coding: utf-8 -*-
"""越ヶ谷製氷冷蔵 株式価値評価書：当社テンプレート（assets/vr_template.pptx）を複製し、
表は公式マスター（VR→）の貼付範囲を貼る。年買法・マルチプル法・前提条件の表はテンプレートの表へ値を流し込む。

    python3 build_vr.py <案件マスター.xlsx> <出力.pptx> <算定日の表記（例：2026年11月15日）>
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.environ.get('MA_SKILL', '/root/.claude/skills/synced/88fd3bc4-f86f-4f70-8539-650c944222ea_4f59f5da-cf81-4443-8419-ba745686a7ff/ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
import im_builder as ib
import vr_builder as vb
import pptx_table as ptab
from im_paste import *
from openpyxl import load_workbook
from pptx.chart.data import CategoryChartData
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

MASTER, OUT, DATE = sys.argv[1], sys.argv[2], sys.argv[3]
wb = load_workbook(MASTER, data_only=True)
v = lambda sh, a: wb[sh][a].value
k = lambda x: (f'△{abs(x) / 1000:,.0f}' if x < 0 else f'{x / 1000:,.0f}')
S1, S2, S3, AP = '1. Valuation Summary', '2. Financial Analysis', '3. Valuation Details', 'Appendix'

prs = vb.open_vr_template()
ib.strip_red_banners(prs)
ib.fix_broken_pictures(prs)
remove_null_pics(prs); remove_watermark(prs)
T = list(prs.slides)
t = lambda n: T[n - 1]

# ---------------------------------------------------------------- 頁構成（評価科目別・Appendix 2頁・要確認事項を追加）
s_item = ib.duplicate_slide(prs, 10, after=10)          # 修正BS［資産］の後ろ → 評価科目別の検討結果
s_chk = ib.duplicate_slide(prs, prs.slides.index(t(21)), after=prs.slides.index(t(21)))     # マルチプル法の後ろ → 評価の前提と追加調整
s_ap1 = ib.duplicate_slide(prs, prs.slides.index(t(6)), after=prs.slides.index(s_chk))      # Appendix 評価手法について
s_ap2 = ib.duplicate_slide(prs, prs.slides.index(t(14)), after=prs.slides.index(s_ap1))     # Appendix 株式価値の考え方
ORDER = [t(1), t(2), t(3), t(4), t(5), t(6), t(7), t(8), t(9), t(10), t(11), s_item, t(12), t(13), t(14),
         t(15), t(16), t(17), t(18), t(19), t(20), t(21), s_chk, s_ap1, s_ap2, t(22)]
lst = prs.slides._sldIdLst
els = {int(e.get('id')): e for e in lst}
keep = [s.slide_id for s in ORDER]
for e in list(lst):
    lst.remove(e)
for sid in keep:
    lst.append(els[sid])


def drop_tables(s):
    for sh in list(s.shapes):
        if sh.has_table:
            sh._element.getparent().remove(sh._element)


def fit(ref, target, hmax=1.0):
    return min(hmax, target / range_height(wb, ref))


def paste(s, ref, x, y, w, fs, target=None, hmax=1.0):
    hs = fit(ref, target, hmax) if target else hmax
    paste_range(s, wb, ref, x, y, w, hscale=hs, fscale=fs)
    return y + range_height(wb, ref, hscale=hs)


def lines(sh, ls, size=9, bold=False):
    """テンプレートのComment本文（●●：●●●）を差し替える。テンプレートの黄色マーカー・太字は外す。"""
    set_text(sh, ls, size=size, bold=bold)
    for h in list(sh._element.iter(qn('a:highlight'))):
        h.getparent().remove(h)


nen, ev, nc = wb['年買法'], wb['EV EBITDAマルチプル'], wb['ネットキャッシュ']
na_fair, na_book = v('修正BS(貸方) ', 'E52'), v('修正BS(貸方) ', 'C52')
cost = [nen[f'{c}6'].value for c in 'KLM']
mkt = [ev[f'{c}7'].value for c in 'KLM']
sens = [nen[f'{c}14'].value for c in 'KLM']
base_op, base_eb = nen['H10'].value, ev['H10'].value
netdebt = nc['F15'].value
debt = v('BS (貸方)', 'F56')
ev_lo, ev_hi = ev['K4'].value, ev['M4'].value

# ================================================================ 表紙・はじめに
vb.set_cover_vr(t(1), '越ヶ谷製氷冷蔵株式会社', DATE)

# ================================================================ P7 株式価値の試算結果
s = t(7)
tb = shape_by_id(s, 15).table
for r, val in enumerate(['越ヶ谷製氷冷蔵株式会社', DATE, '令和8年3月期（2026年3月31日／貸借対照表の基準日）',
                         '令和6年3月期～令和8年3月期（損益計算書の参照期間）']):
    ptab.set_cell(tb.cell(r, 1), val)
set_text(shape_by_id(s, 30), f'{k(cost[0])}　～　{k(cost[2])}（千円）', size=14)
set_text(shape_by_id(s, 4), f'{k(mkt[0])}　～　{k(mkt[2])}（千円）', size=14)
s6 = t(6)
for sh in s6.shapes:
    if sh.has_table:
        for row_ in sh.table.rows:
            for c_ in row_.cells:
                for p_ in c_.text_frame.paragraphs:
                    for r_ in p_.runs:
                        r_.text = r_.text.replace('薄価', '簿価')
set_text(shape_by_id(s6, 13), ['最適な評価手法の選定にあたっては、対象会社の性質や評価目的などを踏まえ、総合的に判断する必要があります。',
                               'なお、各手法の概略につきましては、巻末「Appendix 評価手法について」をご確認ください。'])
ch = shape_by_id(s, 5).chart
cd = CategoryChartData()
cd.categories = ['インカム', 'マーケット', 'コスト']
cd.add_series('min', [0, round(mkt[2] / 1000), round(cost[2] / 1000)])
cd.add_series('差', [0, round((mkt[0] - mkt[2]) / 1000), round((cost[0] - cost[2]) / 1000)])
ch.replace_data(cd)
note(s, 0.49, 7.0, 5.6, ['※基本ケースでは両手法とも算定上マイナス（株式価値は備忘価額程度）。', f'※工場建物を簿価のまま評価した場合、年買法は{k(sens[0])}～{k(sens[2])}千円（P.21）。'], size=7.5)

# ================================================================ P9 過去BSの推移・P10 BS補足［資産］・P12 BS補足［負債・純資産］
for n_ in (9, 10, 12):
    drop_tables(t(n_))
s = t(9)
yb = paste(s, "'BS(借方)'!B3:F61", 0.37, 1.0, 5.47, 0.66, target=6.25)
paste(s, "'BS (貸方)'!B3:F25", 5.85, 1.0, 5.47, 0.66, target=3.2, hmax=0.62)
ref_na = "'BS (貸方)'!B34:F51"
hna = min(0.62, 2.4 / range_height(wb, ref_na))
paste_range(s, wb, ref_na, 5.85, yb - range_height(wb, ref_na, hscale=hna), 5.47, hscale=hna, fscale=0.66)
s = t(10)
paste(s, "'BS(借方)'!B3:F61", 0.37, 1.0, 5.95, 0.66, target=6.25)
place(shape_by_id(s, 14), 6.48, 1.0, 4.72, 6.48); place(shape_by_id(s, 9), 6.47, 1.6, 4.72, 5.8)
lines(shape_by_id(s, 9), [
    f'現預金：{k(v("BS(借方)", "C6"))}→{k(v("BS(借方)", "F6"))}千円。25/3期に定期積金を解約。定期預金は預貯金内訳書19,601千円とBS26,401千円に6,800千円の差（担保差入の有無とあわせて要確認）。',
    f'諸掛金：生命保険の積立金（26/3期{k(v("BS(借方)", "F8"))}千円、大同生命・第一生命）。決算書では現預金の内訳に表示。解約返戻金は未受領。',
    f'売掛金：{k(v("BS(借方)", "C7"))}→{k(v("BS(借方)", "F7"))}千円。月末締め翌月末回収。内訳書は上位先合計が残高を1,724千円上回る（(株)小野田商店の計上区分は要確認）。',
    f'棚卸資産：商品（加工氷・袋氷等）{k(v("BS(借方)", "F9"))}千円、原材料（包装資材）{k(v("BS(借方)", "F10"))}千円。2026年3月31日の実地棚卸による。',
    f'建物：4期とも{k(v("BS(借方)", "F22"))}千円。1991年取得の工場建物で会計上の償却を停止（当期償却不足額21,745千円）。内訳書の簿価280,673千円とも相違（要確認）。',
    f'機械装置：製氷設備・冷凍機等127件。24/3期以降に更新投資（取得）を継続。当期償却不足額3,092千円。',
    f'土地：越谷工場2,145㎡（固定資産税評価額{k(v("不動産(所有)", "J4"))}千円）。',
], size=8.5)
s = t(12)
yb = paste(s, "'BS (貸方)'!B3:F25", 0.37, 1.0, 5.95, 0.66, target=3.0, hmax=0.62)
paste(s, "'BS (貸方)'!B34:F51", 0.37, yb + 0.25, 5.95, 0.66, target=2.6, hmax=0.62)
place(shape_by_id(s, 14), 6.48, 1.0, 4.72, 6.48); place(shape_by_id(s, 9), 6.47, 1.6, 4.72, 5.8)
lines(shape_by_id(s, 9), [
    f'借入金：{k(v("BS (貸方)", "C56"))}→{k(debt)}千円と約定返済で減少。短期借入金80,000千円は継続的に借換え（条件は要確認）。',
    f'借入先：日本政策金融公庫161,252千円（うち資本性ローン105,420千円）、栃木銀行114,101千円、商工中金79,327千円、足利銀行38,013千円、埼玉縣信金27,936千円、埼玉りそな26,395千円、代表者14,000千円。',
    '保証・担保：代表者が連帯保証。栃木銀行・商工中金には代表者個人所有の土地、日本公庫には工場の土地・建物を担保提供。',
    f'未払金：{k(v("BS (貸方)", "F9"))}千円。3月分給与5,721千円・未払消費税1,676千円を含む。',
    '引当金：賞与引当金・退職給付引当金の計上なし（規程は要確認）。',
    f'純資産：{k(v("BS (貸方)", "C50"))}→{k(v("BS (貸方)", "F50"))}千円。資本金は2026年5月に10,000千円へ減資（登記）。',
], size=8.5)

# ================================================================ P11 修正BS［資産］・P13 修正BS［負債・純資産］
s = t(11)
drop_tables(s)
paste(s, paste_ref(wb, '修正BS(借方)'), 0.37, 1.0, 5.95, 0.68, target=6.3)
place(shape_by_id(s, 14), 6.48, 1.0, 4.72, 6.48); place(shape_by_id(s, 9), 6.47, 1.6, 4.72, 5.8)
lines(shape_by_id(s, 9), [
    f'建物：固定資産税評価額{k(v("不動産(所有)", "J5"))}千円へ評価替え（{k(v("修正BS(借方)", "D22"))}千円）。会計上の償却停止で簿価が実態を大きく上回るため。',
    f'土地：固定資産税評価額{k(v("不動産(所有)", "J4"))}千円÷0.7で時価{k(v("土地", "F4"))}千円と推計（＋{k(v("修正BS(借方)", "D27"))}千円）。路線価・鑑定での検証は要確認。',
    f'構築物・機械装置：減価償却計算書の当期償却不足額を減額（{k(v("修正BS(借方)", "D23") + v("修正BS(借方)", "D24"))}千円）。過年度の累積不足額は固定資産台帳で要確認。',
    '保険積立金・売掛金・棚卸資産：資料未受領のため簿価評価（解約返戻金・年齢表・滞留状況は要確認）。',
    f'評価差額合計：{k(v("修正BS(借方)", "D61"))}千円（資産）。',
], size=8.5)
s = t(13)
drop_tables(s)
paste(s, paste_ref(wb, '修正BS(貸方) '), 0.37, 1.0, 5.95, 0.68, target=6.0)
place(shape_by_id(s, 14), 6.48, 1.0, 4.72, 6.48); place(shape_by_id(s, 9), 6.47, 1.6, 4.72, 5.8)
lines(shape_by_id(s, 9), [
    '負債：簿価で評価。賞与引当金・退職給付引当金は規程未確認のため計上していない（要確認）。',
    '税効果：評価差額は純額で評価損。67期に繰越欠損金を使用しており課税所得の見込みが立たないため、繰延税金資産は認識しない（保守的）。回収可能性が認められる場合は約58,331千円の上振れ余地。',
    f'時価純資産：簿価純資産{k(na_book)}千円 → {k(na_fair)}千円。',
], size=8.5)

# ================================================================ 評価科目別の検討結果（追加）
s = s_item
drop_tables(s)
for sh in list(s.shapes):
    if sh.shape_id in (9, 14):
        sh._element.getparent().remove(sh._element)
ib.set_title(s, '評価科目別の検討結果')
yb = paste(s, paste_ref(wb, '評価科目別'), 0.37, 1.0, 10.95, 0.8)
note(s, 0.37, yb + 0.15, 10.95, ['※不動産は越谷工場（土地・建物）のみのため筆別の頁は設けず、科目ごとに評価方法と未評価の理由を一覧化した。',
                                 '※簿価評価とした科目は追加資料の受領により評価額が変動する可能性がある（DDでの追加調整余地）。'], size=8)

# ================================================================ P14 ネットキャッシュ
s = t(14)
drop_tables(s)
yb = paste(s, paste_ref(wb, 'ネットキャッシュ'), 0.37, 1.1, 10.34, 0.85, target=3.1)
place(shape_by_id(s, 19), 0.49, yb + 0.3, 10.71, 1.7); place(shape_by_id(s, 8), 1.90, yb + 0.4, 9.2, 1.5)
lines(shape_by_id(s, 8), [
    '代表者借入金14,000千円：譲渡時に返済を希望するのが通例のためデットライクに含めた（主ケース）。債務免除・相殺で整理する場合はネットデットが同額減少（参考）。',
    '日本公庫の資本性ローン105,420千円：期限一括返済の債務のため有利子負債として扱う。資本とみなす場合はネットデットが同額減少（参考）。',
    '保険積立金：換金可能な非事業用資産としてキャッシュライクに含めた（解約返戻金で要更新）。運転資金相当額は現預金から控除していない。',
], size=8.5)
note(s, 0.49, yb + 2.15, 10.71, [f'※（参考）代表者借入金を除いたネットデット{k(nc["F21"].value)}千円、さらに資本性ローンも除いたネットデット{k(nc["F22"].value)}千円。'], size=8)

# ================================================================ P16 過去PLの推移・P17 販管費の修正・P18 修正PL
s = t(16)
drop_tables(s)
yb = paste(s, 'PL!B2:J56', 0.37, 0.95, 10.83, 0.72, target=6.3)
s = t(17)
drop_tables(s)
paste(s, paste_ref(wb, '修正SGA'), 0.37, 0.85, 11.0, 0.62, target=5.35)
place(shape_by_id(s, 9), 0.49, 6.32, 10.71, 1.12)
lines(shape_by_id(s, 8), ['退職給与：24/3期307千円は臨時的な退職金支出と推察し加算（販管費。規程・対象者は要確認）。'], size=9)
lines(shape_by_id(s, 7), ['役員報酬：代表者の報酬7,200千円／年は標準的な水準と想定し調整なし。支払手数料（66・67期）は内容未確認のため調整せず（要確認）。'], size=9)
s = t(18)
drop_tables(s)
paste(s, paste_ref(wb, '修正PL'), 0.37, 0.85, 10.95, 0.62, target=5.5)
lines(shape_by_id(s, 13), [
    f'製造原価：26/3期の退職給与{k(v("CR", "I17"))}千円を臨時的支出と推察し加算（修正CR）。販管費：24/3期の退職給与307千円を加算。',
    f'修正後営業利益：{k(v("修正PL", "H19"))}／{k(v("修正PL", "K19"))}／{k(v("修正PL", "N19"))}千円。修正後EBITDA：{k(v("修正PL", "F72"))}／{k(v("修正PL", "I72"))}／{k(v("修正PL", "L72"))}千円（企業概要書の調整後EBITDAと一致）。',
    '建物の会計上の償却停止は営業利益を押し上げているが、EBITDAベースでは中立。営業外損益（支払利息・保険返戻金等）は正常収益力に含めない。',
], size=8.5)

# ================================================================ P20 年買法
s = t(20)
def fill(tbl, r, vals, c0=1):
    for j, x in enumerate(vals):
        ptab.set_cell(tbl.cell(r, c0 + j), x if isinstance(x, str) else k(x))


t6 = shape_by_id(s, 6).table
ptab.set_cell(t6.cell(1, 0), '26年3月期')
fill(t6, 2, [na_book]); fill(t6, 3, [nen['C5'].value + nen['C6'].value]); fill(t6, 4, [na_fair])
t24 = shape_by_id(s, 24).table
for j, lab in enumerate(['24年3月期', '25年3月期', '26年3月期']):
    ptab.set_cell(t24.cell(1, 1 + j), lab)
for r_, src in zip((2, 3, 4, 5, 6), (4, 5, 6, 7, 8)):
    fill(t24, r_, [nen[f'{c}{src}'].value for c in 'FGH'])
fill(t24, 8, [base_op], c0=3)
t26 = shape_by_id(s, 26).table
fill(t26, 2, [na_fair] * 3); fill(t26, 3, [nen[f'{c}5'].value for c in 'KLM']); fill(t26, 4, cost)
set_text(shape_by_id(s, 21), '・営業権倍率は中小企業M&Aのボリュームゾーン1～3倍', size=9)
set_text(shape_by_id(s, 20), '・営業利益は直近3期の加重平均（5：3：2）', size=9)
set_text(shape_by_id(s, 8), '・営業権は修正後税考慮後営業利益を基準', size=9)
for sh in s.shapes:
    if sh.shape_type == 6:
        for sub in sh.shapes:
            if sub.has_text_frame and '営業権倍率' in sub.text_frame.text and len(sub.text_frame.text) > 20:
                set_text(sub, ['・営業権倍率は買手のM&A方針・シナジー等で変動',
                               f'・建物を簿価で評価すると{k(sens[0])}～{k(sens[2])}千円とプラスに転じ、建物の評価が結論を左右する'], size=8.5)

# ================================================================ P21 EV/EBITDAマルチプル法
s = t(21)
t5 = shape_by_id(s, 5).table
ptab.set_cell(t5.cell(1, 0), '26年3月期')
fill(t5, 2, [ev['C4'].value]); fill(t5, 3, [ev['C5'].value]); fill(t5, 4, [ev['C6'].value])
t9 = shape_by_id(s, 9).table
for j, lab in enumerate(['24年3月期', '25年3月期', '26年3月期']):
    ptab.set_cell(t9.cell(1, 1 + j), lab)
for r_, src in zip((2, 3, 4, 5, 6), (4, 5, 6, 7, 8)):
    fill(t9, r_, [ev[f'{c}{src}'].value for c in 'FGH'])
fill(t9, 8, [base_eb], c0=3)
t13 = shape_by_id(s, 13).table
fill(t13, 2, [ev[f'{c}4'].value for c in 'KLM']); fill(t13, 3, [ev[f'{c}5'].value or 0 for c in 'KLM'])
fill(t13, 4, [ev['K6'].value] * 3); fill(t13, 5, mkt)
set_text(shape_by_id(s, 8), '・EBITDAは直近3期の加重平均（5：3：2）', size=9)
set_text(shape_by_id(s, 20), '・倍率は中小企業M&Aの実務水準3～5倍（製氷専業の上場会社なし）', size=9)
set_text(shape_by_id(s, 21), '・非流動性ディスカウントは標準形として適用しない', size=9)

# ================================================================ 評価の前提と追加調整の可能性（追加）
s = s_chk
drop_tables(s)
for sh in list(s.shapes):
    if sh.shape_type == 6 or (sh.has_text_frame and not sh.is_placeholder and sh.shape_id not in (15,)):
        if sh.shape_id != 15:
            sh._element.getparent().remove(sh._element)
ib.set_title(s, '評価の前提と追加調整の可能性')
rows = [['No', '論点', '現在の前提', '影響', '確認事項']]
CHK = [('工場建物', '固定資産税評価額へ評価替え（△168,020千円）', '＋', '償却停止の経緯、不動産鑑定・再調達原価による時価。簿価評価なら年買法はプラス'),
       ('税効果', '繰延税金資産を認識しない', '＋', '課税所得の見込み。認める場合は約58,331千円上振れ'),
       ('代表者借入金', 'デットライクに含める（14,000千円）', '＋', f'返済／債務免除・相殺の方針。整理する場合のマルチプル法は{k(ev["K13"].value)}～{k(ev["M13"].value)}千円'),
       ('資本性ローン', '有利子負債として扱う（105,420千円）', '＋', f'承継時の取扱い。資本とみなす場合は{k(ev["K14"].value)}～{k(ev["M14"].value)}千円'),
       ('定期預金', 'BS残高で評価', '△', '内訳書との差異6,800千円、担保差入の有無'),
       ('保険積立金', '簿価で評価', '＋／△', '解約返戻金額'),
       ('退職給付・賞与', '債務を計上しない', '△', '退職金規程・賞与の算定期間'),
       ('退職給与（臨時）', '正常収益力に加算（24/3期・26/3期）', '△', '臨時性（継続的に発生するなら加算を外す）'),
       ('倍率', '年買法1〜3年／EV/EBITDA3〜5倍', '＋／△', '買手のシナジー・類似上場会社の倍率')]
for i, (a, b, c, d) in enumerate(CHK):
    rows.append([str(i + 1), a, b, c, d])
ib.house_table(s, 0.49, 1.05, 10.71, [0.4, 1.6, 3.3, 0.7, 4.7], rows, row_h=0.42, head_h=0.32, size=9,
               aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT])
note(s, 0.49, 5.45, 10.71, ['※影響：＋は株式価値が上振れする方向、△は下振れする方向。いずれも資料の追加受領・ヒアリングで確定させる。'], size=8)
ib.set_section(s, S3)

# ================================================================ Appendix：評価手法について／株式価値の考え方
s = s_ap1
for sh in list(s.shapes):
    if not sh.is_placeholder and not (sh.has_text_frame and abs(sh.left / 914400 - 9.15) < 0.4 and sh.top / 914400 < 0.5):
        sh._element.getparent().remove(sh._element)
ib.set_title(s, '評価手法について')
ib.set_section(s, AP)
ib.set_lead(s, '株式価値の評価手法は、着目する価値の源泉により3つのアプローチに大別される')
rows = [['アプローチ', '代表的な手法', '考え方', '長所', '短所', '本件']]
rows += [['コスト', '時価純資産法\n年買法（時価純資産＋営業権）', '貸借対照表の資産・負債を時価で評価し、超過収益力を営業権として加算', '客観性が高く、中小企業M&Aで最も広く使われる', '将来の収益力を十分に反映しにくい', '採用'],
         ['マーケット', '類似会社比準法\n（EV/EBITDAマルチプル法）', '類似する上場会社・取引事例の倍率を当てはめる', '市場の評価水準を反映でき、説明しやすい', '類似会社の選定で結果が変わる', '採用'],
         ['インカム', 'DCF法・収益還元法\n配当還元法', '将来のキャッシュフロー・利益・配当を現在価値に割り引く', '将来の収益力を直接反映できる', '事業計画の精度に依存する', '不採用']]
ib.house_table(s, 0.49, 1.45, 10.71, [1.1, 2.1, 2.6, 2.1, 1.9, 0.9], rows, row_h=0.95, head_h=0.34, size=9,
               aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
vb.point_box(s, 0.49, 5.2, 10.71, '\n'.join(['・本件は詳細な事業計画が存在しないためインカムアプローチは採用しない。事業計画を策定すればDCF法による検証が可能',
                                   '・コストとマーケットの結果が乖離した場合、乖離そのものを分析結果として扱い、数字を寄せない']), h=1.0)
s = s_ap2
for sh in list(s.shapes):
    if not sh.is_placeholder and not (sh.has_text_frame and abs(sh.left / 914400 - 9.15) < 0.4 and sh.top / 914400 < 0.5):
        sh._element.getparent().remove(sh._element)
ib.set_title(s, '株式価値の考え方')
ib.set_section(s, AP)
boxes = [('事業価値（EV）', f'修正後EBITDA × 倍率\n本件：{k(ev_lo)}～{k(ev_hi)}千円', ib.BLUE),
         ('＋ キャッシュライク', f'現預金・保険積立金等\n本件：{k(ev["C4"].value)}千円', ib.BLUE),
         ('－ デットライク', f'借入金・未払法人税等\n本件：{k(ev["C5"].value)}千円', ib.NAVY),
         ('＝ 株式価値', f'本件：{k(mkt[0])}～{k(mkt[2])}千円\n（算定上マイナス）', ib.NAVY)]
for i, (h, b, col) in enumerate(boxes):
    x = 0.49 + i * 2.72
    hb = ib.rect(s, x, 1.3, 2.5, 0.45, fill=col)
    ib.set_text(hb, h, size=12, bold=True, color=ib.WHITE, align=PP_ALIGN.CENTER)
    bb = ib.rect(s, x, 1.8, 2.5, 1.1, fill=ib.PALE, line=ib.GREY)
    ib.set_text(bb, b, size=10, color=ib.BODY, align=PP_ALIGN.CENTER)
vb.point_box(s, 0.49, 3.3, 10.71, '\n'.join(['・株式価値＝事業価値＋キャッシュライク－デットライク。事業の稼ぐ力（EV）が同じでも、有利子負債が大きいほど株式価値は小さくなる',
                                   '・本件は基準EBITDAに対し有利子負債が約12.8倍と大きく、株式の対価よりも借入金の承継・保証解除が実質的な交渉条件となる']), h=1.1)
vb.point_box(s, 0.49, 4.6, 10.71, '\n'.join(['・役員借入金・資本性ローンの扱い、代表者の退職金の支給は株式価値に直接影響するため、主ケースと参考ケースを分けて示した',
                                   '・最終的な譲渡価格は、譲受企業によるデューデリジェンスの結果および当事者間の交渉により決定される']), h=1.1, label='留意点')

cp = prs.core_properties          # テンプレート作成者の個人名を残さない
cp.author = '株式会社M&A承継機構'; cp.last_modified_by = ''; cp.title = ''
prs.save(OUT)
print('saved', OUT, len(prs.slides))
