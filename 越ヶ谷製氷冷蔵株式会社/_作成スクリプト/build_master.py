# -*- coding: utf-8 -*-
"""越ヶ谷製氷冷蔵：公式フォーマット（M&A承継機構 案件マスター 0925版）へ全データを転記する。

入力：ma-deal-docs スキルの assets/master_template.xlsx、data.py（決算書から転記した原データ・円単位）
出力：master_raw.xlsx（数式のみ。LibreOfficeで再計算 → inject_cache.py で値を書き戻す）

    python3 build_master.py <master_template.xlsx> <出力 master_raw.xlsx>

列：C＝2023年3月期（第64期） D＝2024年3月期 E＝2025年3月期 F＝2026年3月期（第67期・評価基準日） G＝進行期（試算表未受領のため非表示）
"""
import copy
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
SK = os.environ.get('MA_SKILL', '/root/.claude/skills/synced/88fd3bc4-f86f-4f70-8539-650c944222ea_4f59f5da-cf81-4443-8419-ba745686a7ff/ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from master_im_style import (KI, PI, YI, AI, JP, NUM, NAVY, im_sheet, cw, put, row, paste, note,
                             write_paste_index, ungroup_cols, hide_cols, copy_style, row_style)
from data import *

TEMPLATE, OUT = sys.argv[1], sys.argv[2]
wb = load_workbook(TEMPLATE)
REG = []
KIN = '#,##0,;△#,##0,;"-"'          # 円入力・千円表示（△）
P4 = 'CDEF'                          # 4期（BS・PLハイライト等）
PLC = 'CEGI'                         # PL・SGA・CRの帳簿価額列（2列1組）


def mark(ws, rng, page, cell='A1'):
    return paste(ws, rng, page, REG, cell=cell)


def hide_rows(ws, *rows):
    for r in rows:
        ws.row_dimensions[r].hidden = True


def clear(ws, r1, r2, c1, c2, style=False):
    for m in list(ws.merged_cells.ranges):
        if m.min_row <= r2 and m.max_row >= r1 and m.min_col <= c2 and m.max_col >= c1:
            ws.unmerge_cells(str(m))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(r, c)
            x.value = None
            if style:
                x.font = Font(name=JP, size=11); x.fill = PatternFill(); x.border = Border()
                x.alignment = Alignment(); x.number_format = 'General'


def red(ws, cell, t):
    ws[cell] = t
    ws[cell].font = Font(name='游ゴシック Medium', size=9, color='C00000')
    ws[cell].alignment = Alignment(wrap_text=False)


def memo(ws, cell, t, size=9):
    ws[cell] = t
    ws[cell].font = Font(name='游ゴシック Medium', size=size)
    ws[cell].alignment = Alignment(wrap_text=False)


def nf(ws, cells, fmt=KIN):
    for a in cells:
        ws[a].number_format = fmt


# ================================================================ 財務入力→
bsd, bsc, pl, cr, sga = wb['BS(借方)'], wb['BS (貸方)'], wb['PL'], wb['CR'], wb['SGA']
for c, p in zip(P4, PERIODS):
    bsd[f'{c}4'] = p
bsd['G4'] = '進行期\n（試算表未受領）'
bsd['C2'], bsd['D2'], bsd['E2'], bsd['F2'] = '第64期', '第65期', '第66期', '第67期（評価基準日）'
A = {name: vals for name, _, vals in BS_ASSETS}
cash = [sum(A[k][i] for k in ['現金', '当座預金', '普通預金', '定期預金', '定期積金']) for i in range(4)]
BSD_ROWS = {6: ('現金及び預金', cash), 7: ('売掛金', A['売掛金']), 8: ('諸掛金（保険積立金）', A['諸掛金（生命保険積立）']),
            9: ('商品', A['棚卸商品']), 10: ('原材料', A['棚卸原材料']), 11: ('仕掛品', A['仕掛品']),
            12: ('前渡金', A['前渡金']), 13: ('前払費用', A['前払費用']), 14: ('短期貸付金（従業員）', A['貸付金']),
            15: ('預け金', A['預け金']), 16: ('未収入金', A['未収金']), 17: ('仮払金', A['仮払金']),
            22: ('建物', A['建物']), 23: ('構築物', A['構築物']), 24: ('機械装置', A['機械']), 25: ('車両運搬具', A['車両運搬具']),
            26: ('工具器具備品', A['器具備品']), 27: ('土地', A['土地']),
            36: ('特許商標権', A['特許商標権']), 47: ('出資金', A['出資金'])}
for r in list(range(6, 20)) + list(range(22, 35)) + list(range(36, 46)) + list(range(47, 56)) + [57]:
    bsd[f'B{r}'] = None
    for c in P4 + 'G':
        bsd[f'{c}{r}'] = None
for r, (lab, vals) in BSD_ROWS.items():
    bsd[f'B{r}'] = lab
    for c, v in zip(P4, vals):
        bsd[f'{c}{r}'] = v
for r in range(5, 62):
    for c in P4 + 'G':
        if bsd[f'{c}{r}'].number_format != 'General':
            bsd[f'{c}{r}'].number_format = KIN
bsd['B65'] = '決算書記載の資産合計'
bsd['B66'] = '決算書との一致'
for c, v in zip(P4, BS_ASSET_TOTAL):
    bsd[f'{c}65'] = v; bsd[f'{c}65'].number_format = '#,##0'
    bsd[f'{c}66'] = f'=EXACT({c}61,{c}65)'
red(bsd, 'B68', '※第64〜66期は対象会社提供の既存会社説明資料（財務表）、第67期は第67期決算報告書から転記。現金及び預金は現金・当座・普通・定期預金・定期積金の合計。')
red(bsd, 'B69', '※諸掛金は生命保険の積立金（67期内訳：大同生命8,461,965円・第一生命900,529円）。決算書では現金預金の内訳として表示。')
red(bsd, 'B70', '※建物は会計上の減価償却を行っておらず4期同額（67期の償却不足額21,745,406円）。定期預金は預貯金内訳書19,601,032円とBS26,401,032円に6,800,000円の差（要確認）。')
hide_rows(bsd, 18, 19, 28, 29, 30, 31, 32, 33, 34, 37, 38, 39, 40, 41, 42, 43, 44, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60)
hide_cols(bsd, 'G')

L_ = {name: vals for name, _, vals in BS_LIAB}
BSC_ROWS = {6: ('支払手形', L_['支払手形']), 7: ('買掛金', L_['買掛金']), 8: ('短期借入金', L_['短期借入金']),
            9: ('未払金', L_['未払金']), 10: ('預り金', L_['預り金']), 11: ('仮受金', L_['仮受金']),
            12: ('未払法人税等', L_['税金引当金（未払法人税等）']), 19: ('長期借入金', L_['長期借入金'])}
E_ = {name: vals for name, _, vals in BS_EQUITY}
for r in list(range(6, 18)) + list(range(19, 25)) + [36, 37, 38, 40, 41, 42, 44, 45, 46, 47, 48]:
    bsc[f'B{r}'] = None
    for c in P4 + 'G':
        bsc[f'{c}{r}'] = None
for r, (lab, vals) in BSC_ROWS.items():
    bsc[f'B{r}'] = lab
    for c, v in zip(P4, vals):
        bsc[f'{c}{r}'] = v
for r, lab, k in [(36, '資本金', '資本金'), (44, '利益準備金', '利益準備金'), (45, '別途積立金', '別途積立金'), (46, '繰越利益剰余金', '繰越利益剰余金')]:
    bsc[f'B{r}'] = lab
    copy_style(bsc['B45'], bsc[f'B{r}'])
    for c, v in zip(P4, E_[k]):
        bsc[f'{c}{r}'] = v
        copy_style(bsc[f'{c}45'], bsc[f'{c}{r}'])
for r in range(5, 52):
    for c in P4 + 'G':
        if bsc[f'{c}{r}'].number_format != 'General':
            bsc[f'{c}{r}'].number_format = KIN
bsc['C49'] = bsc['D49'] = bsc['E49'] = bsc['F49'] = 0
bsc['B56'] = '有利子負債（短期＋長期借入金）'
for c in P4:
    bsc[f'{c}56'] = f'={c}8+{c}19'; bsc[f'{c}56'].number_format = '#,##0'
bsc['B57'] = 'うち代表者借入金'; bsc['F57'] = 14000000; bsc['F57'].number_format = '#,##0'
bsc['B58'] = 'うち日本政策金融公庫 資本性ローン'; bsc['F58'] = CAPITAL_LOAN_67; bsc['F58'].number_format = '#,##0'
red(bsc, 'B60', '※資本金は2026年5月15日に10,000千円へ減資（登記）。評価基準日（2026年3月31日）時点は11,000千円。')
red(bsc, 'B61', '※借入金（短期・長期）合計461,024千円の内訳は「BSハイライト」参考欄。代表者が連帯保証・個人所有土地の担保提供を行っている。')
hide_rows(bsc, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 37, 38, 40, 41, 42, 47, 48)
hide_cols(bsc, 'G')

# --- PL
P = PL
PLR = {5: ('売上（凍氷・加工氷・商品）', P['売上']), 6: ('保管料（冷蔵保管）', P['保管料']), 7: ('受入手数料', P['受入手数料']),
       12: ('期首商品棚卸高', P['期首商品棚卸高']), 13: ('商品仕入高', P['商品仕入高']),
       14: ('期末商品棚卸高', [-v for v in P['期末商品棚卸高']]),
       21: ('受取利息', P['受取利息']), 22: ('受取配当金', P['受取配当金']), 23: ('雑収入', P['雑収入']),
       29: ('支払利息', P['支払利息']), 30: ('雑損失', P['雑損']), 53: ('法人税等', P['法人税等'])}
for r in [5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 21, 22, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 49, 50, 53, 54, 55]:
    pl[f'B{r}'] = None if r not in (15,) else pl[f'B{r}'].value
    for c in PLC + 'K':
        pl[f'{c}{r}'] = None
for r, (lab, vals) in PLR.items():
    pl[f'B{r}'] = lab
    for c, v in zip(PLC, vals):
        pl[f'{c}{r}'] = v
pl['B15'] = '当期製品製造原価'
copy_style(pl['B14'], pl['B15'])
for c, cc in zip(PLC, PLC):
    pl[f'{c}15'] = f'=CR!{cc}51'
    copy_style(pl[f'{c}14'], pl[f'{c}15'])
pl['B37'] = '特別利益'; pl['B44'] = '特別損失'
for c, v in zip(PLC, PL_CHECK['当期純利益']):
    pl[f'{c}58'] = v; pl[f'{c}58'].number_format = '#,##0'
pl['B58'] = '決算書の当期純利益'; pl['B59'] = 'チェック'
pl['K58'] = None; pl['K59'] = None
pl['B60'] = '決算書の営業利益'; pl['B61'] = 'チェック'
for c, v in zip(PLC, PL_CHECK['営業利益']):
    pl[f'{c}60'] = v; pl[f'{c}60'].number_format = '#,##0'
    pl[f'{c}61'] = f'=EXACT({c}19,{c}60)'
pl['H2'] = None; pl['J2'] = '(単位：千円)'
copy_style(sga['H2'], pl['J2']); pl['J2'].alignment = Alignment(horizontal='right')
hide_cols(pl, 'K', 'L')
pl.column_dimensions['B'].width = 34
for a_, b_ in (('C3', 'D3'), ('E3', 'F3'), ('G3', 'H3'), ('I3', 'J3')):
    pl.merge_cells(f'{a_}:{b_}')
hide_rows(pl, 8, 9, 10, 24, 25, 26, 27, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 49, 50, 54, 55)
red(pl, 'B63', '※決算書では支払手数料（一般業務費）を販管費の外に表示しているが、本マスターでは販管費（SGA 38行）に含めた。営業利益は決算書と一致（61行）。')
red(pl, 'B64', '※65期の雑収入8,482千円・雑損5,192千円の内容は要確認。67期の雑収入には無配当満期定期保険の返戻金2,295,725円を含む（非経常）。特別損益は4期ともなし。')

# --- CR（製造原価報告書）
CRD = {name: dict(items) for name, items in CR}
cr_rows = {
    8: ('期首原材料棚卸高', CRD['原材料費']['原材料期首棚卸高']),
    9: ('原材料仕入高', CRD['原材料費']['原材料仕入高']),
    11: ('期末原材料棚卸高', [-v for v in CRD['原材料費']['原材料期末棚卸高']]),
    15: ('給与', CRD['労務費']['給与']), 16: ('賞与', CRD['労務費']['賞与']), 17: ('退職給与', CRD['労務費']['退職給与']),
    18: ('法定福利費', CRD['労務費']['法定福利費']), 19: ('厚生福利費', CRD['労務費']['厚生福利費']), 20: ('通勤費', CRD['労務費']['通勤費']),
    24: ('外注作業費', CRD['外注加工費']['外注作業費']),
}
ex = CRD['製造経費']
s_ = lambda *ks: [sum(ex[k][i] for k in ks) for i in range(4)]
CR_EXP = [('設備保守料・修繕費', s_('設備保守料', '修繕費')), ('損害保険料', s_('損害保険料')), ('備品・消耗品費', s_('備品費', '消耗品費')),
          ('借上料', s_('借上料')), ('動力費（電力料）', s_('動力費')), ('薬品諸油費・衛生費', s_('薬品諸油費', '衛生費')),
          ('車輛関係費', s_('車輛関係費', '車輛保険料')), ('租税公課', s_('税金公課')), ('減価償却費', s_('減価償却費')),
          ('その他経費', s_('光熱費', '通信費', '新聞図書費', '会議研修費', '組合費', '雑費'))]
for i, (lab, vals) in enumerate(CR_EXP):
    cr_rows[34 + i] = (lab, vals)
CR_DEP_ROW = 42
cr_rows[48] = ('期首仕掛品棚卸高', CRD['製造勘定']['期首仕掛品'])
cr_rows[50] = ('期末仕掛品棚卸高', CRD['製造勘定']['期末仕掛品'])
for r in list(range(8, 14)) + list(range(15, 23)) + list(range(24, 33)) + list(range(34, 46)) + [48, 50]:
    cr[f'B{r}'] = None if r not in (10,) else cr[f'B{r}'].value
    for c in PLC + 'K':
        cr[f'{c}{r}'] = None
for r, (lab, vals) in cr_rows.items():
    cr[f'B{r}'] = lab
    copy_style(cr['B9'], cr[f'B{r}'])
    for c, v in zip(PLC, vals):
        cr[f'{c}{r}'] = v
        cr[f'{c}{r}'].number_format = KIN
for r in (14, 23, 33, 46, 47, 49, 51):
    for c in PLC:
        cr[f'{c}{r}'].number_format = KIN
for c in PLC:   # テンプレートの不具合：期首仕掛品棚卸高(E48等)が SUM(E36:E47) になっていた → 値に置換済み
    pass
cr['B54'] = '決算書の当期製造原価'; cr['B55'] = 'チェック'
for c, v in zip(PLC, PL_CHECK['製造原価']):
    cr[f'{c}54'] = v; cr[f'{c}54'].number_format = '#,##0'
    cr[f'{c}55'] = f'=EXACT({c}51,{c}54)'
cr['H4'] = None; cr['J4'] = '(単位：千円)'; copy_style(sga['H2'], cr['J4']); cr['J4'].alignment = Alignment(horizontal='right')
hide_cols(cr, 'K', 'L')
cr.column_dimensions['B'].width = 30
for a_, b_ in (('C5', 'D5'), ('E5', 'F5'), ('G5', 'H5'), ('I5', 'J5')):
    cr.merge_cells(f'{a_}:{b_}')
hide_rows(cr, 7, 10, 12, 13, 21, 22, 25, 26, 27, 28, 29, 30, 31, 32, 44, 45)
red(cr, 'B57', '※製造経費は決算書の23科目を10区分に集約（設備保守料＋修繕費、備品費＋消耗品費、薬品諸油費＋衛生費、車輛関係費＋車輛保険料、光熱費・通信費・新聞図書費・会議研修費・組合費・雑費＝その他経費）。')
red(cr, 'B58', '※原材料は段ボール・ビニール袋等の包装資材。67期の退職給与6,177,985円は過去3期に発生のない臨時的支出と推察（対象者・規程は要確認）。64期は減価償却費がほぼ未計上（19,763円）。')

# --- SGA
SGD = {}
for g, items in SGA:
    SGD.update(dict(items))
SGA_ROWS = [('役員報酬', '役員報酬'), ('給料手当', '給与'), ('賞与', '賞与'), ('退職給与', '退職給与'), ('法定福利費', '法定福利費'),
            ('退職共済掛金', '退職共済金'), ('福利厚生費', '福利厚生費'), ('通勤費', '通勤費'), ('設備保守料', '設備保守料'),
            ('備品費', '備品費'), ('地代家賃', '地代家賃'), ('借上料', '借上料'), ('修繕費', '修繕費'), ('水道料', '水道料'),
            ('動力費', '動力費'), ('消耗品費', '消耗品費'), ('車輛関係費', '車輛関係費'), ('車輛保険料', '車輛保険料'),
            ('車輛借上料', '車輛借上料'), ('運賃', '運賃'), ('交際接待費', '交際接待費'), ('販売関係費', '販売関係費'),
            ('広告宣伝費', '広告宣伝費'), ('事務関係費', '事務関係費'), ('通信費', '通信費'), ('減価償却費', '減価償却費'),
            ('新聞図書費', '新聞図書費'), ('旅費交通費', '旅費交通費'), ('会議研修費', '会議研修費'), ('組合費', '組合費'),
            ('雑費', '雑費'), ('租税公課', '税金公課'), ('支払手数料', None)]
assert SGA_ROWS[25][0] == '減価償却費'   # 31行目（PLハイライト・EVが参照）
SGA_ROW = {}
for i, (lab, key) in enumerate(SGA_ROWS):
    r = 6 + i
    sga[f'B{r}'] = lab
    vals = SGD[key] if key else P['支払手数料']
    for c, v in zip(PLC, vals):
        sga[f'{c}{r}'] = v
        sga[f'{c}{r}'].number_format = KIN
    SGA_ROW[lab] = r
sga['B39'] = None
for c in PLC + 'K':
    sga[f'{c}39'] = None
for r in range(6, 41):   # 売上高比：1期目も IFERROR で
    sga[f'D{r}'] = f'=IFERROR(C{r}/$C$5,"")'
    for v_, d_ in (('E', 'F'), ('G', 'H'), ('I', 'J')):
        sga[f'{d_}{r}'] = f'=IFERROR({v_}{r}/${v_}$5,"")'
nf(sga, [f'{c}40' for c in PLC])
for r in range(5, 41):   # 進行期（非表示）の売上高比は分母0のため外す
    sga[f'L{r}'] = None
sga['B42'] = '決算書の販管費＋支払手数料'; sga['B43'] = 'チェック'
for c, a, b in zip(PLC, PL_CHECK['販管費'], P['支払手数料']):
    sga[f'{c}42'] = a + b; sga[f'{c}42'].number_format = '#,##0'
    sga[f'{c}43'] = f'=EXACT({c}40,{c}42)'
sga['H2'] = None; sga['J2'] = '(単位：千円)'; copy_style(pl['J2'], sga['J2'])
hide_cols(sga, 'K', 'L', 'M')
sga.column_dimensions['B'].width = 30
for a_, b_ in (('C3', 'D3'), ('E3', 'F3'), ('G3', 'H3'), ('I3', 'J3')):
    sga.merge_cells(f'{a_}:{b_}')
hide_rows(sga, 5, 39)
red(sga, 'B45', '※地代家賃は松戸営業所の店舗賃料1,200,000円と代表者からの駐車場賃借545,460円。支払手数料は決算書では一般業務費として別掲（66期から計上、内容は要確認）。')
red(sga, 'B46', '※第64〜66期は対象会社提供の既存会社説明資料、第67期は決算報告書の販売費及び一般管理費内訳から転記。')
SGA_DEP_ROW = SGA_ROW['減価償却費']

# ================================================================ VR→
mbd, mbc = wb['修正BS(借方)'], wb['修正BS(貸方) ']
mpl, msga, mcr = wb['修正PL'], wb['修正SGA'], wb['修正CR']

# --- 減価償却（建物は固定資産税評価額へ評価替え、構築物・機械は当期償却不足額を減額）
dep = wb['減価償却']
clear(dep, 4, 7, 2, 6)
BLD_EV = REAL_ESTATE_OWNED[1][5]
dep_rows = [('建物（越谷工場）', "='BS(借方)'!F22", f'={BLD_EV}-C4', '固定資産税評価額（91,222千円）へ評価替え。1991年取得の工場建物で会計上の償却を停止しており、償却不足が累積（当期分21,745千円）'),
            ('構築物', "='BS(借方)'!F23", -447749, '当期の償却不足額447,749円を減額（過年度の累積不足は台帳で要確認）'),
            ('機械装置', "='BS(借方)'!F24", -3092403, '当期の償却不足額3,092,403円を減額（過年度の累積不足は台帳で要確認）'),
            ('車両運搬具・工具器具備品', "='BS(借方)'!F25+'BS(借方)'!F26", 0, '償却不足なし（減価償却計算書）。簿価評価')]
for i, (nm, bv, adj, why) in enumerate(dep_rows):
    r = 4 + i
    dep[f'B{r}'] = nm; dep[f'C{r}'] = bv; dep[f'D{r}'] = adj; dep[f'E{r}'] = f'=SUM(C{r},D{r})'; dep[f'F{r}'] = why
    for c in 'CDE':
        dep[f'{c}{r}'].number_format = KIN
nf(dep, ['C8', 'D8', 'E8'])
memo(dep, 'B10', '※出所：第67期 減価償却計算書（185件）。建物の簿価は勘定科目内訳書では280,673,009円と記載があり、BS259,241,627円との差は要確認。')

land = wb['土地']
clear(land, 4, 7, 2, 7)
land['B4'] = '越谷工場／埼玉県越谷市大字平方999-4'; land['C4'] = REAL_ESTATE_OWNED[0][3]
land['D4'] = "='BS(借方)'!F27"; land['F4'] = f'=ROUND({REAL_ESTATE_OWNED[0][5]}/0.7,0)'; land['E4'] = '=F4-D4'
land['G4'] = '固定資産税評価額56,890千円÷0.7にて時価を推計（路線価・鑑定で要確認）'
land['C4'].number_format = AI
for c in 'DEF':
    land[f'{c}4'].number_format = KIN; land[f'{c}8'].number_format = KIN

ins = wb['保険']
clear(ins, 4, 7, 2, 6)
for i, (nm, v) in enumerate([('大同生命保険', 8461965), ('第一生命保険', 900529)]):
    r = 4 + i
    ins[f'B{r}'] = nm; ins[f'C{r}'] = v; ins[f'E{r}'] = f'=C{r}'; ins[f'D{r}'] = f'=E{r}-C{r}'
    ins[f'F{r}'] = '解約返戻金の資料未受領のため簿価評価（要確認）'
    for c in 'CDE':
        ins[f'{c}{r}'].number_format = KIN
nf(ins, ['C8', 'D8', 'E8'])

sec = wb['有価証券']
clear(sec, 4, 8, 2, 7)
sec['B4'] = '該当なし（上場有価証券・投資信託の保有なし）'; sec['D4'] = 0; sec['F4'] = 0; sec['E4'] = '=F4-D4'
sec['G4'] = '出資金820千円（越谷市工業団地組合800千円ほか）は換金性に乏しいが少額のため簿価評価'

rec = wb['売掛・棚卸']
clear(rec, 4, 7, 2, 6); clear(rec, 14, 17, 2, 6)
rec['B4'] = '売掛金（食品卸・氷卸等）'; rec['C4'] = "='BS(借方)'!F7"; rec['D4'] = 0; rec['E4'] = '=SUM(C4,D4)'
rec['F4'] = '年齢表未受領のため簿価評価。月末締め翌月末回収（法人事業概況説明書）。長期滞留の有無は要確認'
rec['B14'] = '棚卸資産（商品・原材料・仕掛品）'; rec['C14'] = "='BS(借方)'!F9+'BS(借方)'!F10+'BS(借方)'!F11"; rec['D14'] = 0; rec['E14'] = '=C14+D14'
rec['F14'] = '実地棚卸（2026年3月31日）に基づく計上。商品は加工氷・袋氷等、原材料は包装資材。滞留・陳腐化の有無は要確認のため簿価評価'
for a in ('C4', 'D4', 'E4', 'C14', 'D14', 'E14', 'C8', 'D8', 'E8', 'C18', 'D18', 'E18'):
    rec[a].number_format = KIN

upay = wb['未払人件費・賞与']
upay['B3'] = '2026年4月支払給与（3月分）'; upay['B4'] = 5720681; upay['C4'] = '月末締め'; upay['D4'] = 0; upay['E4'] = 0; upay['F4'] = 0
upay['G4'] = '月末締め翌月20日払い。3月分給与5,720,681円は未払金に計上済みのため追加計上なし'
upay['B8'] = '賞与（支給規程未確認）'; upay['B9'] = 0; upay['C9'] = '要確認'; upay['D9'] = '要確認'; upay['E9'] = 0; upay['F9'] = 0; upay['G9'] = 0
upay['H9'] = '賞与引当金の計上なし。7月・12月に支給（月別人件費から推察）。算定期間が未確認のため評価上も未計上（要確認）'
ret = wb['退職金']
clear(ret, 4, 10, 2, 10)
ret['C4'] = '全従業員'; ret['H4'] = 0; ret['I4'] = 0; ret['J4'] = '=H4-I4'
ret['C10'] = '合計'; ret['H10'] = '=SUM(H4:H9)'; ret['I10'] = '=SUM(I4:I9)'; ret['J10'] = '=H10-I10'
memo(ret, 'B12', '※退職給付引当金の計上なし。販管費に退職共済掛金120千円／年（中小企業退職金共済と推察）。67期に製造原価の退職給与6,178千円が発生。')
memo(ret, 'B13', '※退職金規程の有無・社内積立部分の要支給額が未確認のため債務は計上していない（要確認）。')

# --- 修正BS
mbd['D22'] = '=減価償却!D4'; mbd['D23'] = '=減価償却!D5'; mbd['D24'] = '=減価償却!D6'; mbd['D27'] = '=土地!E4'
mbd['D8'] = '=保険!D8'
for r in range(5, 62):
    for c in 'CDE':
        mbd[f'{c}{r}'].number_format = KIN
mbd['F22'] = '固定資産税評価額へ評価替え（「減価償却」シート）'
mbd['F23'] = '当期の償却不足額を減額'; mbd['F24'] = '当期の償却不足額を減額'
mbd['F27'] = '固定資産税評価額÷0.7で時価を推計（「土地」シート）'
mbd['F8'] = '解約返戻金未受領のため簿価評価（要確認）'
for a in ('F8', 'F22', 'F23', 'F24', 'F27'):
    mbd[a].font = Font(name='游ゴシック Medium', size=9)
hide_rows(mbd, 18, 19, 28, 29, 30, 31, 32, 33, 34, 37, 38, 39, 40, 41, 42, 43, 44, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59)
mbc['D27'] = '=税効果!E18'
for r in range(5, 56):
    for c in 'CDE':
        mbc[f'{c}{r}'].number_format = KIN
hide_rows(mbc, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 26, 37, 38, 40, 41, 42, 47)

tax = wb['税効果']
clear(tax, 4, 17, 2, 6)
for i, (nm, ok, ref, memo_) in enumerate([
        ('建物', '×', "='修正BS(借方)'!D22", '評価損。繰越欠損金があり課税所得の見込みが立たないため回収可能性を認めず認識しない'),
        ('構築物・機械装置', '×', "='修正BS(借方)'!D23+'修正BS(借方)'!D24", '同上'),
        ('土地', '×', "='修正BS(借方)'!D27", '事業用不動産で売却予定がなく、実現可能性が低いため除外'),
        ('保険積立金', '×', "='修正BS(借方)'!D8", '修正なし')]):
    r = 4 + i
    tax[f'B{r}'] = nm; tax[f'C{r}'] = ok; tax[f'D{r}'] = ref; tax[f'E{r}'] = 0; tax[f'F{r}'] = memo_
    for c in 'DE':
        tax[f'{c}{r}'].number_format = KIN
tax['D18'] = '=SUM(D4:D17)'; tax['E18'] = '=SUM(E4:E17)'
nf(tax, ['D18', 'E18'])
memo(tax, 'B20', '※評価差額は純額で評価損。67期に繰越欠損金を使用しており課税所得の見込みが立たないため、繰延税金資産は認識しない（保守的）。')
memo(tax, 'B21', '※回収可能性が認められる場合は、建物・構築物・機械の評価損（約171,560千円）×34%＝約58,331千円だけ時価純資産が上振れする余地がある。')

# --- 評価科目別の検討結果（評価書「評価科目別の検討結果」の貼付範囲。不動産を筆別に検討する頁の代わり）
vk = im_sheet(wb, '評価科目別', '評価科目別の検討結果（評価書）')
cw(vk, {'B': 2.1, 'C': 1.05, 'D': 1.05, 'E': 1.05, 'F': 5.4})
put(vk, 3, 6, '(単位：千円)', 'unit', border=None)
row(vk, 4, 2, ['科目', '帳簿価額', '評価差額', '評価額', '評価の考え方・調整事由'], 'hdr', height=20)
VK = [('現金及び預金', 6, '残高で評価。定期預金は内訳書との差異6,800千円（要確認）'),
      ('売掛金', 7, '年齢表未受領のため簿価評価（長期滞留の有無は要確認）'),
      ('保険積立金（諸掛金）', 8, '解約返戻金の資料未受領のため簿価評価（要確認）'),
      ('棚卸資産', None, '実地棚卸に基づく計上。滞留・陳腐化は要確認のため簿価評価'),
      ('その他流動資産', None, '前渡金・前払費用・従業員貸付金・預け金・未収入金。簿価評価'),
      ('建物', 22, '固定資産税評価額へ評価替え（会計上の償却停止による過大計上を修正）'),
      ('構築物', 23, '当期の償却不足額を減額（累積不足は台帳で要確認）'),
      ('機械装置', 24, '当期の償却不足額を減額（累積不足は台帳で要確認）'),
      ('車両運搬具・工具器具備品', None, '償却不足なし。簿価評価'),
      ('土地', 27, '固定資産税評価額÷0.7で時価を推計（路線価・鑑定で要確認）'),
      ('特許商標権・出資金', None, '少額のため簿価評価')]
for i, (lab, r_, why) in enumerate(VK):
    rr_ = 5 + i
    if lab == '棚卸資産':
        f_ = [f"='修正BS(借方)'!{c}9+'修正BS(借方)'!{c}10+'修正BS(借方)'!{c}11" for c in 'CDE']
    elif lab == 'その他流動資産':
        f_ = [f"=SUM('修正BS(借方)'!{c}12:{c}17)" for c in 'CDE']
    elif lab == '車両運搬具・工具器具備品':
        f_ = [f"='修正BS(借方)'!{c}25+'修正BS(借方)'!{c}26" for c in 'CDE']
    elif lab == '特許商標権・出資金':
        f_ = [f"='修正BS(借方)'!{c}35+'修正BS(借方)'!{c}46" for c in 'CDE']
    else:
        f_ = [f"='修正BS(借方)'!{c}{r_}" for c in 'CDE']
    hl = lab in ('建物', '構築物', '機械装置', '土地')
    row(vk, rr_, 2, [lab] + f_ + [why], ['lab'] + (['subn'] if hl else ['num']) * 3 + ['txt'], fmts=[None, KIN, KIN, KIN, None], height=20)
n_ = 5 + len(VK)
for j, (lab, why) in enumerate([('賞与引当金（負債）', '支給規程・算定期間が未確認のため未計上（要確認）'),
                                ('退職給付引当金（負債）', '退職金規程が未確認のため未計上（中退共加入。要確認）'),
                                ('繰延税金（税効果）', '評価差額は純額で評価損。課税所得の見込みが立たず回収可能性を認めないため認識しない')]):
    row(vk, n_ + j, 2, [lab, '-', 0, '-', why], ['lab', 'num', 'num', 'num', 'txt'], fmts=[None, KIN, KIN, KIN, None], height=20)
n2 = n_ + 3
row(vk, n2, 2, ['資産の評価差額 合計', f"='修正BS(借方)'!C61", f"='修正BS(借方)'!D61", f"='修正BS(借方)'!E61", ''],
    ['sub', 'subn', 'subn', 'subn', 'sub'], fmts=[None, KIN, KIN, KIN, None], height=20)
mark(vk, f'B3:F{n2}', '株式価値評価書「評価科目別の検討結果」', cell='A2')
vk[f'B{n2 + 2}'] = 'チェック：明細の合計＝修正BS資産合計'
vk[f'C{n2 + 2}'] = f'=SUM(C5:C{n_ - 1})=C{n2}'; vk[f'D{n2 + 2}'] = f'=SUM(D5:D{n_ - 1})=D{n2}'

nc = wb['ネットキャッシュ']
for r in range(4, 13):
    nc[f'C{r}'] = None; nc[f'F{r}'] = None
nc['B4'] = '現金及び預金'; nc['C4'] = "='修正BS(借方)'!E6"
nc['B5'] = '役員貸付金'; nc['C5'] = 0
nc['B6'] = '従業員貸付金'; nc['C6'] = "='修正BS(借方)'!E14"
nc['B7'] = '立替金'; nc['C7'] = 0
nc['B8'] = '保険積立金'; nc['C8'] = "='修正BS(借方)'!E8"
nc['B9'] = '投資有価証券'; nc['C9'] = 0
nc['E4'] = '短期借入金'; nc['F4'] = "='修正BS(貸方) '!E8"
nc['E5'] = '長期借入金（代表者借入を除く）'; nc['F5'] = "='修正BS(貸方) '!E19-'BS (貸方)'!F57"
nc['E6'] = '社債'; nc['F6'] = 0
nc['E7'] = '役員借入金（代表者）'; nc['F7'] = "='BS (貸方)'!F57"
nc['E8'] = '未払法人税'; nc['F8'] = "='修正BS(貸方) '!E12"
nc['F15'] = '=F13-C13'
nc['C17'] = None; nc['F17'] = None
for r in range(4, 16):
    for c in 'CF':
        nc[f'{c}{r}'].number_format = KIN
memo(nc, 'B18', '※代表者借入金14,000千円は譲渡時に返済を希望するのが通例のためデットライクに含めた（主ケース）。債務免除・相殺で整理する場合はネットデットが同額減少（参考ケース）。')
memo(nc, 'B19', '※日本政策金融公庫の資本性ローン105,420千円は期限一括返済の債務のため有利子負債として扱う。資本とみなす場合はネットデットが同額減少。')
nc['B21'] = '（参考）代表者借入金を除いたネットデット'; nc['F21'] = '=F15-F7'
nc['B22'] = '（参考）資本性ローンも除いたネットデット'; nc['F22'] = f"=F21-'BS (貸方)'!F58"
nf(nc, ['F21', 'F22'])

# --- 正常化調整の前提（修正PL下段）：PLハイライト【調整項目】・修正SGA・修正CRの修正額の原本
r0 = 62
mpl.cell(r0, 2, '■ 正常化調整の前提（修正SGA・修正CR・修正PLの修正額、PLハイライト【調整項目】はここを参照）').font = Font(name=JP, size=11, bold=True, color=NAVY)
mpl.cell(r0 + 1, 2, '項目')
for c, f in zip('CFILO', ["='BS(借方)'!C4", "='BS(借方)'!D4", "='BS(借方)'!E4", "='BS(借方)'!F4", "='BS(借方)'!G4"]):
    mpl[f'{c}{r0 + 1}'] = f
mpl.cell(r0 + 1, 18, '調整事由')
ADJ = [
    ('退職給与（製造原価）', ['=CR!C17', '=CR!E17', '=CR!G17', '=CR!I17', None],
     '過去に発生のない臨時的な退職金支出と推察し加算（退職金規程・対象者は要確認）'),
    ('退職給与（販管費）', ['=SGA!C9', '=SGA!E9', '=SGA!G9', '=SGA!I9', None], '同上'),
    ('役員報酬（新体制の標準報酬へ置換）', [0, 0, 0, 0, None],
     '現行の代表者報酬7,200千円／年は経営者報酬として過大ではなく、後任経営者の標準報酬と同水準と想定し調整なし'),
]
for i, (lab, vals, why) in enumerate(ADJ):
    r = r0 + 2 + i
    mpl.cell(r, 2, lab)
    for c, v in zip('CFILO', vals):
        if v is not None:
            mpl[f'{c}{r}'] = v
            mpl[f'{c}{r}'].number_format = KIN
    mpl.cell(r, 18, why)
ADJ_R = {'ret_cr': r0 + 2, 'ret_sga': r0 + 3, 'officer': r0 + 4}
# 修正CR：CRの帳簿価額を参照し、修正額は前提ブロックを参照
MCR_MAP = {5: 8, 6: 9, 8: 11, 12: 15, 13: 16, 14: 17, 15: 18, 16: 19, 17: 20, 19: 24}
for i in range(10):
    MCR_MAP[26 + i] = 34 + i
MCR_MAP[38] = 48; MCR_MAP[40] = 50
for mr, crr in MCR_MAP.items():
    mcr[f'B{mr}'] = f'=CR!B{crr}'
    for bc, src in zip('CFILO', 'CEGIK'):
        mcr[f'{bc}{mr}'] = f'=CR!{src}{crr}'
for bc, ac in zip('CFIL', 'DGJM'):
    mcr[f'{ac}14'] = f'=-修正PL!{bc}{ADJ_R["ret_cr"]}'
for r in range(5, 42):
    for c in 'CDEFGHIJKLMNOPQ':
        mcr[f'{c}{r}'].number_format = KIN
mcr['B7'] = None
mcr['B44'] = 'チェック：修正CRの帳簿価額＝CRの当期製造原価'
for bc, src in zip('CFIL', 'CEGI'):
    mcr[f'{bc}44'] = f'=EXACT({bc}41,CR!{src}51)'
# 修正PL：製造原価の修正額を修正CRに連動
for ac in 'DGJM':
    mpl[f'{ac}15'] = f'=修正CR!{ac}41'
# 修正SGA：退職給与（SGA 9行 → 修正SGA 8行）
# テンプレートの不具合修正：修正SGAの35〜38行がSGAの36〜39行に連動していなかった（科目数が多いと販管費が過少になる）
for mr in range(35, 39):
    msga[f'B{mr}'] = f'=SGA!B{mr + 1}'
    for bc, src in zip('CFILO', 'CEGIK'):
        msga[f'{bc}{mr}'] = f'=SGA!{src}{mr + 1}'
msga['B40'] = 'チェック：修正SGAの帳簿価額＝SGA合計'
for bc, src in zip('CFIL', 'CEGI'):
    msga[f'{bc}40'] = f'=EXACT({bc}39,SGA!{src}40)'
for bc, ac in zip('CFIL', 'DGJM'):
    msga[f'{ac}8'] = f'=-修正PL!{bc}{ADJ_R["ret_sga"]}'
    msga[f'{ac}5'] = f'=-修正PL!{bc}{ADJ_R["officer"]}'
for r in range(5, 40):
    for c in 'CDEFGHIJKLMNOPQ':
        msga[f'{c}{r}'].number_format = KIN
for r in range(5, 57):
    for c in 'CDEFGHIJKLMNOPQ':
        if mpl[f'{c}{r}'].value is not None:
            mpl[f'{c}{r}'].number_format = KIN
for c, v in zip('CFIL', PL_CHECK['当期純利益']):
    mpl[f'{c}58'] = v
mpl['O58'] = None; mpl['O59'] = None
mpl['B58'] = '決算書の当期純利益'; mpl['B59'] = 'チェック'
# 修正後EBITDA（概要書・評価書の共通原本）
mpl['B70'] = '修正後営業利益'; mpl['B71'] = '減価償却費（製造原価＋販管費）'; mpl['B72'] = '修正後EBITDA'
for bc, ec, src in zip('CFIL', 'EHKN', 'CEGI'):
    mpl[f'{bc}70'] = f'={ec}19'
    mpl[f'{bc}71'] = f'=SGA!{src}{SGA_DEP_ROW}+CR!{src}{CR_DEP_ROW}'
    mpl[f'{bc}72'] = f'={bc}70+{bc}71'
    for r in (70, 71, 72):
        mpl[f'{bc}{r}'].number_format = KIN
mpl['B74'] = '※支払利息・雑収入（保険返戻金を含む）・雑損失は営業外のため営業利益ベースの正常収益力に影響しない。特別損益は4期ともなし。'
mpl['B75'] = '※建物の会計上の償却停止は営業利益を押し上げているが、EBITDAベースでは中立。マーケットアプローチ（EBITDA）を併用して補う。支払手数料（66・67期）は内容未確認のため調整せず（要確認）。'
mpl['B76'] = '※代表者からの駐車場賃借（年545千円）は譲渡後も継続利用を前提に調整しない（条件は要協議）。'
for a in ('B74', 'B75', 'B76'):
    mpl[a].font = Font(name='游ゴシック Medium', size=9)

for ws_ in (mpl, msga, mcr):
    hide_cols(ws_, 'C', 'D', 'E', 'O', 'P', 'Q')
hide_rows(mpl, 8, 9, 10, 24, 25, 26, 27, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 49, 50, 54, 55)
hide_rows(msga, 38)
hide_rows(mcr, 7, 9, 10, 20, 21, 22, 23, 24)
mpl.column_dimensions['B'].width = 30; msga.column_dimensions['B'].width = 30
mbd.column_dimensions['F'].width = 46
hide_rows(mbc, 29, 30, 31, 32, 33)
hide_rows(nc, 9, 10, 11, 12)
for a_ in ('B10', 'B11', 'E9', 'E10', 'E11', 'E12'):
    nc[a_] = None
mpl['K2'] = None; mpl['N2'] = '(単位：千円)'; copy_style(sga['J2'], mpl['N2'])
msga['K2'] = None; msga['N2'] = '(単位：千円)'; copy_style(sga['J2'], msga['N2'])
for ws_ in (mpl, msga):
    for a_, b_ in (('F3', 'H3'), ('I3', 'K3'), ('L3', 'N3')):
        ws_.merge_cells(f'{a_}:{b_}')
mpl['A1'] = None; msga['A1'] = None
paste(mpl, 'B3:N56', '株式価値評価書「修正PLの算定」', REG, cell='A1')
paste(msga, 'B3:N39', '株式価値評価書「販売費及び一般管理費の修正」', REG, cell='A1')
paste(mbd, 'B3:E61', '株式価値評価書「修正BS［資産］」', REG, cell='A1')
paste(mbc, 'B3:E53', '株式価値評価書「修正BS［負債・純資産］」', REG, cell='A1')
paste(nc, 'B3:F15', '株式価値評価書「ネットキャッシュの算定」', REG, cell='A1')
for ws_ in (mpl, msga, mbd, mbc, nc):
    ws_['A1'].font = Font(name='游ゴシック Medium', size=9, color='C00000')

nen = wb['年買法']
nen['B3'] = '26年3月期'
nen['B12'] = '※営業権倍率1〜3倍は国内中小企業M&Aのボリュームゾーン。加重は直近期0.5・前期0.3・前々期0.2。税効果は「税効果」シートで判定（本件は認識しない）。'
nen['B13'] = '（感応度）建物を簿価のまま評価した場合の時価純資産'; nen['K13'] = '=K4-減価償却!D4'
nen['B14'] = '（感応度）同 想定株式価値'
for c in 'KLM':
    nen[f'{c}14'] = f'=$K$13+{c}5'
for r in range(4, 15):
    for c in 'CFGHKLM':
        nen[f'{c}{r}'].number_format = KIN
ev = wb['EV EBITDAマルチプル']
ev['B3'] = '26年3月期'
for c, src in zip('FGH', 'EGI'):
    ev[f'{c}7'] = f'=SGA!{src}{SGA_DEP_ROW}+CR!{src}{CR_DEP_ROW}'
ev['K9'] = 0
ev['B12'] = '※倍率3〜5倍は国内中小企業M&Aの実務水準。非流動性ディスカウントは標準形（不適用、K9＝0%）。算定値がマイナスの場合、ネットデットが事業価値を上回ることを示す。'
ev['B13'] = '（参考）代表者借入金を債務免除・相殺で整理した場合の想定株式価値'
ev['B14'] = '（参考）さらに資本性ローンを資本とみなした場合の想定株式価値'
for c in 'KLM':
    ev[f'{c}13'] = f'={c}7+ネットキャッシュ!F7'
    ev[f'{c}14'] = f"={c}13+'BS (貸方)'!F58"
for r in range(4, 15):
    for c in 'CFGHKLM':
        ev[f'{c}{r}'].number_format = KIN
ref = wb['参照企業']
clear(ref, 4, 13, 2, 5)
ref['B4'] = '（未使用）製氷専業の上場会社は存在しない。倍率は国内中小企業M&Aの実務水準（3〜5倍）を採用。低温物流・食品関連の上場会社の倍率は算定日時点の株価で更新が必要（要確認）'
ref['E14'] = None; ref['E15'] = None

# ================================================================ IM→
# --- 会社概要
co = wb['会社概要']
CO = {
    3: '越ヶ谷製氷冷蔵株式会社（純氷ブランド「キングアイス」）',
    4: '埼玉県越谷市蒲生寿町19番15号（登記上の本店）',
    5: '① 越谷工場：埼玉県越谷市大字平方999-4（平方工業団地内）\n② 松戸営業所：千葉県松戸市稔台5-14-3',
    6: '1960年（昭和35年）3月18日（2026年3月期で第67期）',
    7: '=TEXT(\'BS (貸方)\'!F36/1000,"#,##0")&"千円（2026年5月15日に10,000千円へ減資）"',
    8: '小野田 洋',
    9: '小野田 洋（100%）',
    10: '・純氷（角氷・加工氷・袋氷・特殊氷）の製造及び卸小売販売\n・ドライアイス等の商品販売、冷蔵保管（寄託物品の保管）',
    11: '25名（常勤役員1名・従業員24名）\n※2026年3月期末（法人事業概況説明書）',
    12: '営業許可（氷雪製造業）、食品営業許可（氷雪販売業）',
    13: '=TEXT(PL!I11/1000,"#,##0")&"千円（2026年3月期）"',
    14: '=TEXT(AVERAGE(PLハイライト!D25:F25)/1000,"#,##0")&"千円（直近3期平均）"',
    15: '後継者不在',
}
co['B4'] = '本店所在地'; co['B5'] = '拠点'; co['B6'] = '設立'
for r, v in CO.items():
    co[f'C{r}'] = v
    co[f'C{r}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    n = sum(max(1, -(-len(x) // 30)) for x in str(v).split('\n')) if not str(v).startswith('=') else 1
    co.row_dimensions[r].height = max(22.5, 19 * n + 4)
mark(co, 'B2:C15', '基本情報')
co['B18'] = '※発行済株式数は登記22,000株／株主資本等変動計算書注記200株／法人税申告書別表二220株と相違（要確認）。'

# --- PLハイライト（調整項目は修正PLの前提ブロックを参照）
ph = wb['PLハイライト']
tmpl = {k: [copy.copy(ph.cell(k, c)._style) for c in range(2, 8)] for k in (3, 4, 5, 14, 15, 17, 18, 37)}
hts = {k: ph.row_dimensions[k].height for k in tmpl}
clear(ph, 14, 48, 2, 9, style=True)


def put_style(r, k):
    for j, c in enumerate(range(2, 8)):
        ph.cell(r, c)._style = copy.copy(tmpl[k][j])
    ph.row_dimensions[r].height = hts[k]


for c, f in zip('CDEFG', ["='BS(借方)'!C4", "='BS(借方)'!D4", "='BS(借方)'!E4", "='BS(借方)'!F4", "='BS(借方)'!G4"]):
    ph[f'{c}3'] = f
ph.row_dimensions[3].height = 33
ph['C5'] = '-'
for r in (4, 5, 6, 7, 8, 9, 10, 11, 12, 13):   # 進行期（G列・非表示）は試算表未受領のため空欄
    ph[f'G{r}'] = None
ph['D5'] = '=IFERROR(D4/C4,"-")'
put_style(14, 14); put_style(15, 15); put_style(16, 14); put_style(17, 17)
for r in (18, 19):
    put_style(r, 18)
put_style(20, 14)
put_style(21, 37)
ph['B15'] = '減価償却費'
for c, src in zip('CDEF', 'CEGI'):
    ph[f'{c}15'] = f'=SGA!{src}{SGA_DEP_ROW}+CR!{src}{CR_DEP_ROW}'
ph['B17'] = '調整項目 合計'
ph['B18'] = '退職給与（臨時）'
ph['B19'] = '役員報酬'
for x, c in zip('CDEF', 'CFIL'):
    ph[f'{x}18'] = f'=修正PL!{c}{ADJ_R["ret_cr"]}+修正PL!{c}{ADJ_R["ret_sga"]}'
    ph[f'{x}19'] = f'=修正PL!{c}{ADJ_R["officer"]}'
    ph[f'{x}17'] = f'=SUM({x}18:{x}19)'
    ph[f'{x}21'] = f'={x}8+{x}15+{x}17'
ph['B21'] = '調整後EBITDA'
for r in range(4, 22):
    for x in 'CDEF':
        if ph[f'{x}{r}'].number_format not in ('0.0%',) and r not in (5, 7, 9, 11, 13):
            ph[f'{x}{r}'].number_format = KIN
for r in (5, 7, 9, 11, 13):
    ph.row_dimensions[r].height = 14
    for x in 'CDEF':
        ph[f'{x}{r}'].number_format = PI
for r in (14, 16, 20):
    ph.row_dimensions[r].height = 8
for r in (4, 6, 8, 10, 12, 15, 17, 18, 19, 21):
    ph.row_dimensions[r].height = 20
hide_cols(ph, 'G')
ph.column_dimensions['B'].width = 31
mark(ph, 'B2:F21', '財務ハイライト（P/L）')
ph['B24'] = '参考（リード文・Comment用。貼付範囲外）'; ph['B24'].font = Font(name=JP, size=11, bold=True, color=NAVY)
ph['B25'] = '調整後EBITDA（行21の再掲）'
for x in 'CDEF':
    ph[f'{x}25'] = f'={x}21'
ph['B26'] = '調整後EBITDA 直近3期平均'; ph['F26'] = '=AVERAGE(D21:F21)'
ph['B27'] = '調整後EBITDA 直近2期平均'; ph['F27'] = '=AVERAGE(E21:F21)'
ph['B28'] = 'チェック：修正PLの修正後EBITDA（26/3期）と一致'; ph['F28'] = '=ROUND(F21,0)=ROUND(修正PL!L72,0)'
ph['B29'] = 'EBITDA（営業利益＋減価償却費）'
for x in 'CDEF':
    ph[f'{x}29'] = f'={x}8+{x}15'
ph['B30'] = '動力費（製造＋販管）'
ph['B31'] = '運賃'
ph['B32'] = '労務費＋販管人件費'
for x, src in zip('CDEF', 'CEGI'):
    ph[f'{x}30'] = f"=CR!{src}39+SGA!{src}{SGA_ROW['動力費']}"
    ph[f'{x}31'] = f"=SGA!{src}{SGA_ROW['運賃']}"
    ph[f'{x}32'] = f"=CR!{src}23+SUM(SGA!{src}6:{src}13)"
ph['B33'] = '退職給与（製造＋販管）'
for x, src in zip('CDEF', 'CEGI'):
    ph[f'{x}33'] = f'=CR!{src}17+SGA!{src}9'
for r in range(25, 34):
    for x in 'CDEF':
        ph[f'{x}{r}'].number_format = '#,##0;△#,##0;"-"'
ph['F28'].number_format = 'General'
memo(ph, 'B35', '※調整方針は評価書（修正PL）と同一。調整項目は修正PL下段「正常化調整の前提」を参照。23/3期は評価の参照期間外だが同じ方針で算定（該当なし）。')

# --- BSハイライト（直近期末の左右2列。手入力しない）
bh = wb['BSハイライト']
F = lambda r: f"='BS(借方)'!F{r}"
FC = lambda r: f"='BS (貸方)'!F{r}"
LEFT = {5: ('現預金', F(6)), 6: ('保険積立金', F(8)), 7: ('売掛金', F(7)), 8: ('棚卸資産', "='BS(借方)'!F9+'BS(借方)'!F10+'BS(借方)'!F11"),
        9: ('未収入金', F(16)), 10: ('その他', "='BS(借方)'!F12+'BS(借方)'!F13+'BS(借方)'!F14+'BS(借方)'!F15+'BS(借方)'!F17"),
        12: ('建物', F(22)), 13: ('構築物', F(23)), 14: ('機械装置', F(24)), 15: ('車両運搬具', F(25)), 16: ('工具器具備品', F(26)),
        17: ('土地', F(27)), 18: ('特許商標権', F(36)), 19: ('出資金', F(47)), 20: ('その他', 0)}
RIGHT = {5: ('支払手形', FC(6)), 6: ('買掛金', FC(7)), 7: ('短期借入金', FC(8)), 8: ('未払金', FC(9)), 9: ('未払法人税等', FC(12)),
         10: ('その他', "='BS (貸方)'!F10+'BS (貸方)'!F11"), 12: ('長期借入金', FC(19)), 13: (None, None), 14: (None, None),
         16: (None, None), 17: (None, None), 18: ('資本金', FC(36)), 19: ('利益剰余金', "='BS (貸方)'!F43")}
for r, (lab, f) in LEFT.items():
    bh[f'B{r}'] = lab; bh[f'C{r}'] = f
for r, (lab, f) in RIGHT.items():
    bh[f'E{r}'] = lab; bh[f'F{r}'] = f
copy_style(bh['E9'], bh['E10'])
bh['F20'] = '=SUM(F18:F19)'
for r in range(4, 22):
    for c in 'CF':
        bh[f'{c}{r}'].number_format = KIN
bh['B3'] = "=PLハイライト!F3&\"末\""
mark(bh, 'B2:F21', '財務ハイライト（B/S）')
bh['B24'] = '参考：B/S推移・借入先別残高（リード文・Comment用。貼付範囲外）'; bh['B24'].font = Font(name=JP, size=11, bold=True, color=NAVY)
for c, src in zip('CDEF', 'CDEF'):
    bh[f'{c}25'] = f"='BS(借方)'!{src}4"
    bh[f'{c}26'] = f"='BS(借方)'!{src}6"
    bh[f'{c}27'] = f"='BS (貸方)'!{src}56"
    bh[f'{c}28'] = f"='BS (貸方)'!{src}50"
    bh[f'{c}29'] = f"='BS(借方)'!{src}61"
    bh[f'{c}30'] = f'={c}28/{c}29'
    bh[f'{c}31'] = f'={c}27-{c}26'
    for r in (26, 27, 28, 29, 31):
        bh[f'{c}{r}'].number_format = '#,##0,;△#,##0,'
    bh[f'{c}30'].number_format = '0.0%'
for r, lab in zip(range(25, 32), ['項目', '現金及び預金', '借入金（短期・長期）', '純資産', '資産合計', '自己資本比率', 'ネットデット（借入金−現預金）']):
    bh[f'B{r}'] = lab
bh['B33'] = '借入先別残高（2026/3期末）'; bh['C33'] = '期末残高'; bh['D33'] = '支払利息'; bh['E33'] = '担保・備考'
for i, (nm, bal, intr, col_) in enumerate(LOANS_67):
    r = 34 + i
    bh[f'B{r}'] = nm; bh[f'C{r}'] = bal; bh[f'D{r}'] = intr; bh[f'E{r}'] = col_
    bh[f'C{r}'].number_format = '#,##0'; bh[f'D{r}'].number_format = '#,##0'
rr = 34 + len(LOANS_67)
bh[f'B{rr}'] = '合計'; bh[f'C{rr}'] = f'=SUM(C34:C{rr - 1})'; bh[f'D{rr}'] = f'=SUM(D34:D{rr - 1})'
bh[f'C{rr}'].number_format = '#,##0'; bh[f'D{rr}'].number_format = '#,##0'
bh[f'B{rr + 1}'] = 'チェック：BS有利子負債／PL支払利息と一致'
bh[f'C{rr + 1}'] = f"=C{rr}='BS (貸方)'!F56"; bh[f'D{rr + 1}'] = f'=D{rr}=PL!I29'
LOAN_TOTAL_ROW = rr
lws = im_sheet(wb, '借入金', '借入金の状況')
cw(lws, {'B': 3.2, 'C': 1.4, 'D': 1.4, 'E': 5.0})
put(lws, 3, 5, '(単位：千円)', 'unit', border=None)
row(lws, 4, 2, ['借入先', '期末残高', '支払利息（年）', '担保・備考'], 'hdr', height=19)
for i in range(len(LOANS_67)):
    src_r = 34 + i
    row(lws, 5 + i, 2, [f'=BSハイライト!B{src_r}', f'=BSハイライト!C{src_r}', f'=BSハイライト!D{src_r}', f'=BSハイライト!E{src_r}'],
        ['lab', 'num', 'num', 'txt'], fmts=[None, KIN, KIN, None], height=19)
row(lws, 5 + len(LOANS_67), 2, ['合計', f'=BSハイライト!C{rr}', f'=BSハイライト!D{rr}', ''], ['sub', 'subn', 'subn', 'sub'], fmts=[None, KIN, KIN, None], height=19)
LOAN_RANGE = mark(lws, f'B3:E{5 + len(LOANS_67)}', '許認可・取引金融機関（借入金の状況）', cell='A2')


# --- 取引先（期末売掛金残高：年間取引高は未受領のため代替）
tr = wb['取引先']
for c, src in zip('GIK', 'DEF'):
    tr[f'{c}3'] = f"='BS(借方)'!{src}4"
for c in 'GIK':
    tr[f'{c}4'] = '売掛金残高'
CUS = [('(株)日本橋冷凍手島商店', '加工氷等', '東京都中央区', '要確認', '資材の仕入先でもある'),
       ('(株)日本アクセス', '加工氷', '東京都品川区', '食品卸売業', None),
       ('三菱食品(株)', '加工氷', '東京都大田区', '食品卸売業', None),
       ('金木屋商事(株)', '加工氷', '埼玉県川口市', '要確認', None),
       ('(株)小野田商店', '加工氷', '東京都北区', '要確認', '代表者との関係は要確認'),
       ('(株)クールテックサガワ', '加工氷', '福島県本宮市', '要確認', None),
       ('(株)ニチレイ・アイス', '加工氷', '東京都中央区', '氷の製造販売', None)]
ARD = dict(AR_BY_CUSTOMER)
for i, (nm, kind, addr, ind, memo_) in enumerate(CUS):
    r = 5 + i
    tr[f'B{r}'] = i + 1; tr[f'C{r}'] = nm; tr[f'D{r}'] = kind; tr[f'E{r}'] = addr; tr[f'F{r}'] = ind; tr[f'M{r}'] = memo_
    for c, v in zip('GIK', ARD[nm][1:]):
        tr[f'{c}{r}'] = v
for r in (12, 13, 14):
    for c in 'BCDEFGIKM':
        tr[f'{c}{r}'] = None
tr['C15'] = 'その他（複数社）'
for c in 'GIK':
    tr[f'{c}15'] = f'={c}16-SUM({c}5:{c}14)'
tr['C16'] = '合計（期末売掛金）'
tr['G16'] = "='BS(借方)'!D7"; tr['I16'] = "='BS(借方)'!E7"; tr['K16'] = "='BS(借方)'!F7"
for c in 'HJL':
    tr[f'{c}16'] = '100.0%'; tr[f'{c}16'].alignment = Alignment(horizontal='right', vertical='center')
for r in range(5, 17):
    for c in 'GIK':
        tr[f'{c}{r}'].number_format = KIN
        tr[f'{c}{r}'].alignment = Alignment(horizontal='right', vertical='center')
hide_rows(tr, 12, 13, 14)
ungroup_cols(tr)
for c, w_ in zip('BCDEFGHIJKLM', [4.4, 25, 12, 14, 13, 10.5, 7.5, 10.5, 7.5, 10.5, 7.5, 22]):
    tr.column_dimensions[c].width = w_
for r in range(3, 17):
    for c in 'BCDEFGHIJKLM':
        al = copy.copy(tr[f'{c}{r}'].alignment); al.vertical = 'center'; al.wrap_text = True
        tr[f'{c}{r}'].alignment = al
mark(tr, 'B2:M16', '主要販売先の3期推移（期末売掛金残高）')
memo(tr, 'B18', '※取引先別の年間売上高は未受領のため、期末売掛金残高（月末締め翌月末回収）で代替。構成比は期末売掛金合計に対する比率。')
memo(tr, 'B19', '※26/3期の内訳書は上位先の合計が残高を1,723,680円上回るため、その他は残高との差額で表示（(株)小野田商店の計上区分は要確認）。')
tr['B21'] = '参考：上位7社の構成比（26/3期）'; tr['K21'] = '=SUM(K5:K11)/K16'; tr['K21'].number_format = '0.0%'

# --- 仕入先（期末買掛金残高：26/3期のみ）
sp = wb['仕入先']
for c, src in zip('FHJ', 'DEF'):
    sp[f'{c}3'] = f"='BS(借方)'!{src}4"
for c in 'FHJ':
    sp[f'{c}4'] = '買掛金残高'
for i, (nm, kind, v) in enumerate(AP_BY_SUPPLIER_67):
    r = 5 + i
    addr = {'鈴木段ボール工業(株)': '岩手県一関市', '大三紙業(株)': '愛知県豊橋市', '日本液炭(株)': '東京都北区',
            '(株)日本橋冷凍手島商店': '東京都中央区', '越谷松伏水道企業団': '埼玉県越谷市', '(株)シモジマ': '東京都台東区',
            '(株)一陽パック': '千葉県千葉市'}[nm]
    kind = {'資材（段ボール）': '段ボール', '資材': '包装資材', '商品（ドライアイス等）': 'ドライアイス', '資材（包装資材）': '包装資材'}.get(kind, kind)
    sp[f'B{r}'] = i + 1; sp[f'C{r}'] = nm; sp[f'D{r}'] = kind; sp[f'E{r}'] = addr; sp[f'J{r}'] = v
    sp[f'L{r}'] = '販売先でもある' if '手島' in nm else None
for r in (12, 13, 14):
    for c in 'BCDEFHJL':
        sp[f'{c}{r}'] = None
sp['C15'] = 'その他'; sp['J15'] = '=J16-SUM(J5:J14)'
sp['C16'] = '合計（期末買掛金）'; sp['J16'] = "='BS (貸方)'!F7"
sp['K16'] = '100.0%'
for r in range(5, 17):
    sp[f'J{r}'].number_format = KIN
    sp[f'J{r}'].alignment = Alignment(horizontal='right', vertical='center')
hide_rows(sp, 12, 13, 14)
hide_cols(sp, 'F', 'G', 'H', 'I')
for c, w_ in zip('BCDEJKL', [4.4, 27, 15, 15, 12, 9, 20]):
    sp.column_dimensions[c].width = w_
for r in range(3, 17):
    for c in 'BCDEJKL':
        al = copy.copy(sp[f'{c}{r}'].alignment); al.vertical = 'center'; al.wrap_text = True
        sp[f'{c}{r}'].alignment = al
mark(sp, 'B2:L16', '主要仕入先・主要経費')
memo(sp, 'B18', '※仕入先別の年間仕入高は未受領。第67期の勘定科目内訳書（買掛金）の期末残高で代替。原材料（包装資材）の年間仕入高は77,113千円。')
COST = [('原材料仕入高（包装資材等）', "=CR!{s}9"), ('労務費＋販管人件費', "=PLハイライト!{x}32"),
        ('電力料（動力費・製造＋販管）', "=PLハイライト!{x}30"), ('配送運賃', "=PLハイライト!{x}31"),
        ('外注作業費', "=CR!{s}24")]
# 主要費用はレイアウトの都合で N 列以降（仕入先表と列を共有しない）に置く
ws_cost = im_sheet(wb, '主要費用', '主要費用の推移')
cw(ws_cost, {'B': 2.6, 'C': 1.0, 'D': 1.0, 'E': 1.0, 'F': 1.0})
put(ws_cost, 3, 6, '(単位：千円)', 'unit', border=None)
row(ws_cost, 4, 2, ['費目', "='BS(借方)'!C4", "='BS(借方)'!D4", "='BS(借方)'!E4", "='BS(借方)'!F4"], 'hdr', height=19)
for i, (lab, f) in enumerate(COST[:5]):
    vals = [lab] + [f.format(s=s_, x=x_) for s_, x_ in zip('CEGI', 'CDEF')]
    row(ws_cost, 5 + i, 2, vals, ['lab'] + ['num'] * 4, fmts=[None] + [KIN] * 4, height=19)
mark(ws_cost, 'B3:F9', '主要仕入先・主要経費（主要費用の推移）', cell='A2')

# --- 外注先
ex_ = wb['外注先']
clear(ex_, 5, 16, 3, 13)
ex_['C5'] = '外注先別内訳は未受領（要確認）'
ex_['K16'] = '=CR!I24'; ex_['C16'] = '合計（外注作業費）'; ex_['K16'].number_format = KIN
ex_['M5'] = '配送は自社便と委託（倉橋運送・鴻池運輸等）を併用。繁忙期は人材派遣（(株)ヒューマンアイ等）で補完'

# --- 株主・取締役
sh = wb['株主']
sh['E3'] = "=\"役員報酬\"&CHAR(10)&\"（\"&'BS(借方)'!F4&\"）\""
sh['C5'] = '小野田 洋'; sh['D5'] = '代表取締役'; sh['E5'] = "=SGA!I6"; sh['F5'] = 22000; sh['H5'] = '株式数は登記記載（要確認）'
sh['E5'].number_format = KIN
for r in range(6, 15):
    for c in 'CDEFH':
        sh[f'{c}{r}'] = None
sh['E15'].number_format = KIN
ungroup_cols(sh); sh.column_dimensions['H'].width = 26
hide_rows(sh, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14)
sh.row_dimensions[3].height = 33
mark(sh, 'B2:H15', '株主及び役員（株主）')
memo(sh, 'B17', '※発行済株式数は登記上22,000株（2012年4月30日変更）。決算書の株主資本等変動計算書注記は200株、法人税申告書別表二は220株と記載が異なる（株主名簿・定款で要確認）。')
dr = wb['取締役']
dr['F3'] = "=\"役員報酬\"&CHAR(10)&\"（\"&'BS(借方)'!F4&\"）\""
for c, v in zip('CDEFGH', ['小野田 洋', '代表取締役（常勤）', '本人', '=SGA!I6', '経営全般・品質検査', '引継ぎ完了後に退任予定']):
    dr[f'{c}5'] = v
dr['F5'].number_format = KIN; dr['F15'].number_format = KIN
for r in range(6, 15):
    for c in 'CDEFGH':
        dr[f'{c}{r}'] = None
ungroup_cols(dr)
for c, w_ in zip('BCDEFGH', [4.4, 16, 22, 13, 13, 24, 26]):
    dr.column_dimensions[c].width = w_
hide_rows(dr, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14)
dr.row_dimensions[3].height = 33
mark(dr, 'B2:H15', '株主及び役員（役員）')
dr['B17'] = 'チェック：SGA役員報酬（26/3期）と一致'; dr['F17'] = '=F15=SGA!I6'

# --- 従業員（名簿未受領）
em = wb['従業員']
em['B36'] = '※従業員名簿（年齢・勤続・資格・雇用形態）は未受領（要確認）。法人事業概況説明書（第67期）による人員：常勤役員1名・従業員24名（越谷工場19名・松戸営業所5名）、月別従事員数23〜26名。'
em['B36'].font = Font(name='游ゴシック Medium', size=9)

# --- 拠点（本社等）
bp = wb['拠点(本社等)']
clear(bp, 2, 6, 2, 8)
cw(bp, {'B': 1.3, 'C': 3.0, 'D': 1.5, 'E': 0.9, 'F': 0.7, 'G': 3.2})
row(bp, 2, 2, ['拠点', '所在地', '所有区分', '延床面積(㎡)', '従業員', '用途'], 'hdr', height=19)
BP = [('越谷工場', '埼玉県越谷市大字平方999-4（平方工業団地内）', '所有（土地・建物）', 2065, '19名', '製氷・加工・冷蔵保管・配送の中核拠点'),
      ('松戸営業所', '千葉県松戸市稔台5-14-3', '賃借（店舗）', 400, '5名', '角氷・加工氷・凍氷雑品の販売'),
      ('本店（登記）', '埼玉県越谷市蒲生寿町19-15', '代表者住所', 10, '1名', '登記上の本店（管理）。隣接の駐車場を代表者から賃借')]
for i, rec_ in enumerate(BP):
    row(bp, 3 + i, 2, list(rec_), ['lab', 'txt', 'txt', 'num', 'txtc', 'txt'], fmts=[None, None, None, YI, None, None], height=22)
mark(bp, 'B2:G5', '拠点概要')

# --- 不動産（所有・賃借）
ro = wb['不動産(所有)']
clear(ro, 4, 21, 2, 11)
ro['I1'] = '土地は固定資産税評価額÷0.7、建物は固定資産税評価額'
RO = [('越谷工場', '土地）工場敷地', '埼玉県越谷市大字平方999-4', REAL_ESTATE_OWNED[0][3], '宅地', '要確認', "='BS(借方)'!F27", '=J4/0.7', REAL_ESTATE_OWNED[0][5], '固定資産税評価額÷0.7で時価を推計'),
      ('', '建物）工場棟', '同上（家屋番号は要確認）', REAL_ESTATE_OWNED[1][3], '鉄筋コンクリート造\n製氷工場', '1991年3月', "='BS(借方)'!F22", '=J5', REAL_ESTATE_OWNED[1][5], '固定資産税評価額を時価とみなす')]
for i, rec_ in enumerate(RO):
    r = 4 + i
    for c, v in zip('BCDEFGHIJK', rec_):
        ro[f'{c}{r}'] = v
    ro[f'E{r}'].number_format = AI
    for c in 'HIJ':
        ro[f'{c}{r}'].number_format = KIN
    ro.row_dimensions[r].height = 36
ro['C6'] = '合計'
for c in 'HIJ':
    ro[f'{c}6'] = f'=SUM({c}4:{c}5)'; ro[f'{c}6'].number_format = KIN
    copy_style(ro['H4'], ro[f'{c}6']); ro[f'{c}6'].number_format = KIN; ro[f'{c}6'].font = Font(name=NUM, size=12, bold=True)
copy_style(ro['C4'], ro['C6']); ro['C6'].font = Font(name=JP, size=12, bold=True)
for c in 'BDEFGK':
    copy_style(ro[f'{c}5'], ro[f'{c}6'])
hide_rows(ro, *range(7, 22))
for c_, w_ in zip('BCDEFGHIJK', [10, 15, 28, 10, 15, 17, 11, 10, 11, 26]):
    ro.column_dimensions[c_].width = w_
mark(ro, 'B3:K6', '不動産詳細（所有）')
ro['B8'] = 'チェック：BS土地・建物と一致'; ro['H8'] = "=H6='BS(借方)'!F27+'BS(借方)'!F22"
memo(ro, 'B9', '※固定資産税評価額は2025年3月期時点の所有物件資料による（最新年度の評価証明書で要確認）。建物面積は登記2,093.87㎡／内訳書2,065.00㎡で相違（要確認）。')
rl = wb['不動産(賃借) ']
clear(rl, 4, 21, 2, 10)
RL = [('松戸営業所', '建物）店舗', '千葉県松戸市稔台5-14-3', 400, '要確認', '法人（第三者）', 100000, '要確認', '年額1,200千円'),
      ('駐車場', '土地）駐車場', '埼玉県越谷市蒲生寿町19-15', '要確認', '要確認', '代表者', 45455, '－', '年額545,460円。譲渡後の継続条件は要協議')]
for i, rec_ in enumerate(RL):
    r = 4 + i
    for c, v in zip('BCDEFGHIJ', rec_):
        rl[f'{c}{r}'] = v
    rl[f'H{r}'].number_format = KIN
    rl.row_dimensions[r].height = 30
hide_rows(rl, *range(6, 22))
rl['B3'] = '拠点'
for c_, w_ in zip('BCDEFGHIJ', [12, 13, 28, 10, 12, 14, 11, 10, 30]):
    rl.column_dimensions[c_].width = w_
mark(rl, 'B3:J5', '不動産詳細（賃借）')

# --- 設備（科目別の償却資産＋主要設備）
eq = wb['設備']
clear(eq, 2, 22, 2, 11, style=True)
cw(eq, {'B': 1.6, 'C': 0.7, 'D': 1.2, 'E': 1.2, 'F': 1.2, 'G': 1.2, 'H': 3.0})
put(eq, 3, 8, '(単位：千円)', 'unit', border=None)
row(eq, 4, 2, ['科目', '件数', '取得価額', '期末簿価', '当期償却額', '当期償却不足額', '備考'], 'hdr', height=19)
DEP_NOTE = {'建物': '越谷工場（1991年3月取得）ほか。償却を停止', '構築物': '門・塀・構内舗装等', '機械': '製氷設備・冷凍機・袋詰設備等',
            '車両運搬具': '配送用トラック等（期中5件除却）', '器具備品': '事務機器等'}
for i, (nm, n_, acq, bv, dep_, short, acc) in enumerate(DEP_67):
    row(eq, 5 + i, 2, [nm if nm != '機械' else '機械装置', n_, acq, bv, dep_, short, DEP_NOTE[nm]],
        ['lab', 'num', 'num', 'num', 'num', 'num', 'txt'], fmts=[None, '#,##0"件"', KIN, KIN, KIN, KIN, None], height=19)
row(eq, 10, 2, ['合計', '=SUM(C5:C9)', '=SUM(D5:D9)', '=SUM(E5:E9)', '=SUM(F5:F9)', '=SUM(G5:G9)', ''],
    ['sub', 'subn', 'subn', 'subn', 'subn', 'subn', 'sub'], fmts=[None, '#,##0"件"', KIN, KIN, KIN, KIN, None], height=19)
mark(eq, 'B3:H10', '保有設備一覧（償却資産）', cell='A2')
eq['B12'] = 'チェック：減価償却計算書の期末簿価合計＝BS有形固定資産（土地除く）'
eq['E12'] = "=E10='BS(借方)'!F21-'BS(借方)'!F27"
eq['F13'] = '当期償却額＝PL減価償却費'; eq['F12'] = "=F10=PLハイライト!F15"
eqm = im_sheet(wb, '主要設備', '主要設備')
cw(eqm, {'B': 1.7, 'C': 0.6, 'D': 1.5, 'E': 2.4})
row(eqm, 4, 2, ['主要設備', '台数', '能力', '出所'], 'hdr', height=19)
for i, (nm, n_, cap) in enumerate(EQUIPMENT_MAIN):
    row(eqm, 5 + i, 2, [nm, n_, cap, '法人事業概況説明書（第67期）'], ['lab', 'txtc', 'txt', 'txt'], height=19)
EQ_MAIN = mark(eqm, 'B4:E7', '事業概要・純氷の製造工程・保有設備一覧（主要設備）', cell='A2')
memo(eq, 'B20', '※出所：第67期 減価償却計算書（合計表・185件）。資産別の明細はデータルームにて開示予定。主要設備の更新時期・修繕計画は要確認。')

# --- 追加シート：月別売上・許認可・沿革・その他補足
ms = im_sheet(wb, '月別売上', '月別売上（季節性）')
cw(ms, {'B': 1.25, **{L(3 + j): 0.62 for j in range(12)}, 'O': 0.78})
put(ms, 3, 15, '(単位：千円)', 'unit', border=None)
row(ms, 4, 2, ['項目'] + MONTHLY_67['month'] + ['合計'], 'hdr', height=17)
for i, (k_, lab) in enumerate([('凍氷', '売上（凍氷）'), ('保管料', '保管料'), ('原材料仕入', '原材料仕入'), ('人件費', '人件費'), ('従事員数', '従事員数（人）')]):
    r = 5 + i
    vals = MONTHLY_67[k_]
    kinds = ['lab'] + ['num'] * 13
    tot = f'=SUM(C{r}:N{r})' if k_ != '従事員数' else '-'
    row(ms, r, 2, [lab] + vals + [tot], kinds, fmts=[None] + ['#,##0'] * 13, height=17)
mark(ms, 'B3:O9', '月別売上高（季節性）')
ms['B11'] = '参考（リード文用）'; ms['B11'].font = Font(name=JP, size=11, bold=True, color=NAVY)
ms['B12'] = '7〜8月の構成比'; ms['C12'] = '=(F5+G5)/O5'
ms['B13'] = '6〜9月の構成比'; ms['C13'] = '=SUM(E5:H5)/O5'
ms['B14'] = '11〜2月の月平均'; ms['C14'] = '=AVERAGE(J5:M5)'
ms['B15'] = '売上合計と決算書の差'; ms['C15'] = '=O5*1000-PL!I5'
for a in ('C12', 'C13'):
    ms[a].number_format = '0.0%'
ms['C14'].number_format = '#,##0'; ms['C15'].number_format = '#,##0'
memo(ms, 'B17', '※出所：第67期（2026年3月期）法人事業概況説明書の月別売上高等の状況（千円単位で記載のため円換算しない）。')
ms['B20'] = '■ 事業構成（概要書「事業概要」用）'; ms['B20'].font = Font(name=JP, size=11, bold=True, color=NAVY)
for i, (lab, pct_) in enumerate(BUSINESS_MIX_67):
    ms[f'B{21 + i}'] = lab; ms[f'C{21 + i}'] = pct_ / 100; ms[f'C{21 + i}'].number_format = '0.0%'
ms['D21'] = '出所：法人事業概況説明書（第67期）の兼業割合'
ms['B25'] = '拠点別売上高（売上高等の事業所別内訳書）'
for i, (nm, _, cont, sv, _, emp, _) in enumerate(SITE_SALES_67[1:]):
    ms[f'B{26 + i}'] = nm; ms[f'C{26 + i}'] = sv; ms[f'D{26 + i}'] = f'=C{26 + i}/C28'; ms[f'E{26 + i}'] = cont
    ms[f'C{26 + i}'].number_format = '#,##0'; ms[f'D{26 + i}'].number_format = '0.0%'
ms['B28'] = '合計'; ms['C28'] = '=SUM(C26:C27)'; ms['C28'].number_format = '#,##0'
ms['B29'] = 'チェック：PL売上高と一致'; ms['C29'] = '=C28=PL!I11'

lic = im_sheet(wb, '許認可', '許認可')
cw(lic, {'B': 3.2, 'C': 1.8, 'D': 1.6, 'E': 4.3})
row(lic, 4, 2, ['名称', '番号', '許可日', '備考'], 'hdr', height=19)
for i, rec_ in enumerate(LICENSES):
    row(lic, 5 + i, 2, list(rec_), ['lab', 'txt', 'txt', 'txt'], height=22)
mark(lic, 'B4:E6', '許認可・取引金融機関（許認可）')

his = im_sheet(wb, '沿革', '沿革')
cw(his, {'B': 2.6, 'C': 8.3})
row(his, 4, 2, ['年月', '主な出来事'], 'hdr', height=20)
HIST = [('1960年（昭和35年）3月', '越ヶ谷製氷冷蔵株式会社を設立。越谷工場で操業を開始'),
        ('1964年（昭和39年）', '越谷工場に純水装置（イオン交換方式）を導入'),
        ('1972年（昭和47年）', '本社を越谷市蒲生寿町へ移転'),
        ('1990年（平成2年）', '逆浸透膜（RO）ろ過装置を導入'),
        ('1991年（平成3年）3月', '越谷工場（現工場棟）を建設'),
        ('2008年（平成20年）7月', '越谷工場の外壁塗装工事'),
        ('2012年（平成24年）4月', '資本金を11,000千円に変更'),
        ('2020年（令和2年）5月', '株式の譲渡制限に関する規定を変更'),
        ('2026年（令和8年）5月', '資本金を10,000千円に変更（減資）')]
for i, (d, e) in enumerate(HIST):
    row(his, 5 + i, 2, [d, e], ['lab', 'txt'], height=24)
mark(his, 'B4:C13', '沿革')
memo(his, 'B15', '※1960〜1990年は同社公式サイト等の公開情報、1991・2008年は減価償却計算書、2012年以降は登記による。')

oth = im_sheet(wb, 'その他補足', 'その他補足情報')
cw(oth, {'B': 2.2, 'C': 8.7})
row(oth, 4, 2, ['項目', '内容'], 'hdr', height=20)
OT = [('取引条件', '販売：月末締め翌月末回収（掛売99%）／仕入・外注：月末締め翌月20日払い／給与：月末締め翌月20日払い'),
      ('営業時間・休日', '8時〜17時／毎週土日曜日（法人事業概況説明書）'),
      ('帳簿・会計', '会計事務所が決算書・申告書を作成。製品受払簿・原材料受払簿・作業日報を備付'),
      ('取引金融機関', '日本政策金融公庫、栃木銀行、商工組合中央金庫、足利銀行、埼玉縣信用金庫、埼玉りそな銀行'),
      ('電力', '新電力（電力小売事業者）と契約。動力費は製造原価の主要費目（26/3期 33,685千円）'),
      ('加入団体', '越谷市工業団地組合ほか（出資金820千円）')]
for i, (k_, v_) in enumerate(OT):
    row(oth, 5 + i, 2, [k_, v_], ['lab', 'txt'], height=26)
mark(oth, 'B4:C10', 'その他補足情報')

# --- その他（要確認事項・ネクストステップ・社内確認事項）
ot = wb['その他']
ot['B2'] = '要確認事項・ネクストステップ（ヒアリング項目一覧の原本）'; ot['B2'].font = Font(name=JP, size=14, bold=True, color=NAVY)
ot.column_dimensions['B'].width = 6; ot.column_dimensions['C'].width = 16; ot.column_dimensions['D'].width = 100; ot.column_dimensions['E'].width = 8
row(ot, 4, 2, ['No.', '分野', '確認事項', '優先度'], 'hdr', height=19)
QS = [
    ('株主・資本', '発行済株式数の確定（登記22,000株／決算注記200株／別表二220株）と株主名簿・株券の現存状況（株券発行会社）', '高'),
    ('株主・資本', '2026年5月の減資（11,000千円→10,000千円）の目的・手続', '中'),
    ('譲渡条件', '代表者の連帯保証・個人所有土地の担保提供（栃木銀行・商工中金）の解除方法、代表者借入金14,000千円の取扱い（返済／債務免除）', '高'),
    ('譲渡条件', '代表者の引継ぎ期間・退任後の関与形態、代表者所有駐車場の賃貸借の継続', '高'),
    ('財務', '建物の会計上の償却停止の経緯（当期償却不足額21,745千円）と建物簿価（BS259,242千円／内訳書280,673千円）の差異', '高'),
    ('財務', '定期預金残高の内訳差異6,800千円（預貯金内訳19,601千円／BS26,401千円）、担保差入の有無', '高'),
    ('財務', '売掛金内訳の合計超過1,723,680円（(株)小野田商店）と同社との関係（関連当事者か）', '高'),
    ('財務', '退職給与6,178千円（67期）の対象者・退職金規程・中退共加入状況、今後の退職予定者', '高'),
    ('財務', '生命保険（大同生命・第一生命）の解約返戻金額・被保険者', '中'),
    ('財務', '65期の雑収入8,482千円・雑損5,192千円、支払手数料（66期3,156千円・67期1,765千円）の内容', '中'),
    ('財務', '賞与の支給時期・算定期間（賞与引当の要否）', '中'),
    ('事業', '取引先別・商品別（角氷／加工氷／袋氷／ドライアイス／保管）の年間売上高・数量・単価', '高'),
    ('事業', '電力契約の条件と電気料金の感応度、主要設備（製氷設備・冷凍機）の更新時期・修繕計画', '中'),
    ('組織', '従業員名簿（年齢・勤続・資格・雇用形態・給与）、工場長の継続意向、品質検査（代表者担当）の引継ぎ方法', '高'),
    ('許認可', '食品営業許可（氷雪販売業、2021年1月15日・5年更新）の更新状況', '高'),
    ('不動産', '工場土地・建物の最新固定資産税評価額、建物面積（2,093.87㎡／2,065.00㎡）の差異、代表者個人所有土地の明細', '中'),
]
for i, (a_, q_, p_) in enumerate(QS):
    row(ot, 5 + i, 2, [i + 1, a_, q_, p_], ['txtc', 'txt', 'txt', 'txtc'], height=19)
mark(ot, f'B4:E{4 + len(QS)}', '要確認事項・ネクストステップ', cell='A1')
rr = 5 + len(QS) + 1
row(ot, rr, 2, ['', 'プロセス関連（社内確認事項）', '', ''], 'hdr', height=19)
for i, t in enumerate(['※本項は社内確認事項であり、買手候補への開示資料には記載しない。',
                       '・第64〜66期の財務数値は先行して他の仲介会社が作成した企業概要書（財務表）から転記。決算書原本の受領と照合を行う。',
                       '・当該資料の買手候補への配布範囲・先行打診先の有無を確認し、重複打診を避ける。',
                       '・Web掲載画像（公式サイト）の概要書への使用は売主了承済み（担当者確認）。取得にはクラウド環境のネットワーク許可（king-ice.com）が必要。工場外観・工程・商品写真の提供も依頼する。',
                       '・必要資料：決算書（64〜66期原本）、株主名簿、定款、従業員名簿、取引先別売上、固定資産台帳、賃貸借契約書、保険証券、退職金規程']):
    memo(ot, f'C{rr + 1 + i}', t, size=10)

# --- ルール：貼付範囲一覧
ru = wb['ルール']
r = write_paste_index(ru, REG, start_row=17)
for key, rng, pg in [('BS', "'BS(借方)'!B3:F61", '貸借対照表（資産の部）'), ('BS', "'BS (貸方)'!B3:F25", '貸借対照表（負債の部）'),
                     ('BS', "'BS (貸方)'!B34:F51", '貸借対照表（純資産の部）'), ('PL', 'PL!B2:J56', '損益計算書'),
                     ('CR', 'CR!B4:J51', '製造原価報告書'), ('SGA', 'SGA!B2:J40', '販売費及び一般管理費'),
                     ]:
    ru.cell(r, 2, f'   {rng}　→　{pg}').font = Font(name=NUM, size=10)
    r += 1
ru.cell(r + 1, 2, '【本ファイルの出所】第64〜66期＝対象会社提供の既存会社説明資料（財務表）、第67期＝第67期決算報告書・勘定科目内訳書・法人税申告書・減価償却計算書、登記＝履歴事項（2026年9月9日現在）').font = Font(name=JP, size=10)
bsd['A1'] = '■ 貼付範囲 B3:F61 → 企業概要書「貸借対照表（資産の部）」（G列・空行は非表示）'
bsc['A1'] = '■ 貼付範囲 B3:F25（負債）／B34:F51（純資産）→ 企業概要書「貸借対照表」'
pl['A1'] = '■ 貼付範囲 B2:J56 → 企業概要書「損益計算書」（K:L列・空行は非表示）'
cr['A1'] = '■ 貼付範囲 B4:J51 → 企業概要書「製造原価報告書」（売上行・K:L列・空行は非表示）'
sga['A1'] = '■ 貼付範囲 B2:J40 → 企業概要書「販売費及び一般管理費」（売上行・K:M列・空行は非表示）'
for ws_ in (bsd, bsc, pl, cr, sga):
    ws_['A1'].font = Font(name='游ゴシック Medium', size=9, color='C00000')
    ws_['A1'].alignment = Alignment(wrap_text=False)

# シート順：追加シートをIM→の末尾（その他の前）へ
names = wb.sheetnames
for nm in ['主要費用', '主要設備', '借入金', '月別売上', '許認可', '沿革', 'その他補足']:
    wb.move_sheet(nm, offset=names.index('その他') - wb.sheetnames.index(nm))
    names = wb.sheetnames
wb.properties.creator = '株式会社M&A承継機構'; wb.properties.lastModifiedBy = ''
wb.save(OUT)
import json
json.dump({'REG': REG, 'LOAN_RANGE': LOAN_RANGE, 'EQ_MAIN': EQ_MAIN, 'SGA_ROW': SGA_ROW, 'CR_DEP_ROW': CR_DEP_ROW,
           'LOAN_TOTAL_ROW': LOAN_TOTAL_ROW, 'ADJ_R': ADJ_R}, open(os.path.splitext(OUT)[0] + '_index.json', 'w'), ensure_ascii=False, indent=1)
print('saved', OUT, len(REG))
