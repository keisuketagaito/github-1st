# -*- coding: utf-8 -*-
"""株式価値評価書（LUXAS株式会社）を当社フォーマットで作成する。
   数値はすべて案件マスター（Excel）由来。パワポには数字を直接打たない。"""
import openpyxl
from decimal import Decimal, ROUND_HALF_UP
from pptx import Presentation
from pptx.util import Inches, Pt
from pptlib import *

SRC    = 'vr_tpl.pptx'   # 当社フォーマットの評価書テンプレート（ma-valuation/assets）
OUT    = '/home/user/github-1st/LUXAS株式会社/株式価値評価書_LUXAS株式会社_20260923.pptx'
MASTER = '/home/user/github-1st/LUXAS株式会社/案件マスター_LUXAS株式会社_20260923.xlsx'

mb = openpyxl.load_workbook(MASTER, data_only=True)
def K(v):
    """円→千円（四捨五入・半数切上げ）。0は0のまま返す。"""
    if v is None or v == '':
        return None
    return int(Decimal(str(v / 1000.0)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
def g(sheet, addr):
    return K(mb[sheet][addr].value)
def raw(sheet, addr):
    return mb[sheet][addr].value

# ---- 主要な評価結果 ---------------------------------------------------------
NB, EV, NC = '年買法', 'EV EBITDAマルチプル', 'ネットキャッシュ'
BV_NET   = g(NB, 'C4')            # 簿価純資産
ADJ_NET  = g(NB, 'C5')            # 評価差額
TAXEFF   = g(NB, 'C6')            # 税効果
MV_NET   = g(NB, 'C7')            # 時価純資産
OP_BK    = [g(NB, c + '4') for c in 'FGH']
OP_ADJ   = [g(NB, c + '5') for c in 'FGH']
OP_ADJD  = [g(NB, c + '6') for c in 'FGH']
OP_TAX   = [g(NB, c + '7') for c in 'FGH']
OP_AT    = [g(NB, c + '8') for c in 'FGH']
BASE_OP  = g(NB, 'H10')
GW       = [g(NB, c + '5') for c in 'KLM']
NB_VAL   = [g(NB, c + '6') for c in 'KLM']
DEP3     = [g(EV, c + '7') for c in 'FGH']
EBITDA3  = [g(EV, c + '8') for c in 'FGH']
BASE_EB  = g(EV, 'H10')
EV_VAL   = [g(EV, c + '4') for c in 'KLM']
EV_DISC  = [g(EV, c + '5') for c in 'KLM']
NETCASH  = g(EV, 'C6')
EV_EQ    = [g(EV, c + '7') for c in 'KLM']
CASHLIKE = g(NC, 'C13')
DEBTLIKE = g(NC, 'F13')

VS = mb['評価参考']
def vs(r):
    return K(VS.cell(row=r, column=3).value)
PR_EBITDA, PR_OP = vs(8), vs(10)
PR_NET, PR_LO, PR_HI = vs(12), vs(15), vs(16)

BSL, BSR = mb['BS(借方)'], mb['BS (貸方)']
def bsl(row): return [K(BSL.cell(row=row, column=c).value) for c in (4, 5, 6)]
def bsr(row): return [K(BSR.cell(row=row, column=c).value) for c in (4, 5, 6)]
MBD, MBC = mb['修正BS(借方)'], mb['修正BS(貸方) ']
def mbd(row):
    return (K(MBD.cell(row=row, column=3).value), K(MBD.cell(row=row, column=4).value),
            K(MBD.cell(row=row, column=5).value), MBD.cell(row=row, column=7).value)
def mbc(row):
    return (K(MBC.cell(row=row, column=3).value), K(MBC.cell(row=row, column=4).value),
            K(MBC.cell(row=row, column=5).value), MBC.cell(row=row, column=7).value)
def lbl(ws, row):
    return ws.cell(row=row, column=2).value

MPL, MSG = mb['修正PL'], mb['修正SGA']
def pl3(row, col0=6):
    """修正PL：帳簿(F/I/L)・修正額(G/J/M)・修正後(H/K/N) の3期分を返す"""
    return ([K(MPL.cell(row=row, column=c).value) for c in (6, 9, 12)],
            [K(MPL.cell(row=row, column=c).value) for c in (7, 10, 13)],
            [K(MPL.cell(row=row, column=c).value) for c in (8, 11, 14)])
def sg3(row):
    return ([K(MSG.cell(row=row, column=c).value) for c in (6, 9, 12)],
            [K(MSG.cell(row=row, column=c).value) for c in (7, 10, 13)],
            [K(MSG.cell(row=row, column=c).value) for c in (8, 11, 14)],
            MSG.cell(row=row, column=19).value)

VK = mb['評価科目']
VK_ROWS = []
for r in range(5, 40):
    if VK.cell(row=r, column=3).value in (None, '評価差額 合計'):
        break
    VK_ROWS.append((VK.cell(row=r, column=3).value, K(VK.cell(row=r, column=4).value),
                    K(VK.cell(row=r, column=5).value), K(VK.cell(row=r, column=6).value),
                    VK.cell(row=r, column=7).value))

def look(sheet, label, cols=(4, 5, 6)):
    """マスターの科目名で行を引き当てて千円単位の値を返す。"""
    ws = mb[sheet]
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == label:
            return [K(ws.cell(row=r, column=c).value) for c in cols]
    raise KeyError(sheet + '!' + label)

def zebra(table, rows, rgb=(0xF2, 0xF6, 0xF9)):
    """指定行に淡い地色を敷いて行の区切りを見せる。"""
    from pptx.dml.color import RGBColor
    for r in rows:
        for c in range(len(table.columns)):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(*rgb)

def pad_col(table, col, left=0.06, right=0.04):
    """列の左右余白を確保する（隣の数値列と文字が接するのを防ぐ）。"""
    for r in range(len(table.rows)):
        cell = table.cell(r, col)
        cell.margin_left = Inches(left)
        cell.margin_right = Inches(right)

def header_row(table, r, pt=8.5):
    """行をヘッダー体裁（ネイビー地・白文字・太字・中央揃え）に整える。"""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    for c in range(len(table.columns)):
        cell = table.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x0B, 0x30, 0x41)
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(pt)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

def bold_rows(table, rows, on=True):
    """指定行の文字を太字にする（小計・合計行の強調用）。"""
    for r in rows:
        for c in range(len(table.columns)):
            for para in table.cell(r, c).text_frame.paragraphs:
                for run in para.runs:
                    run.font.bold = on

PER  = ['23年11月期', '24年11月期', '25年11月期']
PER4 = PER + ['26年11月期\n（7か月TB）']

def bs4(sheet, label):
    """BSシートから3期＋進行期TBの4期分を取得する"""
    ws = mb[sheet]
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == label:
            return [K(ws.cell(row=r, column=c).value) for c in (4, 5, 6, 7)]
    raise KeyError(sheet + '!' + label)

PLW = mb['PL']
def pl4(row):
    """PLシート：帳簿価額（C/E/G/I列）と売上高比（D/F/H/J列）"""
    v = [K(PLW.cell(row=row, column=c).value) for c in (5, 7, 9, 11)]
    p = [PLW.cell(row=row, column=c).value for c in (6, 8, 10, 12)]
    p = [x if isinstance(x, (int, float)) else None for x in p]
    return v, p

prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
def sec(slide, idx, name):
    set_text(slide.shapes[idx], name)
def rng(lo, hi):
    return '△{} 〜 △{}'.format(num(abs(lo)), num(abs(hi))) if lo < 0 and hi < 0 \
           else '{} 〜 {}'.format(num(lo), num(hi))


# ============================================================ 1. 表紙
s = S(1)
set_text(s.shapes[1], 'LUXAS株式会社　御中')
set_text(s.shapes[2], '2026年9月25日')

# ============================================================ 2. はじめに
s = S(2)
set_lines(s.shapes[5], [
 '本書における試算結果は対象会社より提出を受けた資料及び一般に公開されている情報並びに'
 '対象会社へのヒアリングに基づき算定したものであり、弊社はこれらの資料について監査・調査・'
 '検証を行っておりません。また、資産の実査・鑑定評価、財務デューデリジェンスその他の詳細調査は'
 '実施しておらず、その正確性及び完全性について保証するものではありません。',
 '本書に示す株式価値は、記載の評価基準日および前提条件のもとでの試算値であり、'
 '実際の譲渡価格を保証するものではありません。譲渡価格は、買手候補による事業評価・'
 'デューデリジェンスの結果および当事者間の交渉により決定されます。',
 '本書の存在及び記載されている情報並びに本書に関する全ての連絡・協議は、'
 '貴社と弊社間で締結している秘密保持契約書に定めのある秘密情報に該当致します。',
])

# ============================================================ 4. 目次（Appendixなし）
s = S(4)
set_lines(s.shapes[0], ['Valuation Summary ／ \n試算結果',
                        'Financial Analysis ／ \n財務情報の分析',
                        'Valuation Details ／ \n株式価値の試算'])
drop(s, 8, 7, 5, 4)

# ============================================================ 6. 採用するアプローチの選定
s = S(6)
t = s.shapes[6].table
fill_row(t, 1, ('コストアプローチ', '企業の純資産価値',
                '□ 簿価純資産価額法\n□ 時価純資産価額法\n■ 時価純資産＋営業権法\n　（年買法）', '',
                '時価純資産に会社の超過収益力である営業権を加算することで継続企業価値を算定する手法\n'
                '実質的な財政状態に加え、経営成績をバランスよく反映させることができる「時価純資産価額法」に'
                '「営業権」を加味する手法を採用しました。なお、本手法は、中堅・中小企業のM&Aでの株式評価において、'
                '最も広く用いられている手法です。'))
fill_row(t, 2, ('マーケットアプローチ', '株式市場における株価',
                '□ 市場価額法\n□ 類似業種比準法\n■ 類似会社比準法\n　（EV/EBITDAマルチプル法）', '',
                '規模・業種が類似する企業の株価指標（倍率）をもとに、価値を算定する手法\n'
                '株式市場における客観的な評価尺度である「株価」をベースとし、市場のトレンドを反映できる本手法を'
                '採用しました。なお、類似する上場企業の評価倍率を用いることで、第三者からの客観性と納得性が高い'
                '評価結果が得られる点が特徴です。'))
set_text(s.shapes[4],
  '最適な評価手法の選定にあたっては、対象会社の性質や評価目的などを踏まえ、総合的に判断する必要があります。')
fill_row(t, 3, ('インカムアプローチ', '企業の収益力',
                '□ 収益還元法\n□ DCF法\n　（ディスカウンティド\n　　キャッシュフロー法）\n□ 配当還元法', '',
                '収益還元法やDCF法を適用するには、根拠となる利益やキャッシュフローの計画値が必要となります。'
                '現状、詳細な事業計画が存在せず算定が困難であることから、これらの手法は採用しません。\n'
                '配当還元法は、本来、配当獲得を主眼とする少数株主の視点に立った評価手法です。したがって、'
                '本件取引の性質上、採用は適切ではないと判断しました。'))

# ============================================================ 7. 株式価値の試算結果
s = S(7)
t = s.shapes[4].table
for i, (lab, val) in enumerate([
 ('評価対象', 'LUXAS株式会社'),
 ('算定日', '2026年9月25日'),
 ('基準日', '2025年11月期（貸借対照表の基準日）'),
 ('参照期間', '2023年11月期〜2025年11月期（損益計算書の参照期間）'),
 ('目的', '対象会社の過半数の株式を第三者間で売買取引・M&Aする場合（「本件取引」といいます）の'
          '譲渡価額決定の参考資料とすること（「本件試算目的」といいます）を目的に本株式価値試算報告書は'
          '作成されております。'),
]):
    fill_row(t, i, (lab, val))
set_text(s.shapes[8],  '{} 〜 {}（千円）'.format(num(NB_VAL[0]), num(NB_VAL[2])))
set_text(s.shapes[12], '{} 〜 {}（千円）'.format(num(EV_EQ[0]), num(EV_EQ[2])))

# ============================================================ 9. 過去BSの推移
ASSET = ['【流動資産】', '現金及び預金', '売掛金', '商品', '貯蔵品', '前渡金', '前払費用',
         '未収入金', '未収還付法人税等', '仮払金', '立替金', '【固定資産】', '［有形固定資産］',
         '建物', '建物付属設備', '構築物', '車両運搬具', '工具器具及び備品', '土地',
         'その他の有形固定資産', '［無形固定資産］', '［投資その他資産］', 'リサイクル預託金',
         '長期前払費用', '差入保証金', '営業権', '【繰延資産】', '資産合計']
A_EMPH = {'【流動資産】', '【固定資産】', '［有形固定資産］', '［無形固定資産］',
          '［投資その他資産］', '【繰延資産】', '資産合計'}
DEBT = ['【流動負債】', '買掛金', '短期借入金', '未払金', '未払費用', '未払法人税等',
        '未払消費税等', '預り金', 'カード未払金', '【固定負債】', '長期借入金', '役員借入金',
        '長期未払金', '負債合計']
D_EMPH = {'【流動負債】', '【固定負債】', '負債合計'}
NET = ['【資本金】', '資本金', '【資本剰余金】', '【利益剰余金】', '繰越利益剰余金',
       '【自己株式】', '純資産合計', '負債・純資産合計']
N_EMPH = {'純資産合計', '負債・純資産合計'}

def bs_table(t, rows, emph, sheet, head='資産の部'):
    body = [(head, '', '', '', ''), ('科目',) + tuple(PER4)]
    for lab in rows:
        body.append((lab,) + tuple(num(v) for v in bs4(sheet, lab)))
    fit_rows(t, len(body), 3)
    restyle_rows(t, {i + 2: (2 if lab in emph else 3) for i, lab in enumerate(rows)})
    for i, row in enumerate(body):
        fill_row(t, i, row)
    set_widths(t, [1.72, 0.94, 0.94, 0.94, 0.93])
    table_font(t, 8)
    align_cells(t, (1, 2, 3, 4), 'r', rows=range(2, len(body)))
    align_cells(t, (0,), 'l', rows=range(2, len(body)))
    header_row(t, 1, 8)
    pad_col(t, 0, left=0.08)
    bold_rows(t, [i + 2 for i, lab in enumerate(rows) if lab in emph])
    return body

s = S(9)
bs_table(s.shapes[2].table, ASSET, A_EMPH, 'BS(借方)', '資産の部')
set_heights(s.shapes[2].table, [0.20, 0.38] + [0.205] * len(ASSET))
bs_table(s.shapes[3].table, DEBT, D_EMPH, 'BS (貸方)', '負債の部')
set_heights(s.shapes[3].table, [0.20, 0.38] + [0.205] * len(DEBT))
bs_table(s.shapes[4].table, NET, N_EMPH, 'BS (貸方)', '純資産の部')
set_heights(s.shapes[4].table, [0.20, 0.38] + [0.205] * len(NET))
place(s.shapes[4], top=4.96)

# ============================================================ 10. BS補足［資産］
s = S(10)
bs_table(s.shapes[4].table, ASSET, A_EMPH, 'BS(借方)', '資産の部')
set_heights(s.shapes[4].table, [0.20, 0.38] + [0.205] * len(ASSET))
set_lines(s.shapes[0], [
 '現預金：10,140千円（2025年11月期末）。月商の約0.2か月分と薄く、進行期は14,792千円へ回復。資金繰りの実態を要確認。',
 '売掛金：27,628千円の内訳は㈲ワンラブ（FC本部）17,373千円、カード・PayPay等の決済会社10,255千円。いずれも翌月回収で滞留債権は認められない。',
 '商品：106,017千円と総資産の45%を占める。内訳は店舗8店52,017千円、本社32,000千円、バックヤード12,000千円、外部保管10,000千円。'
 '本社・バックヤード分は2026年3月に全額廃棄されており、基準日時点で既に滞留・陳腐化していたと判断し評価減する（P.13）。',
 '仮払金：1,650千円。回収を前提とした支出ではなく資産性が認められないため全額評価減する。',
 '土地：2,974千円（駐車場用地1筆・89.4㎡）。路線価図・固定資産税評価証明が未受領のため簿価評価（要確認）。',
 '車両運搬具：1,058千円（事業用3台）。中古車市場価値が簿価を上回る可能性があるが査定書が未受領のため簿価評価（要確認）。',
 '差入保証金：27,371千円。賃貸借契約書が未受領で償却条項・返還条件が不明のため簿価評価。進行期は2,024千円減少しており、'
 'うち2,000千円は撤退店舗の敷金返還と推察（要確認）。',
 '長期前払費用：1,712千円。内訳が未受領のため簿価評価。信用保証料等の前払であれば換金価値は認められない（要確認）。',
 '営業権：1,500千円。仕入先デザイナーズワン㈱に係るもの。年買法で営業権を別途算定するため二重計上となり、全額評価減する。',
])
font_size(s.shapes[0], 8.5)

# ============================================================ 11. 評価科目別の検討結果
s = S(11)
set_text(s.shapes[0], '評価科目別の検討結果')
set_text(s.shapes[2], '評価科目')
t = s.shapes[3].table
VKH = [('№', '科目', '帳簿価額', '評価差額', '評価額', '検討結果・調整事由')]
VKB = [(str(i + 1), nm, num(bk), num(dv), num(av), note)
       for i, (nm, bk, dv, av, note) in enumerate(VK_ROWS)]
VKT = [('', '評価差額 合計', '－', num(ADJ_NET), '－', '')]
ALLV = VKH + VKB + VKT
fit_rows(t, len(ALLV), 2)
for i, row in enumerate(ALLV):
    fill_row(t, i, row)
place(s.shapes[3], top=1.35, width=10.95)
set_heights(t, [0.26] + [0.315] * (len(ALLV) - 1))
set_widths(t, [0.26, 2.00, 0.88, 0.88, 0.88, 6.05])
table_font(t, 8)
align_cells(t, (0,), 'c', rows=range(1, len(ALLV)))
align_cells(t, (1, 5), 'l', rows=range(1, len(ALLV)))
align_cells(t, (2, 3, 4), 'r', rows=range(1, len(ALLV)))
header_row(t, 0, 8)
pad_col(t, 1)
pad_col(t, 5, left=0.10)
bold_rows(t, [len(ALLV) - 1])

# ============================================================ 12/14. 修正BS
def adj_table(t, rows, emph, ws, head):
    body = [('　', '', '', '(単位：千円)'), (head, ''),
            ('科目', PER[2], '修正額', '修正後')]
    for r in rows:
        bk = K(ws.cell(row=r, column=3).value)
        dv = K(ws.cell(row=r, column=4).value)
        af = K(ws.cell(row=r, column=5).value)
        body.append((lbl(ws, r), num(bk), num(dv), num(af)))
    fit_rows(t, len(body), 4)
    restyle_rows(t, {i + 3: (3 if r in emph else 4) for i, r in enumerate(rows)})
    unspan_row(t, 2)
    for i, row in enumerate(body):
        fill_row(t, i, row)
    table_font(t, 8)
    align_cells(t, (0,), 'l', rows=range(3, len(body)))
    align_cells(t, (1, 2, 3), 'r', rows=range(3, len(body)))
    header_row(t, 2, 8)
    pad_col(t, 0, left=0.08)
    bold_rows(t, [i + 3 for i, r in enumerate(rows) if r in emph])
    return body

s = S(12)
ASSET_R = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 20, 21, 22, 23, 24, 25, 26, 27, 28,
           35, 46, 47, 48, 49, 50, 56, 61]
AR_EMPH = {5, 20, 21, 35, 46, 56, 61}
_b = adj_table(s.shapes[4].table, ASSET_R, AR_EMPH, MBD, '資産の部')
set_widths(s.shapes[4].table, [2.08, 1.18, 1.18, 1.14])
set_heights(s.shapes[4].table, [0.16, 0.22, 0.32] + [0.205] * len(ASSET_R))
set_lines(s.shapes[0], [
 '商品：本社32,000千円・バックヤード12,000千円は2026年3月に全額廃棄。基準日時点で既に滞留・陳腐化していたと判断し44,000千円を評価減。',
 '仮払金：回収を前提とした支出ではなく資産性が認められないため全額評価減。',
 '営業権：年買法で営業権を別途算定するため二重計上となること、および換金価値が認められないことから全額評価減。',
 'その他の科目：資料未受領により評価替えできないものは簿価評価とし、調整余地をP.12に整理している。',
])
font_size(s.shapes[0], 8.5)

s = S(14)
DEBT_R = [5, 6, 7, 8, 9, 10, 11, 12, 13, 18, 19, 20, 21, 27, 28,
          35, 36, 39, 43, 44, 48, 49, 50, 51, 52, 53]
DR_EMPH = {5, 18, 27, 28, 35, 39, 43, 48, 49, 52, 53}
adj_table(s.shapes[4].table, DEBT_R, DR_EMPH, MBC, '負債・純資産の部')
set_widths(s.shapes[4].table, [1.97, 1.17, 1.17, 1.16])
set_heights(s.shapes[4].table, [0.16, 0.22, 0.32] + [0.215] * len(DEBT_R))
set_lines(s.shapes[0], [
 '負債側に評価差額は計上していない。資産側の評価減{}千円が純資産に反映され、時価純資産は{}千円となる。'.format(
     num(ADJ_NET), num(MV_NET)),
 '借入金：長期借入金191,858千円（9本）。民間6本は代表者の連帯保証および信用保証協会の保証付き（公庫3本の保証条件は要確認）。'
 '譲渡実行時の連帯保証解除が譲渡条件の中心となる。役員借入金7,566千円は譲渡時の返済を希望。',
 '未払費用：賃金台帳では当月締・翌月25日払であり、11月分給与が未払計上されているか要確認。'
 '未計上であれば社会保険料概算15%を含め約10,000千円超の追加計上が必要。',
 '未計上の引当金等：賞与引当金・退職給付引当金・役員退職慰労引当金は規程が未受領のため、'
 '資産除去債務は撤退2店舗の原状回復費用が未確認のため、いずれも計上していない。',
 '税効果：繰越欠損金を有し課税所得の見込みが立たないため繰延税金資産の回収可能性を認めず、評価差額に対する税効果は認識していない。'
 '回収可能性が認められる場合は評価差額の34%相当（約16,031千円）だけ時価純資産が上振れする。',
])
font_size(s.shapes[0], 8.5)

# ============================================================ 13. BS補足［負債・純資産］
s = S(13)
bs_table(s.shapes[4].table, DEBT, D_EMPH, 'BS (貸方)', '負債の部')
set_heights(s.shapes[4].table, [0.20, 0.38] + [0.205] * len(DEBT))
bs_table(s.shapes[5].table, NET, N_EMPH, 'BS (貸方)', '純資産の部')
set_heights(s.shapes[5].table, [0.20, 0.38] + [0.205] * len(NET))
place(s.shapes[5], top=5.00)
set_lines(s.shapes[0], [
 '長期借入金：2024年11月期に大型犬専門店2店舗の出店等で218,455千円まで増加し、2025年11月期末は191,858千円。'
 '進行期の月次では期首残高のまま据え置かれており、返済額は決算時に一括計上されていると推察（要確認）。',
 '役員借入金：2025年11月期末7,566千円。進行期は2026年6月末で3,403千円へ減少している。譲渡時の返済を希望。',
 '未払消費税等：3,776千円→11,784千円と大幅に増加。納付時期・分割納付の有無を要確認。',
 'カード未払金：3,224千円（2025年11月期末）から進行期は7,790千円へ増加。仕入の決済手段の変化か要確認。',
 '純資産：2025年11月期は当期純損失20,359千円により△14,057千円の債務超過に転じている。'
 '進行期は棚卸資産廃棄損44,000千円の計上により2026年6月末で△52,244千円まで拡大。',
])
font_size(s.shapes[0], 8.5)

# ============================================================ 15. ネットキャッシュの算定
s = S(15)
t = s.shapes[2].table
NCL = [('現金及び預金', g(NC, 'C4'))]
NCD = [('長期借入金', g(NC, 'F5')), ('役員借入金', g(NC, 'F7')), ('未払法人税等', g(NC, 'F8'))]
fill_row(t, 1, ('キャッシュライクアイテム', '金額', '', 'デットライクアイテム', '金額'))
for i in range(3):
    l = NCL[i] if i < len(NCL) else ('', None)
    d = NCD[i] if i < len(NCD) else ('', None)
    fill_row(t, i + 2, (l[0] or '　', num(l[1]) if l[1] is not None else '　', '　',
                        d[0] or '　', num(d[1]) if d[1] is not None else '　'))
fill_row(t, 5, ('合計', num(CASHLIKE), '　', '合計', num(DEBTLIKE)))
fill_row(t, 6, ('　', '　', '　', '　', '　'))
fill_row(t, 7, ('ネットキャッシュ', num(NETCASH), '　', '　', '　'))
set_widths(t, [2.85, 1.55, 0.35, 4.65, 1.55])
table_font(t, 9)
align_cells(t, (1, 4), 'r', rows=range(2, 8))
bold_rows(t, [5, 7])
_nt = s.shapes.add_textbox(Inches(0.37), Inches(4.60), Inches(10.95), Inches(0.80))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※評価基準日（2025年11月30日）の残高。保険積立金・投資有価証券・役員貸付金の残高はない。',
 '※役員借入金は譲渡時の返済を希望しているため、有利子負債と同様にデットライクアイテムとして控除している。',
 '※長期未払金2,346千円は分割払いの未払金であり有利子負債ではないためデットライクに含めていない（要確認）。含める場合、想定株式価値は同額だけ下振れする。',
 '※事業運営に必要な運転資金相当額を現預金から控除する考え方もあるが、基準日の現預金水準（月商の約0.2か月分）に照らし、本評価では控除していない。',
]):
    pp = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    pp.text = line
    for rr in pp.runs:
        rr.font.size = Pt(8.5)
        rr.font.name = '游ゴシック Medium'

# ============================================================ 17. 過去PLの推移
s = S(17)
t = s.shapes[2].table
PLROWS = [(5, '生体売上高'), (6, 'サービス売上高'), (7, '物販売上高'), (11, '売上高'),
          (12, '期首棚卸高'), (13, '商品仕入高'), (14, '外注費'), (16, '期末棚卸高'),
          (17, '売上原価'), (18, '売上総利益'), (19, '販売費及び一般管理費'), (20, '営業利益'),
          (21, '営業外収益'), (22, '受取利息'), (24, '雑収入'), (29, '営業外費用'),
          (30, '支払利息'), (31, '雑損失'), (37, '経常利益'), (45, '特別損失'),
          (48, '棚卸資産廃棄損'), (52, '税引前当期純利益'), (53, '法人税等'), (57, '税引後当期純利益')]
PL_EMPH = {11, 17, 18, 19, 20, 21, 29, 37, 45, 52, 57}
body = [('', '', '', '', '', '', '', '', '(単位：千円)'),
        ('科目', PER4[0], '', PER4[1], '', PER4[2], '', PER4[3], ''),
        ('', '帳簿価額', '売上高比', '帳簿価額', '売上高比',
         '帳簿価額', '売上高比', '帳簿価額', '売上高比')]
for r, lab in PLROWS:
    v, p = pl4(r)
    row = [lab]
    for i in range(4):
        row += [num(v[i]), pct(p[i])]
    body.append(tuple(row))
fit_rows(t, len(body), 4)
restyle_rows(t, {i + 3: (3 if r in PL_EMPH else 4) for i, (r, _l) in enumerate(PLROWS)})
for i, row in enumerate(body):
    fill_row(t, i, row)
set_widths(t, [2.35, 1.10, 0.78, 1.10, 0.78, 1.10, 0.78, 1.10, 0.78])
set_heights(t, [0.16, 0.28, 0.24] + [0.235] * len(PLROWS))
table_font(t, 8)
align_cells(t, tuple(range(1, 9)), 'r', rows=range(3, len(body)))
bold_rows(t, [i + 3 for i, (r, _l) in enumerate(PLROWS) if r in PL_EMPH])

# ============================================================ 18/19. 販管費・修正PLの算定
def mod3_table(t, rows, emph, ws, labfn, extra=None):
    body = [('', '', '', '', '', '', '', '', '', '(単位：千円)'),
            ('科目', PER[0], '', '', PER[1], '', '', PER[2]),
            ('', '帳簿価額', '修正額', '修正後') * 1 + ('帳簿価額', '修正額', '修正後',
                                                        '帳簿価額', '修正額', '修正後')]
    body[2] = ('', '帳簿価額', '修正額', '修正後', '帳簿価額', '修正額', '修正後',
               '帳簿価額', '修正額', '修正後')
    for r in rows:
        v = [K(ws.cell(row=r, column=c).value) for c in (6, 7, 8, 9, 10, 11, 12, 13, 14)]
        body.append((labfn(ws, r),) + tuple(num(x) for x in v))
    if extra:
        body += extra
    fit_rows(t, len(body), 4)
    for i, row in enumerate(body):
        fill_row(t, i, row)
    set_widths(t, [2.15, 0.98, 0.98, 0.98, 0.98, 0.98, 0.98, 0.98, 0.98, 0.96])
    table_font(t, 8)
    align_cells(t, tuple(range(1, 10)), 'r', rows=range(3, len(body)))
    return body

s = S(18)
SGAROWS = list(range(5, 31)) + [39]
b = mod3_table(s.shapes[4].table, SGAROWS, {39}, MSG, lbl)
set_heights(s.shapes[4].table, [0.16, 0.26, 0.24] + [0.185] * len(SGAROWS))
bold_rows(s.shapes[4].table, [len(b) - 1])
set_text(s.shapes[0],
  '従業員給与：現場責任者（淺野崇氏）の人件費を役員報酬へ振替。役員報酬：代表者の退任分を戻し入れ、'
  '後任役員の同水準報酬を控除（差引ゼロ）。')
set_text(s.shapes[5],
  'リース料：非事業用車両のリース料を控除。接待交際費：適正水準4,000千円を超える部分を控除。'
  '保険料：事業関連性の低い生命保険料を控除。')
font_size(s.shapes[0], 8.5); font_size(s.shapes[5], 8.5)

s = S(19)
MPLROWS = [11, 17, 18, 19, 20, 21, 24, 29, 37]
BRIDGE = [('【正常収益力への橋渡し】', '', '', '', '', '', '', '', '', '')]
for r, lab in ((62, '修正後営業利益（修正P/L）'), (63, '　経常的雑収入（社員割引）の加算'),
               (64, '正常収益力（修正後営業利益ベース）')):
    v = [K(MPL.cell(row=r, column=c).value) for c in (8, 11, 14)]
    BRIDGE.append((lab, '', '', num(v[0]), '', '', num(v[1]), '', '', num(v[2])))
b = mod3_table(s.shapes[4].table, MPLROWS, set(), MPL, lbl, extra=BRIDGE)
set_heights(s.shapes[4].table, [0.16, 0.26, 0.24] + [0.32] * (len(MPLROWS) + len(BRIDGE)))
bold_rows(s.shapes[4].table, [3 + len(MPLROWS) + len(BRIDGE) - 1])
set_text(s.shapes[2],
  '雑収入：社員割引に伴い事業へ付随して継続的に発生する部分は経常的な損益として正常収益力に加算する。'
  'なお雑収入は営業外収益であるため修正P/L上の営業利益には含まれず、加算後の金額が企業概要書の調整項目合計と一致する。\n'
  '支払利息：資金の調達方法から独立した会社財産の収益獲得力を評価するため、正常収益力は営業利益ベースで算定している。\n'
  '特別損失：固定資産売却損968千円・除却損330千円は非経常的な損益であり、営業利益に含まれないため調整していない。'
  '2026年3月の棚卸資産廃棄損44,000千円は基準日後の事象であり、修正B/Sで評価差額として反映している。')
font_size(s.shapes[2], 8)

# ============================================================ 21. 年買法
s = S(21)
t = s.shapes[10].table                                  # 時価純資産（5x2）
for i, (lab, v) in enumerate([('', '(単位：千円)'), (PER[2], '金額'),
                              ('簿価純資産額', num(BV_NET)), ('修正額', num(ADJ_NET)),
                              ('時価純資産', num(MV_NET))]):
    fill_row(t, i, (lab, v))
table_font(t, 9); align_cells(t, (1,), 'r', rows=range(2, 5)); bold_rows(t, [4])

t = s.shapes[11].table                                  # 基準営業利益（9x4）
rows = [('', '', '', '(単位：千円)'), ('　',) + tuple(PER),
        ('簿価営業利益',) + tuple(num(x) for x in OP_BK),
        ('修正額',) + tuple(num(x) for x in OP_ADJ),
        ('修正後営業利益',) + tuple(num(x) for x in OP_ADJD),
        ('税考慮額（実効税率34%）',) + tuple(num(x) for x in OP_TAX),
        ('修正後税考慮後営業利益',) + tuple(num(x) for x in OP_AT),
        ('加重', '20%', '30%', '50%'),
        ('', '', '基準営業利益(三期加重平均)', num(BASE_OP))]
for i, row in enumerate(rows):
    fill_row(t, i, row)
table_font(t, 9); align_cells(t, (1, 2, 3), 'r', rows=range(2, 9)); bold_rows(t, [4, 6, 8])

t = s.shapes[12].table                                  # 想定株式価値（5x4）
for i, row in enumerate([('', '', '', '(単位：千円)'),
                         ('営業権倍率', 'x1', 'x2', 'x3'),
                         ('時価純資産', num(MV_NET), '　', '　'),
                         ('営業権',) + tuple(num(x) for x in GW),
                         ('想定株式価値',) + tuple(num(x) for x in NB_VAL)]):
    fill_row(t, i, row)
table_font(t, 9); align_cells(t, (1, 2, 3), 'r', rows=(3, 4)); align_cells(t, (1,), 'c', rows=(2,))
bold_rows(t, [4])
set_text(s.shapes[8],
  '・営業利益は過去3年の加重平均を採用（25年11月期：24年11月期：23年11月期　＝５：３：２）')
set_text(s.shapes[9],
  '・営業権倍率は国内の中小企業M&Aの実例におけるボリュームゾーンである1〜3倍を適用\n'
  '・繰越欠損金があり課税所得の見込みが立たないため繰延税金資産は不認識')
font_size(s.shapes[8], 8); font_size(s.shapes[9], 8); font_size(s.shapes[2], 8)

# ============================================================ 22. EV/EBITDAマルチプル法
s = S(22)
t = s.shapes[5].table
for i, (lab, v) in enumerate([('', '(単位：千円)'), (PER[2], '金額'),
                              ('キャッシュライク', num(CASHLIKE)),
                              ('デットライク', num(DEBTLIKE)),
                              ('ネットキャッシュ', num(NETCASH))]):
    fill_row(t, i, (lab, v))
table_font(t, 9); align_cells(t, (1,), 'r', rows=range(2, 5)); bold_rows(t, [4])

t = s.shapes[6].table
rows = [('', '', '', '(単位：千円)'), ('　',) + tuple(PER),
        ('簿価営業利益',) + tuple(num(x) for x in OP_BK),
        ('修正額',) + tuple(num(x) for x in OP_ADJ),
        ('修正後営業利益',) + tuple(num(x) for x in OP_ADJD),
        ('減価償却費',) + tuple(num(x) for x in DEP3),
        ('修正後EBITDA',) + tuple(num(x) for x in EBITDA3),
        ('加重', '20%', '30%', '50%'),
        ('', '', '基準EBITDA', num(BASE_EB))]
for i, row in enumerate(rows):
    fill_row(t, i, row)
table_font(t, 9); align_cells(t, (1, 2, 3), 'r', rows=range(2, 9)); bold_rows(t, [6, 8])

t = s.shapes[7].table
for i, row in enumerate([('', '', '', '(単位：千円)'),
                         ('マルチプル倍率', 'x3', 'x4', 'x5'),
                         ('EV（EBITDA×倍率）',) + tuple(num(x) for x in EV_VAL),
                         ('ネットキャッシュ', num(NETCASH), '　', '　'),
                         ('想定株式価値',) + tuple(num(x) for x in EV_EQ)]):
    fill_row(t, i, row)
table_font(t, 9); align_cells(t, (1, 2, 3), 'r', rows=(2, 4)); align_cells(t, (1,), 'c', rows=(3,))
bold_rows(t, [4])

# ============================================================ 追加① 参考｜進行期の年換算
s = dup_slide(prs, 11)
set_text(s.shapes[0], '参考｜進行期実績を年換算した場合の試算')
sec(s, 1, '1. Valuation Summary')
set_text(s.shapes[2], '前提｜不採算2店舗の撤退効果が通期で発現し、進行期7か月の水準が維持されること')
t = s.shapes[3].table
t_ = t; fit_cols(t_, 4); unmerge_v(t_)
REF = [('区分', '項目', '金額（千円）', '前提・算定方法'),
       ('正常収益力', '進行期7か月 営業利益', num(vs(5)), '月次では減価償却費が未計上のため実質的にEBITDAベース'),
       ('', '　正常収益力調整（7か月相当）', num(vs(6)), '第5期の調整項目合計{}千円を7か月按分'.format(num(OP_ADJ[2]))),
       ('', '進行期7か月 修正後EBITDA', num(vs(7)), ''),
       ('', '年換算 修正後EBITDA', num(PR_EBITDA), '7か月実績を12か月へ単純年換算'),
       ('', '　減価償却費（通期実績）', num(vs(9)), '第5期の実績額を通期見込みとして控除'),
       ('', '年換算 修正後営業利益', num(PR_OP), ''),
       ('株式価値', '時価純資産（2026年6月末ベース）', num(PR_NET),
        '進行期末の純資産から営業権・仮払金を控除。棚卸資産の廃棄損は進行期に計上済'),
       ('', '営業権（税考慮後営業利益×1倍）', num(vs(13)), '実効税率34%を考慮'),
       ('', '営業権（税考慮後営業利益×3倍）', num(vs(14)), '同上'),
       ('', '想定株式価値（営業権1〜3倍）', '{} 〜 {}'.format(num(PR_LO), num(PR_HI)), '時価純資産＋営業権', '', '')]
fit_rows(t, len(REF), 2)
restyle_rows(t, {i: 2 for i in range(1, len(REF))})
unmerge_v(t)
for i, row in enumerate(REF):
    fill_row(t, i, row)
place(s.shapes[3], top=1.35, width=10.95)
set_heights(t, [0.30] + [0.40] * (len(REF) - 1))
set_widths(t, [1.25, 3.05, 1.45, 5.20])
table_font(t, 9)
align_cells(t, (0, 1, 3), 'l', rows=range(1, len(REF)))
align_cells(t, (2,), 'r', rows=range(1, len(REF)))
header_row(t, 0, 9)
pad_col(t, 1); pad_col(t, 3, left=0.14)
bold_rows(t, [3, 4, 6, len(REF) - 1])
_nt = s.shapes.add_textbox(Inches(0.37), Inches(6.30), Inches(10.95), Inches(0.80))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※本試算は「進行期7か月の水準が通期で維持されること」を唯一の前提とした感応度分析であり、事業計画に基づくものではない。実績が確定するまでは評価結果として用いない。',
 '※進行期は月次で減価償却費・法定福利費が未計上であり、通期ではその分だけ利益が減少する。また2026年5月に租税公課4,969千円を一括計上しており、単月では5〜6月が営業赤字である。',
 '※マーケットアプローチはネットキャッシュ（2026年6月末で△180,469千円）の影響が大きく、年換算ベースでも株式価値はマイナスにとどまるため、本ページではコストアプローチのみを示している。',
]):
    pp = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    pp.text = line
    for rr in pp.runs:
        rr.font.size = Pt(8.5); rr.font.name = '游ゴシック Medium'
REF_SLIDE = len(prs.slides.__iter__.__self__._sldIdLst)

# ============================================================ 追加② 評価の前提と追加調整の可能性
s = dup_slide(prs, 11)
set_text(s.shapes[0], '評価の前提と追加調整の可能性')
sec(s, 1, '3. Valuation Details ')
set_text(s.shapes[2], '影響の方向｜＋＝株式価値の上振れ要因／△＝下振れ要因')
t = s.shapes[3].table
fit_cols(t, 4); unmerge_v(t)
TODO = [
 ('区分', '確認を要する事項', '方向', '評価への影響'),
 ('棚卸資産', '外部保管在庫10,000千円の実在性', '△', '実在しない場合、時価純資産が同額だけ下振れする'),
 ('未払費用', '11月分給与（当月締・翌月25日払）が未払計上されているか', '△',
  '未計上であれば社会保険料概算15%を含め約10,000千円超の追加計上が必要'),
 ('引当金', '賞与・退職金・役員退職慰労金の各規程の有無と要支給額', '△',
  '規程がある場合、要支給額を引当計上するため時価純資産が下振れする'),
 ('原状回復', '撤退2店舗の原状回復費用・違約金、稼働6店舗の資産除去債務', '△',
  '発生額を負債計上するため時価純資産が下振れする'),
 ('差入保証金', '賃貸借契約書の償却条項・契約解除時の返還額（基準日残高27,371千円）', '△',
  '償却条項がある場合、返還額まで評価減が必要'),
 ('長期前払費用', '内訳（信用保証料等の前払かどうか／基準日残高1,712千円）', '△',
  '換金価値が認められない場合は全額評価減'),
 ('土地・車両', '駐車場用地の路線価、事業用車両3台の査定額', '＋／△',
  '時価が簿価を上回る場合は時価純資産が上振れする'),
 ('税効果', '将来の課税所得の見通し', '＋',
  '回収可能性が認められる場合、評価差額の34%相当（約16,031千円）だけ時価純資産が上振れする'),
 ('正常収益力', '進行期（2026年11月期）の着地。減価償却費・法定福利費の通期見込み', '＋／△',
  '撤退効果が通期で発現すれば正常収益力は回復する（P.8の感応度分析を参照）'),
 ('正常収益力', '共通部門に残る本社仕入40,994千円・外注費5,610千円の店舗への配賦', '＋／△',
  '店舗別の収益性評価に影響する。全社の正常収益力の水準は変わらない'),
]
fit_rows(t, len(TODO), 2)
restyle_rows(t, {i: 2 for i in range(1, len(TODO))})
unmerge_v(t)
for i, row in enumerate(TODO):
    fill_row(t, i, row)
place(s.shapes[3], top=1.35, width=10.95)
set_heights(t, [0.30] + [0.44] * (len(TODO) - 1))
set_widths(t, [1.20, 4.20, 0.70, 4.85])
table_font(t, 8.5)
align_cells(t, (0, 1, 3), 'l', rows=range(1, len(TODO)))
align_cells(t, (2,), 'c', rows=range(1, len(TODO)))
header_row(t, 0, 8.5)
pad_col(t, 1); pad_col(t, 3, left=0.12)
_nt = s.shapes.add_textbox(Inches(0.37), Inches(6.55), Inches(10.95), Inches(1.00))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '【本評価書の位置づけ】',
 '・本評価書に示す想定株式価値は、記載の評価基準日および前提条件のもとでの試算値であり、実際の譲渡価格を保証するものではない。',
 '・いずれの手法でも想定株式価値はマイナスとなる。実務上は株式価値を備忘的な水準（ゼロ近傍）と置いたうえで、'
 '代表者の連帯保証（長期借入金191,858千円）の解除および役員借入金7,566千円の返済を譲渡条件の中心に据える整理が現実的である。',
 '・一方、正常収益力（基準営業利益{}千円・基準EBITDA{}千円）自体は確保されており、不採算2店舗の撤退効果が通期で発現すれば評価は改善する。'.format(
     num(BASE_OP), num(BASE_EB)),
]):
    pp = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    pp.text = line
    for rr in pp.runs:
        rr.font.size = Pt(8.5)
        rr.font.name = '游ゴシック' if i == 0 else '游ゴシック Medium'
        rr.font.bold = (i == 0)
TODO_SLIDE = len(prs.slides.__iter__.__self__._sldIdLst)

# ---------------------------------------------------------------- ページ構成
KEEP = [1, 2, 3, 4,
        5, 6, 7, REF_SLIDE,
        8, 9, 10, 11, 12, 13, 14, 15,
        16, 17, 18, 19,
        20, 21, 22, TODO_SLIDE,
        23]
prune_and_order(prs, KEEP)
prs.save(OUT)
print('saved:', OUT, '/ slides:', len(prs.slides.__iter__.__self__._sldIdLst))
