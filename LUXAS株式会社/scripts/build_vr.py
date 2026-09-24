# -*- coding: utf-8 -*-
"""株式価値評価書（LUXAS株式会社）を当社フォーマットで作成する。
   数値はすべて案件マスター（Excel）由来。パワポには数字を直接打たない。"""
import openpyxl
from decimal import Decimal, ROUND_HALF_UP
from pptx import Presentation
from pptx.util import Inches, Pt
from pptlib import *

SRC    = 'torizuka_im.pptx'
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

PER = ['23年11月期', '24年11月期', '25年11月期']
prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
def sec(slide, idx, name):
    set_text(slide.shapes[idx], name)
def rng(lo, hi):
    return '△{} 〜 △{}'.format(num(abs(lo)), num(abs(hi))) if lo < 0 and hi < 0 \
           else '{} 〜 {}'.format(num(lo), num(hi))

# ============================================================ 1. 表紙
s = S(1)
_p0 = s.shapes[0].text_frame.paragraphs[0]
_p0.runs[0].text = '株式価値評価書  '
_p0.runs[1].text = '<LUXAS株式会社>'
set_text(s.shapes[2], 'Valuation Report　')

# ============================================================ 2. はじめに
s = S(2)
set_text(s.shapes[2], 'はじめに')
set_lines(s.shapes[4], [
 '当該株式価値評価書（以下、「本評価書」といいます）は、株式会社M&A承継機構（以下「弊社」といいます）が、'
 '対象会社の発行済株式の取得もしくは事業その他資産の譲受（以下「本取引」といいます）に関する仲介者として、'
 '対象会社から提供された情報に基づき、対象会社の株式価値に関する参考情報を提供するために作成したものです。',
 '本評価書の存在及び記載されている情報並びに本評価書に関する全ての連絡・協議は、'
 '貴社と弊社間で締結している秘密保持契約書に定めのある秘密情報に該当致します。',
 '本評価書は、対象会社から提供された決算書・試算表その他の資料に依拠して作成しており、'
 '弊社はこれらの資料について監査・調査・検証を行っておりません。'
 'また、資産の実査・鑑定評価、財務デューデリジェンスその他の詳細調査は実施しておらず、'
 '情報提供者は本評価書に記載されているいかなる情報及び意見についても、'
 '明示・黙示にかかわらずその正確性及び完全性を表明または保証するものではありません。',
 '本評価書に示す株式価値は、記載の評価基準日および前提条件のもとでの試算値であり、'
 '実際の譲渡価格を保証するものではありません。譲渡価格は、買手候補による事業評価・'
 'デューデリジェンスの結果および当事者間の交渉により決定されます。',
 'いかなる場合でも、本取引に関して、対象会社の役職員（元役職員を含む）及び取引先に対して、'
 '貴社及び貴社のアドバイザー等本取引に関与する者が直接、間接を問わずコンタクトをとることを'
 '禁止させていただくと共に、全てのご連絡は必ず弊社を介していただきますようお願い申し上げます。',
])

# ============================================================ 3. 株価評価にあたっての注意点
s = S(61)
set_text(s.shapes[0], '株価評価にあたっての注意点')
sec(s, 1, '　はじめに')
set_text(s.shapes[2],
  '本評価書は2025年11月30日を評価基準日とし、対象会社から受領した第3期〜第5期の決算書一式、'
  '進行期の試算表・店舗別試算表、棚卸表および従業員名簿に依拠して作成している。'
  '以下の前提・限界を前提としてご覧いただきたい。')
t = s.shapes[3].table
NOTE_ROWS = [
 ('項目', '内容'),
 ('評価基準日', '2025年11月30日（第5期末＝直近決算日）。'
              '進行期の試算表（2026年6月末）は決算整理（減価償却・棚卸評価等）が未了のため評価の基礎としていない。'
              'ただし基準日後に判明した事実は評価差額に反映している'),
 ('評価の目的', '本取引における譲渡価格の検討にあたっての参考値の提示。'
              '税務上の時価・会計上の公正価値の算定を目的とするものではない'),
 ('採用したアプローチ', 'コストアプローチ（時価純資産＋営業権法）およびマーケットアプローチ（EV/EBITDAマルチプル法）。'
                        '選定理由はP.6を参照'),
 ('依拠した資料', '決算書・勘定科目内訳明細書・減価償却内訳明細書・法人税等申告書（第3期〜第5期）、'
                '勘定科目残高推移表（2025年12月〜2026年6月）、部門別損益計算書、棚卸表、従業員名簿、賃金台帳'),
 ('実施していない手続', '監査・財務デューデリジェンス、資産の実査・立会、不動産および動産の鑑定評価、'
                        '取引先・金融機関への確認。契約書（賃貸借・フランチャイズ）は未受領'),
 ('本評価書の位置づけ', '本評価書の試算値は交渉の出発点であり、最終的な譲渡価格は買手候補の事業評価と'
                        '当事者間の交渉により決定される。要確認事項（P.22）の解消により試算値は変動しうる'),
]
add_rows(t, 2, len(NOTE_ROWS) - len(trs(t)), after_idx=2)
for i, row in enumerate(NOTE_ROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=1.90, width=10.94)
set_heights(t, [0.30] + [0.62] * (len(NOTE_ROWS) - 1))
set_widths(t, [2.10, 8.84])
table_font(t, 9)

# ============================================================ 4. 目次
s = S(3)
set_lines(s.shapes[0], ['Valuation Summary', 'Financial Analysis', 'Valuation Details'])
drop(s, 4, 5)

# ============================================================ 5. Section1 扉
s = S(4)
set_text(s.shapes[0], 'Valuation Summary')

# ============================================================ 6. 採用するアプローチの選定
s = S(14)
set_text(s.shapes[0], '採用するアプローチの選定')
sec(s, 1, 'Valuation Summary')
set_text(s.shapes[2],
  '中小企業のM&Aで一般的に用いられるコストアプローチ（時価純資産＋営業権法）と'
  'マーケットアプローチ（EV/EBITDAマルチプル法）の2手法を採用する。'
  'インカムアプローチは、詳細な事業計画が存在せず将来キャッシュ・フローの合理的な見積りが困難であるため採用しない。')
t = s.shapes[3].table
fit_cols(t, 4)
unmerge_v(t)
APP_ROWS = [
 ('アプローチ', '手法', '採否', '採否の理由'),
 ('コストアプローチ', '時価純資産＋営業権法（年買法）', '採用',
  '純資産を時価評価したうえで正常収益力に基づく営業権を加算する手法。'
  '中小企業M&Aで最も一般的に用いられ、当事者にとって検証可能性が高い'),
 ('マーケットアプローチ', '類似会社比準法（EV/EBITDAマルチプル法）', '採用',
  '上場類似会社の株価倍率を参照して事業価値を算定する手法。'
  '市場環境を反映した客観的な検証手段として併用する'),
 ('インカムアプローチ', 'DCF法・収益還元法', '不採用',
  '将来の事業計画が策定されておらず、将来キャッシュ・フローを合理的に見積ることが困難であるため'),
 ('インカムアプローチ', '配当還元法', '不採用',
  '少数株主の視点に立つ手法であり、発行済株式の100%譲渡という本取引の性質に適合しないため'),
]
fit_rows(t, len(APP_ROWS), 2)
for i, row in enumerate(APP_ROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=2.05, width=10.83)
set_heights(t, [0.34] + [1.05] * (len(APP_ROWS) - 1))
set_widths(t, [1.90, 2.90, 0.90, 5.13])
table_font(t, 9.5)
align_cells(t, (0, 1, 3), 'l', rows=range(1, len(APP_ROWS)))
align_cells(t, (2,), 'c', rows=range(1, len(APP_ROWS)))
header_row(t, 0, 9.5)
for _c in (0, 1, 3):
    pad_col(t, _c, left=0.12)
drop(s, 4)

# ============================================================ 7. 株式価値の試算結果
s = S(13)
set_text(s.shapes[0], '株式価値の試算結果')
sec(s, 1, 'Valuation Summary')
set_text(s.shapes[2],
  '想定株式価値はコストアプローチで{}千円、マーケットアプローチで{}千円となり、'
  'いずれの手法でもマイナスとなる（評価基準日：2025年11月30日）。'.format(
      rng(NB_VAL[0], NB_VAL[2]), rng(EV_EQ[0], EV_EQ[2])))
t = s.shapes[3].table
for _r in range(len(t.rows)):
    unspan_row(t, _r)
fit_cols(t, 4)
SUM_ROWS = [
 ('評価手法', '想定株式価値（下限）', '想定株式価値（上限）', '算定の内訳（単位：千円）'),
 ('コストアプローチ\n（時価純資産＋営業権法）', num(NB_VAL[0]), num(NB_VAL[2]),
  '時価純資産 {} ＋ 営業権 {}〜{}（基準営業利益{}×1〜3倍）'.format(
      num(MV_NET), num(GW[0]), num(GW[2]), num(BASE_OP))),
 ('マーケットアプローチ\n（EV/EBITDAマルチプル法）', num(EV_EQ[0]), num(EV_EQ[2]),
  'EV {}〜{}（基準EBITDA{}×3〜5倍）＋ ネットキャッシュ {}'.format(
      num(EV_VAL[0]), num(EV_VAL[2]), num(BASE_EB), num(NETCASH))),
 ('参考｜進行期の年換算ベース\n（コストアプローチ）', num(PR_LO), num(PR_HI),
  '撤退効果が通期で発現した場合の仮説試算。時価純資産 {} ＋ 営業権（年換算修正後営業利益{}×1〜3倍）※P.8参照'.format(
      num(PR_NET), num(PR_OP))),
]
fit_rows(t, len(SUM_ROWS), 2)
for i, row in enumerate(SUM_ROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=1.60, width=10.85)
set_heights(t, [0.34, 0.62, 0.62, 0.62])
set_widths(t, [2.45, 1.50, 1.50, 5.40])
table_font(t, 9)
align_cells(t, (1, 2), 'r', rows=range(1, len(SUM_ROWS)))
# 試算の前提条件（当社フォーマットの定型ブロック）
_pre = s.shapes.add_textbox(Inches(0.35), Inches(3.92), Inches(10.85), Inches(0.62))
_pre.text_frame.word_wrap = True
for _i, _ln in enumerate([
 '【試算の前提条件】　評価対象：LUXAS株式会社　／　算定日：2026年9月23日　／　'
 '基準日：2025年11月期（貸借対照表の基準日）　／　参照期間：2023年11月期〜2025年11月期（損益計算書の参照期間）',
 '目的：対象会社の過半数の株式を第三者間で売買取引・M&Aする場合（「本件取引」といいます）の譲渡価額決定の'
 '参考資料とすること（「本件試算目的」といいます）を目的に本株式価値試算報告書は作成されております。',
]):
    _pp = _pre.text_frame.paragraphs[0] if _i == 0 else _pre.text_frame.add_paragraph()
    _pp.text = _ln
    for _rr in _pp.runs:
        _rr.font.size = Pt(8.5)
        _rr.font.name = '游ゴシック Medium'
set_text(s.shapes[6], '評価結果の読み方　―　株式価値は借入金の水準に規定されている')
set_text(s.shapes[8], 'コストアプローチ（採用・主たる手法）')
set_text(s.shapes[10], '時価純資産{}千円（うち評価差額{}千円）'.format(num(MV_NET), num(ADJ_NET)))
set_text(s.shapes[12], '営業権は{}〜{}千円（基準営業利益{}×1〜3倍）'.format(
    num(GW[0]), num(GW[2]), num(BASE_OP)))
set_text(s.shapes[14], '債務超過額を営業権で埋め切れずマイナスとなる')
set_text(s.shapes[16], 'マーケットアプローチ（検証手法）')
set_text(s.shapes[18], '基準EBITDA {}千円に対しEVは{}〜{}千円'.format(
    num(BASE_EB), num(EV_VAL[0]), num(EV_VAL[2])))
set_text(s.shapes[20], 'ネットキャッシュ{}千円が株式価値を押し下げる'.format(num(NETCASH)))
set_text(s.shapes[22], '事業価値は確保できるが借入超過分を打ち消せない')
set_text(s.shapes[24],
  'いずれの手法でも想定株式価値はマイナスであり、実務上は株式価値を備忘的な水準（ゼロ近傍）と置いたうえで、'
  '代表者の連帯保証（長期借入金{}千円）の解除と役員借入金{}千円の返済を譲渡条件の中心に据える整理が現実的である。'
  '一方、撤退効果が通期で発現すれば正常収益力は回復する。'.format(
      num(bsr(19)[2]), num(bsr(20)[2])))
set_text(s.shapes[30], '株式価値は2手法ともマイナス。\n借入金191,858千円が最大の要因')
set_text(s.shapes[32], '正常収益力（基準営業利益{}千円）\n自体は確保されている'.format(num(BASE_OP)))
set_text(s.shapes[35], '進行期の年換算では上限が\n{}千円まで回復する（P.8）'.format(num(PR_HI)))

# ============================================================ 8. 参考｜進行期の年換算ベース
s = S(15)
set_text(s.shapes[0], '参考｜進行期実績を年換算した場合の試算')
sec(s, 1, 'Valuation Summary')
set_text(s.shapes[2],
  '不採算2店舗の撤退効果が通期で発現し、進行期7か月の水準が年間を通じて維持されることを前提とした仮説試算。'
  '評価結果ではないが、撤退が正常収益力に与える影響の大きさを示すものとして併記する。')
t = s.shapes[3].table
REF_ROWS = [
 ('区分', '項目', '金額（千円）', '前提・算定方法'),
 ('正常収益力', '進行期7か月 営業利益', num(vs(5)), '月次では減価償却費が未計上のため実質的にEBITDAベース'),
 ('', '　正常収益力調整（7か月相当）', num(vs(6)), '第5期の調整項目合計{}千円を7か月按分'.format(num(OP_ADJ[2]))),
 ('', '進行期7か月 修正後EBITDA', num(vs(7)), ''),
 ('', '年換算 修正後EBITDA', num(PR_EBITDA), '7か月実績を12か月へ単純年換算'),
 ('', '　減価償却費（通期実績）', num(vs(9)), '第5期の実績額を通期見込みとして控除'),
 ('', '年換算 修正後営業利益', num(PR_OP), ''),
 ('株式価値', '時価純資産（2026年6月末ベース）', num(PR_NET),
  '進行期末の純資産から営業権・仮払金を控除。棚卸資産の廃棄損は進行期に計上済のため追加の評価減はない'),
 ('', '営業権（税考慮後営業利益×1倍）', num(vs(13)), '実効税率34%を考慮'),
 ('', '営業権（税考慮後営業利益×3倍）', num(vs(14)), '同上'),
 ('', '想定株式価値（営業権1〜3倍）', '{} 〜 {}'.format(num(PR_LO), num(PR_HI)), '時価純資産＋営業権'),
]
fit_rows(t, len(REF_ROWS), 2)
for i, row in enumerate(REF_ROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=1.93, width=10.95)
set_heights(t, [0.32] + [0.44] * (len(REF_ROWS) - 1))
set_widths(t, [1.25, 3.20, 1.60, 4.90])
table_font(t, 9)
align_cells(t, (0, 1, 3), 'l', rows=range(1, len(REF_ROWS)))
align_cells(t, (2,), 'r', rows=range(1, len(REF_ROWS)))
header_row(t, 0, 9)
pad_col(t, 1)
pad_col(t, 3, left=0.16)
bold_rows(t, [3, 4, 6, len(REF_ROWS) - 1])
note = s.shapes.add_textbox(Inches(0.37), Inches(7.00), Inches(10.95), Inches(0.80))
note.text_frame.word_wrap = True
for i, line in enumerate([
 '※本試算は「進行期7か月の水準が通期で維持されること」を唯一の前提とした感応度分析であり、事業計画に基づくものではない。実績が確定するまでは評価結果として用いない。',
 '※進行期は月次で減価償却費・法定福利費が未計上であり、通期ではその分だけ利益が減少する。また2026年5月に租税公課4,969千円を一括計上しており、単月では5〜6月が営業赤字である。',
 '※マーケットアプローチはネットキャッシュ（2026年6月末で△180,469千円）の影響が大きく、年換算ベースでも株式価値はマイナスにとどまるため、本ページではコストアプローチのみを示している。',
]):
    p = note.text_frame.paragraphs[0] if i == 0 else note.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8.5)
        r.font.name = '游ゴシック Medium'

# ============================================================ 9. Section2 扉
s = S(10)
set_text(s.shapes[0], 'Financial Analysis　― 時価純資産')

# ============================================================ 10. 過去B/Sの推移
s = S(55)
set_text(s.shapes[0], '過去B/Sの推移')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '総資産の45%を棚卸資産が占める資産構成。2024年11月期の出店で借入金が218,455千円まで増加し、'
  '2025年11月期は当期純損失{}千円により純資産が{}千円の債務超過に転じている。'.format(
      num(abs(look('BS (貸方)', '繰越利益剰余金')[1] - look('BS (貸方)', '繰越利益剰余金')[2])),
      num(BV_NET)))
ASSET = ['【流動資産】', '現金及び預金', '売掛金', '商品', '貯蔵品', '前渡金', '前払費用',
         '未収入金', '未収還付法人税等', '仮払金', '【固定資産】', '［有形固定資産］', '建物',
         '車両運搬具', '工具器具及び備品', '土地', 'その他の有形固定資産', '［無形固定資産］',
         '［投資その他資産］', 'リサイクル預託金', '長期前払費用', '差入保証金', '営業権',
         '【繰延資産】', '資産合計']
t = s.shapes[3].table
del_rows(t, range(2 + len(ASSET), 39))
_AE = {0, 10, 11, 17, 18, 23, 24}
restyle_rows(t, {i + 2: (2 if i in _AE else 3) for i in range(len(ASSET))})
place(s.shapes[3], top=1.58)
fill_row(t, 0, ('資産の部', '', '', ''))
fill_row(t, 1, ('科目', PER[0], PER[1], PER[2]))
for i, lab in enumerate(ASSET):
    fill_row(t, i + 2, [lab] + [num(v) for v in look('BS(借方)', lab)])
set_heights(t, [0.20, 0.26] + [0.21] * len(ASSET))
set_widths(t, [1.65, 0.95, 0.95, 0.95])

DEBT = ['【流動負債】', '買掛金', '未払金', '未払費用', '未払法人税等', '未払消費税等', '預り金',
        'カード未払金', '【固定負債】', '長期借入金', '役員借入金', '長期未払金', '負債合計']
t = s.shapes[4].table
add_rows(t, 3, 1, after_idx=3)
_DE = {0, 8, 12}
restyle_rows(t, {i + 2: (2 if i in _DE else 3) for i in range(len(DEBT))})
place(s.shapes[4], top=1.58)
fill_row(t, 0, ('負債の部', '', '', ''))
fill_row(t, 1, ('科目', PER[0], PER[1], PER[2]))
for i, lab in enumerate(DEBT):
    fill_row(t, i + 2, [lab] + [num(v) for v in look('BS (貸方)', lab)])
set_heights(t, [0.20, 0.26] + [0.23] * len(DEBT))
set_widths(t, [1.65, 0.95, 0.95, 0.95])

NET = ['【資本金】', '資本金', '【資本剰余金】', '【利益剰余金】', '繰越利益剰余金',
       '【自己株式】', '純資産合計', '負債・純資産合計']
t = s.shapes[5].table
del_rows(t, [9])
fill_row(t, 0, ('純資産の部', '', '', ''))
for i, lab in enumerate(NET):
    fill_row(t, i + 1, [lab] + [num(v) for v in look('BS (貸方)', lab)])
set_heights(t, [0.20] + [0.19] * len(NET))
set_widths(t, [1.65, 0.95, 0.95, 0.95])
place(s.shapes[5], top=5.40)
_nt = s.shapes.add_textbox(Inches(0.49), Inches(7.30), Inches(3.00), Inches(0.20))
_nt.text_frame.text = '※単位：千円。評価基準日は2025年11月30日（第5期末）。'
for _r in _nt.text_frame.paragraphs[0].runs:
    _r.font.size = Pt(8.5); _r.font.name = '游ゴシック Medium'

# ============================================================ 11. 評価科目別の検討結果
s = S(31)
set_text(s.shapes[0], '評価科目別の検討結果')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '評価差額は合計{}千円で、うち44,000千円は2026年3月に廃棄した棚卸資産である。'
  '資料が未受領で評価できなかった科目は簿価評価とし、追加調整の余地をあわせて記載している。'.format(num(ADJ_NET)))
t = s.shapes[3].table
VKH = [('№', '科目', '帳簿価額', '評価差額', '評価額', '検討結果・調整事由')]
VKB = [(str(i + 1), nm, num(bk), num(dv), num(av), note)
       for i, (nm, bk, dv, av, note) in enumerate(VK_ROWS)]
VKT = [('', '評価差額 合計', '－', num(ADJ_NET), '－', '')]
ALLV = VKH + VKB + VKT
fit_rows(t, len(ALLV), 2)
for i, row in enumerate(ALLV):
    fill_row(t, i, row)
place(s.shapes[3], top=1.72, width=10.96)
set_heights(t, [0.26] + [0.315] * (len(ALLV) - 1))
set_widths(t, [0.26, 2.00, 0.88, 0.88, 0.88, 6.06])
table_font(t, 8)
align_cells(t, (0,), 'c', rows=range(1, len(ALLV)))
align_cells(t, (2, 3, 4), 'r', rows=range(1, len(ALLV)))
pad_col(t, 5)

# ============================================================ 12. 修正B/S［資産］
s = S(7)
set_text(s.shapes[0], '修正B/S［資産の部］')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '資産合計{}千円に対し評価差額{}千円を計上し、修正後の資産合計は{}千円となる。'
  '評価減はいずれも換金価値・資産性が認められない科目に限っている。'.format(
      num(mbd(61)[0]), num(mbd(61)[1]), num(mbd(61)[2])))
t = s.shapes[7].table
ASSET_R = [5, 6, 7, 8, 9, 11, 13, 14, 20, 21, 22, 25, 26, 27, 28, 46, 47, 48, 49, 50, 61]
A_EMPH = {5, 20, 21, 46, 61}
rows = [('科目', '帳簿価額', '評価差額', '評価額')]
for r in ASSET_R:
    bk, dv, av, _ = mbd(r)
    rows.append((lbl(MBD, r), num(bk), num(dv), num(av)))
fit_rows(t, len(rows), 3)
_sp = {i + 1: (2 if r in A_EMPH else 3) for i, r in enumerate(ASSET_R)}
_sp[0] = 1
restyle_rows(t, _sp)
unspan_row(t, 0)
for i, row in enumerate(rows):
    fill_row(t, i, row)
place(s.shapes[7], top=1.76, width=6.32)
set_heights(t, [0.30] + [0.235] * (len(rows) - 1))
set_widths(t, [2.42, 1.30, 1.30, 1.30])
table_font(t, 8.5)
align_cells(t, (1, 2, 3), 'r', rows=range(1, len(rows)))
header_row(t, 0)
place(s.shapes[4], top=1.76, height=4.10)
place(s.shapes[5], top=2.32, height=3.50)
place(s.shapes[3], top=6.02, height=1.62)
place(s.shapes[8], top=6.10, height=1.52)
set_lines(s.shapes[5], [
 ('head', '≪棚卸資産（商品）≫'),
 '・基準日残高106,017千円のうち、本社32,000千円・バックヤード12,000千円は2026年3月に全額廃棄されている。'
 '基準日時点で既に滞留・陳腐化していたと判断し44,000千円を評価減。',
 '・残る外部保管在庫10,000千円は実在性が未確認であり、さらに評価減となる可能性がある（要確認）。',
 ('head', '≪営業権≫'),
 '・年買法で営業権を別途算定するため、B/S計上分を残すと二重計上となる。換金価値も認められないため全額評価減。',
 ('head', '≪仮払金≫'),
 '・回収を前提とした支出ではなく資産性が認められないため全額評価減。',
 ('head', '≪簿価評価とした主な科目≫'),
 '・土地（路線価資料未受領）、差入保証金（賃貸借契約書未受領）、長期前払費用（内訳未受領）、'
 '車両運搬具（査定書未受領）。いずれも評価替えの余地があり、P.22に要確認事項として整理している。',
])
font_size(s.shapes[5], 8.5)
set_lines(s.shapes[8], [
 ('head', '【税効果の取扱い】'),
 '評価差額に対する繰延税金資産は認識していない。対象会社は繰越欠損金を有し、'
 '評価差額が将来の課税所得を減額する見込みが立たないため、回収可能性が認められないと判断した。',
 '※保守的な取扱いであり、課税所得の見通しが立つ場合は最大で評価差額の34%相当'
 '（約16,031千円）だけ時価純資産が上振れする余地がある。',
])
font_size(s.shapes[8], 8.5)
set_text(s.shapes[6],
  '※評価基準日：2025年11月30日。※単位：千円。※千円未満の端数処理により合計が一致しない場合があります。')

# ============================================================ 13. 修正B/S［負債・純資産］
s = S(8)
set_text(s.shapes[0], '修正B/S［負債・純資産の部］')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '資産側の評価減{}千円が純資産に反映され、時価純資産は{}千円となる。'
  '賞与引当金・退職給付引当金・資産除去債務は規程・見積が未受領のため未計上である。'.format(
      num(ADJ_NET), num(MV_NET)))
t = s.shapes[4].table
DEBT_R = [5, 6, 8, 9, 10, 11, 12, 13, 18, 19, 20, 21, 28, 35, 43, 44, 49, 50, 51, 52, 53]
D_EMPH = {5, 18, 28, 49, 52, 53}
rows = [('科目', '帳簿価額', '評価差額', '評価額', '　')]
for r in DEBT_R:
    bk, dv, av, _ = mbc(r)
    rows.append((lbl(MBC, r), num(bk), num(dv), num(av), '　'))
fit_rows(t, len(rows), 3)
_sp = {i + 1: (2 if r in D_EMPH else 3) for i, r in enumerate(DEBT_R)}
_sp[0] = 1
restyle_rows(t, _sp)
unspan_row(t, 0)
for i, row in enumerate(rows):
    fill_row(t, i, row)
place(s.shapes[4], top=1.76, width=5.47)
set_heights(t, [0.30] + [0.235] * (len(rows) - 1))
set_widths(t, [1.97, 1.15, 1.15, 1.15, 0.05])
table_font(t, 8.5)
align_cells(t, (1, 2, 3), 'r', rows=range(1, len(rows)))
place(s.shapes[3], top=1.76, height=5.88)
place(s.shapes[5], top=2.36, height=5.20)
set_lines(s.shapes[5], [
 ('head', '≪借入金≫'),
 '・長期借入金191,858千円（9本）。民間6本は代表者の連帯保証および信用保証協会の保証付き'
 '（公庫3本の保証条件は要確認）。譲渡実行時の連帯保証解除が譲渡条件の中心となる。',
 '・役員借入金7,566千円は譲渡時の返済を希望。ネットキャッシュ算定ではデットライクアイテムとして控除している。',
 ('head', '≪未払費用≫'),
 '・賃金台帳では当月締・翌月25日払であり、11月分給与が未払計上されているか要確認。'
 '未計上であれば社会保険料概算15%を含め約10,000千円超の追加計上が必要となる。',
 ('head', '≪未計上の引当金等≫'),
 '・賞与引当金：賞与制度・支給実績が未確認のため未計上。',
 '・退職給付引当金／役員退職慰労引当金：退職金規程・役員退職慰労金規程が未受領のため未計上。',
 '・資産除去債務：店舗内装の原状回復義務。撤退2店舗の原状回復費用・違約金が未確認のため未計上。',
 ('head', '≪長期未払金≫'),
 '・2,346千円は分割払いの未払金であり有利子負債ではないため、ネットキャッシュのデットライクに含めていない（要確認）。',
 ('head', '≪時価純資産≫'),
 '・簿価純資産{}千円 ＋ 評価差額{}千円 ＋ 税効果{} ＝ {}千円。'.format(
     num(BV_NET), num(ADJ_NET), '－', num(MV_NET)),
])
font_size(s.shapes[5], 8.5)

# ============================================================ 14. ネットキャッシュの算定
s = S(53)
set_text(s.shapes[0], 'ネットキャッシュの算定')
sec(s, 1, 'Financial Analysis')
_sub = clone_shape(S(31), 2, s)
set_text(_sub,
  'マーケットアプローチにおいて事業価値（EV）から株式価値へ橋渡しするネットキャッシュは{}千円。'
  '現預金10,140千円に対し有利子負債が199,666千円あり、株式価値を大きく押し下げる要因となっている。'.format(
      num(NETCASH)))
t = s.shapes[2].table
fit_cols(t, 4)
del_rows(t, [0])
unmerge_v(t)
NCROWS = [
 ('区分', '科目', '金額（千円）', '取扱いの考え方'),
 ('キャッシュライク', '現金及び預金', num(g(NC, 'C4')), '基準日残高。運転資金として必要な水準の切り分けは行っていない'),
 ('', 'キャッシュライク 小計', num(CASHLIKE), '保険積立金・投資有価証券・役員貸付金の残高はない'),
 ('デットライク', '長期借入金', num(g(NC, 'F5')), '9本。すべて代表者の連帯保証付き'),
 ('', '役員借入金', num(g(NC, 'F7')), '譲渡時の返済を希望しているため有利子負債と同様に取扱う'),
 ('', '未払法人税等', num(g(NC, 'F8')), '事業運営に伴う支払義務として控除'),
 ('', 'デットライク 小計', num(DEBTLIKE), ''),
 ('ネットキャッシュ', 'キャッシュライク － デットライク', num(NETCASH),
  'EV/EBITDAマルチプル法における株式価値への調整額（P.21）'),
]
fit_rows(t, len(NCROWS), 2)
restyle_rows(t, {i: 1 for i in range(2, len(NCROWS))})
for i, row in enumerate(NCROWS):
    fill_row(t, i, row)
place(s.shapes[2], top=1.90, width=10.96)
set_heights(t, [0.34] + [0.52] * (len(NCROWS) - 1))
set_widths(t, [1.55, 3.15, 1.40, 4.86])
table_font(t, 9)
align_cells(t, (0, 1, 3), 'l', rows=range(1, len(NCROWS)))
align_cells(t, (2,), 'r', rows=range(1, len(NCROWS)))
header_row(t, 0, 9)
pad_col(t, 1)
pad_col(t, 3, left=0.16)
zebra(t, range(1, len(NCROWS), 2))
bold_rows(t, [2, 6, len(NCROWS) - 1])
_nt = s.shapes.add_textbox(Inches(0.34), Inches(6.30), Inches(10.96), Inches(0.70))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※評価基準日（2025年11月30日）の残高。',
 '※長期未払金2,346千円は分割払いの未払金であり有利子負債ではないためデットライクに含めていない（要確認）。含める場合、想定株式価値は同額だけ下振れする。',
 '※事業運営に必要な運転資金相当額を現預金から控除する考え方もあるが、基準日の現預金水準（月商の約0.2か月分）に照らし、本評価では控除していない。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8.5)
        r.font.name = '游ゴシック Medium'

# ============================================================ 15. Section2 扉（P/L）
s = S(20)
set_text(s.shapes[0], 'Financial Analysis　― 正常収益力')

# ============================================================ 16. 過去P/Lの推移
s = S(56)
set_text(s.shapes[0], '過去P/Lの推移')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '売上高は2期で17.7%成長した一方、2025年11月期は粗利率が59.1%へ低下し、'
  '販管費の増加とあわせて営業損失{}千円を計上した。'
  '正常収益力の算定はこの簿価営業利益を出発点とする。'.format(num(abs(OP_BK[2]))))
t = s.shapes[3].table
fit_cols(t, 4)
del_rows(t, [0, 2])
unmerge_v(t)
PLROWS = [
 (5, '生体売上高'), (6, 'サービス売上高'), (7, '物販売上高'), (11, '売上高'),
 (17, '売上原価'), (18, '売上総利益'), (19, '販売費及び一般管理費'), (20, '営業利益'),
 (21, '営業外収益'), (24, '　うち雑収入'), (29, '営業外費用'), (30, '　うち支払利息'),
 (37, '経常利益'), (45, '特別損失'),
 (52, '税引前当期純利益'), (53, '法人税等'), (57, '税引後当期純利益'),
]
P_EMPH = {11, 17, 18, 19, 20, 21, 29, 37, 45, 52, 57}
rows = [('科目', PER[0], PER[1], PER[2])]
for r, lab in PLROWS:
    bk, _, _ = pl3(r)
    rows.append((lab, num(bk[0]), num(bk[1]), num(bk[2])))
fit_rows(t, len(rows), 2)
restyle_rows(t, {i + 1: (2 if r in P_EMPH else 3) for i, (r, _l) in enumerate(PLROWS)})
for i, row in enumerate(rows):
    fill_row(t, i, row)
place(s.shapes[3], top=1.62, width=6.90)
set_heights(t, [0.32] + [0.285] * (len(rows) - 1))
set_widths(t, [2.40, 1.50, 1.50, 1.50])
table_font(t, 9)
align_cells(t, (0,), 'l', rows=range(1, len(rows)))
align_cells(t, (1, 2, 3), 'r', rows=range(1, len(rows)))
header_row(t, 0, 9)
bold_rows(t, [i + 1 for i, (r, _l) in enumerate(PLROWS) if r in P_EMPH])
_nt = s.shapes.add_textbox(Inches(7.50), Inches(1.62), Inches(3.80), Inches(5.00))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '【正常収益力算定上の着眼点】',
 '・営業利益は3期を通じて売上高比0.5%前後にとどまり、2025年11月期は不採算2店舗の影響で営業損失に転じている。',
 '・役員報酬・非事業用車両のリース料・事業関連性の低い生命保険料など、本件実行後に発生しない費用が販管費に含まれている（P.17）。',
 '・特別損失に計上された固定資産売却損968千円・除却損330千円は非経常的な損益であり、営業利益に含まれないため調整していない。',
 '・支払利息は債権者への分配前の収益力を見る観点から営業利益ベースで評価しており、あらためての控除は行っていない。',
 '・2026年3月の棚卸資産廃棄損44,000千円は基準日後の事象であり、P/Lではなく修正B/Sで評価差額として反映している（P.12）。',
 '※単位：千円。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(9) if i == 0 else Pt(8.5)
        r.font.name = '游ゴシック' if i == 0 else '游ゴシック Medium'
        r.font.bold = (i == 0)
    p.space_after = Pt(5)

# ============================================================ 17. 販管費の修正
s = S(57)
set_text(s.shapes[0], '販管費の修正')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  'M&A実行後に継続的に発生しない費用を戻し入れ、継続的に必要となる費用を控除する。'
  '2025年11月期の修正額は{}千円で、修正後の販管費は{}千円となる。'.format(
      num(sg3(39)[1][2]), num(sg3(39)[2][2])))
t = s.shapes[3].table
fit_cols(t, 5)
del_rows(t, [0, 2])
unmerge_v(t)
SGAROWS = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
           24, 25, 26, 27, 28, 29, 30, 39]
rows = [('科目', '帳簿価額', '修正額', '修正後', '調整事由')]
for r in SGAROWS:
    bk, ad, af, note = sg3(r)
    rows.append((lbl(MSG, r), num(bk[2]), num(ad[2]), num(af[2]), note or '－'))
fit_rows(t, len(rows), 2)
restyle_rows(t, {len(rows) - 1: 2})
for i, row in enumerate(rows):
    fill_row(t, i, row)
place(s.shapes[3], top=1.66, width=10.96)
set_heights(t, [0.28] + [0.192] * (len(rows) - 1))
set_widths(t, [1.85, 1.05, 1.05, 1.05, 5.96])
table_font(t, 8)
align_cells(t, (1, 2, 3), 'r', rows=range(1, len(rows)))
align_cells(t, (0, 4), 'l', rows=range(1, len(rows)))
header_row(t, 0, 8)
pad_col(t, 4)
bold_rows(t, [len(rows) - 1])
_nt = s.shapes.add_textbox(Inches(0.34), Inches(7.14), Inches(10.96), Inches(0.40))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※2025年11月期（評価基準期）の金額。第3期・第4期の修正額はそれぞれ△7,535千円・△19,814千円（案件マスター「修正SGA」シート参照）。単位：千円。',
 '※役員報酬は、代表者（淺野ゆう子氏）の退任分14,400千円を戻し入れ、後任役員が同水準の報酬を受ける前提で同額を控除しているため、差引の修正額はゼロとなる。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8)
        r.font.name = '游ゴシック Medium'

# ============================================================ 18. 修正P/Lの算定
s = S(23)
set_text(s.shapes[0], '修正P/Lの算定（正常収益力）')
sec(s, 1, 'Financial Analysis')
set_text(s.shapes[2],
  '販管費の修正を反映した修正後営業利益に、経常的に発生する雑収入（社員割引）を加算して正常収益力とする。'
  '2025年11月期の正常収益力は{}千円、3期を通じては{}千円〜{}千円の水準である。'.format(
      num(OP_ADJD[2]), num(min(OP_ADJD)), num(max(OP_ADJD))))
t = s.shapes[3].table
fit_cols(t, 7)
del_rows(t, [0])
unmerge_v(t)
MPLROWS = [(11, '売上高'), (17, '売上原価'), (18, '売上総利益'),
           (19, '販売費及び一般管理費'), (20, '営業利益'),
           (21, '営業外収益'), (24, '　うち雑収入'), (29, '営業外費用'), (37, '経常利益')]
M_EMPH = {11, 18, 20, 37}
rows = [('科目', PER[0], PER[1], '{}\n帳簿価額'.format(PER[2]),
         '{}\n修正額'.format(PER[2]), '{}\n修正後'.format(PER[2]), '調整事由')]
for r, lab in MPLROWS:
    bk, ad, af = pl3(r)
    rows.append((lab, num(bk[0]), num(bk[1]), num(bk[2]), num(ad[2]), num(af[2]),
                 MPL.cell(row=r, column=19).value or '－'))
BRIDGE = [(62, '修正後営業利益（修正P/L）'), (63, '　経常的雑収入（社員割引）の加算'),
          (64, '正常収益力（修正後営業利益ベース）')]
rows.append(('【正常収益力への橋渡し】', '', '', '', '', '', ''))
for r, lab in BRIDGE:
    v = [K(MPL.cell(row=r, column=c).value) for c in (8, 11, 14)]
    rows.append((lab, num(v[0]), num(v[1]), '－', '－', num(v[2]),
                 MPL.cell(row=r, column=19).value or '－'))
fit_rows(t, len(rows), 2)
restyle_rows(t, {i + 1: (2 if r in M_EMPH else 3) for i, (r, _l) in enumerate(MPLROWS)})
restyle_rows(t, {len(rows) - 4: 2, len(rows) - 1: 2})
for i, row in enumerate(rows):
    fill_row(t, i, row)
place(s.shapes[3], top=1.90, width=10.98)
set_heights(t, [0.46] + [0.34] * (len(rows) - 1))
set_widths(t, [2.10, 1.02, 1.02, 1.02, 0.96, 1.02, 3.84])
table_font(t, 8.5)
align_cells(t, (1, 2, 3, 4, 5), 'r', rows=range(1, len(rows)))
bold_rows(t, [i + 1 for i, (r, _l) in enumerate(MPLROWS) if r in M_EMPH] + [len(rows) - 1])
pad_col(t, 6)
_nt = s.shapes.add_textbox(Inches(0.34), Inches(7.10), Inches(10.98), Inches(0.50))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※第3期・第4期についても同一の方針で修正しており、修正後営業利益はそれぞれ{}千円・{}千円、正常収益力は{}千円・{}千円となる。'.format(
     num(K(MPL.cell(row=62, column=8).value)), num(K(MPL.cell(row=62, column=11).value)),
     num(OP_ADJD[0]), num(OP_ADJD[1])),
 '※雑収入は営業外収益であるため修正P/L上は営業利益に含まれないが、社員割引に伴い事業へ付随して継続的に発生するため正常収益力には加算する。'
 'この加算後の金額が企業概要書の調整項目合計と一致する。単位：千円。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8)
        r.font.name = '游ゴシック Medium'

# ============================================================ 19. Section3 扉
s = S(54)
set_text(s.shapes[0], 'Valuation Details')

# ============================================================ 20. 年買法による算出
s = S(38)
set_text(s.shapes[0], '年買法（時価純資産＋営業権法）による算出')
sec(s, 1, 'Valuation Details')
set_text(s.shapes[2],
  '正常収益力の加重平均（直近期50%・前期30%・前々期20%）に実効税率34%を考慮した基準営業利益は{}千円。'
  '営業権1〜3倍を時価純資産{}千円に加算した想定株式価値は{}千円となる。'.format(
      num(BASE_OP), num(MV_NET), rng(NB_VAL[0], NB_VAL[2])))
t = s.shapes[3].table
fit_cols(t, 5)
NBROWS = [
 ('区分', '項目', PER[0], PER[1], PER[2]),
 ('① 正常収益力', '簿価営業利益', num(OP_BK[0]), num(OP_BK[1]), num(OP_BK[2])),
 ('', '修正額（P.17・P.18）', num(OP_ADJ[0]), num(OP_ADJ[1]), num(OP_ADJ[2])),
 ('', '修正後営業利益', num(OP_ADJD[0]), num(OP_ADJD[1]), num(OP_ADJD[2])),
 ('', '税考慮額（実効税率34%）', num(OP_TAX[0]), num(OP_TAX[1]), num(OP_TAX[2])),
 ('', '修正後税考慮後営業利益', num(OP_AT[0]), num(OP_AT[1]), num(OP_AT[2])),
 ('', '加重', '20%', '30%', '50%'),
 ('', '基準営業利益（加重平均）', '', '', num(BASE_OP)),
 ('② 時価純資産', '簿価純資産額（2025年11月期末）', '', '', num(BV_NET)),
 ('', '評価差額（P.11・P.12）', '', '', num(ADJ_NET)),
 ('', '税効果（繰延税金資産は不認識）', '', '', '－'),
 ('', '時価純資産', '', '', num(MV_NET)),
 ('③ 想定株式価値', '営業権倍率', '1.0倍', '2.0倍', '3.0倍'),
 ('', '営業権（①×倍率）', num(GW[0]), num(GW[1]), num(GW[2])),
 ('', '想定株式価値（②＋営業権）', num(NB_VAL[0]), num(NB_VAL[1]), num(NB_VAL[2])),
]
fit_rows(t, len(NBROWS), 2)
restyle_rows(t, {3: 2, 5: 2, 7: 2, 11: 2, len(NBROWS) - 1: 2})
for i, row in enumerate(NBROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=1.90, width=9.60)
set_heights(t, [0.32] + [0.345] * (len(NBROWS) - 1))
set_widths(t, [1.72, 3.28, 1.52, 1.52, 1.56])
table_font(t, 9)
align_cells(t, (2, 3, 4), 'r', rows=range(1, len(NBROWS)))
bold_rows(t, [3, 5, 7, 11, len(NBROWS) - 1])
_nt = s.shapes.add_textbox(Inches(0.37), Inches(7.10), Inches(10.83), Inches(0.60))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※営業権倍率1〜3倍は国内中小企業M&Aのボリュームゾーン。※単位：千円。',
 '※修正額は企業概要書の調整項目合計と同一（販管費の修正＋経常的雑収入）。修正P/L上の修正後営業利益との差額は当該雑収入である（P.18の橋渡しを参照）。',
 '※評価差額に対する繰延税金資産は、繰越欠損金を有し課税所得の見込みが立たないため回収可能性を認めず、認識していない（保守的な取扱い）。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8)
        r.font.name = '游ゴシック Medium'

# ============================================================ 21. EV/EBITDAマルチプル法
s = S(59)
set_text(s.shapes[0], 'EV/EBITDAマルチプル法による算出')
sec(s, 1, 'Valuation Details')
set_text(s.shapes[2],
  '修正後EBITDAの加重平均による基準EBITDAは{}千円。事業価値（EV）にネットキャッシュ{}千円を'
  '加算した想定株式価値は{}千円となる。'.format(
      num(BASE_EB), num(NETCASH), rng(EV_EQ[0], EV_EQ[2])))
t = s.shapes[3].table
fit_cols(t, 5)
EVROWS = [
 ('区分', '項目', PER[0], PER[1], PER[2]),
 ('① 正常収益力', '修正後営業利益（P.18）', num(OP_ADJD[0]), num(OP_ADJD[1]), num(OP_ADJD[2])),
 ('', '減価償却費（販管費計上分）', num(DEP3[0]), num(DEP3[1]), num(DEP3[2])),
 ('', '修正後EBITDA', num(EBITDA3[0]), num(EBITDA3[1]), num(EBITDA3[2])),
 ('', '加重', '20%', '30%', '50%'),
 ('', '基準EBITDA（加重平均）', '', '', num(BASE_EB)),
 ('② ネットキャッシュ', 'キャッシュライク（P.14）', '', '', num(CASHLIKE)),
 ('', 'デットライク（P.14）', '', '', num(-DEBTLIKE)),
 ('', 'ネットキャッシュ', '', '', num(NETCASH)),
 ('③ 想定株式価値', 'マルチプル倍率', '3.0倍', '4.0倍', '5.0倍'),
 ('', '事業価値EV（①×倍率）', num(EV_VAL[0]), num(EV_VAL[1]), num(EV_VAL[2])),
 ('', 'ネットキャッシュ（②）', num(NETCASH), num(NETCASH), num(NETCASH)),
 ('', '想定株式価値', num(EV_EQ[0]), num(EV_EQ[1]), num(EV_EQ[2])),
]
fit_rows(t, len(EVROWS), 2)
restyle_rows(t, {3: 2, 5: 2, 8: 2, len(EVROWS) - 1: 2})  # 修正後EBITDA・基準EBITDA・ネットキャッシュ・想定株式価値
for i, row in enumerate(EVROWS):
    fill_row(t, i, row)
place(s.shapes[3], top=1.90, width=9.60)
set_heights(t, [0.32] + [0.375] * (len(EVROWS) - 1))
set_widths(t, [1.72, 3.28, 1.52, 1.52, 1.56])
table_font(t, 9)
align_cells(t, (2, 3, 4), 'r', rows=range(1, len(EVROWS)))
bold_rows(t, [3, 5, 8, len(EVROWS) - 1])
_nt = s.shapes.add_textbox(Inches(0.37), Inches(7.10), Inches(10.83), Inches(0.60))
_nt.text_frame.word_wrap = True
for i, line in enumerate([
 '※修正後EBITDA＝修正後営業利益＋減価償却費であり、企業概要書の調整後EBITDAと同一である。※単位：千円。',
 '※想定株式価値＝事業価値EV＋ネットキャッシュ。当社フォーマットでは非流動性ディスカウントを適用していない。',
 '※類似上場会社の選定にあたっては、ペット関連小売・生体販売を主要事業とする国内上場会社を想定しているが、'
 '同社は債務超過かつ売上規模が大きく異なるため、倍率レンジ3〜5倍は保守的に設定している（要確認）。',
]):
    p = _nt.text_frame.paragraphs[0] if i == 0 else _nt.text_frame.add_paragraph()
    p.text = line
    for r in p.runs:
        r.font.size = Pt(8)
        r.font.name = '游ゴシック Medium'

# ============================================================ 22. 評価の前提と要確認事項
s = S(60)
set_text(s.shapes[0], '評価の前提と追加調整の可能性')
sec(s, 1, 'Valuation Details')
drop(s, 5)                                    # 前案件の顔写真
set_text(s.shapes[2],
  '本評価は受領済資料の範囲で行っており、以下の事項が確認されれば試算値は変動する。'
  '影響の方向（＋＝株式価値の上振れ要因／△＝下振れ要因）をあわせて示す。')
t = s.shapes[3].table
fit_cols(t, 4)
TODO = [
 ('区分', '確認を要する事項', '方向', '評価への影響'),
 ('棚卸資産', '外部保管在庫10,000千円の実在性', '△',
  '実在しない場合、時価純資産が同額だけ下振れする'),
 ('未払費用', '11月分給与（当月締・翌月25日払）が未払計上されているか', '△',
  '未計上であれば社会保険料概算15%を含め約10,000千円超の追加計上が必要'),
 ('引当金', '賞与・退職金・役員退職慰労金の各規程の有無と要支給額', '△',
  '規程がある場合、要支給額を引当計上するため時価純資産が下振れする'),
 ('原状回復', '撤退2店舗の原状回復費用・違約金、および稼働6店舗の資産除去債務', '△',
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
for i, row in enumerate(TODO):
    fill_row(t, i, row)
place(s.shapes[3], top=1.72, width=10.94)
set_heights(t, [0.30] + [0.44] * (len(TODO) - 1))
set_widths(t, [1.25, 4.35, 0.70, 4.64])
table_font(t, 8.5)
align_cells(t, (2,), 'c', rows=range(1, len(TODO)))
set_lines(s.shapes[4], [
 ('head', '【本評価書の位置づけ】'),
 '・本評価書に示す想定株式価値は、記載の評価基準日および前提条件のもとでの試算値であり、実際の譲渡価格を保証するものではない。',
 '・いずれの手法でも想定株式価値はマイナスとなる。実務上は株式価値を備忘的な水準（ゼロ近傍）と置いたうえで、'
 '代表者の連帯保証の解除および役員借入金の返済を譲渡条件の中心に据える整理が現実的である。',
 '・一方、正常収益力（基準営業利益{}千円・基準EBITDA{}千円）自体は確保されており、'
 '不採算2店舗の撤退効果が通期で発現すれば評価は改善する。進行期の着地がレンジを左右する最大の変数である。'.format(
     num(BASE_OP), num(BASE_EB)),
 '・最終的な譲渡価格は、買手候補による事業評価・デューデリジェンスの結果および当事者間の交渉により決定される。',
])
place(s.shapes[4], top=6.52, width=10.94, height=1.10)
font_size(s.shapes[4], 8.5)

# ============================================================ 23. 裏表紙
s = S(63)

# ---------------------------------------------------------------- ページ構成
KEEP = [1, 2, 61, 3,
        4, 14, 13, 15,
        10, 55, 31, 7, 8, 53,
        20, 56, 57, 23,
        54, 38, 59, 60,
        63]
prune_and_order(prs, KEEP)
prs.save(OUT)
print('saved:', OUT, '/ slides:', len(prs.slides.__iter__.__self__._sldIdLst))
