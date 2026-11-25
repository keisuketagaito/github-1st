"""ジャパニアス株式会社 化学・バイオ事業（事業譲渡）：公式フォーマットの案件マスターを作る。

入力（MA_WORK/up/）：
  pl.xlsx         化学・バイオ損益計算書（事業部PL、月次、2024/11期〜2026/11期、ファイル日付261020）
  customers.xlsx  顧客別売上高（2024/11期・2025/11期・2026/11期 12〜10月実績）
  employees.xlsx  従業員一覧（2026年10月末在籍、技術者33名＋管理社員（営業職）1名）
出力：MA_WORK/work/master_raw.xlsx（数式のみ）→ LibreOfficeで再計算 → inject_cache で値を書き戻す

本件は上場会社の一事業部門の事業譲渡であり、対象事業単独の貸借対照表・決算書は存在しない。
そのため公式マスターの PL/SGA（会社の決算書形式）には転記せず、受領した事業部PL（月次）を
「事業部PL月次」に転記し、IM用の各シートはそこを数式で参照する。株主・役員・不動産・設備・BS・VR→は使用しない。
"""
import os, sys, datetime, copy, json
from collections import OrderedDict
REPO = os.environ.get('REPO', os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
SK = os.path.join(REPO, '.claude', 'skills', 'ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border
from openpyxl.utils import get_column_letter as CL
from master_im_style import (im_sheet, cw, put, row, paste, note, write_paste_index, hide_cols,
                             KI, PI, YI, JP, NUM, NAVY, NOTE)

W = os.environ.get('MA_WORK', '.')
wb = load_workbook(os.path.join(SK, 'assets', 'master_template.xlsx'))
REG = []

# ======================================================================== 受領データ
src = load_workbook(f'{W}/up/pl.xlsx', data_only=True)['Sheet1']
MCOLS = list(range(2, 14)) + list(range(16, 28)) + list(range(30, 42))   # B..M, P..AA, AD..AO（36か月）
MONTHS = [src.cell(3, c).value for c in MCOLS]                            # Excel日付（シリアル値またはdatetime）
MONTHS = [(datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(m))) if isinstance(m, (int, float)) else m for m in MONTHS]
FLAG = [src.cell(4, c).value for c in MCOLS]                              # 実績 / 見込/実績


def mrow(r):
    out = []
    for c in MCOLS:
        v = src.cell(r, c).value
        out.append(v if isinstance(v, (int, float)) else None)
    return out


cus = load_workbook(f'{W}/up/customers.xlsx', data_only=True).active
emp = load_workbook(f'{W}/up/employees.xlsx', data_only=True).active


def clear_sheet(ws):
    """公式シートの見本（値・書式・結合・行高・列幅）を消して作り直す。"""
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    for r in ws.iter_rows():
        for c in r:
            c.value = None
            c.font = Font(name=JP, size=11); c.fill = PatternFill(); c.border = Border()
            c.alignment = Alignment(); c.number_format = 'General'
    for k in list(ws.row_dimensions.keys()):
        ws.row_dimensions[k].height = None
        ws.row_dimensions[k].hidden = False
    for k in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[k]
    ws.sheet_view.showGridLines = False


def title(ws, t):
    ws['A1'] = t
    ws['A1'].font = Font(name=JP, size=14, bold=True, color=NAVY)


def mark(ws, rng, page, cell='A2'):
    return paste(ws, rng, page, REG, cell=cell)


def red(ws, cell, t):
    ws[cell] = t
    ws[cell].font = Font(name=NOTE, size=9, color='C00000')


def new_sheet(name, ttl, before):
    ws = wb.create_sheet(name, index=wb.sheetnames.index(before))
    ws.sheet_view.showGridLines = False
    title(ws, ttl)
    return ws


# ======================================================================== 財務入力：事業部PL月次（受領資料の転記）
SRC = '事業部PL月次'
pm = new_sheet(SRC, '化学・バイオ事業 事業部損益計算書（月次）— 受領資料「化学・バイオ損益計算書 261020 v2」の転記（円）', 'VR→ ')
red(pm, 'A2', '■ 転記元：化学・バイオ損益計算書（事業部PL、ファイル日付2026/10/20）。2026年10・11月は「見込/実績」表示で見込値を含む。'
               '本シートは貼付範囲ではない（IM用シートが数式で参照する）。円単位で入力。')
pm.column_dimensions['A'].width = 2
pm.column_dimensions['B'].width = 34
C0 = 3                                    # C列＝2023年12月
MC = [CL(C0 + i) for i in range(36)]      # 36か月の列
for i, c in enumerate(MC):
    pm.column_dimensions[c].width = 11
    d = MONTHS[i]
    pm[f'{c}4'] = d.strftime('%Y/%m') if hasattr(d, 'strftime') else str(d)
    pm[f'{c}5'] = '見込' if '見込' in str(FLAG[i]) else '実績'
    for rr in (4, 5):
        pm[f'{c}{rr}'].font = Font(name=JP, size=9, bold=True, color='FFFFFF')
        pm[f'{c}{rr}'].fill = PatternFill('solid', fgColor=NAVY)
        pm[f'{c}{rr}'].alignment = Alignment(horizontal='center')
pm['B4'] = '科目（月）'; pm['B5'] = '区分'
for a in ('B4', 'B5'):
    pm[a].font = Font(name=JP, size=9, bold=True, color='FFFFFF'); pm[a].fill = PatternFill('solid', fgColor=NAVY)
# 期間合計の列
PER = OrderedDict([('FY24', (0, 11, '2024/11期')), ('FY25', (12, 23, '2025/11期')), ('FY26', (24, 35, '2026/11期見込')),
                   ('CUR', (24, 33, '2025/12〜2026/9')), ('PY', (12, 21, '2024/12〜2025/9'))])
PC = {}
for j, (k, (a, b, lab)) in enumerate(PER.items()):
    c = CL(C0 + 36 + 1 + j)
    PC[k] = c
    pm.column_dimensions[c].width = 14
    pm[f'{c}4'] = lab; pm[f'{c}5'] = '合計'
    for rr in (4, 5):
        pm[f'{c}{rr}'].font = Font(name=JP, size=9, bold=True, color='FFFFFF')
        pm[f'{c}{rr}'].fill = PatternFill('solid', fgColor='0889C9')
        pm[f'{c}{rr}'].alignment = Alignment(horizontal='center')

ITEMS = [  # (key, 表示名, 受領資料の行 or None=数式, 区分)
    ('sales_d', '派遣売上高', 5, 'in'), ('sales_u', '請負売上高', 6, 'in'), ('sales', '純売上高', None, 'f'),
    ('c_sal', '技術者給与手当', 10, 'in'), ('c_hw', '技術者法定福利費', 11, 'in'), ('c_bon', '技術者賞与引当金繰入額', 12, 'in'),
    ('c_ret', '退職給付費用（概算）', 13, 'in'), ('c_lab', '労務費', None, 'f'), ('c_trv', '旅費交通費（原価）', 16, 'in'),
    ('cogs', '当期総製造費用（売上原価）', None, 'f'), ('gp', '売上総利益', None, 'f'),
    ('s_sal', '本社給与手当', 23, 'in'), ('s_hw', '法定福利費', 24, 'in'), ('s_bon', '賞与引当金繰入額', 25, 'in'),
    ('s_ret', '退職給付費用', 26, 'in'), ('s_rec', '採用紹介料', 27, 'in'), ('s_ad', '広告宣伝費', 28, 'in'),
    ('s_trv', '旅費交通費', 29, 'in'), ('s_wel', '福利厚生費', 30, 'in'), ('s_ent', '接待交際費', 31, 'in'),
    ('s_off', '事務所家賃', 32, 'in'), ('s_land', '地代家賃', 33, 'in'), ('s_utl', '水道光熱費', 34, 'in'),
    ('s_eq', '備品費', 35, 'in'), ('s_rent', '賃借料', 36, 'in'), ('s_tax', '租税公課', 37, 'in'),
    ('s_sta', '文具消耗品', 38, 'in'), ('s_news', '新聞図書費', 39, 'in'), ('s_com', '運搬通信費', 40, 'in'),
    ('s_cons', '委託費', 41, 'in'), ('s_fee', '支払手数料', 42, 'in'), ('s_bank', '振込手数料', 43, 'in'),
    ('s_ins', '保険料', 44, 'in'), ('s_edu', '教育研修費', 45, 'in'), ('s_dues', '諸会費', 46, 'in'),
    ('s_rep', '修繕費', 47, 'in'), ('s_dep', '減価償却費', 48, 'in'), ('s_dpa', '敷金償却費', 49, 'in'),
    ('s_mtg', '会議費', 50, 'in'),
    ('sga', '販売費及び一般管理費計（事業部直接費）', None, 'f'), ('op0', '事業所利益（共通費配賦前）', None, 'f'),
    ('alloc_rate', '共通費配賦率（売上高比）', 53, 'in'), ('alloc', '共通販管費（本社共通費の配賦）', 54, 'in'),
    ('op', '営業利益（共通費配賦後）', None, 'f'),
    ('x_sal', '［PL外］人件費（加算）：給与手当', 61, 'in'), ('x_hw', '［PL外］人件費（加算）：法定福利費', 62, 'in'),
    ('x_bon', '［PL外］人件費（加算）：賞与引当金', 63, 'in'), ('x_n', '［PL外］同上の人数：所長', 64, 'in'),
    ('x_tot', '［PL外］人件費（加算）合計', None, 'f'),
    ('h_all', '在籍（技術者）', 76, 'in'), ('h_reg', '　うち正社員', 77, 'in'), ('h_con', '　うち契約社員', 78, 'in'),
    ('h_op', '稼働数（BP除く）', 79, 'in'), ('h_sb', '未稼働数（スタンバイ／会社指示休業）', 80, 'in'),
    ('h_lv', '休職者数（産育休／私傷病）', 81, 'in'), ('h_in', '入社数', 84, 'in'), ('h_out', '退職数', 85, 'in'),
]
SGA_KEYS = [k for k, _, _, _ in ITEMS if k.startswith('s_')]
RR = {}
r = 6
for key, lab, srow, kind in ITEMS:
    RR[key] = r
    pm[f'B{r}'] = lab
    pm[f'B{r}'].font = Font(name=JP, size=10, bold=(kind == 'f'))
    r += 1
    if key in ('sales', 'gp', 'op', 'x_tot'):
        r += 1
FORM = {
    'sales': lambda c: f'={c}{RR["sales_d"]}+{c}{RR["sales_u"]}',
    'c_lab': lambda c: f'=SUM({c}{RR["c_sal"]}:{c}{RR["c_ret"]})',
    'cogs': lambda c: f'={c}{RR["c_lab"]}+{c}{RR["c_trv"]}',
    'gp': lambda c: f'={c}{RR["sales"]}-{c}{RR["cogs"]}',
    'sga': lambda c: f'=SUM({c}{RR["s_sal"]}:{c}{RR["s_mtg"]})',
    'op0': lambda c: f'={c}{RR["gp"]}-{c}{RR["sga"]}',
    'op': lambda c: f'={c}{RR["op0"]}-{c}{RR["alloc"]}',
    'x_tot': lambda c: f'=SUM({c}{RR["x_sal"]}:{c}{RR["x_bon"]})',
}
for key, lab, srow, kind in ITEMS:
    rr = RR[key]
    vals = mrow(srow) if srow else None
    for i, c in enumerate(MC):
        if kind == 'in':
            v = vals[i]
            pm[f'{c}{rr}'] = v
        else:
            pm[f'{c}{rr}'] = FORM[key](c)
        pm[f'{c}{rr}'].number_format = '0.0%' if key == 'alloc_rate' else '#,##0;△#,##0;"-"'
        pm[f'{c}{rr}'].font = Font(name=NUM, size=10, bold=(kind == 'f'))
    for k, (a, b, lab_) in PER.items():
        c = PC[k]
        if key == 'alloc_rate':
            pm[f'{c}{rr}'] = f'=IFERROR({c}{RR["alloc"]}/{c}{RR["sales"]},"")'
            pm[f'{c}{rr}'].number_format = '0.0%'
        elif key == 'x_n':
            pm[f'{c}{rr}'] = f'=MAX({MC[a]}{rr}:{MC[b]}{rr})'
        elif key == 'h_all' or key.startswith('h_') and key in ('h_reg', 'h_con'):
            pm[f'{c}{rr}'] = f'=SUM({MC[a]}{rr}:{MC[b]}{rr})'     # 延べ人月
            pm[f'{c}{rr}'].number_format = '#,##0'
        else:
            pm[f'{c}{rr}'] = f'=SUM({MC[a]}{rr}:{MC[b]}{rr})'
            pm[f'{c}{rr}'].number_format = '#,##0;△#,##0;"-"'
        pm[f'{c}{rr}'].font = Font(name=NUM, size=10, bold=(kind == 'f'))
r_chk = RR['h_out'] + 3
pm[f'B{r_chk}'] = '■ 検算（受領資料の通期合計との一致）'; pm[f'B{r_chk}'].font = Font(name=JP, size=10, bold=True, color=NAVY)
CHK = [('純売上高', 'sales', 7), ('売上総利益', 'gp', 21), ('販管費計', 'sga', 51), ('事業所利益', 'op0', 52), ('営業利益', 'op', 55)]
for j, (lab, key, srow_) in enumerate(CHK):
    rr = r_chk + 1 + j
    pm[f'B{rr}'] = f'受領資料の通期{lab}（FY24/FY25/FY26）'
    for k, scol in (('FY24', 14), ('FY25', 28), ('FY26', 42)):
        c = PC[k]
        pm[f'{c}{rr}'] = src.cell(srow_, scol).value
        pm[f'{c}{rr}'].number_format = '#,##0;△#,##0'
        cc = CL(C0 + 36 + 1 + 5 + list(PER).index(k))
        pm[f'{cc}{rr}'] = f'=ROUND({c}{rr}-{c}{RR[key]},0)=0'
pm[f'B{r_chk + 7}'] = ('※受領資料の「一人当たり売上」（通期）は、売上高に見込月（10・11月）を含む一方、稼働数は9月までの合計のため過大。'
                       '本マスターでは12〜9月の実績期間で再計算した（人員推移シート）。受領資料の57〜60行は#REF!（参照切れ）のため転記していない。')
pm[f'B{r_chk + 8}'] = ('※［PL外］人件費（加算）は受領資料の営業利益の計算に含まれていない欄。所長（=専任営業担当）1名分と推察。'
                       '販管費の本社給与手当等との関係（重複の有無）は要確認。')
for a in (f'B{r_chk + 7}', f'B{r_chk + 8}'):
    pm[a].font = Font(name=NOTE, size=9, color='C00000')
pm.freeze_panes = 'C6'


def P(key, per):
    """事業部PL月次の期間合計セル参照"""
    return f"{SRC}!{PC[per]}{RR[key]}"


# ======================================================================== IM→ 対象事業の概要（公式「会社概要」）
ws = wb['会社概要']
clear_sheet(ws)
cw(ws, {'B': 1.35, 'C': 5.06})
row(ws, 2, 2, ['項目', '内容'], ['hdrl', 'hdrl'], height=19)
pl_ = 'PLハイライト'
ROWS = [
    ('売主', 'ジャパニアス株式会社（東京証券取引所グロース市場上場）', 19),
    ('対象事業', '化学・バイオ事業（売主の一事業部門。子会社化されていない）', 19),
    ('事業内容', '化学・バイオ領域の研究開発・分析・検査等に従事する技術者の派遣\n（事業部PL上、売上高の全額が派遣売上高。請負売上高の計上なし）', 34),
    ('所属拠点', '大宮事業所（埼玉県さいたま市大宮区）\n※技術者は全員が同事業所所属。拠点の承継可否は要確認', 34),
    ('就業先エリア', '東京都・神奈川県・埼玉県・大阪府（2026/11期の顧客の勤務地）', 19),
    ('事業開始', '2023年11月頃（従業員一覧上の最初の技術者の入社。事業部PLは2023年12月から）', 19),
    ('従業員数', '（人員構成シートで設定）', 34),
    ('顧客数', '=顧客分析!C40', 34),
    ('許認可', '労働者派遣事業（派14-300005）※売主の許可。事業譲渡では承継されないため\n譲受会社における許可が必要', 34),
    ('売上高', f'=TEXT({pl_}!E4/1000,"#,##0")&"千円（2026年11月期見込：12〜9月実績＋10〜11月見込）"', 19),
    ('売上総利益', f'=TEXT({pl_}!E6/1000,"#,##0")&"千円（同上、売上総利益率"&TEXT({pl_}!E7,"0.0%")&"）"', 19),
    ('営業利益', f'=TEXT({pl_}!E11/1000,"#,##0")&"千円（本社共通費配賦後。配賦前の事業所利益は"&TEXT({pl_}!E8/1000,"#,##0")&"千円）"', 19),
    ('譲渡スキーム', '事業譲渡（希望譲渡対価：2,000万円）', 19),
    ('譲渡理由', '経営資源の選択と集中（不採算事業からの撤退ではない旨を売主より聴取）', 19),
]
for i, (k, v, h) in enumerate(ROWS):
    rr = 3 + i
    put(ws, rr, 2, k, 'lab', size=10.5)
    put(ws, rr, 3, v, 'txt', size=10.5)
    ws.row_dimensions[rr].height = h
mark(ws, 'B2:C16', '基本情報（対象事業の概要）', cell='A1')

# ======================================================================== IM→ 売主会社の概要
ws = new_sheet('売主概要', '売主会社の概要（ジャパニアス株式会社）', '会社概要')
cw(ws, {'B': 1.35, 'C': 5.3})
row(ws, 4, 2, ['項目', '内容'], ['hdrl', 'hdrl'], height=19)
SROWS = [
    ('会社名', 'ジャパニアス株式会社', 19),
    ('設立', '1999年12月22日', 19),
    ('本社', '神奈川県横浜市西区みなとみらい2-2-1　横浜ランドマークタワー18F', 19),
    ('上場区分', '東京証券取引所 グロース市場（証券コード：9558）、2022年上場', 19),
    ('経営理念', '先端テクノロジーで日本の明日に新たな価値を提供する', 19),
    ('事業内容', 'AI・IoT・クラウドを中心とした先端エンジニアリング事業\n（開示セグメントは「先端エンジニアリング事業」の単一セグメント）', 34),
    ('社員数', '1,851名（2025年2月末現在）', 19),
    ('事業所', '首都圏第一〜第五・大宮・宇都宮・仙台・札幌・名古屋・関西・福岡の各事業所、\n横浜開発センター（受託開発拠点）、研修施設（J-college等）', 34),
    ('認証資格', '労働者派遣事業（派14-300005）、プライバシーマーク、ISO27001', 19),
    ('主要取引先', '大手電機・精密・自動車・ITの各グループ等、約200社（売主全体。会社案内）', 19),
    ('業績（全社）', '2025年11月期：売上高 約120.8億円、営業利益 約9.8億円（公表決算）', 19),
]
for i, (k, v, h) in enumerate(SROWS):
    rr = 5 + i
    put(ws, rr, 2, k, 'lab', size=10.5); put(ws, rr, 3, v, 'txt', size=10.5)
    ws.row_dimensions[rr].height = h
mark(ws, 'B4:C15', '売主会社の概要')
note(ws, 17, '※出所：売主会社案内（化学・バイオ領域の営業資料、2025年4月）、公表決算。全社の数値であり、対象事業（化学・バイオ事業）の数値ではない。')
# 売上高推移（会社案内「経営状況」の棒グラフ読取り、百万円。2025年は公表決算）
ws['B20'] = '■ 売主全社の売上高推移（百万円）— グラフ元データ（貼付範囲外）'
ws['B20'].font = Font(name=JP, size=10, bold=True, color=NAVY)
HIST = [(2015, 3227), (2016, 3870), (2017, 4428), (2018, 5205), (2019, 6055), (2020, 6669), (2021, 7243),
        (2022, 8324), (2023, 9885), (2024, 11211), (2025, 12080)]
for i, (y, v) in enumerate(HIST):
    ws.cell(21 + i, 2, f'{y}/11期'); ws.cell(21 + i, 3, v).number_format = '#,##0'
note(ws, 33, '※2015〜2024年は会社案内の「経営状況」グラフの表示値。2025年11月期は公表決算（約120.8億円）を百万円で表示。')

# ======================================================================== IM→ PLハイライト（公式シートを事業部PL用に組み直し）
ws = wb['PLハイライト']
clear_sheet(ws)
cw(ws, {'B': 2.05, 'C': 0.92, 'D': 0.92, 'E': 0.92, 'F': 0.98, 'G': 0.98})
put(ws, 2, 7, '(単位：千円)', 'unit', border=None)
row(ws, 3, 2, ['', '2024/11期\n実績', '2025/11期\n実績', '2026/11期\n見込', '2025/12〜\n2026/9実績', '前年同期\n(24/12〜25/9)'],
    'hdr', size=9, height=30)
PCOL = {'C': 'FY24', 'D': 'FY25', 'E': 'FY26', 'F': 'CUR', 'G': 'PY'}
LINES = [  # (行, ラベル, kind, key or formula)
    (4, '売上高', 'v', 'sales'), (5, '前期比', 'r', None), (6, '売上総利益', 'v', 'gp'), (7, '売上総利益率', 'r', None),
    (8, '事業所利益（共通費配賦前）', 'v', 'op0'), (9, '利益率', 'r', None), (10, '共通販管費（本社共通費の配賦）', 'v', 'alloc'),
    (11, '営業利益（共通費配賦後）', 'v', 'op'), (12, '営業利益率', 'r', None),
]
for rr, lab, kind, key in LINES:
    if kind == 'v':
        big = key in ('sales', 'op')
        put(ws, rr, 2, lab, 'gryb' if key == 'op' else ('sub' if key == 'gp' else 'lab'), border='box' if key == 'op' else 'std')
        for c, per in PCOL.items():
            put(ws, rr, ord(c) - 64, f'={P(key, per)}', 'grybn' if key == 'op' else ('subn' if key == 'gp' else 'num'), fmt=KI,
                border='box' if key == 'op' else 'std')
        ws.row_dimensions[rr].height = 19
    else:
        put(ws, rr, 2, lab, 'rlab')
        for c in PCOL:
            if rr == 5:
                f = {'C': '-', 'D': '=D4/C4', 'E': '=E4/D4', 'F': '=F4/G4', 'G': '-'}[c]
            else:
                base = {7: 6, 9: 8, 12: 11}[rr]
                f = f'={c}{base}/{c}4'
            put(ws, rr, ord(c) - 64, f, 'rnum', fmt='0.0%;△0.0%;"-"' if rr != 5 else '0.0%')
        ws.row_dimensions[rr].height = 14
ws.row_dimensions[13].height = 8
row(ws, 14, 2, ['期末在籍技術者数（名）', '=事業部PL月次!N' + str(RR['h_all']), '=事業部PL月次!Z' + str(RR['h_all']),
                '=事業部PL月次!AJ' + str(RR['h_all']), '=事業部PL月次!AJ' + str(RR['h_all']), '=事業部PL月次!X' + str(RR['h_all'])],
    ['lab'] + ['num'] * 5, fmts=[None] + [YI] * 5, height=19)
put(ws, 14, 5, '=事業部PL月次!AJ' + str(RR['h_all']), 'num', fmt='#,##0"※"')
row(ws, 15, 2, ['延べ稼働人月', f'={P("h_op", "FY24")}', f'={P("h_op", "FY25")}', '-', f'={P("h_op", "CUR")}', f'={P("h_op", "PY")}'],
    ['lab'] + ['num'] * 5, fmts=[None] + [YI] * 5, height=19)
row(ws, 16, 2, ['稼働1人月当たり売上高', '=C4/C15', '=D4/D15', '-', '=F4/F15', '=G4/G15'], ['lab'] + ['num'] * 5, fmts=[None] + [KI] * 5, height=19)
row(ws, 17, 2, ['稼働1人月当たり売上総利益', '=C6/C15', '=D6/D15', '-', '=F6/F15', '=G6/G15'], ['lab'] + ['num'] * 5, fmts=[None] + [KI] * 5, height=19)
mark(ws, 'B2:G17', '財務ハイライト（P/L）', cell='A1')
# 参考（貼付範囲外）：リード文・Comment用
REF = [
    (20, '■ 参考（貼付範囲外）', None),
    (21, '売上高の伸び（2026/11期見込÷2024/11期）', '=E4/C4'),
    (22, '12〜9月実績の前年同期比', '=F4/G4-1'),
    (23, '直近3か月（2026/7〜9）の月平均売上高', f'=AVERAGE({SRC}!AH{RR["sales"]}:AJ{RR["sales"]})'),
    (24, '同×12（年換算の試算値）', '=C23*12'),
    (25, '共通費配賦率（2026/11期見込）', f'={SRC}!{PC["FY26"]}{RR["alloc_rate"]}'),
    (26, '［PL外］所長人件費（2026/11期見込、年額）', f'={P("x_tot", "FY26")}'),
    (27, '販管費のうち本社給与手当・法定福利費・賞与引当金（2026/11期見込）', f'={P("s_sal", "FY26")}+{P("s_hw", "FY26")}+{P("s_bon", "FY26")}'),
    (28, '採用紹介料＋広告宣伝費（2024/11期）', f'={P("s_rec", "FY24")}+{P("s_ad", "FY24")}'),
    (29, '採用紹介料＋広告宣伝費（2025/11期）', f'={P("s_rec", "FY25")}+{P("s_ad", "FY25")}'),
    (30, '採用紹介料＋広告宣伝費（2026/11期見込）', f'={P("s_rec", "FY26")}+{P("s_ad", "FY26")}'),
    (31, '売上総利益率：2025/12〜2026/9実績', '=F6/F4'),
    (32, '10・11月見込の売上高合計', f'={SRC}!AK{RR["sales"]}+{SRC}!AL{RR["sales"]}'),
    (33, '顧客別売上（2025/12〜2026/10）合計', '=取引先!J21'),
    (34, '事業部PLの2025/12〜2026/10売上高（10月は見込/実績表示）', f'=SUM({SRC}!AA{RR["sales"]}:AK{RR["sales"]})'),
    (35, '差額（顧客別売上−事業部PL）', '=C33-C34'),
    (36, '2026/11期見込の労務費（月平均）', f'={P("c_lab", "FY26")}/12'),
    (37, '2026/9単月の労務費', f'={SRC}!AJ{RR["c_lab"]}'),
]
for rr, lab, f in REF:
    ws.cell(rr, 2, lab).font = Font(name=JP, size=10, bold=(f is None), color=NAVY if f is None else '000000')
    if f:
        ws.cell(rr, 3, f).number_format = '0.0%' if ('比' in lab or '率' in lab or '伸び' in lab) else '#,##0'
red(ws, 'B39', '※2026/11期見込は2025/12〜2026/9の実績＋2026/10〜11の見込（受領資料の「見込/実績」表示）。在籍技術者数の2026/11期欄は2026/9末（※）。')
red(ws, 'B40', '※面談時の説明（年間売上高約1.5億円・売上総利益約3,400万円・営業利益約500万円・粗利率約22.7%）との差異は「その他」シート。')

# ======================================================================== IM→ 損益計算書（事業部PL、財務情報詳細）
ws = new_sheet('事業部PL', '財務情報詳細｜事業部損益計算書', 'BSハイライト')
cw(ws, {'B': 2.55, 'C': 0.95, 'D': 0.62, 'E': 0.95, 'F': 0.62, 'G': 0.95, 'H': 0.62, 'I': 0.95, 'J': 0.62})
put(ws, 3, 10, '(単位：千円)', 'unit', border=None); ws.merge_cells('I3:J3')
row(ws, 4, 2, ['科目', '2024/11期 実績', '', '2025/11期 実績', '', '2026/11期 見込', '', '2025/12〜2026/9 実績', ''], 'hdr', size=10, height=19)
row(ws, 5, 2, ['', '金額', '売上高比', '金額', '売上高比', '金額', '売上高比', '金額', '売上高比'], 'hdr', size=9, height=16)
for a, b in (('C4', 'D4'), ('E4', 'F4'), ('G4', 'H4'), ('I4', 'J4')):
    ws.merge_cells(f'{a}:{b}')
ws.merge_cells('B4:B5')
DPL = [('派遣売上高', 'sales_d', 'lab'), ('売上高', 'sales', 'sub'),
       ('技術者給与手当', 'c_sal', 'lab'), ('技術者法定福利費', 'c_hw', 'lab'), ('技術者賞与引当金繰入額', 'c_bon', 'lab'),
       ('退職給付費用（概算）', 'c_ret', 'lab'), ('【労務費】', 'c_lab', 'gry'), ('旅費交通費', 'c_trv', 'lab'),
       ('売上原価', 'cogs', 'sub'), ('売上総利益', 'gp', 'sub'),
       ('販売費及び一般管理費（事業部直接費）', 'sga', 'lab'), ('事業所利益（共通費配賦前）', 'op0', 'sub'),
       ('共通販管費（本社共通費の配賦）', 'alloc', 'lab'), ('営業利益（共通費配賦後）', 'op', 'gryb')]
PP = [('C', 'D', 'FY24'), ('E', 'F', 'FY25'), ('G', 'H', 'FY26'), ('I', 'J', 'CUR')]
for i, (lab, key, kd) in enumerate(DPL):
    rr = 6 + i
    nk = {'lab': 'num', 'sub': 'subn', 'gry': 'gryn', 'gryb': 'grybn'}[kd]
    bd = 'box' if kd == 'gryb' else 'std'
    put(ws, rr, 2, lab, kd, border=bd)
    for a, b, per in PP:
        put(ws, rr, ord(a) - 64, f'={P(key, per)}', nk, fmt=KI, border=bd)
        put(ws, rr, ord(b) - 64, f'=IFERROR({a}{rr}/{a}$7,"")', nk, fmt='0.0%;△0.0%;"-"', size=9, border=bd)
    ws.row_dimensions[rr].height = 18
mark(ws, 'B3:J19', '財務情報詳細｜損益計算書')
note(ws, 21, '※事業部PL（売主の管理会計）。対象事業単独の貸借対照表はない。減価償却費の計上なし。')

# ======================================================================== IM→ 販管費（事業部直接費）
ws = new_sheet('事業部販管費', '財務情報詳細｜販売費及び一般管理費（事業部直接費）', 'BSハイライト')
cw(ws, {'B': 2.55, 'C': 0.95, 'D': 0.62, 'E': 0.95, 'F': 0.62, 'G': 0.95, 'H': 0.62, 'I': 0.95, 'J': 0.62})
put(ws, 3, 10, '(単位：千円)', 'unit', border=None); ws.merge_cells('I3:J3')
row(ws, 4, 2, ['科目', '2024/11期 実績', '', '2025/11期 実績', '', '2026/11期 見込', '', '2025/12〜2026/9 実績', ''], 'hdr', size=10, height=19)
row(ws, 5, 2, ['', '金額', '売上高比', '金額', '売上高比', '金額', '売上高比', '金額', '売上高比'], 'hdr', size=9, height=16)
for a, b in (('C4', 'D4'), ('E4', 'F4'), ('G4', 'H4'), ('I4', 'J4')):
    ws.merge_cells(f'{a}:{b}')
ws.merge_cells('B4:B5')
SGA_SHOW = [k for k in SGA_KEYS]
lab_of = {k: l for k, l, _, _ in ITEMS}
rr = 6
hide = []
for key in SGA_SHOW:
    put(ws, rr, 2, lab_of[key], 'lab')
    for a, b, per in PP:
        put(ws, rr, ord(a) - 64, f'={P(key, per)}', 'num', fmt=KI)
        put(ws, rr, ord(b) - 64, f'=IFERROR({a}{rr}/{P("sales", per)},"")', 'num', fmt='0.0%;△0.0%;"-"', size=9)
    ws.row_dimensions[rr].height = 15
    # 3期とも0の科目は非表示（貼付しない）
    tot = sum(abs(x or 0) for x in mrow([s for k_, l_, s, _ in ITEMS if k_ == key][0]))
    if tot == 0:
        ws.row_dimensions[rr].hidden = True
        hide.append(lab_of[key])
    rr += 1
put(ws, rr, 2, '販売費及び一般管理費 合計', 'sub')
for a, b, per in PP:
    put(ws, rr, ord(a) - 64, f'=SUM({a}6:{a}{rr - 1})', 'subn', fmt=KI)
    put(ws, rr, ord(b) - 64, f'=IFERROR({a}{rr}/{P("sales", per)},"")', 'subn', fmt='0.0%;△0.0%;"-"', size=9)
ws.row_dimensions[rr].height = 18
SGA_TOT_ROW = rr
for j, (a, b, per) in enumerate(PP):
    ws.cell(rr + 1, ord(a) - 64, f'=ROUND({a}{rr}-{P("sga", per)},0)=0')
ws.cell(rr + 1, 2, 'チェック（事業部PL月次の販管費計と一致）').font = Font(name=NOTE, size=9)
mark(ws, f'B3:J{rr}', '財務情報詳細｜販売費及び一般管理費')
note(ws, rr + 3, '※3期とも計上のない科目（' + '・'.join(hide) + '）は非表示。共通販管費（本社共通費の配賦）は含まない。')

# ======================================================================== IM→ 収益構造（共通費配賦と譲渡後の見方）
ws = new_sheet('収益構造', '収益構造｜事業部PLの費用区分と譲渡後の取扱い（想定）', 'BSハイライト')
cw(ws, {'B': 0.75, 'C': 2.55, 'D': 0.82, 'E': 0.82, 'F': 0.82, 'G': 0.62, 'H': 4.6})
put(ws, 3, 8, '(単位：千円)', 'unit', border=None)
row(ws, 4, 2, ['区分', '科目', '2024/11期', '2025/11期', '2026/11期\n見込', '売上高比\n(26/11期)', '譲渡後の取扱い（当社の想定。要確認）'],
    'hdr', size=9.5, height=30)
SROW = [
    ('売上', '派遣売上高', ['sales'], 'lab', '顧客との派遣契約の承継（顧客の同意が前提）により継続'),
    ('売上原価', '技術者人件費（給与手当・法定福利費・\n賞与引当金・退職給付費用）', ['c_lab'], 'lab',
     '転籍により譲受会社が負担。給与・賞与・退職給付等の労働条件の承継内容は要確認'),
    ('', '旅費交通費', ['c_trv'], 'lab', '技術者の通勤交通費等と推察。継続して発生'),
    ('', '売上総利益', ['gp'], 'sub', ''),
    ('販管費', '本社給与手当・法定福利費・賞与引当金', ['s_sal', 's_hw', 's_bon'], 'lab',
     '事業部に直課された管理人件費。内容（下記所長人件費の配賦分か等）は要確認'),
    ('', '採用紹介料・広告宣伝費', ['s_rec', 's_ad'], 'lab', '増員ペースと譲受会社の採用手法により変動'),
    ('', '事務所家賃・その他', [k for k in SGA_KEYS if k not in ('s_sal', 's_hw', 's_bon', 's_rec', 's_ad')], 'lab',
     '大宮事業所の家賃等の配分と推察。拠点の承継有無により変動（要確認）'),
    ('', '事業所利益（共通費配賦前）', ['op0'], 'sub', ''),
    ('共通費', '共通販管費（本社共通費の配賦）', ['alloc'], 'lab',
     '売主の本社機能（管理部門等）の費用を売上高比で配賦。譲渡後は売主の配賦はなくなる一方、\n給与計算・労務・経理・営業管理等は譲受会社の体制で必要'),
    ('', '営業利益（共通費配賦後）', ['op'], 'gryb', ''),
]
for i, (kb, lab, keys, kd, txt) in enumerate(SROW):
    rr = 5 + i
    nk = {'lab': 'num', 'sub': 'subn', 'gryb': 'grybn'}[kd]
    bd = 'box' if kd == 'gryb' else 'std'
    put(ws, rr, 2, kb, 'gry' if kd == 'lab' else kd, size=9.5, border=bd)
    put(ws, rr, 3, lab, kd, size=9.5, border=bd)
    for c, per in (('D', 'FY24'), ('E', 'FY25'), ('F', 'FY26')):
        put(ws, rr, ord(c) - 64, '=' + '+'.join(P(k, per) for k in keys), nk, fmt=KI, border=bd)
    put(ws, rr, 7, f'=F{rr}/F$5', nk, fmt='0.0%;△0.0%;"-"', size=9, border=bd)
    put(ws, rr, 8, txt, 'txt', size=8.5, border=bd)
    ws.row_dimensions[rr].height = 30 if ('\n' in lab or '\n' in txt) else 21
rr = 5 + len(SROW)
ws.row_dimensions[rr].height = 8
rr += 1
put(ws, rr, 2, '参考', 'gry', size=9.5)
put(ws, rr, 3, '［PL外記載］所長（専任営業）人件費\n（給与手当・法定福利費・賞与引当金）', 'lab', size=9.5)
for c, per in (('D', 'FY24'), ('E', 'FY25'), ('F', 'FY26')):
    put(ws, rr, ord(c) - 64, f'={P("x_tot", per)}', 'num', fmt=KI)
put(ws, rr, 7, f'=F{rr}/F$5', 'num', fmt='0.0%', size=9)
put(ws, rr, 8, '事業部PLの欄外に「人件費（加算）」として記載があり、上記の営業利益には含まれていない。\n販管費の本社給与手当等との重複の有無を含め、計上範囲は要確認', 'txt', size=8.5)
ws.row_dimensions[rr].height = 30
X_ROW = rr
mark(ws, f'B3:H{rr}', '収益構造と共通費配賦')
note(ws, rr + 2, '※金額は事業部PL（月次）の合計。2026/11期見込は10・11月の見込を含む。譲渡後の取扱いは現時点の当社の想定であり、譲渡対象の範囲とともに要確認。')

# ======================================================================== 月次推移（グラフ元データ＋四半期表）
ws = new_sheet('月次推移', '月次業績・人員の推移（グラフ元データ）', 'BSハイライト')
cw(ws, {'B': 0.9, 'C': 0.55, 'D': 0.9, 'E': 0.9, 'F': 0.9, 'G': 0.9, 'H': 0.7, 'I': 0.7, 'J': 0.6, 'K': 0.6})
row(ws, 4, 2, ['年月', '区分', '売上高', '売上総利益', '事業所利益', '営業利益', '在籍', '稼働', '入社', '退職'], 'hdr', size=9, height=19)
for i, c in enumerate(MC):
    rr = 5 + i
    vals = [f'={SRC}!{c}4', f'={SRC}!{c}5'] + [f'={SRC}!{c}{RR[k]}' for k in ('sales', 'gp', 'op0', 'op')] + \
           [f'=IF({SRC}!{c}{RR[k]}="","",{SRC}!{c}{RR[k]})' for k in ('h_all', 'h_op', 'h_in', 'h_out')]
    row(ws, rr, 2, vals, ['txtc', 'txtc'] + ['num'] * 8, fmts=[None, None] + [KI] * 4 + [YI] * 4, size=9, height=15)
red(ws, 'A2', '■ グラフ元データ（B4:K40）→ 企業概要書「月次業績の推移」「人員の推移」のグラフ。四半期表 B43:G56 → 「月次業績の推移」に貼付')
# 四半期表（貼付範囲）。列幅はこの表に合わせる（上のグラフ元データは貼付しない）
cw(ws, {'B': 0.95, 'C': 0.8, 'D': 0.72, 'E': 0.8, 'F': 0.6, 'G': 0.6, 'H': 0.6, 'I': 0.6, 'J': 0.6, 'K': 0.6})
put(ws, 43, 6, '(単位：千円)', 'unit', border=None); ws.merge_cells('F43:G43')
row(ws, 44, 2, ['四半期', '売上高', '前年同期比', '売上総利益', '粗利率', '期末在籍'], 'hdr', size=9, height=19)
Q = []
for fy, base in (('24/11期', 0), ('25/11期', 12), ('26/11期', 24)):
    for q in range(4):
        Q.append((f'{fy} {q + 1}Q', base + q * 3))
for i, (lab, m0) in enumerate(Q):
    rr = 45 + i
    c1, c3 = MC[m0], MC[m0 + 2]
    put(ws, rr, 2, lab + ('*' if m0 + 2 >= 34 else ''), 'lab', size=9)
    put(ws, rr, 3, f'=SUM({SRC}!{c1}{RR["sales"]}:{c3}{RR["sales"]})', 'num', fmt=KI, size=9)
    put(ws, rr, 4, f'=IFERROR(C{rr}/C{rr - 4},"-")' if i >= 4 else '-', 'num', fmt='0.0%;△0.0%;"-"', size=9)
    put(ws, rr, 5, f'=SUM({SRC}!{c1}{RR["gp"]}:{c3}{RR["gp"]})', 'num', fmt=KI, size=9)
    put(ws, rr, 6, f'=E{rr}/C{rr}', 'num', fmt='0.0%;△0.0%', size=9)
    put(ws, rr, 7, f'=IF({SRC}!{c3}{RR["h_all"]}="","-",{SRC}!{c3}{RR["h_all"]})', 'num', fmt=YI, size=9)
    ws.row_dimensions[rr].height = 15
mark_q = paste(ws, 'B43:G56', '月次業績の推移（四半期表）', REG, cell='A3')
red(ws, 'A3', '■ 貼付範囲 B43:G56 → 企業概要書「月次業績の推移」（四半期表）')

# ======================================================================== 顧客（公式「取引先」）
ws = wb['取引先']
clear_sheet(ws)
# 顧客の匿名化（記号は2026/11期 12〜10月の売上順。取引終了先はP〜T）
CUS = OrderedDict()
blocks = [('FY24', 3, 14), ('FY25', 17, 34), ('FY26', 38, 52)]
for k, a, b in blocks:
    for r_ in range(a, b + 1):
        nm = cus.cell(r_, 2).value
        CUS.setdefault(nm, {})[k] = cus.cell(r_, 3).value
        if k == 'FY26':
            CUS[nm]['loc'] = cus.cell(r_, 5).value
ORDER26 = [cus.cell(r_, 2).value for r_ in range(38, 53)]
ENDED = [n for n in CUS if n not in ORDER26]
ENDED.sort(key=lambda n: -(CUS[n].get('FY25', 0) + CUS[n].get('FY24', 0)))
CODE = {}
for i, n in enumerate(ORDER26 + ENDED):
    CODE[n] = chr(ord('A') + i) + '社'
IND = {  # 業種（公開情報等による当社の推察）、業種区分
    '理研計器株式会社': ('ガス検知・計測機器メーカー', '精密・計測・光学機器'),
    '富士フイルムビジネスエキスパート株式会社': ('大手メーカーグループの機能会社', 'メーカー系グループ会社'),
    'タカノ株式会社': ('メーカー（業種は要確認）', 'その他・要確認'),
    '日東富士製粉株式会社': ('製粉・食品メーカー', '食品'),
    'トーヨーケム株式会社': ('化学メーカー（樹脂・接着剤等）', '化学・素材・エネルギー'),
    'HOYA株式会社': ('光学ガラス・精密機器メーカー', '精密・計測・光学機器'),
    '鹿島建設株式会社': ('総合建設会社', '建設'),
    '東曹産業株式会社': ('化学メーカーのグループ会社', '化学・素材・エネルギー'),
    '武蔵塗料株式会社': ('塗料メーカー', '化学・素材・エネルギー'),
    '東芝環境ソリューション株式会社': ('環境分析・環境関連サービス', '試験・分析受託'),
    '株式会社消費科学研究所': ('試験・検査機関', '試験・分析受託'),
    'インターステム株式会社': ('バイオ関連企業', 'バイオ・医療'),
    'コスモ石油株式会社': ('石油精製・元売', '化学・素材・エネルギー'),
    '株式会社フカガワ': ('業種は要確認', 'その他・要確認'),
    '株式会社三菱ケミカルリサーチ': ('化学系の調査・分析受託', '試験・分析受託'),
    '株式会社デイラボ': ('業種は要確認', 'その他・要確認'),
    '二光光学株式会社': ('光学部品メーカー', '精密・計測・光学機器'),
    '株式会社セブンケミカル': ('化学関連（業種は要確認）', '化学・素材・エネルギー'),
    '株式会社ニコン・セル・イノベーション': ('細胞医薬品等の受託製造', 'バイオ・医療'),
    '株式会社ファスマック': ('遺伝子解析等の受託検査', '試験・分析受託'),
}
GROUPS = ['化学・素材・エネルギー', '精密・計測・光学機器', '試験・分析受託', '食品', 'バイオ・医療', '建設', 'メーカー系グループ会社', 'その他・要確認']
# 従業員の就業先 → 顧客記号
EMP_CLIENT = {}
for r_ in range(3, 36):
    cl = emp.cell(r_, 15).value
    EMP_CLIENT.setdefault(cl, 0)
    EMP_CLIENT[cl] += 1


def pref(s):
    if not s:
        return '－'
    for t in ('都', '道', '府', '県'):
        i = s.find(t)
        if 0 < i <= 4:
            return s[:i + 1]
    return s


cw(ws, {'B': 0.3, 'C': 0.62, 'D': 2.2, 'E': 0.75, 'F': 0.82, 'G': 0.6, 'H': 0.82, 'I': 0.6, 'J': 0.82, 'K': 0.6, 'L': 0.55, 'M': 1.7})
# 列：B№ C顧客 D業種 E勤務地 F/G FY24 H/I FY25 J/K FY26 L配属人数 M備考（N:業種区分 O:実名 P/Q/R:有無フラグ は貼付範囲外）
put(ws, 2, 13, '(単位：千円)', 'unit', border=None)
row(ws, 3, 2, ['№', '顧客', '業種（推察）', '勤務地', '2024/11期', '', '2025/11期', '', '2026/11期（12〜10月）', '', '配属\n人数', '備考'],
    'hdr', size=9, height=19)
row(ws, 4, 2, ['', '', '', '', '売上高', '構成比', '売上高', '構成比', '売上高', '構成比', '', ''], 'hdr', size=9, height=15)
for a, b in (('F3', 'G3'), ('H3', 'I3'), ('J3', 'K3'), ('B3', 'B4'), ('C3', 'C4'), ('D3', 'D4'), ('E3', 'E4'), ('L3', 'L4'), ('M3', 'M4')):
    ws.merge_cells(f'{a}:{b}')
TOT = 21
lines = ORDER26 + ENDED      # 5〜19：FY26取引先15社、22〜26：取引終了先（範囲外の明細）
for i, n in enumerate(ORDER26):
    rr = 5 + i
    d = CUS[n]
    first = 'FY24' if 'FY24' in d else ('FY25' if 'FY25' in d else 'FY26')
    nplace = EMP_CLIENT.get(n, 0)
    remark = {'FY24': '2024/11期から取引', 'FY25': '2025/11期から取引', 'FY26': '2026/11期から取引'}[first]
    if nplace == 0:
        remark = remark.replace('取引', '・10月末配属なし')
    vals = [i + 1, CODE[n], IND[n][0], pref(d.get('loc')), d.get('FY24'), f'=IFERROR(F{rr}/F${TOT},"")', d.get('FY25'),
            f'=IFERROR(H{rr}/H${TOT},"")', d.get('FY26'), f'=IFERROR(J{rr}/J${TOT},"")', nplace if nplace else '－', remark]
    row(ws, rr, 2, vals, ['txtc', 'txt', 'txt', 'txt', 'num', 'num', 'num', 'num', 'num', 'num', 'txtc', 'txt'],
        fmts=[None, None, None, None, KI, PI, KI, PI, KI, PI, YI, None], size=9.5, height=17)
    ws.cell(rr, 14, IND[n][1]); ws.cell(rr, 15, n)
ED0 = 24
put(ws, 20, 2, '', 'gry', size=9.5)
put(ws, 20, 3, f'取引終了先（{len(ENDED)}社）', 'gry', size=9.5)
put(ws, 20, 4, '化学・精密・バイオ等（明細は顧客分析）', 'gry', size=8.5)
put(ws, 20, 5, '－', 'gry', size=9.5)
for c in 'FHJ':
    put(ws, 20, ord(c) - 64, f'=SUM({c}{ED0}:{c}{ED0 + len(ENDED) - 1})', 'gryn', fmt=KI, size=9.5)
for c, b in (('G', 'F'), ('I', 'H'), ('K', 'J')):
    put(ws, 20, ord(c) - 64, f'=IFERROR({b}20/{b}${TOT},"")', 'gryn', fmt=PI, size=9.5)
put(ws, 20, 12, '－', 'gry', size=9.5); put(ws, 20, 13, '2026/11期は取引なし', 'gry', size=9)
ws.row_dimensions[20].height = 17
put(ws, TOT, 2, '', 'sub'); put(ws, TOT, 3, '合計', 'sub'); put(ws, TOT, 4, '', 'sub'); put(ws, TOT, 5, '', 'sub')
for c in 'FHJ':
    put(ws, TOT, ord(c) - 64, f'=SUM({c}5:{c}20)', 'subn', fmt=KI, size=9.5)
for c in 'GIK':
    put(ws, TOT, ord(c) - 64, 1, 'subn', fmt=PI, size=9.5)
put(ws, TOT, 12, '=SUM(L5:L19)', 'subn', fmt=YI, size=9.5); put(ws, TOT, 13, '', 'sub')
ws.row_dimensions[TOT].height = 17
mark(ws, 'B2:M21', '主要顧客の3期推移', cell='A1')
ws.cell(ED0 - 1, 2, '■ 取引終了先の明細（貼付範囲外。合計は20行に集計）').font = Font(name=JP, size=10, bold=True, color=NAVY)
for i, n in enumerate(ENDED):
    rr = ED0 + i
    d = CUS[n]
    for j, v in enumerate([CODE[n], IND[n][0], '－', d.get('FY24'), None, d.get('FY25'), None, d.get('FY26')]):
        if v is not None:
            ws.cell(rr, 3 + j, v).number_format = '#,##0,' if isinstance(v, (int, float)) else 'General'
    ws.cell(rr, 14, IND[n][1]); ws.cell(rr, 15, n)
for rr in list(range(5, 20)) + list(range(ED0, ED0 + len(ENDED))):
    for src_c, dst_c in (('F', 'P'), ('H', 'Q'), ('J', 'R')):
        ws[f'{dst_c}{rr}'] = f'=IF({src_c}{rr}="","",{src_c}{rr})'
ws['P4'] = 'FY24(集計用)'; ws['Q4'] = 'FY25(集計用)'; ws['R4'] = 'FY26(集計用)'
ws.cell(30, 2, '■ 検算（貼付範囲外）').font = Font(name=JP, size=10, bold=True, color=NAVY)
ws.cell(31, 2, '事業部PLの売上高'); ws.cell(31, 6, f'={P("sales", "FY24")}'); ws.cell(31, 8, f'={P("sales", "FY25")}')
ws.cell(31, 10, f'=SUM({SRC}!AA{RR["sales"]}:AK{RR["sales"]})')
ws.cell(32, 2, '差額（顧客別合計−事業部PL）')
for c in 'FHJ':
    ws[f'{c}31'].number_format = '#,##0'
    ws[f'{c}32'] = f'={c}21-{c}31'; ws[f'{c}32'].number_format = '#,##0;△#,##0'
ws.cell(33, 2, '配属人数の合計（休職1名・B社グループ会社1名を除く）'); ws.cell(33, 12, '=L21')
note(ws, 35, '※N列＝業種区分（顧客分析の集計キー）、O列＝実名（社内用。貼付範囲外・概要書には載せない）。')
for c in ('N', 'O'):
    ws.column_dimensions[c].width = 26

# ======================================================================== 顧客分析（業種構成・集中度・継続性）
ws = new_sheet('顧客分析', '顧客基盤の分析（業種構成・集中度・取引継続）', '仕入先')
cw(ws, {'B': 1.95, 'C': 0.85, 'D': 0.6, 'E': 0.85, 'F': 0.6, 'G': 0.85, 'H': 0.6, 'I': 0.62, 'J': 0.62})
put(ws, 3, 9, '(単位：千円)', 'unit', border=None); ws.merge_cells('I3:J3')
row(ws, 4, 2, ['業種区分（推察）', '2024/11期', '', '2025/11期', '', '2026/11期（12〜10月）', '', '社数\n(26/11期)', '配属\n人数'], 'hdr', size=9, height=19)
row(ws, 5, 2, ['', '売上高', '構成比', '売上高', '構成比', '売上高', '構成比', '', ''], 'hdr', size=9, height=15)
for a, b in (('C4', 'D4'), ('E4', 'F4'), ('G4', 'H4'), ('B4', 'B5'), ('I4', 'I5'), ('J4', 'J5')):
    ws.merge_cells(f'{a}:{b}')
TR = '取引先'
for i, g in enumerate(GROUPS):
    rr = 6 + i
    put(ws, rr, 2, g, 'lab', size=9.5)
    for c, sc in (('C', 'F'), ('E', 'H'), ('G', 'J')):
        put(ws, rr, ord(c) - 64, f'=SUMIF({TR}!$N$5:$N$28,$B{rr},{TR}!{sc}$5:{sc}$28)', 'num', fmt=KI, size=9.5)
        nc_ = chr(ord(c) + 1)
        put(ws, rr, ord(nc_) - 64, f'=IFERROR({c}{rr}/{c}$14,"")', 'num', fmt=PI, size=9.5)
    put(ws, rr, 9, f'=COUNTIFS({TR}!$N$5:$N$19,$B{rr},{TR}!$J$5:$J$19,">0")', 'num', fmt=YI, size=9.5)
    put(ws, rr, 10, f'=SUMIF({TR}!$N$5:$N$19,$B{rr},{TR}!$L$5:$L$19)', 'num', fmt=YI, size=9.5)
    ws.row_dimensions[rr].height = 17
put(ws, 14, 2, '合計', 'sub', size=9.5)
for c in 'CEG':
    put(ws, 14, ord(c) - 64, f'=SUM({c}6:{c}13)', 'subn', fmt=KI, size=9.5)
for c in 'DFH':
    put(ws, 14, ord(c) - 64, f'=SUM({c}6:{c}13)', 'subn', fmt=PI, size=9.5)
put(ws, 14, 9, '=SUM(I6:I13)', 'subn', fmt=YI, size=9.5); put(ws, 14, 10, '=SUM(J6:J13)', 'subn', fmt=YI, size=9.5)
ws.row_dimensions[14].height = 17
mark(ws, 'B3:J14', '顧客基盤の分析（業種構成）')
# 集中度・継続（別範囲。縦に離して置く）
row(ws, 18, 2, ['指標', '2024/11期', '2025/11期', '2026/11期\n(12〜10月)'], 'hdr', size=9, height=28)
FL = {'FY24': 'P', 'FY25': 'Q', 'FY26': 'R'}
TRR = f'{TR}!$F$5:$F$28'


def rng(per):
    return f'{TR}!${FL[per]}$5:${FL[per]}$28'


IND2 = [
    ('取引社数（社）', lambda p: f'=COUNTIF({rng(p)},">0")', YI),
    ('上位1社の売上構成比', lambda p: f'=LARGE({rng(p)},1)/SUM({rng(p)})', PI),
    ('上位3社の売上構成比', lambda p: f'=SUM(LARGE({rng(p)},{{1,2,3}}))/SUM({rng(p)})', PI),
    ('上位5社の売上構成比', lambda p: f'=SUM(LARGE({rng(p)},{{1,2,3,4,5}}))/SUM({rng(p)})', PI),
    ('前期から継続して取引（社）', None, YI),
    ('当期に新規取引（社）', None, YI),
    ('前期取引があり当期取引なし（社）', None, YI),
    ('前期からの継続顧客の売上構成比', None, PI),
]
prev = {'FY24': None, 'FY25': 'FY24', 'FY26': 'FY25'}
for i, (lab, fn, fm) in enumerate(IND2):
    rr = 19 + i
    put(ws, rr, 2, lab, 'lab', size=9.5)
    for j, p in enumerate(('FY24', 'FY25', 'FY26')):
        c = 'CDE'[j]
        pv = prev[p]
        if fn:
            f = fn(p)
        elif pv is None:
            f = '-'
        elif i == 4:
            f = f'=COUNTIFS({rng(p)},">0",{rng(pv)},">0")'
        elif i == 5:
            f = f'=COUNTIFS({rng(p)},">0",{rng(pv)},"")'
        elif i == 6:
            f = f'=COUNTIFS({rng(pv)},">0",{rng(p)},"")'
        else:
            f = f'=SUMIFS({rng(p)},{rng(pv)},">0")/SUM({rng(p)})'
        put(ws, rr, 3 + j, f, 'num', fmt=fm, size=9.5)
    ws.row_dimensions[rr].height = 17
mark2 = paste(ws, 'B18:E26', '顧客基盤の分析（集中度・継続）', REG, cell='A16')
note(ws, 27, '※構成比の分母は顧客別売上の合計。2024/11期以前（事業開始前）の取引はないため、2024/11期の継続・新規は「－」。')
# 取引終了先の明細（別範囲。列幅を確保するためL〜O列に置く）
cw(ws, {'L': 0.8, 'M': 2.0, 'N': 0.85, 'O': 0.85})
row(ws, 30, 12, ['取引終了先', '業種（推察）', '2024/11期', '2025/11期'], 'hdr', size=9, height=17)
for i, n in enumerate(ENDED):
    rr = 31 + i
    d = CUS[n]
    row(ws, rr, 12, [CODE[n], IND[n][0], d.get('FY24'), d.get('FY25')], ['lab', 'txt', 'num', 'num'], fmts=[None, None, KI, KI], size=9.5, height=16)
paste(ws, f'L30:O{30 + len(ENDED)}', '顧客基盤の分析（取引終了先）', REG, cell='A29')
# 参考（リード文・Comment・会社概要用）
ws['B38'] = '■ 参考（貼付範囲外）'; ws['B38'].font = Font(name=JP, size=10, bold=True, color=NAVY)
ws['B39'] = '2026年10月末に技術者の配属がある顧客数（B社グループ会社を除く）'; ws['C39'] = f'=COUNTIF({TR}!L5:L19,">0")'
ws['B40'] = '会社概要「顧客数」欄'
ws['C40'] = (f'=TEXT(E19,"0")&"社（2025/12〜2026/10の取引実績）。うち2026年10月末に技術者が\n配属されている顧客は"&TEXT(C39,"0")&"社（ほか1社のグループ会社へ1名配属）"')
ws['B41'] = '2026/11期の上位3社'; ws['C41'] = '=E21'
ws['B42'] = '2024/11期最大顧客（取引終了）の構成比'; ws['C42'] = f'=MAX({TR}!F24:F28)/{TR}!F21'
ws['B43'] = '2025/11期から2026/11期へ継続した顧客の売上構成比'; ws['C43'] = '=E26'

# ======================================================================== 従業員（公式「従業員」）
ws = wb['従業員']
clear_sheet(ws)
EDU = {  # 学歴区分, 専攻（学校名は載せない）, 専攻系統
    104642: ('大学院卒', '応用化学', '化学系'), 105652: ('大卒', '畜産学', '生物・生命科学系'),
    105503: ('大卒', '応用生命科学', '生物・生命科学系'), 105394: ('大卒', '応用生物科学', '生物・生命科学系'),
    105398: ('大学院卒', '医科学', '医・薬・栄養系'), 105536: ('大卒', 'アニマルサイエンス', '生物・生命科学系'),
    105917: ('大卒', '化学・生物化学', '化学系'), 106094: ('大卒', '環境科学', '生物・生命科学系'),
    105032: ('大卒', '環境応用化学', '化学系'), 105685: ('大卒', '応用化学', '化学系'),
    105053: ('大卒', 'バイオサイエンス', '生物・生命科学系'), 104792: ('大学院卒', '医学', '医・薬・栄養系'),
    105097: ('大卒', '化学', '化学系'), 105545: ('大卒', '応用化学', '化学系'),
    105164: ('専門卒', 'バイオテクノロジー', '生物・生命科学系'), 105356: ('大学院卒', '総合人間自然科学', 'その他理系'),
    105557: ('高卒', '化学工業', '化学系'), 105216: ('大卒', '応用生物科学', '生物・生命科学系'),
    105339: ('大学院卒', '薬学', '医・薬・栄養系'), 105619: ('大卒', '管理栄養士養成', '医・薬・栄養系'),
    105908: ('大卒', '生物資源学', '生物・生命科学系'), 106153: ('専門卒', '応用生物科学', '生物・生命科学系'),
    105367: ('大卒', '環境科学', '生物・生命科学系'), 105079: ('大卒', '応用化学', '化学系'),
    105098: ('大卒', 'バイオサイエンス', '生物・生命科学系'), 105258: ('大卒', 'アニマルサイエンス', '生物・生命科学系'),
    105482: ('大卒', '薬科学', '医・薬・栄養系'), 105485: ('専門卒', '応用化学', '化学系'),
    106085: ('大学中退', '物質材料工学', '化学系'), 106105: ('短大卒', '保育', '理系以外'),
    106139: ('大卒', '健康栄養', '医・薬・栄養系'), 106167: ('専門卒', 'バイオテクノロジー', '生物・生命科学系'),
    106006: ('大学院卒', '応用理科学', '化学系'), 200184: ('大卒', '経済学', '理系以外'),
}
QUAL = {'毒物劇物取扱者試験': '毒物劇物取扱者', '有機溶剤作業主任者': '有機溶剤作業主任者',
        '特定化学物質及び四アルキル鉛等作業主任者': '特定化学物質作業主任者', '危険物取扱者試験乙種4類': '危険物乙種4類',
        '品質管理検定（QC検定）3級': 'QC検定3級', '上級バイオ技術者認定試験': '上級バイオ技術者',
        '中級バイオ技術者認定試験': '中級バイオ技術者'}
staff = []
for r_ in list(range(3, 36)) + [39]:
    g = lambda c: emp.cell(r_, c).value
    q = g(16) if r_ != 39 else '-'
    ql = [] if (q in (None, '-')) else [QUAL.get(x.strip(), x.strip()) for x in str(q).replace('_x000D_', '').split('\n') if x.strip()]
    if '中級バイオ技術者' in ql and '上級バイオ技術者' in ql:
        ql.remove('中級バイオ技術者')
    cl = g(15) if r_ != 39 else None
    code = {None: '－'}.get(cl) if cl is None else (
        CODE.get(cl) or ('休職中' if '休職' in cl else ('B社グループ会社' if '富士フイルムビジネスイノベーション' in cl else '要確認')))
    lc = '－' if r_ == 39 else (pref(CUS.get(cl, {}).get('loc')) if cl in CUS else ('－' if '休職' in cl else '要確認'))
    staff.append(dict(no=g(1), ini=g(2), typ=g(3), sex=g(4), age=g(5), join=g(6), base=g(7), bonus=g(8),
                      addr=g(9), edu=EDU[g(1)], site=g(11), tech=(g(12) if r_ != 39 else '－'), ind=(g(13) if r_ != 39 else '－'),
                      role=(g(14) if r_ != 39 else '営業（所長）'), client=cl, code=code, loc=lc, qual='、'.join(ql) if ql else '－',
                      nq=len(ql), sales=(r_ == 39)))
tech = [s for s in staff if not s['sales']]
ordkey = lambda s: (s['code'] if s['code'] not in ('休職中', 'B社グループ会社') else ('Z' + s['code']), s['join'])
tech.sort(key=ordkey)
roster = [s for s in staff if s['sales']] + tech
cw(ws, {'B': 0.3, 'C': 0.3, 'D': 0.3, 'E': 0.75, 'F': 0.35, 'G': 0.3, 'H': 0.4, 'I': 0.62, 'J': 0.45, 'K': 0.48,
        'L': 0.82, 'M': 1.05, 'N': 1.55, 'O': 0.9, 'P': 0.62, 'Q': 2.0})
ws['F2'] = '基準日'; ws['G2'] = datetime.datetime(2026, 10, 31); ws['G2'].number_format = 'yyyy/mm/dd'
row(ws, 4, 2, ['№', '氏名', 'イニシャル', '雇用形態', '性別', '生年月日', '年齢', '入社年月', '勤続\n(年)', '技術', '役割',
               '業界', '最終学歴（専攻）', '就業先', '勤務地', '保有資格'], 'hdr', size=9, height=26)
for i, s in enumerate(roster):
    rr = 5 + i
    typ = {'正社員': '正社員', '有期契約社員': '有期契約'}[s['typ']]
    vals = [i + 1, '', '', typ, s['sex'], '', s['age'], s['join'].strftime('%Y/%m'), f'=IF(AA{rr}="","",DATEDIF(AA{rr},$G$2,"m")/12)',
            s['tech'], s['role'].replace('/', '・'), s['ind'], f'{s["edu"][0]}（{s["edu"][1]}）', s['code'], s['loc'], s['qual']]
    row(ws, rr, 2, vals, ['txtc', 'txt', 'txt', 'txtc', 'txtc', 'txt', 'num', 'txtc', 'num', 'txtc', 'txt', 'txt', 'txt', 'txtc', 'txtc', 'txt'],
        fmts=[None, None, None, None, None, None, YI, 'yyyy/mm', '0.0', None, None, None, None, None, None, None], size=9, height=17)
    # 貼付範囲外：集計用の生データ（社内用）
    for j, v in enumerate([s['base'], s['bonus'], pref(s['addr']), s['edu'][0], s['edu'][2], s['nq'], s['ini'], s['client']]):
        ws.cell(rr, 19 + j, v)
    ws.cell(rr, 19).number_format = '#,##0'
    ws.cell(rr, 27, s['join']).number_format = 'yyyy/mm/dd'   # AA：入社年月日（勤続の計算用）
EMP_LAST = 5 + len(roster) - 1
for c in ('S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'):
    ws.column_dimensions[c].width = 14
for j, h in enumerate(['基本給（円）', '2026年5月賞与', '居住都道府県', '学歴区分', '専攻系統', '資格数', 'イニシャル', '就業先（実名）']):
    ws.cell(4, 19 + j, h).font = Font(name=JP, size=9, bold=True, color='C00000')
ws.cell(3, 19, '■ 集計用の生データ（社内用・貼付範囲外。概要書には個人別の給与・氏名・就業先実名を載せない）').font = Font(name=NOTE, size=9, color='C00000')
hide_cols(ws, 'C', 'D', 'G')
mark(ws, f'B4:Q{EMP_LAST}', '従業員一覧（①B4:Q21、②B4:Q4＋B22:Q' + str(EMP_LAST) + '）', cell='A1')
note(ws, EMP_LAST + 2, '※出所：従業員一覧（2026年10月末在籍）。氏名・学校名・個人別の給与は掲載しない。勤務地は就業先の所在都県（顧客別売上資料の勤務地）。')

# ======================================================================== 人員構成（集計）
ws = new_sheet('人員構成', '人員構成（技術者33名の集計）', '拠点(本社等)')
cw(ws, {'B': 0.95, 'C': 1.75, 'D': 0.6, 'E': 0.65})
E_ = f'従業員!$#$6:$#${EMP_LAST}'


def er(col):
    return E_.replace('#', col)


def cnt(*conds):
    return '=COUNTIFS(' + ','.join(f'{er(c)},"{v}"' for c, v in conds) + ')'


def block(ws, r0, title_, items, n_ref):
    row(ws, r0, 2, ['区分', '項目', '人数', '構成比'], 'hdr', size=9, height=17)
    rr = r0 + 1
    for grp, its in items:
        first = rr
        for lab, f in its:
            put(ws, rr, 2, grp if rr == first else '', 'gry', size=9)
            put(ws, rr, 3, lab, 'lab', size=9)
            put(ws, rr, 4, f, 'num', fmt='#,##0"名"', size=9)
            put(ws, rr, 5, f'=IFERROR(D{rr}/{n_ref},"")', 'num', fmt='0%', size=9)
            ws.row_dimensions[rr].height = 14.5
            rr += 1
        if rr - first > 1:
            ws.merge_cells(f'B{first}:B{rr - 1}')
    return rr - 1


ws['G4'] = '技術者数'; ws['H4'] = f'=COUNTA(従業員!$B$6:$B${EMP_LAST})'
NREF = '$H$4'
left = [
    ('雇用形態', [('正社員（無期雇用）', cnt(('E', '正社員'))), ('有期契約社員', cnt(('E', '有期契約')))]),
    ('性別', [('男性', cnt(('F', '男'))), ('女性', cnt(('F', '女')))]),
    ('年齢', [('20代', f'=COUNTIFS({er("H")},">=20",{er("H")},"<30")'), ('30代', f'=COUNTIFS({er("H")},">=30",{er("H")},"<40")'),
            ('40代', f'=COUNTIFS({er("H")},">=40",{er("H")},"<50")'), ('50代以上', f'=COUNTIFS({er("H")},">=50")')]),
    ('勤続', [('1年未満', f'=COUNTIFS({er("J")},"<1")'), ('1年以上2年未満', f'=COUNTIFS({er("J")},">=1",{er("J")},"<2")'),
            ('2年以上', f'=COUNTIFS({er("J")},">=2")')]),
    ('学歴', [('大学院卒', cnt(('V', '大学院卒'))), ('大卒', cnt(('V', '大卒'))),
            ('専門・短大・高卒等', f'=H4-D16-D17')]),
    ('専攻系統', [('化学系', cnt(('W', '化学系'))), ('生物・生命科学系', cnt(('W', '生物・生命科学系'))),
              ('医・薬・栄養系', cnt(('W', '医・薬・栄養系'))), ('その他理系', cnt(('W', 'その他理系'))), ('理系以外', cnt(('W', '理系以外')))]),
]
endL = block(ws, 4, '属性', left, NREF)
mark(ws, f'B4:E{endL}', '人員構成（属性）')
r1 = endL + 4
right = [
    ('技術区分', [('化学', cnt(('K', '化学'))), ('バイオ', cnt(('K', 'バイオ')))]),
    ('役割', [('研究開発', cnt(('L', '研究開発'))), ('分析補助', cnt(('L', '分析補助'))), ('検査・評価', cnt(('L', '検査・評価'))),
            ('検査', cnt(('L', '検査')))]),
    ('業界', [('化学関連', cnt(('M', '化学関連'))), ('製薬・医薬品関連', cnt(('M', '製薬・医薬品関連'))), ('食品関連', cnt(('M', '食品関連')))]),
    ('勤務地', [('東京都', cnt(('P', '東京都'))), ('神奈川県', cnt(('P', '神奈川県'))), ('埼玉県', cnt(('P', '埼玉県'))),
             ('大阪府', cnt(('P', '大阪府'))), ('休職中・要確認', f'=H4-SUM(D{r1 + 10}:D{r1 + 13})')]),
    ('保有資格', [('資格保有者（実人数）', f'=COUNTIFS({er("X")},">0")'),
              ('有機溶剤作業主任者', f'=COUNTIFS({er("Q")},"*有機溶剤*")'),
              ('特定化学物質作業主任者', f'=COUNTIFS({er("Q")},"*特定化学物質*")'),
              ('危険物取扱者乙種4類', f'=COUNTIFS({er("Q")},"*危険物*")'),
              ('毒物劇物取扱者', f'=COUNTIFS({er("Q")},"*毒物劇物*")'),
              ('その他（QC検定・バイオ技術者）', f'=COUNTIFS({er("Q")},"*QC*")+COUNTIFS({er("Q")},"*バイオ技術者*")')]),
]
endR = block(ws, r1, '専門性', right, NREF)
mark_r = paste(ws, f'B{r1}:E{endR}', '人員構成（専門性・勤務地・資格）', REG, cell=f'A{r1 - 1}')
red(ws, f'A{r1 - 1}', f'■ 貼付範囲 B{r1}:E{endR} → 企業概要書「人員構成」（右表）')
# 参考（リード文・会社概要用）
R = endR + 3
ws[f'B{R}'] = '■ 参考（貼付範囲外）'; ws[f'B{R}'].font = Font(name=JP, size=10, bold=True, color=NAVY)
REFS = [
    ('平均年齢（技術者）', f'=AVERAGE({er("H")})', '0.0'),
    ('平均勤続（技術者、年）', f'=AVERAGE({er("J")})', '0.0'),
    ('20代比率', f'=D{6 + 3}/H4', '0%'),
    ('正社員（技術者）平均基本給（円/月）', f'=AVERAGEIFS({er("S")},{er("E")},"正社員")', '#,##0'),
    ('正社員（技術者）基本給 最低', f'=_xlfn.MINIFS({er("S")},{er("E")},"正社員")', '#,##0'),
    ('正社員（技術者）基本給 最高', f'=_xlfn.MAXIFS({er("S")},{er("E")},"正社員")', '#,##0'),
    ('2026年5月賞与 支給者数', f'=COUNT({er("T")})', '0'),
    ('2026年5月賞与 支給者平均（円）', f'=AVERAGE({er("T")})', '#,##0'),
    ('有期契約社員の時給（円）', f'=SUMIFS({er("S")},{er("E")},"有期契約")', '#,##0'),
    ('居住地：東京都', cnt(('U', '東京都')), '0'), ('居住地：神奈川県', cnt(('U', '神奈川県')), '0'),
    ('居住地：埼玉県', cnt(('U', '埼玉県')), '0'), ('居住地：千葉県', cnt(('U', '千葉県')), '0'), ('居住地：群馬県', cnt(('U', '群馬県')), '0'),
    ('理系の学歴（専攻系統が理系以外以外）', f'=H4-D{endL}', '0'),
    ('稼働中（顧客へ配属）', f'=H4-COUNTIFS({er("O")},"休職中")', '0'),
    ('休職中', cnt(('O', '休職中')), '0'),
    ('会社概要「従業員数」欄', '="34名（技術者33名：正社員"&TEXT(D5,"0")&"名・有期契約社員"&TEXT(D6,"0")&"名、営業1名）\n2026年10月末。技術者の平均年齢"&TEXT(AVERAGE(' + er("H") + '),"0.0")&"歳"', None),
]
for i, (lab, f, fm) in enumerate(REFS):
    rr = R + 1 + i
    ws[f'B{rr}'] = lab; ws[f'C{rr}'] = f
    if fm:
        ws[f'C{rr}'].number_format = fm
REF_EMP = {lab: f'C{R + 1 + i}' for i, (lab, f, fm) in enumerate(REFS)}
# 会社概要が参照するセル（C40）を固定位置へ
json.dump(REF_EMP, open(f'{W}/work/ref_emp.json', 'w'), ensure_ascii=False)
wb['会社概要']['C9'] = '=人員構成!' + REF_EMP['会社概要「従業員数」欄']

# ======================================================================== 人員推移（年度別の稼働・採用・定着・採算）
ws = new_sheet('人員推移', '人員の推移と1人当たり採算', '拠点(本社等)')
cw(ws, {'B': 2.75, 'C': 0.95, 'D': 0.95, 'E': 1.05})
put(ws, 3, 5, '(単位：千円)', 'unit', border=None)
row(ws, 4, 2, ['指標', '2024/11期', '2025/11期', '2025/12〜\n2026/9'], 'hdr', size=9, height=28)
PER3 = [('C', 'FY24', 'N'), ('D', 'FY25', 'Z'), ('E', 'CUR', 'AJ')]
INDS = [
    ('期末在籍技術者数（名）', lambda c, p, m: f'={SRC}!{m}{RR["h_all"]}', YI),
    ('延べ在籍人月', lambda c, p, m: f'={P("h_all", p)}', YI),
    ('延べ稼働人月', lambda c, p, m: f'={P("h_op", p)}', YI),
    ('稼働率（稼働人月÷在籍人月）', lambda c, p, m: f'={c}7/{c}6', PI),
    ('未稼働（スタンバイ等）の延べ人月', lambda c, p, m: f'={P("h_sb", p)}', YI),
    ('休職者の延べ人月', lambda c, p, m: f'={P("h_lv", p)}', YI),
    ('入社数（名）', lambda c, p, m: f'={P("h_in", p)}', YI),
    ('退職数（名）', lambda c, p, m: f'={P("h_out", p)}', YI),
    ('退職率（退職数÷（期首在籍＋入社数））', None, PI),
    ('稼働1人月当たり売上高', lambda c, p, m: f'={P("sales", p)}/{c}7', KI),
    ('稼働1人月当たり労務費・旅費（売上原価）', lambda c, p, m: f'={P("cogs", p)}/{c}7', KI),
    ('稼働1人月当たり売上総利益', lambda c, p, m: f'={P("gp", p)}/{c}7', KI),
    ('採用関連費（採用紹介料＋広告宣伝費）', lambda c, p, m: f'={P("s_rec", p)}+{P("s_ad", p)}', KI),
    ('入社1名当たり採用関連費', lambda c, p, m: f'={c}17/{c}11', KI),
]
first_in = {'C': f'{SRC}!C{RR["h_all"]}', 'D': f'{SRC}!N{RR["h_all"]}', 'E': f'{SRC}!Z{RR["h_all"]}'}
for i, (lab, fn, fm) in enumerate(INDS):
    rr = 5 + i
    put(ws, rr, 2, lab, 'lab', size=9.5)
    for c, p, m in PER3:
        if fn:
            f = fn(c, p, m)
        else:
            # 期首在籍：2024/11期は期首（2023/12）の在籍3名を期首とみなす
            st = f'{SRC}!C{RR["h_all"]}' if c == 'C' else first_in[c]
            f = f'={c}12/({st}+{c}11)'
        put(ws, rr, ord(c) - 64, f, 'num', fmt=fm, size=9.5)
    ws.row_dimensions[rr].height = 17
mark(ws, 'B3:E18', '人員の推移と1人当たり採算')
note(ws, 20, '※事業部PLの人員欄（月末在籍・稼働・未稼働・休職・入退社）による。2026/11期は人員欄が9月までのため2025/12〜2026/9で表示。'
             '2024/11期の期首在籍は2023/12の在籍数。')

# ======================================================================== 譲渡対象の範囲
ws = new_sheet('譲渡対象', '譲渡対象の範囲（現時点の想定）', 'その他')
cw(ws, {'B': 1.0, 'C': 2.6, 'D': 1.25, 'E': 6.0})
row(ws, 4, 2, ['区分', '項目', '取扱い', '内容・確認事項'], 'hdr', size=9.5, height=19)
TGT = [
    ('人員', '技術者33名・営業担当1名（計34名）', '承継対象（想定）', '転籍による承継を想定。従業員の個別同意が前提。給与・賞与・退職給付・有給休暇残日数・勤続年数の通算等の承継条件は要確認'),
    ('', '事業部長・専務取締役（売主）', '対象外の可能性', '現在の事業運営に関与。譲渡後の関与（引継ぎ期間・顧客への挨拶等）は要確認'),
    ('顧客契約', '顧客との労働者派遣基本契約・個別契約', '承継対象（想定）', '契約上の地位の移転・再締結には顧客の同意が必要。譲渡制限条項・契約期間・更新条件は契約書で要確認'),
    ('許認可', '労働者派遣事業許可（派14-300005）', '承継されない', '許可は法人単位。譲受会社が自社の許可で受け入れる（派遣元責任者の選任・事業所の届出等を含む）'),
    ('拠点', '大宮事業所（埼玉県さいたま市大宮区）', '要確認', '売主の他事業と共用している可能性。賃借契約の承継・転貸・移転の要否は要確認'),
    ('設備・備品', '事業用の設備・備品', '要確認', '事業部PLに減価償却費の計上はない。PC・携帯電話等の貸与品の扱いは要確認'),
    ('運転資金', '売掛金・未払給与・預り金等', '対象外（想定）', '譲渡日前の売上の回収・給与の支払は売主で精算する想定。譲受会社は譲渡日以降の運転資金を手当て'),
    ('採用基盤', '求人媒体・スカウト媒体のアカウント、応募者情報、紹介会社との契約', '要確認', '個人情報の移転手続を含め、承継の可否・範囲は要確認'),
    ('商号・ブランド', '「ジャパニアス」「JNA」の商号・商標', '対象外（想定）', '譲渡後は譲受会社の名称で取引'),
    ('研修', '売主の研修施設・研修プログラム（J-college等）', '要確認', '対象事業の技術者が利用しているかを含め要確認'),
    ('データ', '勤怠・給与・顧客管理等のデータ', '要確認', '移行範囲・方法は要確認'),
]
for i, (a, b, c, d) in enumerate(TGT):
    rr = 5 + i
    put(ws, rr, 2, a, 'gry', size=9.5); put(ws, rr, 3, b, 'lab', size=9.5); put(ws, rr, 4, c, 'txtc', size=9.5)
    put(ws, rr, 5, d, 'txt', size=9)
    ws.row_dimensions[rr].height = 30
mark(ws, f'B4:E{4 + len(TGT)}', '譲渡対象の範囲')

# ======================================================================== 事業譲渡の実務上の確認事項
ws = new_sheet('確認事項', '事業譲渡に伴う実務上の確認事項', 'その他')
cw(ws, {'B': 1.0, 'C': 2.45, 'D': 7.4})
row(ws, 4, 2, ['分野', '項目', '内容'], 'hdr', size=9.5, height=19)
CHKS = [
    ('従業員', '転籍の同意', '34名それぞれの個別同意が必要。労働条件（給与水準・賞与・退職給付・福利厚生）の提示内容が同意の前提になる'),
    ('', '無期雇用派遣の維持', '技術者32名は正社員（無期雇用）で、派遣可能期間の制限の対象外。譲受会社で無期雇用として受け入れれば同様の取扱い（譲受会社の雇用形態による）'),
    ('', '有期契約社員（1名）', '個人単位の派遣可能期間（同一組織で3年）の制限が適用。契約期間・更新の扱いを確認'),
    ('', '賞与・退職給付の精算', '事業部PLは技術者賞与・退職給付を毎月引当（概算）。譲渡日時点の賞与計算期間・退職給付の精算方法は売主と協議（要確認）'),
    ('顧客', '契約の承継・再締結', '15社との基本契約・個別契約の承継または再締結。派遣先への派遣元変更の通知、派遣先管理台帳等の更新'),
    ('', '事業所単位の期間制限', '有期雇用の派遣労働者を受け入れる派遣先は抵触日の通知が必要（無期雇用の派遣労働者は対象外）'),
    ('許認可', '譲受会社の派遣許可', '譲受会社が労働者派遣事業の許可を保有していること。派遣元責任者・事業所の要件を確認'),
    ('資金', '運転資金', '派遣料は月末締め翌月末払い（会社案内の取引条件）。給与の支払が入金に先行するため、労務費の約1か月分を目安に運転資金が必要と想定'),
    ('', '消費税', '譲渡対価（営業権等）に係る消費税の取扱いは要確認'),
]
for i, (a, b, c) in enumerate(CHKS):
    rr = 5 + i
    put(ws, rr, 2, a, 'gry', size=9.5); put(ws, rr, 3, b, 'lab', size=9.5); put(ws, rr, 4, c, 'txt', size=9)
    ws.row_dimensions[rr].height = 30
mark(ws, f'B4:D{4 + len(CHKS)}', '事業譲渡に伴う確認事項')
note(ws, 16, '※労務費（2026/11期見込の月平均）はPLハイライトの参考欄C36。法令上の取扱いは譲受会社の顧問社労士・弁護士の確認を前提とする。')

# ======================================================================== 取引条件と人材モデル（会社案内）
ws = new_sheet('取引条件', '取引条件と人材モデル（化学・バイオ領域の営業資料より）', 'その他')
cw(ws, {'B': 1.55, 'C': 4.0, 'D': 3.6})
row(ws, 4, 2, ['項目', '内容（会社案内「お取引に関わる条件面」等）', '事業部PL・従業員一覧との照合'], 'hdr', size=9.5, height=19)
TERMS = [
    ('契約形態', '労働者派遣契約（準委任・請負は別途相談）\n労働者派遣基本契約書＋個別契約書', '事業部PLの売上は全額が派遣売上高（請負売上高0）'),
    ('料金', '時間給制（残業は1分単位で請求）。月給制も相談可', '顧客別・技術者別の請求単価は未受領（要確認）'),
    ('請求単価の水準', '時間給2,000円台後半〜3,000円台後半（競合比較表の記載）', '稼働1人月当たり売上高は約40万円（事業部PL）'),
    ('支払条件', '月末締め翌月末払い', '－'),
    ('契約期間', '1か月以上。終了の場合は45日前に連絡', '個別契約の期間・更新頻度は要確認'),
    ('人材の雇用形態', '正社員（無期雇用）が基本。派遣期間に制限がなく長期就業が可能', '技術者33名中32名が正社員（従業員一覧）'),
    ('人材の要件', '理系の学歴（工・理・農・医・薬学部、理工系専門学校等）で、\n実験器具・溶剤の基礎知識を有する人材', '専攻系統が理系の技術者は人員構成シート参照'),
    ('採用方法', '営業担当が面接に入り、性格診断データ等で見極め。\nスカウトサービスで個別にスカウトした人材から選考', '採用紹介料・広告宣伝費を事業部PLに計上'),
    ('年齢層', '20代中盤〜40代前半が中心', '技術者の平均年齢は人員構成シート参照'),
]
for i, (a, b, c) in enumerate(TERMS):
    rr = 5 + i
    put(ws, rr, 2, a, 'lab', size=9.5); put(ws, rr, 3, b, 'txt', size=9); put(ws, rr, 4, c, 'txt', size=9)
    ws.row_dimensions[rr].height = 30 if '\n' in b else 20
mark(ws, f'B4:D{4 + len(TERMS)}', '取引条件と人材モデル')
note(ws, 15, '※会社案内は化学・バイオ領域の営業資料（2025年4月、大宮事業所）。記載の条件は標準的な取引条件であり、個別の契約内容は契約書で要確認。')

# ======================================================================== 業務領域（売主の対応分野と対象事業の実績）
ws = new_sheet('業務領域', '業務領域｜売主の対応分野と対象事業の実績', 'その他')
cw(ws, {'B': 1.25, 'C': 3.9, 'D': 4.0, 'E': 1.2})
row(ws, 4, 2, ['領域', '売主の対応分野（公式サイト・会社案内）', '対象事業で確認できた実績（従業員一覧・顧客別売上）', '確認状況'], 'hdr', size=9, height=26)
AREAS = [
    ('開発', '素材・化合物の研究開発、有機合成、化学合成・試薬調製、\n試作・評価、既存製品の改良、プロセス改善',
     '="役割「研究開発」"&TEXT(人員構成!D' + str(r1 + 3) + ',"0")&"名（化学メーカー・計測機器メーカー・食品メーカー・\n大手メーカーグループ会社等へ配属）"', '実績あり'),
    ('分析', '化学分析・機器分析・手分析、品質管理、理化学検査\n（HPLC・GC・ICP・SEM・FT-IR等）',
     '="役割「分析補助」"&TEXT(人員構成!D' + str(r1 + 4) + ',"0")&"名・「検査・評価」"&TEXT(人員構成!D' + str(r1 + 5) + ',"0")&"名\n（環境分析・試験検査機関・化学・建設等）。使用機器の内訳は要確認"', '実績あり\n（機器は要確認）'),
    ('動物実験', '薬効薬理試験、解剖、経口投与・静脈注射・採血等', '従業員一覧からは確認できない', '要確認'),
    ('細胞実験', '細胞培養、DNA・タンパク質精製、免疫染色、ELISA、\nウエスタンブロッティング、遺伝子発現・DNA/RNA解析',
     '="技術区分「バイオ」"&TEXT(人員構成!D' + str(r1 + 2) + ',"0")&"名（役割はいずれも「検査」。バイオ関連企業・\n試験検査機関等へ配属）。細胞実験への該当は要確認"', '一部要確認'),
]
for i, (a, b, c, d) in enumerate(AREAS):
    rr = 5 + i
    put(ws, rr, 2, a, 'lab', size=9.5); put(ws, rr, 3, b, 'txt', size=8.5); put(ws, rr, 4, c, 'txt', size=8.5)
    put(ws, rr, 5, d, 'txtc', size=8.5)
    ws.row_dimensions[rr].height = 32
mark(ws, 'B4:E8', '業務領域（売主の対応分野と対象事業の実績）')
note(ws, 10, '※売主の対応分野は公式サイト（化学・バイオ領域）・会社案内による売主全体の記載であり、対象事業の実績を示すものではない。')
row(ws, 13, 2, ['業務例', '業務内容', '使用機器・手法', '出所'], 'hdr', size=9, height=19)
CASES = [
    ('大手化学メーカー\n粘着テープの試作・評価', '粘着樹脂・添加剤・有機溶剤の秤量・調合、塗工・乾燥・裁断、\n引張試験等による性能評価、試験データの取りまとめ', '引張試験、GPC、IR、顕微鏡'),
    ('大手医薬品メーカー\n消毒剤の品質管理（機器分析）', '液体・粉体原料の定量試験・分析、サンプル前処理\n（秤量・抽出・希釈・ろ過・焼成）、試験結果の入力・報告', 'HPLC、滴定装置、GC、UV、FT-IR'),
    ('大手素材メーカー\n燃料電池の触媒合成', '電極・電解質等の機器分析・物性評価、粉体・液体の秤量・混合・\n加熱・乾燥・焼成、データ処理', '各種分析機器'),
    ('バイオベンチャー\n新規抗体の取得と評価', '細胞ライブラリの作成、フローサイトメトリーでの解析、\n培養上清を用いたELISAスクリーニング、試薬調整・報告書作成', 'フローサイトメトリー、ELISA'),
]
for i, (a_, b_, c_) in enumerate(CASES):
    rr = 14 + i
    put(ws, rr, 2, a_, 'lab', size=8.5); put(ws, rr, 3, b_, 'txt', size=8.5); put(ws, rr, 4, c_, 'txt', size=8.5)
    put(ws, rr, 5, '会社案内\n(2025/4)', 'txtc', size=8)
    ws.row_dimensions[rr].height = 32
paste(ws, 'B13:E17', '業務領域（営業資料の業務例）', REG, cell='A12')
note(ws, 19, '※会社案内（化学・バイオ領域の営業資料）に「あくまで一例」として掲載された業務例。対象事業の現在の顧客・技術者の業務との対応は要確認。')

# ======================================================================== 沿革
ws = new_sheet('沿革', '沿革（売主・対象事業）', 'その他')
cw(ws, {'B': 1.45, 'C': 1.0, 'D': 8.2})
row(ws, 4, 2, ['年月', '区分', '内容'], 'hdr', size=9.5, height=19)
HIS = [
    ('1999年12月', '売主', 'ジャパニアス株式会社を横浜にて設立'),
    ('2019年', '売主', '大宮営業所（現 大宮事業所）を開設'),
    ('2022年', '売主', '東京証券取引所グロース市場に上場'),
    ('2023年11月', '対象事業', '化学・バイオ事業の技術者の入社（従業員一覧上の最初の入社。事業の立上げ時期は要確認）'),
    ('2023年12月', '対象事業', '事業部PLの計上開始（在籍技術者3名）'),
    ('2024年11月期', '対象事業', '="売上高"&TEXT(PLハイライト!C4/1000,"#,##0")&"千円、期末在籍技術者"&TEXT(PLハイライト!C14,"0")&"名、顧客"&TEXT(顧客分析!C19,"0")&"社"'),
    ('2025年4月', '対象事業', '化学・バイオ領域の営業資料（会社案内）を作成'),
    ('2025年11月期', '対象事業', '="売上高"&TEXT(PLハイライト!D4/1000,"#,##0")&"千円（前期比"&TEXT(PLハイライト!D5,"0.0倍")&"）、期末在籍技術者"&TEXT(PLハイライト!D14,"0")&"名、顧客"&TEXT(顧客分析!D19,"0")&"社"'),
    ('2025年11月期', '売主', '売上高 約120.8億円、営業利益 約9.8億円（全社、公表決算）'),
    ('2026年10月', '対象事業', '在籍技術者33名（従業員一覧）'),
    ('2026年11月期', '対象事業', '="売上高"&TEXT(PLハイライト!E4/1000,"#,##0")&"千円（見込）、売上総利益"&TEXT(PLハイライト!E6/1000,"#,##0")&"千円（見込）"'),
]
for i, (a, b, c) in enumerate(HIS):
    rr = 5 + i
    put(ws, rr, 2, a, 'lab', size=9.5); put(ws, rr, 3, b, 'txtc', size=9.5); put(ws, rr, 4, c, 'txt', size=9.5)
    ws.row_dimensions[rr].height = 19
ws['D12'].value = '="売上高"&TEXT(PLハイライト!D4/1000,"#,##0")&"千円（前期比"&TEXT(PLハイライト!D4/PLハイライト!C4,"0.0")&"倍）、期末在籍技術者"&TEXT(PLハイライト!D14,"0")&"名、顧客"&TEXT(顧客分析!D19,"0")&"社"'
mark(ws, f'B4:D{4 + len(HIS)}', '沿革')
note(ws, 17, '※出所：会社案内（経営状況）、事業部PL、従業員一覧、公表決算。顧客数は顧客別売上資料の取引社数。')

# ======================================================================== 参考｜譲渡対価と人材・顧客基盤の対比
ws = new_sheet('対価対比', '参考｜譲渡希望対価と人材・顧客基盤・採算の対比', 'その他')
cw(ws, {'B': 3.0, 'C': 1.2, 'D': 5.6})
row(ws, 4, 2, ['指標', '数値', '算出方法・出所'], 'hdr', size=9.5, height=19)
EMPS = '人員構成'
VAL = [
    ('希望譲渡対価', 20000000, '譲渡希望条件', KI + '"千円"' if False else '#,##0,"千円"'),
    ('承継予定人員（技術者＋営業）', 34, '従業員一覧（2026年10月末）。転籍は個別同意が前提', '#,##0"名"'),
    ('1名当たり対価', '=C5/C6', '希望譲渡対価÷承継予定人員', '#,##0,"千円"'),
    ('採用関連費（採用紹介料＋広告宣伝費）3期累計', '=人員推移!C17+人員推移!D17+人員推移!E17', '事業部PL（2024/11期〜2026/9）。広告宣伝費の内容は要確認', '#,##0,"千円"'),
    ('同期間の入社数', '=人員推移!C11+人員推移!D11+人員推移!E11', '事業部PLの人員欄', '#,##0"名"'),
    ('入社1名当たり採用関連費', '=C8/C9', '採用関連費÷入社数（退職者の採用費を含む）', '#,##0,"千円"'),
    ('売上総利益（2026/11期見込）', '=PLハイライト!E6', '事業部PL', '#,##0,"千円"'),
    ('希望譲渡対価÷売上総利益', '=C5/C11', '', '0.00"倍"'),
    ('事業所利益（共通費配賦前、2026/11期見込）', '=PLハイライト!E8', '事業部PL。譲受会社側の管理コストは控除していない', '#,##0,"千円"'),
    ('希望譲渡対価÷事業所利益', '=C5/C13', '', '0.00"倍"'),
]
for i, (a, b, c, fm) in enumerate(VAL):
    rr = 5 + i
    kd = 'sub' if i in (2, 5) else 'lab'
    put(ws, rr, 2, a, kd, size=9.5); put(ws, rr, 3, b, 'subn' if kd == 'sub' else 'num', fmt=fm, size=9.5)
    put(ws, rr, 4, c, 'txt', size=8.5)
    ws.row_dimensions[rr].height = 19
mark(ws, f'B4:D{4 + len(VAL)}', '参考｜譲渡希望対価の位置づけ')
note(ws, 16, '※当社算定の参考指標。採用関連費は採用活動の費用であり、稼働中の顧客契約・就業実績の価値を含まない。転籍同意の状況により承継人数は変動しうる。')

# ======================================================================== 要確認事項・ネクストステップ（買手向け）
ws = new_sheet('要確認事項', '要確認事項・ネクストステップ（買手向け。社内限の情報は「プロセス関連」シートに隔離）', 'その他')
cw(ws, {'B': 1.2, 'C': 8.6, 'D': 0.8})
row(ws, 4, 2, ['分野', '確認事項', '優先度'], 'hdr', size=9.5, height=19)
TODO = [
    ('収益性', '共通販管費（本社共通費の配賦、売上高比約12%）の内訳と配賦基準。譲渡後に不要となる費用・譲受会社で必要となる費用の切り分け', '高'),
    ('収益性', '事業部PL欄外の所長人件費（年約8.6百万円）と販管費の本社給与手当等（同約4.5百万円）の関係（重複・配賦の有無）', '高'),
    ('収益性', '2026年10・11月の見込値の実績化（顧客別売上の10月実績は事業部PLの見込を約3.5百万円下回る）', '高'),
    ('人員', '転籍の条件（給与・賞与・退職給付・有給残日数・勤続通算）と、従業員への説明・同意取得の進め方', '高'),
    ('人員', '事業部長・専務取締役の現在の役割（顧客対応・採用・技術者フォロー）と、譲渡後の引継ぎ期間', '高'),
    ('人員', '専任営業担当（所長）の業務範囲（顧客開拓・契約・面接・技術者フォロー・勤怠管理）と継続意向', '高'),
    ('顧客', '顧客別の請求単価・契約期間・更新頻度、契約書の譲渡制限条項、派遣元変更への顧客の意向', '高'),
    ('顧客', '取引終了先（5社）の終了理由、10月末に配属のない顧客（3社）の今後の見込み', '中'),
    ('顧客', 'B社グループ会社への配属（有期契約社員1名）の売上計上先', '中'),
    ('業務', '対象事業での動物実験・細胞実験の実績の有無、使用機器（HPLC・GC等）の経験の内訳', '中'),
    ('譲渡対象', '大宮事業所の承継有無、採用媒体アカウント・応募者情報の承継、貸与品・データの移行範囲', '中'),
    ('採用', '採用経路別の入社実績（スカウト・紹介会社・求人広告）と広告宣伝費の内容', '中'),
]
for i, (a, b, c) in enumerate(TODO):
    rr = 5 + i
    put(ws, rr, 2, a, 'lab', size=9.5); put(ws, rr, 3, b, 'txt', size=9); put(ws, rr, 4, c, 'txtc', size=9.5)
    ws.row_dimensions[rr].height = 21
mark(ws, f'B4:D{4 + len(TODO)}', '要確認事項・ネクストステップ')

# ======================================================================== その他（面談情報との差異）・プロセス関連（社内確認事項）
ws = new_sheet('プロセス関連', 'プロセス関連（社内確認事項）— 本項は社内確認事項であり、買手候補への開示資料には記載しない', 'その他')
ws['A1'].font = Font(name=JP, size=14, bold=True, color='C00000')
cw(ws, {'B': 2.6, 'C': 1.4, 'D': 1.4, 'E': 1.0, 'F': 4.5})
ws['B3'] = '■ 面談時の説明と事業部PLの差異（資料間照合）'; ws['B3'].font = Font(name=JP, size=11, bold=True, color=NAVY)
row(ws, 4, 2, ['項目', '面談時の説明', '事業部PL（2026/11期見込）', '差異', '備考'], 'hdr', size=9.5, height=19)
DIFF = [
    ('年間売上高', 150000000, '=PLハイライト!E4', '面談値は概数。PLは10・11月の見込を含む。直近3か月平均×12の試算は約150百万円'),
    ('売上総利益', 34000000, '=PLハイライト!E6', '概ね一致'),
    ('営業利益', 5000000, '=PLハイライト!E11', 'PLの営業利益は共通費配賦後。所長人件費（PL外）は含まない'),
    ('売上総利益率', 0.227, '=PLハイライト!E7', '面談値は1.5億円で割った率と推察'),
    ('営業利益率', 0.033, '=PLハイライト!E12', '同上'),
    ('従業員数', 34, 34, '一致（技術者33名＋営業1名、2026年10月末）'),
]
for i, (a, b, c, d) in enumerate(DIFF):
    rr = 5 + i
    pc = i in (3, 4)
    put(ws, rr, 2, a, 'lab', size=9.5)
    put(ws, rr, 3, b, 'num', fmt=PI if pc else ('#,##0"名"' if i == 5 else KI), size=9.5)
    put(ws, rr, 4, c, 'num', fmt=PI if pc else ('#,##0"名"' if i == 5 else KI), size=9.5)
    put(ws, rr, 5, f'=D{rr}-C{rr}', 'num', fmt=('0.0%;△0.0%' if pc else ('#,##0;△#,##0' if i == 5 else '#,##0,;△#,##0,')), size=9.5)
    put(ws, rr, 6, d, 'txt', size=9)
    ws.row_dimensions[rr].height = 19
ws['B13'] = '■ プロセス関連（社内確認事項）'; ws['B13'].font = Font(name=JP, size=11, bold=True, color='C00000')
INTERNAL = [
    '売主は早期の譲渡を希望している。',
    '別のM&A仲介会社も本件に関与しており、既に概要書を作成して候補先にアプローチしている（規模感の不一致等で見送りが多いとのこと）。',
    '売主の実質的な許容条件：仲介手数料等を踏まえて追加の持ち出しが発生しない条件であれば許容可能との意向。',
    'トップ面談は松島専務が対応する想定。',
    '面談メモ「IM：アドレイ送付、来週中に回答予定」— 現在の進捗は要確認。',
    'アドバイザリー契約は非専任を前提に検討。',
    '当社の買手候補ロングリスト約150社を作成済み（WDB、テクノプロ・R&D、ワールドインテック、アドバンテック、パーソルテンプスタッフ等）。',
    '概要書の顧客名は匿名（A社〜T社）。実名との対応は「取引先」シートO列。NDA締結後の実名開示の範囲は売主と要協議。',
    '公式サイト（jna.co.jp）は作成環境から閲覧できなかったため、サイトの記載内容は依頼者の要約に基づく。2025年11月期の全社業績も依頼者提示値（公表資料で要照合）。',
]
for i, t in enumerate(INTERNAL):
    ws.cell(14 + i, 2, f'・{t}').font = Font(name=JP, size=10)

# ======================================================================== ルール：貼付範囲一覧
rule = wb['ルール']
write_paste_index(rule, REG, start_row=17)
rule['B16'] = ('■ 本件（事業譲渡・事業部門）の運用：公式シートのうち 会社概要／PLハイライト／取引先／従業員 を事業部門用に組み直して使用。'
               'BS・株主・取締役・仕入先・外注先・拠点・不動産・設備・VR→（評価）は対象事業単独の資料がないため未使用。財務の正本は「事業部PL月次」。')
rule['B16'].font = Font(name=JP, size=10, bold=True, color='C00000')
os.makedirs(f'{W}/work', exist_ok=True)
json.dump({'REG': REG, 'RR': RR, 'PC': PC, 'EMP_LAST': EMP_LAST, 'r1': r1, 'endL': endL, 'endR': endR,
           'SGA_TOT_ROW': SGA_TOT_ROW, 'X_ROW': X_ROW, 'ENDED': len(ENDED), 'CODE': CODE},
          open(f'{W}/work/meta.json', 'w'), ensure_ascii=False, indent=1)
wb.save(f'{W}/work/master_raw.xlsx')
print('saved', len(REG), 'ranges')
for x in REG:
    print(x)
