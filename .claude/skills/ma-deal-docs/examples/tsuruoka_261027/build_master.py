import os
"""鶴岡中央青果：公式フォーマット（M&A承継機構 案件マスター 0925版）へ全データを転記する。
入力：up5/b.xlsx（ユーザー記入途中の公式マスター）、旧マスター（261019、値）、追加資料の集計（bunrui.json）
出力：master_raw.xlsx（数式のみ。LibreOfficeで再計算→inject_cacheで値を書き戻す）
"""
import sys, json, copy, datetime
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts'))
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
from master_im_style import im_sheet, cw, put, row, paste, note, KI, PI, YI, AI, JP, NUM, NAVY

S = os.environ.get('MA_WORK', '.')   # 作業フォルダ（up5/=受領ファイル、master/・im/=出力）
OLD = os.environ.get('REPO', '.') + '/株式会社鶴岡中央青果/案件マスター_鶴岡中央青果_261019.xlsx'
wb = load_workbook(f'{S}/up5/b.xlsx')
old = load_workbook(OLD, data_only=True)
REG = []   # 貼付範囲一覧


# ---------------------------------------------------------------- helpers
def cs(ws, src, dst):
    """セル書式のコピー（同一シート内／別シート可：src, dstはセル）"""
    dst.font = copy.copy(src.font); dst.fill = copy.copy(src.fill); dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment); dst.number_format = src.number_format


def row_style(ws, src_r, dst_r, c1, c2, height=True):
    for c in range(c1, c2 + 1):
        cs(ws, ws.cell(src_r, c), ws.cell(dst_r, c))
    if height:
        ws.row_dimensions[dst_r].height = ws.row_dimensions[src_r].height


def clear(ws, r1, r2, c1, c2, style=False):
    for m in list(ws.merged_cells.ranges):   # 範囲にかかる結合は解除してから消す
        if m.min_row <= r2 and m.max_row >= r1 and m.min_col <= c2 and m.max_col >= c1:
            ws.unmerge_cells(str(m))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(r, c)
            x.value = None
            if style:
                x.font = Font(name=JP, size=11); x.fill = PatternFill(); x.border = Border()
                x.alignment = Alignment(); x.number_format = 'General'


def ungroup_cols(ws, upto=40):
    """テンプレートの列幅定義は C:G のように範囲でまとまっていることがある。
    1列だけ非表示にすると範囲全体が消えるため、列ごとの定義に分解する。"""
    from openpyxl.worksheet.dimensions import ColumnDimension
    for key, dim in list(ws.column_dimensions.items()):
        if dim.min and dim.max and dim.max > dim.min and dim.min <= upto:
            w, hid = dim.width, dim.hidden
            lo, hi = dim.min, min(dim.max, upto)
            del ws.column_dimensions[key]
            for i in range(lo, hi + 1):
                ws.column_dimensions[L(i)] = ColumnDimension(ws, index=L(i), width=w, hidden=hid)


def hide_cols(ws, *cols):
    ungroup_cols(ws)
    for c in cols:
        ws.column_dimensions[c].hidden = True


def hide_rows(ws, *rows):
    for r in rows:
        ws.row_dimensions[r].hidden = True


def mark(ws, rng, page, cell='A1'):
    return paste(ws, rng, page, REG, cell=cell)


def note_red(ws, cell, t):
    ws[cell] = t; ws[cell].font = Font(name='游ゴシック Medium', size=9, color='C00000')


def copy_sheet(src, dst):
    """旧マスターのシートを値・書式ごと複製（数式は持たない参考シート用）。"""
    for r in src.iter_rows():
        for c in r:
            d = dst.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font = copy.copy(c.font); d.fill = copy.copy(c.fill); d.border = copy.copy(c.border)
                d.alignment = copy.copy(c.alignment); d.number_format = c.number_format
    for k, dim in src.column_dimensions.items():
        dst.column_dimensions[k].width = dim.width
    for k, dim in src.row_dimensions.items():
        dst.row_dimensions[k].height = dim.height
    for m in src.merged_cells.ranges:
        dst.merge_cells(str(m))
    dst.sheet_view.showGridLines = False


TB = '27年3月期\n4〜9月TB'

# ================================================================ 財務入力→
bsd = wb['BS(借方)']; bsc = wb['BS (貸方)']; pl = wb['PL']; sga = wb['SGA']
bsd['G4'] = TB
bsd['G2'] = '進行期'
G = {6: 393410846, 7: 22835432, 8: 1137638, 9: 14405198, 10: None, 11: 3045753, 12: None, 13: 101324, 14: None, 15: -203000,
     22: 92065694, 23: 11399836, 24: 284395, 25: 2, 26: 15, 27: 18546, 28: 93137240, 29: -3654000,
     36: 2155595, 37: 371800, 47: 1600000, 48: 75000, 49: 74500, 50: 33302569, 51: None}
for r, v in G.items():
    bsd.cell(r, 7).value = v
GC = {6: 18258090, 7: 3564068, 8: 6109017, 9: 327720, 10: None, 11: 3023000, 19: 526264000, 20: 8980000,
      36: 40000000, 44: 2080000, 45: 56958488}
for r, v in GC.items():
    bsc.cell(r, 7).value = v
# 決算書合計との照合（範囲外）
bsd['B65'] = '決算書・試算表の資産合計'; bsd['B66'] = 'チェック'
for c, v in zip('DEFG', [651947559, 567528964, 485190104, 665564383]):
    bsd[f'{c}65'] = v; bsd[f'{c}65'].number_format = '#,##0'
    bsd[f'{c}66'] = f'=EXACT({c}61,{c}65)'
note_red(bsd, 'B68', '※進行期（G列）は2026年9月末の合計残高試算表（4〜9月の6か月累計）。2026年8〜9月に長期借入金が約1.65億円増加（資金使途は要確認）。')
hide_cols(bsd, 'C'); hide_cols(bsc, 'C')
hide_rows(bsd, 16, 17, 18, 19, 30, 31, 32, 33, 34, 38, 39, 40, 41, 42, 43, 44, 45, 52, 53, 54, 55, 58, 59, 60)
hide_rows(bsc, 12, 13, 14, 15, 16, 17, 21, 22, 23, 24, 37, 38, 39, 40, 41, 42, 47, 48, 49)
for ws_ in (bsd, bsc):
    for r in range(26, 80):
        pass

# --- PL 進行期（2026/4〜9 試算表）
PLK = {5: 568871022, 6: 5369372, 7: 9473775, 8: 18416492, 9: 19958506, 10: -757740,
       12: 14596695, 13: 514967990, 14: -14405198,
       21: 161697, 22: 30950, 23: 58794, 24: 1184146, 25: 163350, 26: 1231436,
       29: 2090167, 30: None, 31: None}
for r, v in PLK.items():
    pl.cell(r, 11).value = v
pl['K58'] = 21165907
pl['H2'] = None; pl['L2'] = '(単位：千円)'; cs(pl, sga['H2'], pl['L2']); pl['L2'].alignment = Alignment(horizontal='right')
hide_cols(pl, 'C', 'D')
pl.column_dimensions['B'].width = 36
for a_, b_ in (('E3', 'F3'), ('G3', 'H3'), ('I3', 'J3'), ('K3', 'L3')):   # 期の見出しを2列に結合（貼付時に折り返さない）
    pl.merge_cells(f'{a_}:{b_}'); sga.merge_cells(f'{a_}:{b_}')
hide_rows(pl, 15, 27, 32, 33, 34, 35, 40, 41, 42, 43, 48, 49, 50, 53, 54, 55)
note_red(pl, 'B61', '※進行期（K列）は2026年4〜9月の試算表（6か月累計）。前年同期（2025年4〜9月）はPLハイライトの参考欄。決算書では売上高を「商品売上高」1本で表示（内訳は試算表のみ）。')
note_red(pl, 'B62', '※受託販売は2019年10月以降、受託手数料のみの純額表示（個別注記表）。2026/3期の前期損益修正損20,900千円は過年度保険解約訂正額（非経常）。')

# --- SGA：2026/3期（I列）の賞与〜広告宣伝費が1行ずれていたのを修正（決算書の販管費内訳と照合）
for r, v in {8: None, 9: 3026000, 10: 746960, 11: 13295870, 12: 460243, 13: 89300}.items():
    sga.cell(r, 9).value = v
SGAK = {6: 6150000, 7: 30948228, 8: None, 9: 1139000, 10: None, 11: 5327882, 12: 124395, 13: 10000, 14: 9490,
        15: 502790, 16: 4665234, 17: 12047802, 18: 96382, 19: 2223299, 20: 1059003, 21: 2318547, 22: None,
        23: 3616634, 24: 1290018, 25: 116000, 26: 7000, 27: 3082216, 28: 325790, 29: 702130, 30: 1408266,
        31: 3654000, 32: 1899133, 33: None, 34: 3023000}
for r, v in SGAK.items():
    sga.cell(r, 11).value = v
sga['B18'] = '振込手数料負担金'
sga['B33'] = None   # 振込手数料負担金は18行に統合（決算・試算表とも同一科目）
sga['H2'] = None; sga['L2'] = '(単位：千円)'; cs(sga, pl['L2'], sga['L2'])
sga['B42'] = '決算書・試算表の販管費計'
for c, v in zip('EGIK', [214823441, 215718454, 193151054, 85746239]):
    sga[f'{c}42'] = v; sga[f'{c}42'].number_format = '#,##0'
    sga[f'{c}43'] = f'=EXACT({c}40,{c}42)'
sga['B43'] = 'チェック'
hide_cols(sga, 'C', 'D')
sga.column_dimensions['B'].width = 36
for r in range(5, 41):   # 1期目（未使用・非表示）の売上高比は分母0のため数式を外す
    sga[f'D{r}'] = None
hide_rows(sga, 5, 33, 35, 36, 37, 38, 39)
note_red(sga, 'B45', '※2026/3期（I列）の賞与〜広告宣伝費は記入時に1行ずれていたため決算書の内訳に合わせて修正（賞与0、退職金共済掛金3,026千円、退職金747千円、法定福利費13,296千円、福利厚生費460千円、広告宣伝費89千円）。')
note_red(sga, 'B46', '※消費税は税込経理。決算では納付額を租税公課、試算表では消費税引当金に計上。進行期（K列）は2026年4〜9月の6か月累計。')

# ================================================================ VR→
mpl = wb['修正PL']; msga = wb['修正SGA']
mpl['O58'] = 21165907
# 正常化調整の前提（PLハイライト【調整項目】・修正SGAの原本）
r0 = 62
mpl.cell(r0, 2, '■ 正常化調整の前提（修正SGA・修正PLの修正額、PLハイライト【調整項目】はここを参照）').font = Font(name=JP, size=11, bold=True, color=NAVY)
hdr = mpl.cell(r0 + 1, 2, '項目')
for c, f in zip('FILO', ["='BS(借方)'!D4", "='BS(借方)'!E4", "='BS(借方)'!F4", "='BS(借方)'!G4"]):
    mpl[f'{c}{r0 + 1}'] = f
mpl.cell(r0 + 1, 18, '調整事由')
items = [
    ('退任役員の役員報酬（実績）', [17400000, 17400000, 11160000, 4050000],
     '譲渡時に退任する役員（代表者・配偶者・監査役）と2026/5退任済みの常務取締役の報酬を全額控除。取締役〔代表者の弟〕は残留のため対象外。進行期は4〜9月の実績'),
    ('代表者の顧問料（退任後）', [2160000, 2160000, 2160000, 1080000],
     '代表者の現行報酬（月36万円）の半額＝月18万円（2026/10/16面談で合意）。顧問料のため社会保険料は連動させない'),
    ('社会保険料率（概算）', [0.15, None, None, None], '退任役員の報酬控除に連動して社会保険料（概算15%）も控除'),
    ('賞与の正常化（社保込）', [0, 0, '=-AVERAGE(SGA!E8,SGA!G8)*(1+$F$66)', '=L67*6/12'],
     '賞与規定はなく業績に応じて支給。2026/3期は支給なし。継続支給を前提に前2期平均（社保込）を費用計上（保守的）。進行期は6/12'),
    ('寄付金', ['=SGA!E26', '=SGA!G26', '=SGA!I26', '=SGA!K26'], '譲渡後に発生しない費用として加算'),
    ('受取奨励金（営業外収益）', ['=PL!E24', '=PL!G24', '=PL!I24', '=PL!K24'], '毎期経常的に発生する事業関連収益として営業利益に振替（内容は要確認）'),
    ('出荷奨励金（営業外費用）', ['=PL!E30', '=PL!G30', '=PL!I30', '=PL!K30'], '受取奨励金と対になる支出として同様に振替'),
]
for i, (lab, vals, why) in enumerate(items):
    r = r0 + 2 + i
    mpl.cell(r, 2, lab)
    for c, v in zip('FILO', vals):
        if v is not None:
            mpl[f'{c}{r}'] = v
            mpl[f'{c}{r}'].number_format = '0%' if lab.startswith('社会保険料率') else '#,##0;△#,##0;"-"'
    mpl.cell(r, 18, why)
# 修正SGAの修正額（SGA行r → 修正SGA行r-1）。列：F/G=24/3、I/J=25/3、L/M=26/3、O/P=進行期
for bc, ac in zip('FILO', 'GJMP'):
    msga[f'{ac}5'] = f'=-修正PL!{bc}64+修正PL!{bc}65'            # 役員報酬
    msga[f'{ac}7'] = f'=-修正PL!{bc}67'                          # 賞与
    msga[f'{ac}10'] = f'=-修正PL!{bc}64*修正PL!$F$66'             # 法定福利費
    msga[f'{ac}25'] = f'=-修正PL!{bc}68'                         # 寄付金
    msga[f'{ac}34'] = f'=-(修正PL!{bc}69-修正PL!{bc}70)'          # 受取奨励金（純額）の振替
    mpl[f'{ac}24'] = f'=-修正PL!{bc}69'
    mpl[f'{ac}30'] = f'=-修正PL!{bc}70'
msga['B34'] = '受取奨励金（純額）の振替'
# テンプレートの不具合修正：修正SGAの修正額が修正PLの販管費に連動していなかった
for ac in 'DGJMP':
    mpl[f'{ac}18'] = f'=修正SGA!{ac}39'
# 修正後EBITDA（参考）
mpl['B72'] = '修正後営業利益'; mpl['B73'] = '減価償却費'; mpl['B74'] = '修正後EBITDA'
for bc, ec, sg in zip('FILO', 'HKNQ', 'EGIK'):
    mpl[f'{bc}72'] = f'={ec}19'; mpl[f'{bc}73'] = f'=SGA!{sg}31'; mpl[f'{bc}74'] = f'={bc}72+{bc}73'
    for r in (72, 73, 74):
        mpl[f'{bc}{r}'].number_format = '#,##0;△#,##0;"-"'
mpl['B76'] = '※支払利息は営業外費用のため営業利益ベースの正常収益力には含めない。特別損益（前期損益修正損・固定資産売却損益・貸倒引当金繰入/戻入）は非経常のため対象外。'
mpl['B77'] = '※雑収入（ゴミ処理料・冷蔵庫保管料等の経常部分を含む）は内訳資料がないため加算していない（要確認）。支払保険料・接待交際費のオーナー帰属分も未確認（要確認）。'

# --- 修正BS
mbd = wb['修正BS(借方)']; mbc = wb['修正BS(貸方) ']
mbd['D15'] = 203000                 # 貸倒引当金の戻入
mbd['D28'] = '=土地!E8'
mbd['D36'] = -2155595               # 電話加入権
mbd['D47'] = '=有価証券!E9'
mbc['D27'] = '=税効果!E18'          # 繰延税金負債
# テンプレートの不具合修正：繰延税金負債を「修正額」と「修正税効果」の両方で差し引いていた
mbc['D50'] = "='修正BS(借方)'!D61-'修正BS(借方)'!D60-(D28-D27)"
mbc['D51'] = "='修正BS(借方)'!D60-D27"
mbd['F15'] = '個別に回収可能性を判断するため全額戻入（貸倒実績は要確認）'
mbd['F28'] = '固定資産税評価額（R8年度）÷0.7で時価を推計（「土地」シート）'
mbd['F36'] = '換金価値なしとしてゼロ評価'
mbd['F47'] = 'フィデアHD株式を時価評価（「有価証券」シート）'

land = wb['土地']
clear(land, 4, 7, 2, 7)
for i, (nm, area, book, ref) in enumerate([
        ('本社・市場／鶴岡市宝田一丁目4-2', 3100.22, 70757240, "='不動産(所有)'!I4"),
        ('パッケージ場／鶴岡市宝田一丁目5-1', 900.5, 21380000, "='不動産(所有)'!I10"),
        ('旧倉庫敷地（駐車場）／東田川郡庄内町長沼字宮前304-7', 261.81, 1000000, "='不動産(所有)'!I13")]):
    r = 4 + i
    land[f'B{r}'] = nm; land[f'C{r}'] = area; land[f'D{r}'] = book; land[f'F{r}'] = ref
    land[f'E{r}'] = f'=F{r}-D{r}'; land[f'G{r}'] = '固定資産税評価額（R8年度）÷0.7にて時価を推計（鑑定・路線価で要確認）'
    land[f'C{r}'].number_format = '#,##0.00'
sec = wb['有価証券']
clear(sec, 4, 8, 2, 7)
sec['B4'] = '株式）フィデアホールディングス'; sec['C4'] = 3000; sec['D4'] = 1200000; sec['F4'] = '=2814*C4'
sec['E4'] = '=F4-D4'; sec['G4'] = '2026年9月18日終値（2,814円）を元に時価を算出（算定日終値で要更新）'
sec['B5'] = '株式）(株)庄果・出資金）鶴岡信用金庫ほか'; sec['C5'] = '-'; sec['D5'] = 1400000; sec['F5'] = '=D5'
sec['E5'] = '=F5-D5'; sec['G5'] = '取引関係上の保有。市場価格なし（簿価評価）'
ins = wb['保険']
clear(ins, 4, 7, 2, 6)
ins['B4'] = '事業保険積立金'; ins['C4'] = 0; ins['E4'] = 0; ins['D4'] = '=E4-C4'
ins['F4'] = '2026/3期に全額取崩し済（前期損益修正損20,900千円）。簿外の解約返戻金の有無は要確認'
rec = wb['売掛・棚卸']
clear(rec, 4, 7, 2, 6); clear(rec, 14, 17, 2, 6)
rec['B4'] = '売掛金（青果物）'; rec['C4'] = "='BS(借方)'!F7"; rec['D4'] = 0; rec['F4'] = '年齢表未受領のため簿価評価（長期滞留の有無は要確認）'
rec['B5'] = '資材肥料売掛金'; rec['C5'] = "='BS(借方)'!F8"; rec['D5'] = 0; rec['F5'] = '同上'
rec['E4'] = '=SUM(C4,D4)'; rec['E5'] = '=SUM(C5,D5)'
rec['B14'] = '棚卸資産（青果物・資材）'; rec['C14'] = "='BS(借方)'!F9"; rec['D14'] = 0; rec['E14'] = '=C14+D14'
rec['F14'] = '青果物は短期回転。廃棄を要する滞留在庫はなく帳簿は概ね実態どおり（第2回面談）。資材の滞留有無は要確認'
dep = wb['減価償却']
clear(dep, 4, 7, 2, 6)
for i, (nm, r_) in enumerate([('建物', 22), ('建物付属設備', 23), ('構築物', 24), ('車両運搬具・工具器具備品等', None)]):
    r = 4 + i
    dep[f'B{r}'] = nm
    dep[f'C{r}'] = f"='BS(借方)'!F{r_}" if r_ else "='BS(借方)'!F25+'BS(借方)'!F26+'BS(借方)'!F27"
    dep[f'D{r}'] = 0; dep[f'E{r}'] = f'=SUM(C{r},D{r})'
    dep[f'F{r}'] = '固定資産台帳未受領。別表十六で償却不足なしを確認し簿価評価' if i == 0 else '簿価評価'
upay = wb['未払人件費・賞与']
upay['B4'] = 2837000; upay['C4'] = '3/16〜3/31分'; upay['D4'] = 0; upay['E4'] = 0; upay['F4'] = 0
upay['B3'] = '2026年4月支払給与（3月後半分）'
upay['G4'] = '未払給与2,837千円は未払金に計上済のため追加計上なし'
upay['B9'] = 0; upay['C9'] = '-'; upay['D9'] = '-'; upay['E9'] = 0; upay['F9'] = 0
upay['B8'] = '賞与（規定なし・業績連動）'; upay['H9'] = '賞与規定なし。2026/3期は支給なし。引当不要と判断（要確認）'
ret = wb['退職金']
clear(ret, 4, 10, 2, 10)
ret['C4'] = '全従業員'; ret['H4'] = 0; ret['I4'] = 0; ret['J4'] = '=H4-I4'
ret['B6'] = '※中小企業退職金共済（掛金3,026千円/年）・オリックス生命で外部積立。退職時の会社の追加持出しは基本的にないとの認識（退職金規程で要確認）。'
ret['C10'] = '合計'; ret['H10'] = '=SUM(H4:H9)'; ret['I10'] = '=SUM(I4:I9)'; ret['J10'] = '=H10-I10'
tax = wb['税効果']
clear(tax, 4, 17, 2, 6)
for i, (nm, ok, ref, memo) in enumerate([
        ('貸倒引当金', '〇', "='修正BS(借方)'!D15", '戻入益に課税'),
        ('土地', '×', "='修正BS(借方)'!D28", '事業用不動産で売却予定がなく、実現可能性が低いため除外'),
        ('電話加入権', '×', "='修正BS(借方)'!D36", '流動性がないため除外'),
        ('投資有価証券', '〇', "='修正BS(借方)'!D47", '評価益に課税')]):
    r = 4 + i
    tax[f'B{r}'] = nm; tax[f'C{r}'] = ok; tax[f'D{r}'] = ref
    tax[f'E{r}'] = f'=-D{r}*-0.34' if ok == '〇' else 0
    tax[f'F{r}'] = memo
tax['D18'] = '=SUM(D4:D17)'; tax['E18'] = '=SUM(E4:E17)'
tax['F18'] = '繰延税金負債（評価益×実効税率34%）'
nc = wb['ネットキャッシュ']
for r in range(4, 13):
    nc[f'C{r}'] = None; nc[f'F{r}'] = None
nc['C4'] = "='修正BS(借方)'!E6"; nc['C9'] = '=有価証券!F4'
nc['B9'] = '投資有価証券（上場株式・時価）'
nc['F5'] = "='修正BS(貸方) '!E19"; nc['F8'] = "='修正BS(貸方) '!E10"
nc['C15'] = '=C13-F13'
nc['F15'] = '=F13-C13'      # テンプレートの不具合修正（I13を参照していた）
nc['C17'] = None; nc['F17'] = None
nc['B19'] = '※預り保証金8,980千円・保証金33,303千円は事業上の預託・預り金としてネットキャッシュに含めていない（取扱いは要確認）。(株)庄果株式・信金出資は取引関係上の保有のため除外。'

nen = wb['年買法']
nen['B3'] = '26年3月期'
nen['C4'] = "='修正BS(貸方) '!C52"
nen['C5'] = "='修正BS(貸方) '!D50"
nen['C6'] = "='修正BS(貸方) '!D51"      # 税効果は「税効果」シートで判定したもののみ（テンプレートは修正額全体に34%）
nen['C7'] = '=SUM(C4:C6)'
for c, bc in zip('FGH', 'FIL'):
    nen[f'{c}3'] = f"=修正PL!{bc}3"
    nen[f'{c}4'] = f"=修正PL!{bc}19"
    ac = {'F': 'G', 'I': 'J', 'L': 'M'}[bc]
    nen[f'{c}5'] = f"=修正PL!{ac}19"
nen['G8'] = '=SUM(G6:G7)'; nen['H8'] = '=SUM(H6:H7)'   # テンプレートの不具合修正（修正額を二重計上していた）
nen['K4'] = '=C7'
nen['B12'] = '※営業権倍率1〜3倍は国内中小企業M&Aのボリュームゾーン。加重は直近期0.5・前期0.3・前々期0.2。'
ev = wb['EV EBITDAマルチプル']
ev['B3'] = '26年3月期'
ev['C4'] = '=ネットキャッシュ!C13'; ev['C5'] = '=ネットキャッシュ!F13'
for c, bc, sg in zip('FGH', 'FIL', 'EGI'):
    ac = {'F': 'G', 'I': 'J', 'L': 'M'}[bc]
    ev[f'{c}3'] = f"=修正PL!{bc}3"; ev[f'{c}4'] = f"=修正PL!{bc}19"; ev[f'{c}5'] = f"=修正PL!{ac}19"
    ev[f'{c}7'] = f'=SGA!{sg}31'
ev['B12'] = '※倍率3〜5倍は国内中小企業M&Aの実務水準。非流動性ディスカウント30%。算定値がマイナスの場合、ネットデットが事業価値を上回ることを示す。'
ref = wb['参照企業']
clear(ref, 4, 15, 2, 5)
ref['B4'] = '（未使用）倍率は国内中小企業M&Aの実務水準（3〜5倍）を採用。卸売市場関連の上場会社の倍率は算定日時点の株価で更新が必要（要確認）'

# ================================================================ IM→
# --- 会社概要
co = wb['会社概要']
CO = {
    3: '株式会社鶴岡中央青果（http://www.turu-chu.co.jp/）',
    4: '山形県鶴岡市宝田一丁目4番20号',
    5: '① パッケージ場：山形県鶴岡市宝田一丁目5-1\n② 旧倉庫敷地（駐車場）：山形県東田川郡庄内町長沼字宮前304-7',
    6: '1973年（昭和48年）11月22日（第54期）',
    7: '=TEXT(\'BS (貸方)\'!F36/1000,"#,##0")&"千円"',
    8: '岡部 寛',
    9: '岡部 寛（1,510株／37.8%）ほか7名\n代表者一族4名で75.8%を保有',
    10: '・青果物（野菜・果実）の卸売（地方卸売市場の開設・運営、買付販売・受託販売）\n・パッケージ（小分け・包装）、配送（量販店の配送受託を含む）\n・包装資材の販売',
    11: '24名（正社員17名、嘱託・再雇用・試用3名、パート・アルバイト4名）\n※2026年9月1日時点。別途、常勤役員2名',
    12: '地方卸売市場の認定（卸売市場法・山形県知事）\n2020年6月認定、2026年2月変更認定',
    13: '=TEXT(PL!I11/1000,"#,##0")&"千円（2026年3月期）"',
    14: '=TEXT(AVERAGE(PLハイライト!D25:F25)/1000,"#,##0")&"千円（過去3期平均）"',
    15: '後継者不在およびさらなる企業の発展',
}
co['B5'] = 'その他拠点'; co['B6'] = '設立'
for r, v in CO.items():
    co[f'C{r}'] = v
    co[f'C{r}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    n = sum(max(1, -(-len(x) // 27)) for x in str(v).split('\n')) if not str(v).startswith('=') else 1
    co.row_dimensions[r].height = max(22.5, 19 * n + 4)
mark(co, 'B2:C15', '基本情報')

# --- PLハイライト（行を再構成：調整項目は修正PLの前提を参照）
ph = wb['PLハイライト']
tmpl = {k: [copy.copy(ph.cell(k, c)._style) for c in range(2, 8)] for k in (3, 4, 5, 14, 15, 17, 18, 37)}
hts = {k: ph.row_dimensions[k].height for k in tmpl}
clear(ph, 14, 48, 2, 9, style=True)
ph['I4'] = None


def put_style(r, k):
    for j, c in enumerate(range(2, 8)):
        ph.cell(r, c)._style = copy.copy(tmpl[k][j])
    ph.row_dimensions[r].height = hts[k]


for c, f in zip('CDEFG', ["='BS(借方)'!C4", "='BS(借方)'!D4", "='BS(借方)'!E4", "='BS(借方)'!F4", "='BS(借方)'!G4"]):
    ph[f'{c}3'] = f
ph.row_dimensions[3].height = 33
ph['G5'] = '=IFERROR(G4/G42,"-")'
put_style(14, 14); put_style(15, 15); put_style(16, 14); put_style(17, 17)
for r in range(18, 24):
    put_style(r, 18)
put_style(24, 14); put_style(25, 37)
ph['B15'] = '減価償却費'
for c, sg in zip('CDEFG', ['C', 'E', 'G', 'I', 'K']):
    ph[f'{c}15'] = f'=SGA!{sg}31'
ph['C15'] = 0
ph['B17'] = '調整項目 合計'
adj = [('役員報酬（退任役員）', '=修正PL!{c}64'), ('顧問料（代表者）', '=-修正PL!{c}65'),
       ('社会保険料', '={x}18*修正PL!$F$66'), ('賞与の正常化', '=修正PL!{c}67'),
       ('寄付金', '=修正PL!{c}68'), ('受取奨励金（純額）', '=修正PL!{c}69-修正PL!{c}70')]
for i, (lab, f) in enumerate(adj):
    r = 18 + i
    ph[f'B{r}'] = lab
    for x, c in zip('DEFG', 'FILO'):
        ph[f'{x}{r}'] = f.format(c=c, x=x)
for x in 'DEFG':
    ph[f'{x}17'] = f'=SUM({x}18:{x}23)'
    ph[f'{x}25'] = f'={x}8+{x}15+{x}17'
ph['B25'] = '調整後EBITDA'
for r in (17, 25):
    ph[f'C{r}'] = None
hide_cols(ph, 'C')
for r in range(4, 14):   # 1期目（未使用・非表示）は空欄に
    ph[f'C{r}'] = None
ph['D5'] = '-'
for r in (5, 7, 9, 11, 13):          # 比率行は低く（鳥塚の概要書と同じ見え方）
    ph.row_dimensions[r].height = 14
for r in (14, 16, 24):
    ph.row_dimensions[r].height = 8
for r in (4, 6, 8, 10, 12, 15, 17, 18, 19, 20, 21, 22, 23, 25):
    ph.row_dimensions[r].height = 20
ph.column_dimensions['B'].width = 31
mark(ph, 'B2:G25', '財務ハイライト（P/L）')
# 参考（貼付範囲外）
ph['B28'] = '参考（リード文・Comment用。貼付範囲外）'; ph['B28'].font = Font(name=JP, size=11, bold=True, color=NAVY)
ref_rows = [
    (29, '調整後EBITDA 3期平均', '=AVERAGE(D25:F25)'), (30, '調整後EBITDA 直近2期平均', '=AVERAGE(E25:F25)'),
    (31, 'チェック：修正PLの修正後EBITDA（26/3期）と一致', '=ROUND(F25,0)=ROUND(修正PL!L74,0)'),
    (32, 'EBITDA（営業利益＋減価償却費）', None), (33, '販売費及び一般管理費', None),
]
for r, lab, f in ref_rows:
    ph[f'B{r}'] = lab
    if f:
        ph[f'C{r}'] = f
ph['B32'] = 'EBITDA（営業利益＋減価償却費）'
for x in 'DEFG':
    ph[f'{x}32'] = f'={x}8+{x}15'
for x, sg in zip('DEFG', 'EGIK'):
    ph[f'{x}33'] = f'=SGA!{sg}40'
ph['B35'] = '前年同期（2025年4〜9月 試算表）'; ph['B35'].font = Font(name=JP, size=11, bold=True)
PY = [('前年同期 売上高', 667639313), ('前年同期 売上総利益', 108315825), ('前年同期 販管費', 98862164),
      ('前年同期 営業利益', 9453661), ('前年同期 経常利益', 10012691), ('前年同期 税引前当期純利益', 9880669),
      ('前年同期 減価償却費', 6240000), ('前年同期 人件費（役員報酬・給料・法定福利等）', 50027381), ('前年同期 運賃', 17023172)]
# 前年同期（売上高は42行目：G5の前期比が参照）
ph['B36'], ph['G36'] = '前年同期 売上総利益', 108315825
ph['B37'], ph['G37'] = '前年同期 販管費', 98862164
ph['B38'], ph['G38'] = '前年同期 営業利益', 9453661
ph['B39'], ph['G39'] = '前年同期 経常利益', 10012691
ph['B40'], ph['G40'] = '前年同期 税引前当期純利益', 9880669
ph['B41'], ph['G41'] = '前年同期 減価償却費', 6240000
ph['B42'], ph['G42'] = '前年同期 売上高', 667639313
ph['B43'], ph['G43'] = '前年同期 人件費（役員報酬・給料・法定福利等）', 50027381
ph['B44'], ph['G44'] = '前年同期 運賃', 17023172
ph['B45'] = '当期 人件費（役員報酬・給料・法定福利等）'; ph['G45'] = '=SUM(SGA!K6:K12)'
for r in range(29, 46):
    for x in 'CDEFG':
        ph[f'{x}{r}'].number_format = '#,##0;△#,##0;"-"'
ph['C31'].number_format = 'General'
ph['B47'] = '※進行期は2026年4〜9月の試算表（6か月累計。年換算しない）。前期比は前年同期との比較。進行期の調整項目は同じ方針で6か月分を算定（賞与の正常化は年額の6/12）。'
ph['B47'].font = Font(name='游ゴシック Medium', size=9)

# --- BSハイライト（手入力値→BSシートへの数式に置換）
bh = wb['BSハイライト']
F = lambda r: f"='BS(借方)'!F{r}"
FC = lambda r: f"='BS (貸方)'!F{r}"
BHV = {'C5': F(6), 'C6': "='BS(借方)'!F7+'BS(借方)'!F8", 'C7': F(9), 'C8': F(11), 'C9': F(12),
       'C10': "='BS(借方)'!F10+'BS(借方)'!F13+'BS(借方)'!F14+'BS(借方)'!F15",
       'C12': F(22), 'C13': F(23), 'C14': F(24), 'C15': F(28), 'C16': F(36), 'C17': F(37), 'C18': F(47), 'C19': F(50),
       'C20': "='BS(借方)'!F25+'BS(借方)'!F26+'BS(借方)'!F27+'BS(借方)'!F29+'BS(借方)'!F48+'BS(借方)'!F49+'BS(借方)'!F51+'BS(借方)'!F56",
       'F5': FC(6), 'F6': FC(7), 'F7': FC(9), 'F8': FC(10), 'F9': FC(11), 'F10': FC(8),
       'F12': FC(19), 'F13': FC(20), 'F18': FC(36), 'F19': FC(43)}
for k, v in BHV.items():
    bh[k] = v
bh['E10'] = 'その他'; cs(bh, bh['E9'], bh['E10'])
bh['B6'] = '売掛金'
bh['B3'] = "=PLハイライト!F3&\"末\""
mark(bh, 'B2:F21', '財務ハイライト（B/S）')
bh['B24'] = '参考：B/S推移（リード文・Comment用。貼付範囲外）'; bh['B24'].font = Font(name=JP, size=11, bold=True, color=NAVY)
for j, (c, src) in enumerate(zip('CDEF', 'DEFG')):
    bh[f'{c}25'] = f"='BS(借方)'!{src}4"
    bh[f'{c}26'] = f"='BS(借方)'!{src}6"
    bh[f'{c}27'] = f"='BS (貸方)'!{src}19"
    bh[f'{c}28'] = f"='BS (貸方)'!{src}50"
    bh[f'{c}29'] = f"='BS(借方)'!{src}61"
    bh[f'{c}30'] = f'={c}28/{c}29'
    bh[f'{c}31'] = f'={c}27-{c}26'
    for r in (26, 27, 28, 29, 31):
        bh[f'{c}{r}'].number_format = '#,##0,;△#,##0,'
    bh[f'{c}30'].number_format = '0.0%'
for r, lab in zip(range(25, 32), ['項目', '現金及び預金', '長期借入金', '純資産', '資産合計', '自己資本比率', 'ネットデット（借入金−現預金）']):
    bh[f'B{r}'] = lab
bh['B33'] = '借入先別残高（2026/3期末）'
for i, (nm, v) in enumerate([('鶴岡信用金庫 本店営業部', 182840000), ('山形銀行 鶴岡支店', 70250000),
                             ('日本政策金融公庫 山形支店（1）', 60730000), ('日本政策金融公庫 山形支店（2）', 39380000),
                             ('日本政策金融公庫 山形支店（3）', 7858000)]):
    bh[f'B{34 + i}'] = nm; bh[f'C{34 + i}'] = v; bh[f'C{34 + i}'].number_format = '#,##0'
bh['B39'] = '合計'; bh['C39'] = '=SUM(C34:C38)'; bh['B40'] = 'チェック：BS長期借入金と一致'; bh['C40'] = "=C39='BS (貸方)'!F19"
bh['B42'] = '※2026年9月末は長期借入金526,264千円（3月末比＋165,206千円）、現預金393,411千円。新規借入の借入先・資金使途・条件は要確認。金利・返済条件・担保・保証も要確認。'
bh['B42'].font = Font(name='游ゴシック Medium', size=9)

# --- 取引先（買参人別順位表：年間売上金額）
tr = wb['取引先']
for c, src in zip('GIK', 'DEF'):
    tr[f'{c}3'] = f"='BS(借方)'!{src}4"
BUY = [  # 名称, 取引内容, 所在地, 業種等, FY24, FY25, FY26, 備考
    ('(株)主婦の店（鶴岡店）', '青果物', '山形県鶴岡市', '食品スーパー', 308383428, 302458433, 345226753, '配送受託あり'),
    ('生活協同組合共立社', '青果物', '山形県鶴岡市', '生活協同組合', 96150823, 99248896, 85615684, '毎日納品'),
    ('長井中央青果(株)', '青果物', '山形県長井市', '青果卸売業', '=36988620+51092214', 77319990, 75394815, '仕入先でもある'),
    ('松浦青果(株)', '青果物', '要確認', '要確認', 13988200, 20012345, 32186104, None),
    ('生活協同組合連合会アイチョイス', '青果物', '愛知県名古屋市', '生協連合会', 25340515, 28050058, 22375855, None),
    ('(株)新庄丸果青果', '青果物', '山形県新庄市', '青果卸売業', 25221720, 23896155, 20030390, None),
    ('(株)マルタ', '青果物', '要確認', '要確認', 26352085, 32060257, 19175010, None),
    ('小野寺食品 やま亀', '青果物', '要確認', '食品加工', 20114508, 16807815, 18891770, None),
    ('(株)R&Cながの青果', '青果物', '長野県', '青果卸売業', 27716363, 25776810, 18856750, None),
    ('(株)マルト直治郎商店', '青果物', '山形県三川町', '青果卸売業', '－', 6678890, 17633870, '仕入先でもある'),
]
row_style(tr, 16, 17, 2, 13); row_style(tr, 15, 16, 2, 13)
for i, rec_ in enumerate(BUY):
    r = 5 + i
    tr[f'B{r}'] = i + 1
    for c, v in zip('CDEFGIKM', [rec_[0], rec_[1], rec_[2], rec_[3], rec_[4], rec_[5], rec_[6], rec_[7]]):
        tr[f'{c}{r}'] = v
tr['B15'] = None; tr['C15'] = '社内勘定（近郷・パッケージ）'; tr['D15'] = '部門別勘定'
tr['E15'] = '－'; tr['F15'] = '－'; tr['M15'] = '性質は要確認（注記参照）'
tr['G15'] = '=SUM(G22:G28)'; tr['I15'] = '=SUM(I22:I28)'; tr['K15'] = '=SUM(K22:K28)'
tr['C16'] = 'その他'
for c in 'GIK':
    tr[f'{c}16'] = f'={c}17-SUM({c}5:{c}15)'
tr['C17'] = '合計（決算書の売上高）'
tr['G17'] = '=PL!E11'; tr['I17'] = '=PL!G11'; tr['K17'] = '=PL!I11'
for r in range(5, 18):
    for c, d in zip('HJL', 'GIK'):
        tr[f'{c}{r}'] = f'=IFERROR({d}{r}/{d}$17,"－")'
        cs(tr, tr['H5'], tr[f'{c}{r}']) if r < 15 else None
    for c in 'GIK':
        if r < 15:
            cs(tr, tr['G5'], tr[f'{c}{r}'])
for r in range(5, 18):
    tr.row_dimensions[r].height = 23.45
    for c in 'GIK':
        tr[f'{c}{r}'].alignment = Alignment(horizontal='right', vertical='center')
for c in 'GIK':
    tr[f'{c}17'].number_format = '#,##0,'
tr['H17'] = '-'; tr['J17'] = '-'; tr['L17'] = '-'
for c in 'HJL':
    tr[f'{c}17'] = '100.0%'
    tr[f'{c}17'].alignment = Alignment(horizontal='right', vertical='center')
ungroup_cols(tr)
for c, w_ in zip('BCDEFGHIJKLM', [4.4, 25, 13, 14, 16, 10.5, 7.5, 10.5, 7.5, 10.5, 7.5, 20]):
    tr.column_dimensions[c].width = w_
for r in range(3, 18):
    for c in 'BCDEFGHIJKLM':
        al = copy.copy(tr[f'{c}{r}'].alignment); al.vertical = 'center'; al.wrap_text = True
        tr[f'{c}{r}'].alignment = al
mark(tr, 'B2:M17', '主要得意先の3期推移')
tr['B19'] = '※出所：買参人別順位表（販売管理システム、上位20位）の年間売上金額。上位20位圏外の年度は「－」。構成比は決算書の売上高に対する比率（集計基準が異なるため参考値）。'
tr['B20'] = '社内勘定（集計）'
INT_B = [('KK 地物果実', 6919565, 110231725, 99079935), ('KK マルカ', 31997559, 67036327, 67277546),
         ('KK 開発', 64841026, 33259997, 39250623), ('KK 社長車', 40489545, 38577365, 38330905),
         ('KK メロン（元常務取扱）', 65506775, None, None), ('パッケージ根菜', 46495688, 36837944, 36540635),
         ('パッケージ果実・蔬菜', '=34618043+15961990', '=26594781+21465797', '=31116065+21893443')]
for i, (nm, a, b, c_) in enumerate(INT_B):
    r = 22 + i
    tr[f'C{r}'] = nm; tr[f'G{r}'] = a; tr[f'I{r}'] = b; tr[f'K{r}'] = c_
    for c in 'GIK':
        tr[f'{c}{r}'].number_format = '#,##0'
tr['C30'] = '上位20位計（順位表）'; tr['G30'] = 863326289; tr['I30'] = 1034424284; tr['K30'] = 1035467601
tr['C31'] = '※24/3期の上位20位計は25/3期の順位表の前年欄（25/3期の上位20先の前年実績）。'

# --- 仕入先（荷主別順位表：年間仕入金額）
sp = wb['仕入先']
for c, src in zip('FHJ', 'DEF'):
    sp[f'{c}3'] = f"='BS(借方)'!{src}4"
SUP = [
    ('丸果庄内青果(株)', '青果物', '山形県三川町', 237536555, 246284304, 315032146, '公設庄内青果物市場の卸売業者'),
    ('(株)庄果', '青果物', '山形県三川町', 113535109, 114398359, 98621582, '同上'),
    ('長井中央青果(株)', '青果物', '山形県長井市', 29926325, 35262749, 62586074, '販売先でもある'),
    ('(株)マルト直治郎商店', '青果物', '山形県三川町', 30113414, 56151367, 21974990, '販売先でもある'),
    ('成田食品(株)', '青果・加工品', '福島県相馬市', 14127795, 14916067, 15635130, None),
    ('庄内みどり農業協同組合', '受託品', '山形県酒田市', 7118174, 7593661, 8879410, '受託販売の委託者'),
    ('マルショウ青果(有)', '青果物', '要確認', 14839869, 3207289, 7847408, None),
    ('生産者A（波の会）', '地場産青果物', '庄内地方', 5608122, 6263216, 6934170, '個人生産者のため匿名'),
    ('生産者B（波の会）', '地場産青果物', '庄内地方', 8173503, 7441032, 6490104, '個人生産者のため匿名'),
    ('(株)丸越', '青果物', '要確認', '－', 6056888, 6406749, None),
]
row_style(sp, 16, 17, 2, 12); row_style(sp, 15, 16, 2, 12)
for i, rec_ in enumerate(SUP):
    r = 5 + i
    sp[f'B{r}'] = i + 1
    for c, v in zip('CDEFHJL', rec_):
        sp[f'{c}{r}'] = v
sp['B15'] = None; sp['C15'] = '社内勘定（近郷・パッケージ）'; sp['D15'] = '部門別勘定'; sp['E15'] = '－'
sp['L15'] = '性質は要確認（注記参照）'
sp['F15'] = '=SUM(F22:F30)'; sp['H15'] = '=SUM(H22:H30)'; sp['J15'] = '=SUM(J22:J30)'
sp['C16'] = 'その他'
for c in 'FHJ':
    sp[f'{c}16'] = f'={c}17-SUM({c}5:{c}15)'
sp['C17'] = '合計（決算書の仕入高）'
sp['F17'] = '=PL!E13'; sp['H17'] = '=PL!G13'; sp['J17'] = '=PL!I13'
for r in range(5, 18):
    for c, d in zip('GIK', 'FHJ'):
        sp[f'{c}{r}'] = f'=IFERROR({d}{r}/{d}$17,"－")'
        if r < 15:
            cs(sp, sp['G5'], sp[f'{c}{r}'])
    for c in 'FHJ':
        if r < 15:
            cs(sp, sp['F5'], sp[f'{c}{r}'])
        sp[f'{c}{r}'].alignment = Alignment(horizontal='right', vertical='center')
    sp.row_dimensions[r].height = 23.45
for c in 'FHJ':
    sp[f'{c}17'].number_format = '#,##0,'
for c in 'GIK':
    sp[f'{c}17'] = '100.0%'; sp[f'{c}17'].alignment = Alignment(horizontal='right', vertical='center')
ungroup_cols(sp)
for c, w_ in zip('BCDEFGHIJKL', [4.4, 25, 15, 16, 10.5, 7.5, 10.5, 7.5, 10.5, 7.5, 24]):
    sp.column_dimensions[c].width = w_
for r in range(3, 18):
    for c in 'BCDEFGHIJKL':
        al = copy.copy(sp[f'{c}{r}'].alignment); al.vertical = 'center'; al.wrap_text = True
        sp[f'{c}{r}'].alignment = al
mark(sp, 'B2:L17', '主要仕入先の3期推移')
sp['B19'] = '※出所：荷主別順位表（販売管理システム、上位20位）の年間仕入金額。上位20位圏外の年度は「－」。構成比は決算書の仕入高に対する比率（集計基準が異なるため参考値）。'
sp['B20'] = '社内勘定（集計）'
INT_S = [('KK 地物果実', 7439366, 127941370, 80678993), ('KK マルカ', 28693863, 69250898, 66961587),
         ('KK 開発', 80299497, 33616540, 38502541), ('KK 社長車', 43810598, 38396325, 38548107),
         ('KK メロン・大根・常務（担当者別）', '=107800157+11451728+9226201', 13058625, None),
         ('パッケージ根菜', 48689621, 36903178, 36570360), ('パッケージ果実', 35852672, 23616079, 32629978),
         ('パッケージ蔬菜', 16356642, 21396557, 21921243), ('パッケージ生花', 16190861, 13344110, 12407120)]
for i, (nm, a, b, c_) in enumerate(INT_S):
    r = 22 + i
    sp[f'C{r}'] = nm; sp[f'F{r}'] = a; sp[f'H{r}'] = b; sp[f'J{r}'] = c_
    for c in 'FHJ':
        sp[f'{c}{r}'].number_format = '#,##0'
sp['C32'] = '上位20件計（順位表）'; sp['F32'] = 728315067; sp['H32'] = 892035615; sp['J32'] = 889128701

# --- 外注先
ex = wb['外注先']
clear(ex, 5, 16, 3, 13)
ex['C5'] = '該当なし（外注加工はなし）'
ex['M5'] = '配送は原則自社。東京向け出荷・一部の地元スーパー向けのみ外部運送会社を利用（運賃の委託先別内訳は要確認）'

# --- 株主
sh = wb['株主']
sh['E3'] = "=\"役員報酬\"&CHAR(10)&\"（\"&'BS(借方)'!F4&\"）\""
SH = [('岡部 寛', '代表取締役', 4320000, 1510, ''), ('岡部 一', '取締役（代表者の弟）', 4200000, 720, ''),
      ('岡部 和子', '代表者の親族', 0, 500, ''), ('梅津 恵一', '－', 0, 360, ''),
      ('小野寺 則和', '監査役', 120000, 350, ''), ('岡部 竹古', '取締役（代表者の配偶者）', 120000, 300, ''),
      ('佐藤 令子', '－', 0, 200, ''), ('菅原 圭吾', '元常務取締役', 6600000, 60, '2026年5月29日退任')]
for i, (nm, rl, pay, n, memo) in enumerate(SH):
    r = 5 + i
    sh[f'C{r}'] = nm; sh[f'D{r}'] = rl; sh[f'E{r}'] = pay; sh[f'F{r}'] = n; sh[f'H{r}'] = memo or None
hide_rows(sh, 4, 13, 14)
ungroup_cols(sh); sh.column_dimensions['D'].width = 25
sh.row_dimensions[3].height = 33
mark(sh, 'B2:H15', '株主及び役員（株主）')
sh['B17'] = 'チェック：発行済株式総数4,000株'; sh['F17'] = '=F15=4000'
sh['B18'] = 'うち代表者一族（岡部家4名）'; sh['F18'] = '=F5+F6+F7+F10'; sh['G18'] = '=F18/F15'
sh['B19'] = '※出所：法人税申告書別表二（2026/3期）。株主名簿・名義株の有無・親族外株主の取得経緯は要確認。'

# --- 取締役
dr = wb['取締役']
dr['F3'] = "=\"役員報酬\"&CHAR(10)&\"（\"&'BS(借方)'!F4&\"）\""
DR = [('岡部 寛', '代表取締役（常勤）', '本人', 4320000, '経営全般・販売', '譲渡後は顧問（最大1年）'),
      ('岡部 一', '取締役（常勤）', '弟', 4200000, '買付・販売', '譲渡後も継続勤務予定'),
      ('岡部 竹古', '取締役（非常勤）', '配偶者', 120000, '業務への関与なし', '譲渡時に退任予定'),
      ('小野寺 則和', '監査役（非常勤）', '－', 120000, '会計監査', '譲渡時に退任予定'),
      ('菅原 圭吾', '常務取締役（常勤）', '－', 6600000, '－', '2026年5月29日退任')]
for i, rec_ in enumerate(DR):
    r = 5 + i
    for c, v in zip('CDEFGH', rec_):
        dr[f'{c}{r}'] = v
hide_rows(dr, 4, 10, 11, 12, 13, 14)
ungroup_cols(dr)
for c, w_ in zip('BCDEFGH', [4.4, 16, 22, 13, 13, 24, 26]):
    dr.column_dimensions[c].width = w_
dr.row_dimensions[3].height = 33
mark(dr, 'B2:H15', '株主及び役員（役員）')
dr['B17'] = 'チェック：SGA役員報酬（26/3期）と一致'; dr['F17'] = '=F15=SGA!I6'

# --- 従業員
em = wb['従業員']
em['G2'] = datetime.datetime(2026, 9, 1)
for r in range(5, 35):
    v = em[f'I{r}'].value
    if isinstance(v, str) and v.count('/') == 2:     # 入社年月日が文字列で入力されている → 日付に変換
        y, m, d = map(int, v.split('/'))
        em[f'I{r}'] = datetime.datetime(y, m, d); em[f'I{r}'].number_format = 'yyyy/mm/dd'
    if em[f'E{r}'].value == 'パートタイム労働者':
        em[f'E{r}'] = 'パート'
    if em[f'K{r}'].value == '野菓営業部':
        em[f'K{r}'] = '野菜営業部'     # 記入時の誤字を修正
em['M4'] = '26年9月分総支給'
hide_cols(em, 'C', 'D', 'G', 'I', 'M', 'N', 'O', 'P')
em.column_dimensions['E'].width = 17
for r in range(5, 29):
    em[f'H{r}'].number_format = '0'; em[f'J{r}'].number_format = '0'
    # テンプレートの「÷365」はうるう年分だけ年齢・勤続が過大になるため、満年数（DATEDIF）に置換
    em[f'H{r}'] = f'=IF(G{r}="","",DATEDIF(G{r},$G$2,"Y"))'
    em[f'J{r}'] = f'=IF(I{r}="","",DATEDIF(I{r},$G$2,"Y"))'
hide_rows(em, 29, 30, 31, 32, 33, 34)
mark(em, 'B4:Q28', '従業員一覧')
em['B37'] = '人数'; em['E37'] = '=COUNTA(E5:E34)'
em['B38'] = '平均年齢'; em['E38'] = '=AVERAGE(H5:H28)'
em['B39'] = '平均勤続年数'; em['E39'] = '=AVERAGE(J5:J28)'
em['B40'] = '60歳以上'; em['E40'] = '=COUNTIF(H5:H28,">=60")'
em['B41'] = '正社員'; em['E41'] = '=COUNTIF(E5:E28,"正社員")'
em['B42'] = 'セリ人資格者'; em['E42'] = '=COUNTIF(Q5:Q28,"*セリ人*")'
em['B44'] = '※出所：職員名簿（2026年9月1日現在）。氏名・イニシャル・生年月日・入社日は概要書に貼らない（非表示列）。賃金台帳（2026/4〜10）は氏名表記のため名簿№との対応確認後に総支給額を転記（要確認）。'

# --- 拠点(本社等)
bp = wb['拠点(本社等)']
hdr_st = [copy.copy(bp.cell(2, c)._style) for c in range(2, 8)]
body_st = [copy.copy(bp.cell(3, c)._style) for c in range(2, 8)]
for j, v in enumerate(['拠点', '所在地', '所有区分', '土地面積\n(㎡)', '主な建物', '用途']):
    bp.cell(2, 2 + j, v)
BP = [('本社・市場', '山形県鶴岡市宝田一丁目4番20号\n（地番：宝田一丁目4-2）', '法人', 3100.22, '事務所（管理棟）・市場・倉庫・容器庫・貯蔵庫', '事務、セリ・相対取引、保管'),
      ('パッケージ場', '山形県鶴岡市宝田一丁目5-1', '法人', 900.5, '事務所兼倉庫（パッケージ場）・倉庫', '小分け・包装、保管'),
      ('旧倉庫敷地', '山形県東田川郡庄内町長沼字宮前304-7', '法人', 261.81, 'なし（倉庫は老朽化により解体済み）', '月極駐車場として近隣へ賃貸'),
      ('賃借物件', 'なし（地代家賃の計上なし）', '－', None, '－', '－')]
for i, rec_ in enumerate(BP):
    r = 3 + i
    for j, v in enumerate(rec_):
        x = bp.cell(r, 2 + j, v)
        x._style = copy.copy(body_st[j])
    bp.cell(r, 5).number_format = '#,##0.00'
    bp.cell(r, 5).alignment = Alignment(horizontal='right', vertical='center')
    bp.row_dimensions[r].height = 39
for c, w_ in zip('BCDEFG', [12, 30, 9, 10, 34, 26]):
    bp.column_dimensions[c].width = w_
mark(bp, 'B2:G6', '拠点概要（本社・市場）')
bp['B8'] = '※営業時間は早朝5時から（終業時刻は要確認）、定休日は水・日・祝日。車両は自己所有とリースを併用（リース料 年6,227千円）。'

# --- 不動産(所有)
re_ = wb['不動産(所有)']
body = [copy.copy(re_.cell(4, c)._style) for c in range(2, 12)]
clear(re_, 4, 21, 2, 11, style=True)
for m in list(re_.merged_cells.ranges):
    if m.min_row >= 4:
        re_.unmerge_cells(str(m))
LAND_BLD = [
    ('本社・市場', '土地）市場敷地', '鶴岡市宝田一丁目4-2', 3100.22, '宅地', '要確認', 70757240, '=J4/0.7', 35861205, '固定資産税評価額を0.7で割戻し、時価を算出'),
    (None, '建物）事務所', '宝田一丁目4-2', 365.63, '鉄骨造 2階', '1981年新築', None, None, 11218061, ''),
    (None, '建物）市場', '宝田一丁目4-2', 320.22, '鉄骨造 1階', '1981年新築', None, None, 4712536, ''),
    (None, '建物）倉庫', '宝田一丁目4-2', 327.13, '鉄骨造 1階', '1981年新築', None, None, 4560792, ''),
    (None, '建物）容器庫', '宝田一丁目4-2', 198.74, '軽量鉄骨造 2階', '1981年新築', None, None, 1237049, ''),
    (None, '建物）貯蔵庫', '宝田一丁目4-2', 34.48, '軽量鉄骨造 1階', '1981年新築', None, None, 566907, ''),
    ('パッケージ場', '土地）作業場敷地', '鶴岡市宝田一丁目5-1', 900.5, '宅地', '要確認', 21380000, '=J10/0.7', 10560298, '固定資産税評価額を0.7で割戻し、時価を算出'),
    (None, '建物）事務所兼倉庫', '宝田一丁目5-1', 627.4, '鉄骨造 2階', '1982年新築', None, None, 13110977, ''),
    (None, '建物）倉庫', '宝田一丁目5-1', 139.08, '鉄骨造 1階', '2012年新築', None, None, 11655678, ''),
    ('旧倉庫敷地', '土地）駐車場', '東田川郡庄内町長沼字宮前304-7', 261.81, '宅地', '要確認', 1000000, '=J13/0.7', 1056005, '倉庫は解体済み。月極駐車場として賃貸'),
    ('建物計', '建物・付属設備（7棟）', '－', '=SUM(E5:E9,E11:E12)', '－', '－', "='BS(借方)'!F22+'BS(借方)'!F23", '=H14', '=SUM(J5:J9,J11:J12)', '減価償却過不足なし（別表十六）、簿価＝時価として評価。棟別簿価は台帳で要確認'),
]
# 行の並び：4土地,5-9建物,10土地,11-12建物,13土地,14建物計 → 上の配列は4〜14
for i, rec_ in enumerate(LAND_BLD):
    r = 4 + i
    for j, v in enumerate(rec_):
        x = re_.cell(r, 2 + j, v)
        x._style = copy.copy(body[j])
    re_.row_dimensions[r].height = 27
    for c in (8, 9, 10):
        re_.cell(r, c).number_format = '#,##0,;△#,##0,;"-"'
        re_.cell(r, c).alignment = Alignment(horizontal='right', vertical='center')
    re_.cell(r, 5).number_format = '#,##0.00'
    re_.cell(r, 5).alignment = Alignment(horizontal='right', vertical='center')
r = 15
for j, v in enumerate(['合計', '', '', '=E4+E10+E13', '', '', '=SUM(H4:H14)', '=SUM(I4:I14)', '=J4+J10+J13+J14', '']):
    x = re_.cell(r, 2 + j, v if v != '' else None)
    x._style = copy.copy(body[j])
    x.font = Font(name=NUM if j >= 3 else JP, size=12, bold=True)
re_.cell(15, 5).number_format = '#,##0.00'
for c in (8, 9, 10):
    re_.cell(15, c).number_format = '#,##0,;△#,##0,;"-"'
re_.row_dimensions[15].height = 27
re_['H3'] = '前期末簿価\n(千円)'; re_['I3'] = '時価\n(千円)'; re_['J3'] = '固定資産税\n評価額(千円)'
ungroup_cols(re_)
for c, w_ in zip('BCDEFGHIJK', [11, 22, 24, 9, 17, 14, 10, 10, 11, 26]):
    re_.column_dimensions[c].width = w_
mark(re_, 'B3:K15', '不動産詳細')
re_['B17'] = 'チェック：土地簿価＝BS土地'; re_['H17'] = "=H4+H10+H13='BS(借方)'!F28"
re_['B18'] = '※面積は登記地積・床面積（R8年度固定資産税課税明細書）。主要棟は1981〜82年築であり、新耐震基準への適合・アスベスト・修繕履歴は要確認。担保設定は登記簿謄本で要確認。'
re_['B19'] = '※償却資産の課税標準額9,779千円。R8年度固定資産税・都市計画税の年税額1,488千円。'

# --- 不動産(賃借)
rl = wb['不動産(賃借) ']
clear(rl, 4, 21, 2, 10)
rl['B4'] = '該当なし'; rl['C4'] = '賃借物件なし（地代家賃の計上なし）'

# --- 設備（固定資産内訳書：建物の内訳＋科目別）
eq = wb['設備']
st_h = [copy.copy(eq.cell(2, c)._style) for c in range(2, 12)]
st_b = [copy.copy(eq.cell(3, c)._style) for c in range(2, 12)]
st_t = [copy.copy(eq.cell(13, c)._style) for c in range(2, 12)]
clear(eq, 2, 40, 2, 11, style=True)
for j, v in enumerate(['№', '科目', '資産名（用途）', '償却方法', '取得年', '取得月', '耐用年数', '取得原価', '期末帳簿価額\n(千円)', '所在・面積等']):
    x = eq.cell(2, 2 + j, v); x._style = copy.copy(st_h[j])
eq.row_dimensions[2].height = 33
old_eq = old['設備']
items_eq = [(old_eq[f'B{r}'].value, old_eq[f'C{r}'].value, old_eq[f'D{r}'].value, old_eq[f'E{r}'].value) for r in range(6, 32)]
for i, (nm, loc, area, bv) in enumerate(items_eq):
    r = 3 + i
    vals = [i + 1, '建物', nm, None, ('2025' if '2025年5月' in str(nm) else None), ('5' if '2025年5月' in str(nm) else None), None, None, bv,
            f'{loc}' + (f'（{area:,.1f}㎡）' if area else '')]
    for j, v in enumerate(vals):
        x = eq.cell(r, 2 + j, v); x._style = copy.copy(st_b[j])
    eq.cell(r, 10).number_format = '#,##0,;△#,##0,;"-"'
    eq.cell(r, 10).alignment = Alignment(horizontal='right', vertical='center')
    eq.row_dimensions[r].height = 19.35
rT = 3 + len(items_eq)
for j, v in enumerate([None, '建物 計', None, None, None, None, None, None, f'=SUM(J3:J{rT - 1})', None]):
    x = eq.cell(rT, 2 + j, v); x._style = copy.copy(st_t[j])
eq.cell(rT, 10).number_format = '#,##0,;△#,##0,;"-"'
eq.row_dimensions[rT].height = 19.35
hide_cols(eq, 'E', 'F', 'G', 'H', 'I')
eq.column_dimensions['D'].width = 24; eq.column_dimensions['K'].width = 24; eq.column_dimensions['C'].width = 9
mark(eq, f'B2:K{rT}', '保有設備一覧（建物内訳）')
eq[f'B{rT + 1}'] = 'チェック：BS建物と一致'; eq[f'J{rT + 1}'] = f"=J{rT}='BS(借方)'!F22"
# 科目別（別範囲）
r1 = rT + 4
for j, v in enumerate(['科目', '期末帳簿価額\n(千円)', '備考']):
    x = eq.cell(r1, [4, 10, 11][j], v); x._style = copy.copy(st_h[[1, 8, 9][j]])
eq.row_dimensions[r1].height = 33
KM = [('建物', 22, '左表の内訳のとおり。改修工事を資産計上'), ('建物付属設備', 23, '冷蔵庫ユニット・シャッター等（内訳は台帳で要確認）'),
      ('構築物', 24, ''), ('機械装置', 25, '備忘価額のみ（償却済）'), ('車両運搬具', 26, '備忘価額のみ（償却済）。車両は自己所有とリースを併用'),
      ('工具器具備品', 27, ''), ('ソフトウェア', 37, '販売管理システム等と推察（要確認）')]
for i, (nm, br, memo) in enumerate(KM):
    r = r1 + 1 + i
    for c, v, k in [(4, nm, 1), (10, f"='BS(借方)'!F{br}", 8), (11, memo or None, 9)]:
        x = eq.cell(r, c, v); x._style = copy.copy(st_b[k])
    eq.cell(r, 10).number_format = '#,##0,;△#,##0,;"-"'
    eq.cell(r, 10).alignment = Alignment(horizontal='right', vertical='center')
    eq.row_dimensions[r].height = 19.35
rK = r1 + 1 + len(KM)
for c, v, k in [(4, '合計', 1), (10, f'=SUM(J{r1 + 1}:J{rK - 1})', 8), (11, None, 9)]:
    x = eq.cell(rK, c, v); x._style = copy.copy(st_t[k])
eq.cell(rK, 10).number_format = '#,##0,;△#,##0,;"-"'
eq.row_dimensions[rK].height = 19.35
REG.append(('設備', f'設備!D{r1}:K{rK}', '保有設備一覧（科目別）'))
eq['A1'] = eq['A1'].value + f'／D{r1}:K{rK}（科目別）'
eq[f'B{rK + 2}'] = '※固定資産内訳書（2026/3期末）。固定資産台帳は未受領のため、償却方法・取得年月・取得原価、冷蔵設備・フォークリフト・車両等の明細は受領後に追加（非表示列）。'
EQ_REF = (f'設備!B2:K{rT}', f'設備!D{r1}:K{rK}')

# ================================================================ 追加シート（業種に合わせたページ用）
# --- 売上構成（収益源別：進行期試算表）
sc = im_sheet(wb, '売上構成', '収益源別売上高（進行期 2026年4〜9月 試算表）')
cw(sc, {'B': 2.2, 'C': 1.15, 'D': 0.8, 'E': 5.2})
put(sc, 4, 5, '(単位：千円)', 'unit', border=None)
row(sc, 5, 2, ['収益区分（2026年4〜9月）', '金額', '構成比', '内容'], 'hdr', height=19)
SC = [('商品売上高', '=PL!K5', '青果物の販売（買付販売は総額、受託販売は2019年10月以降手数料のみの純額で計上）'),
      ('委託手数料収入', '=PL!K9', '受託販売の手数料。生産者の商品を預かり、販売後に販売額の8.5%を収受'),
      ('主婦センター費（配送受託）', '=PL!K8', '(株)主婦の店の仕入商品の配送受託。配送額の約10%を収受'),
      ('資材売上高', '=PL!K6+PL!K7', '包装資材の販売（自社利用分の振替を含む。肥料の取扱いはほぼなし）'),
      ('売上割戻（完納奨励金）', '=PL!K10', '期限内に代金を支払った八百屋等へ支給（市場の商慣習）')]
for i, (lab, f, memo) in enumerate(SC):
    r = 6 + i
    row(sc, r, 2, [lab, f, f'=C{r}/C$11', memo], ['lab', 'num', 'num', 'txt'], fmts=[None, KI, PI, None], height=21)
row(sc, 11, 2, ['売上高 計', '=SUM(C6:C10)', '=C11/C$11', ''], ['sub', 'subn', 'subn', 'sub'], fmts=[None, KI, PI, None], height=19)
paste(sc, 'B4:E11', '事業の仕組みと収益源', REG)
sc['B13'] = 'チェック：PL売上高と一致'; sc['C13'] = '=C11=PL!K11'
sc['B15'] = '参考（貼付範囲外）'
sc['B16'] = '受託手数料率（第2回面談）'; sc['C16'] = 0.085
sc['B17'] = '受託販売の取扱高（推計：手数料÷率、6か月）'; sc['C17'] = '=C7/C16'; sc['C17'].number_format = '#,##0'
sc['B18'] = 'うち資材売上（自社利用分の振替）'; sc['C18'] = '=PL!K7'; sc['C18'].number_format = '#,##0'
sc['B20'] = '※決算書では売上高を「商品売上高」1本で表示しているため、収益源別の内訳は試算表（2026年4〜9月）で代替。資材売上（自社）は同額を資材仕入（自社）に計上（振替）。'

# --- 品目別売上（分類別商品売上実績表）
bj = json.load(open(f'{S}/docs/bunrui.json'))
fy = {'202403': None, '202503': None, '202603': None}
for k in bj:
    for y in fy:
        if y in k:
            fy[y] = bj[k]
GROUPS = [  # 区分, 表示名, 大分類コード, 主な品目
    ('野菜', '葉茎菜類', ['311'], 'ねぎ・ほうれん草・小松菜'),
    ('野菜', '果菜類', ['341'], 'ミニトマト・胡瓜・なす'),
    ('野菜', '土物類', ['361'], 'たまねぎ・馬鈴薯・さつま芋'),
    ('野菜', 'きのこ・山菜類', ['381'], '生しいたけ・なめこ・わらび'),
    ('野菜', '豆類', ['351'], 'だだちゃ豆・枝豆・いんげん'),
    ('野菜', '根菜類', ['301'], '人参・たけのこ・ごぼう'),
    ('野菜', 'その他野菜', ['371', '391', '398', '701'], 'しょうが・みょうが・食用菊'),
    ('果実', 'メロン類', ['471'], 'アンデスメロン・赤アンデス'),
    ('果実', 'かき類', ['435'], '庄内柿'),
    ('果実', 'なし類', ['430'], '幸水・豊水・ラ・フランス'),
    ('果実', '柑橘類', ['401', '411'], 'みかん・ネーブル・デコポン'),
    ('果実', 'キウイ・バナナ等', ['491'], 'キウイ・バナナ・アボカド'),
    ('果実', 'りんご類', ['420'], 'ふじ・つがる'),
    ('果実', 'その他果実', ['439', '440', '444', '445', '446', '447', '450', '457', '460', '480', '495', '702'], 'いちご・すいか・桃・さくらんぼ'),
    ('花き', '花き類', ['600'], '生花（花束等）'),
    ('加工品', '加工品・その他', ['511', '531', '541', '551', '561', '581', '586', '591', '599', '999'], 'もやし・干物・漬物等'),
]
pj = im_sheet(wb, '品目別売上', '品目別（大分類）売上高の3期推移（分類別商品売上実績表）')
cw(pj, {'B': 0.8, 'C': 1.45, 'D': 0.95, 'E': 0.62, 'F': 0.95, 'G': 0.62, 'H': 0.95, 'I': 0.62, 'J': 2.4})
put(pj, 4, 10, '(単位：千円)', 'unit', border=None)
row(pj, 5, 2, ['区分', '品目（大分類）', "='BS(借方)'!D4", '', "='BS(借方)'!E4", '', "='BS(借方)'!F4", '', '主な品目（2026/3期の上位）'], 'hdr', height=17)
row(pj, 6, 2, ['', '', '売上高', '構成比', '売上高', '構成比', '売上高', '構成比', ''], 'hdr', height=15)
for a, b in [('B5', 'B6'), ('C5', 'C6'), ('J5', 'J6'), ('D5', 'E5'), ('F5', 'G5'), ('H5', 'I5')]:
    pj.merge_cells(f'{a}:{b}')
# 生データ（大分類別）：行40〜
RAW0 = 40
pj.cell(RAW0 - 1, 2, '大分類別の生データ（分類別商品売上実績表の「大分類計」当年欄の合計売上金額）').font = Font(name=JP, size=10, bold=True)
for j, h in enumerate(['大分類', '名称', '集計区分', '2024/3期', '2025/3期', '2026/3期']):
    pj.cell(RAW0, 2 + j, h).font = Font(name=JP, size=10, bold=True)
codes = sorted(set(k[:3] for d in fy.values() for k in d))
code2grp = {c: g[1] for g in GROUPS for c in g[2]}
for i, cd in enumerate(codes):
    r = RAW0 + 1 + i
    nm = next((k[4:] for d in fy.values() for k in d if k[:3] == cd), '')
    pj.cell(r, 2, cd); pj.cell(r, 3, nm); pj.cell(r, 4, code2grp.get(cd, '加工品・その他'))
    for j, y in enumerate(['202403', '202503', '202603']):
        v = next((val[0] for k, val in fy[y].items() if k[:3] == cd), 0)
        x = pj.cell(r, 5 + j, round(v)); x.number_format = '#,##0'
rawN = RAW0 + len(codes)
for i, (kbn, nm, cds, main) in enumerate(GROUPS):
    r = 7 + i
    vals = [kbn, nm]
    for col in 'EFG':
        vals.append(f'=SUMIF($D${RAW0 + 1}:$D${rawN},$C{r},{col}${RAW0 + 1}:{col}${rawN})')
    row(pj, r, 2, [kbn, nm, vals[2], f'=D{r}/D$23', vals[3], f'=F{r}/F$23', vals[4], f'=H{r}/H$23', main],
        ['lab', 'labw', 'num', 'rnum', 'num', 'rnum', 'num', 'rnum', 'txt'], fmts=[None, None, KI, PI, KI, PI, KI, PI, None], height=17)
    for c in 'EGI':
        pj[f'{c}{r}'].font = Font(name=NUM, size=9)
    pj[f'J{r}'].font = Font(name=JP, size=9)
for kb, (a, b) in {'野菜': (7, 13), '果実': (14, 20)}.items():
    pj.merge_cells(f'B{a}:B{b}')
row(pj, 23, 2, ['合計', '', '=SUM(D7:D22)', '=D23/D$23', '=SUM(F7:F22)', '=F23/F$23', '=SUM(H7:H22)', '=H23/H$23', ''],
    ['sub', 'sub', 'subn', 'subn', 'subn', 'subn', 'subn', 'subn', 'sub'], fmts=[None, None, KI, PI, KI, PI, KI, PI, None], height=19)
pj.merge_cells('B23:C23')
paste(pj, 'B4:J23', '品目別売上高の3期推移', REG)
pj['B25'] = '野菜 計'; pj['B26'] = '果実 計'; pj['B27'] = '県内／県外（2026/3期、売上金額）'
for c in 'DFH':
    pj[f'{c}25'] = f'=SUM({c}7:{c}13)'; pj[f'{c}26'] = f'=SUM({c}14:{c}20)'
    pj[f'{c}25'].number_format = '#,##0'; pj[f'{c}26'].number_format = '#,##0'
pj['D27'] = 829653005; pj['F27'] = 464272577; pj['D27'].number_format = '#,##0'; pj['F27'].number_format = '#,##0'
pj['B28'] = 'チェック：生データ合計と一致'; pj['D28'] = f'=ROUND(D23-SUM(E{RAW0 + 1}:E{rawN}),0)=0'
pj['B30'] = '※出所：分類別商品売上実績表（販売管理システム、各期3月末現在の年間累計）。決算書の商品売上高とは集計範囲が異なる（受託販売分を販売額の総額で集計していると推察：要確認）。'
pj['B31'] = '※「県内／県外」の区分が産地か販売先かは要確認。'

# --- 不動産概況（不動産(所有)から数式で集計）
gk = im_sheet(wb, '不動産概況', '不動産概況（不動産(所有)シートから集計）')
cw(gk, {'B': 1.05, 'C': 2.25, 'D': 0.85, 'E': 0.75, 'F': 1.75, 'G': 0.95, 'H': 0.95, 'I': 1.05, 'J': 0.75})
put(gk, 4, 10, '(単位：千円)', 'unit', border=None)
row(gk, 5, 2, ['拠点', '所在地', '土地面積\n(㎡)', '土地の\n所有区分', '建物の所有区分', '土地簿価', '土地時価\n（推計）', '建物固定資産税\n評価額', '担保'], 'hdr', height=30)
GK = [('本社・市場', "='不動産(所有)'!D4", "='不動産(所有)'!E4", '法人', '法人（5棟／延床1,246.20㎡）', "='不動産(所有)'!H4", "='不動産(所有)'!I4", "=SUM('不動産(所有)'!J5:J9)", '要確認'),
      ('パッケージ場', "='不動産(所有)'!D10", "='不動産(所有)'!E10", '法人', '法人（2棟／延床766.48㎡）', "='不動産(所有)'!H10", "='不動産(所有)'!I10", "=SUM('不動産(所有)'!J11:J12)", '要確認'),
      ('旧倉庫敷地', "='不動産(所有)'!D13", "='不動産(所有)'!E13", '法人', '建物なし（解体済み）', "='不動産(所有)'!H13", "='不動産(所有)'!I13", 0, '要確認')]
for i, rec_ in enumerate(GK):
    r = 6 + i
    row(gk, r, 2, list(rec_), ['lab', 'txt', 'num', 'txtc', 'txt', 'num', 'num', 'num', 'txtc'], fmts=[None, None, AI, None, None, KI, KI, KI, None], height=24)
row(gk, 9, 2, ['合計', '', '=SUM(D6:D8)', '', '法人 7棟', '=SUM(G6:G8)', '=SUM(H6:H8)', '=SUM(I6:I8)', ''],
    ['sub', 'sub', 'subn', 'sub', 'sub', 'subn', 'subn', 'subn', 'sub'], fmts=[None, None, AI, None, None, KI, KI, KI, None], height=19)
paste(gk, 'B4:J9', '不動産概況', REG)
gk['B11'] = 'チェック：土地簿価＝BS土地'; gk['G11'] = "=G9='BS(借方)'!F28"
gk['B12'] = '※土地時価＝固定資産税評価額÷0.7（公示価格水準への割戻し。推計）。建物評価額は令和8年度固定資産税評価額。担保設定は登記簿謄本（未受領）で確認予定。'

# --- 許認可
kk = im_sheet(wb, '許認可', '許認可一覧')
cw(kk, {'B': 2.0, 'C': 2.0, 'D': 3.3, 'E': 2.05, 'F': 1.6})
row(kk, 5, 2, ['許認可等', '根拠法令・所管', '内容', '認定日・効力発生日', '備考'], 'hdr', height=19)
KK = [('地方卸売市場の認定', '卸売市場法\n（山形県知事）', '卸売市場の名称：株式会社鶴岡中央青果\n所在地：山形県鶴岡市宝田一丁目4番20号',
       '2020年6月19日認定\n（効力発生日 2020年6月21日）', '改正卸売市場法の施行に伴う認定'),
      ('地方卸売市場の変更の認定', '卸売市場法第14条\n（山形県知事）', '認定事項の変更（変更内容は要確認）', '2026年2月13日認定\n（効力発生日 2026年4月1日）', '認定証の写しを受領'),
      ('倉庫業', '倉庫業法', '定款の目的に記載', '要確認', '登録の有無は要確認')]
for i, rec_ in enumerate(KK):
    row(kk, 6 + i, 2, list(rec_), ['lab', 'txt', 'txt', 'txt', 'txt'], height=36)
paste(kk, 'B5:F8', '許認可一覧', REG)
kk['B10'] = '※出所：地方卸売市場の認定証・変更の認定証（山形県）、履歴事項全部証明書。認定の条件（取扱品目・開場日等）は業務規程で要確認。'

# --- 旧マスターから複製（値）：月別売上・沿革・その他補足・その他
for nm in ['月別売上', '沿革', 'その他補足', 'その他']:
    dst = wb.create_sheet(nm if nm != 'その他' else 'その他_tmp')
    copy_sheet(old[nm], dst)
REG.append(('月別売上', '月別売上!B4:O8', '月別売上高（季節性）'))
REG.append(('沿革', '沿革!B4:C16', '沿革'))
REG.append(('その他補足', 'その他補足!B4:C15', 'その他補足情報'))
# 沿革の更新（認定証の受領）
en = wb['沿革']
HIST = [('1973年11月', '株式会社鶴岡中央青果を設立'),
        ('1981〜82年', '現本社地（宝田一丁目）に事務所・市場・倉庫・パッケージ場（事務所兼倉庫）を建築'),
        ('1984年11月', '住居表示の実施により本店所在地の表記が「鶴岡市宝田一丁目4番20号」となる'),
        ('2006年', '会社法施行に伴い取締役会設置会社・監査役設置会社の登記'),
        ('2012年', '定款の目的を変更（現行：青果物の販売及び卸売、倉庫業ほか）。宝田一丁目5-1に倉庫（鉄骨造139㎡）を新築'),
        ('2019年10月', '受託販売の会計表示を総額から純額（受託手数料のみ）へ変更'),
        ('2020年6月', '改正卸売市場法の施行に伴い、山形県知事より地方卸売市場の認定を受ける'),
        ('2025年5月', '資材倉庫の外壁・シャッターを改修'),
        ('2026年1月', '鶴岡市内のもう1社の青果市場が撤退し、市内で唯一の青果市場となる'),
        ('2026年2月', '地方卸売市場の変更の認定（2026年4月1日効力発生。変更内容は要確認）'),
        ('2026年5月', '役員改選（取締役3名重任、常務取締役退任）'),
        ('時期要確認', '庄内町長沼の倉庫を老朽化により解体（跡地は駐車場として賃貸）')]
st_l = copy.copy(en['B5']._style); st_r = copy.copy(en['C5']._style)
for i, (a, b) in enumerate(HIST):
    r = 5 + i
    en[f'B{r}'] = a; en[f'C{r}'] = b
    en[f'B{r}']._style = copy.copy(st_l); en[f'C{r}']._style = copy.copy(st_r)
    en.row_dimensions[r].height = en.row_dimensions[5].height
en['B17'] = None
en['A2'] = '■ 貼付範囲 B4:C16 → 企業概要書「沿革」（範囲をコピーしてパワポに貼り付け。数字はここを直す）'
en['B18'] = '※出所：履歴事項全部証明書、地方卸売市場の認定証、固定資産税課税明細書、個別注記表、固定資産内訳書。創業の経緯・代表者の経歴は売主様へ確認中。'
en['B18'].font = Font(name='游ゴシック Medium', size=9)
ho = wb['その他補足']
ho['C5'] = '・地方卸売市場（卸売市場法）：2020年6月に山形県知事の認定、2026年2月に変更の認定（「許認可一覧」参照）\n・倉庫業：定款目的に記載（登録の有無は要確認）'

# --- その他（要確認事項・社内確認事項）：テンプレートの空シートへ移す
ot = wb['その他']
src_ot = wb['その他_tmp']
copy_sheet(src_ot, ot)
wb.remove(src_ot)
ot['A2'] = '最終更新：2026年10月27日（追加資料〔9月試算表・買参人別／荷主別順位表・分類別商品売上実績表・認定証・賃金台帳〕を反映）。★＝重要論点'
upd = {
    6: ('一部回答', '回答資料受領（買参人別・荷主別順位表、上位20位×3期）。(株)主婦の店鶴岡店は2026/3期345,227千円（決算書売上高の28.9%）で依存度が高い。社内勘定（KK・パッケージ）の性質は要確認'),
    24: ('', '受領（10/27）：2026年4〜9月の営業利益20,426千円・経常利益21,166千円（前年同期9,454千円・10,013千円）。売上高は前年同期比△6.9%、販管費△13.3%（人件費・運賃・減価償却費）'),
    5: ('', '回答資料受領：地方卸売市場の認定証（2020/6/19認定）・変更の認定証（2026/2/13認定、2026/4/1効力）。変更内容・仲卸業者数・売買参加者数は要確認'),
    36: ('', '分類別商品売上実績表（3期）を受領し「品目別売上」シートに集計。決算書の商品売上高とは集計範囲が異なる（差異の内訳は要確認）'),
}
for r, (_, txt) in upd.items():
    ot[f'F{r}'] = txt
new_issues = [
    ('財務', '★2026年8〜9月に長期借入金が約1.65億円増加（3月末361,058千円→9月末526,264千円）、現預金は393,411千円。借入先・資金使途・条件', '高', '要確認：評価基準日後の借入。ネットデット・株式価値への影響を確認'),
    ('事業', '★買参人・荷主の双方に計上される社内勘定（KK 地物果実・マルカ・開発・社長車、パッケージ根菜・果実・蔬菜・生花）の性質（集荷担当別の勘定か、部門間の振替か）', '高', '要確認：売上高・仕入高を両建てで膨らませていないか（2026/3期 販売側333,489千円・仕入側328,220千円）'),
    ('事業', '分類別商品売上実績表の合計（約12.9億円）と決算書の商品売上高（約12.0億円）の差異', '中', '要確認：受託販売分の総額集計等と推察'),
    ('人事', '賃金台帳（2026/4〜10）の氏名と職員名簿№の対応（従業員一覧に月額総支給を記載するため）', '中', '要確認。台帳では2026年9月に2名退職（近郷部）'),
    ('財務', '投資有価証券が2026/3末2,600千円→9月末1,600千円に減少（(株)庄果株式の処分有無）', '中', '要確認'),
]
last = max(r for r in range(4, 46) if ot[f'B{r}'].value is not None and isinstance(ot[f'B{r}'].value, int))
# 既存の問題一覧（〜41行）の直後には Web出所（43-44行）があるため、出所行を下へずらして追記する
web1, web2 = ot['C43'].value, ot['C44'].value
proc = [[ot.cell(r, c).value for c in range(2, 7)] for r in range(46, 64)]
proc_st = [[copy.copy(ot.cell(r, c)._style) for c in range(2, 7)] for r in range(46, 64)]
clear(ot, 42, 70, 2, 6, style=True)
st_row = [copy.copy(ot.cell(41, c)._style) for c in range(2, 7)]
for i, (k, d, p, f) in enumerate(new_issues):
    r = 42 + i
    for j, v in enumerate([last + 1 + i, k, d, p, f]):
        x = ot.cell(r, 2 + j, v); x._style = copy.copy(st_row[j])
    ot.row_dimensions[r].height = ot.row_dimensions[41].height
rr = 42 + len(new_issues) + 1
ot[f'C{rr}'] = web1; ot[f'C{rr + 1}'] = web2
rp = rr + 3
for i, (vals, sts) in enumerate(zip(proc, proc_st)):
    for j, (v, s_) in enumerate(zip(vals, sts)):
        x = ot.cell(rp + i, 2 + j, v); x._style = s_
ot.column_dimensions['A'].width = 2

# ================================================================ シート順・ルール
order = wb.sheetnames
def move_after(name, after):
    sh_ = wb[name]; wb._sheets.remove(sh_)
    wb._sheets.insert(wb._sheets.index(wb[after]) + 1, sh_)
move_after('売上構成', 'BSハイライト')
move_after('品目別売上', '売上構成')
move_after('月別売上', '外注先')
move_after('不動産概況', '拠点(本社等)')
move_after('許認可', '設備')
move_after('沿革', '許認可')
move_after('その他補足', '沿革')

ru = wb['ルール']
ru['B16'] = '■ 本案件の運用（株式会社鶴岡中央青果）'; ru['B16'].font = Font(name=JP, size=11, bold=True, color=NAVY)
RULES = ['本ファイルが唯一の数値ソース。企業概要書・評価書の数字は必ず本ファイルから貼る（パワポに直接打たない）。',
         '財務入力：D列=24年3月期（第51期）、E列=25年3月期、F列=26年3月期（直近期・評価基準日）、G列=進行期（2026年4〜9月の試算表、6か月累計）。1期目（C列）は未使用のため非表示。',
         '非表示の行・列は概要書に貼らない（空行・1期目・売上行・氏名等）。貼付スクリプト（im_paste.paste_range）も非表示を自動で除く。',
         '未確認事項は推測で埋めず「要確認」。推測は「推察」「想定」「（仮説）」と明記し、「その他」シートに集約。社内限の情報は「その他」の「プロセス関連（社内確認事項）」のみに記載。',
         '更新：2026/10/27 公式フォーマットへ移行（旧マスター261019の内容を転記）。9月試算表・順位表・分類別売上・認定証を反映。SGA 26/3期の行ずれ、テンプレートの数式不具合（年買法G8/H8、ネットキャッシュF15、修正BS(貸方)D50）を修正。']
for i, t in enumerate(RULES):
    ru[f'B{17 + i}'] = f'{i + 1}. {t}'; ru[f'B{17 + i}'].font = Font(name=JP, size=10)
r = 23
ru[f'B{r}'] = '■ 企業概要書への貼付範囲一覧（各IMシートのA1/A2にも表記）'; ru[f'B{r}'].font = Font(name=JP, size=11, bold=True, color=NAVY)
for key, rng, pg in REG + [('BS', "'BS(借方)'!B3:G61", '貸借対照表（資産の部）'), ('BS', "'BS (貸方)'!B3:G25", '貸借対照表（負債の部）'),
                           ('BS', "'BS (貸方)'!B34:G51", '貸借対照表（純資産の部）'), ('PL', 'PL!B2:L56', '損益計算書'), ('SGA', 'SGA!B2:L40', '販売費及び一般管理費')]:
    r += 1
    ru[f'B{r}'] = f'   {rng}　→　{pg}'; ru[f'B{r}'].font = Font(name=NUM, size=10)
bsd['A1'] = '■ 貼付範囲 B3:G61 → 企業概要書「貸借対照表（資産の部）」（C列・空行は非表示）'
bsc['A1'] = '■ 貼付範囲 B3:G25（負債）／B34:G51（純資産）→ 企業概要書「貸借対照表」'
pl['A1'] = '■ 貼付範囲 B2:L56 → 企業概要書「損益計算書」（C:D列・空行は非表示）'
sga['A1'] = '■ 貼付範囲 B2:L40 → 企業概要書「販売費及び一般管理費」（売上行・C:D列・空行は非表示）'
for ws_ in (bsd, bsc, pl, sga):
    ws_['A1'].font = Font(name='游ゴシック Medium', size=9, color='C00000')
    ws_['A1'].alignment = Alignment(wrap_text=False)

out = f'{S}/master/master_raw.xlsx'
import os; os.makedirs(f'{S}/master', exist_ok=True)
wb.save(out)
json.dump({'REG': REG, 'EQ': EQ_REF}, open(f'{S}/master/paste_index.json', 'w'), ensure_ascii=False)
print('saved', out, len(REG))
