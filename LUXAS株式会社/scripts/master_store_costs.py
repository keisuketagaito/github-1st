# -*- coding: utf-8 -*-
"""案件マスターに『店舗別費用』シートを追加する。

出所：店舗別（部門別）勘定科目残高推移表 2025年12月〜2026年6月（2026年9月18日出力・11部門）。
既存の『店舗別損益』は売上高／売上総利益／販管費／営業利益の4指標のみだったため、
販管費を科目別に分解したシートを追加する。全科目の部門合計が全社PL（2025年12月〜
2026年6月の勘定科目残高推移表）と一致することを検算済み。

このスクリプトは冪等。既に『店舗別費用』があれば作り直す。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

MASTER = '/home/user/github-1st/LUXAS株式会社/案件マスター_LUXAS株式会社_20260923.xlsx'
SHEET  = '店舗別費用'

JP, JPM, NUM = '游ゴシック', '游ゴシック Medium', 'Arial'
NAVY   = 'FF0B3041'
HDRFIL = PatternFill('solid', fgColor='FF0B3041')
SUBFIL = PatternFill('solid', fgColor='FFE1F3FB')
FMT_K  = '#,##0,;[RED]"△ "#,##0,;\\－'
FMT_P  = '0.0%;\\-0.0%;\\－'
THIN   = Side(style='thin', color='FFBFBFBF')
BOX    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# (№, 部門, 状況, 売上高, 従業員給与, 法定福利費, 役員報酬, ロイヤリティ, 地代家賃,
#  広告宣伝費, 支払手数料, 販管費計)   ※単位：円。その他は販管費計からの差額で算出する
ROWS = [
 (1, 'ワンラブ カインズホーム明和店',        '稼働中',
     25108810, 11070780, -1054506, 0, 3467132,        0, 2310000,  655922, 18976890),
 (2, 'ワンラブ バローミタス伊勢店',          '稼働中',
     40221277,  8234159,  -967806, 0, 4826556,        0, 1540000,  110550, 14171474),
 (3, 'ワンラブ ホームセンターバロー久居店',  '稼働中',
     33840444,  7922592, -1058795, 0, 4399257,        0, 1540000,  155900, 13333435),
 (4, 'ワンラブ ドン・キホーテ緑店',          '稼働中',
     36963203,  6677386,  -709546, 0,       0,  5407317,  770000, 1043292, 13900076),
 (5, 'おっきなもふもふ応援隊 岐阜総本店',    '稼働中',
     60022946, 13862144, -1525287, 0,       0,  5160083, 2459820, 2012786, 22723774),
 (6, 'おっきなもふもふ応援隊 大阪総本店',    '稼働中',
     58446704, 11461033, -1133275, 0,       0,  6935929, 2073500, 2100129, 22221605),
 (7, 'ワンラブ アクアウォーク大垣店',        '2026年2月に撤退',
      6112430,  1059512,  -175480, 0,  916864,    30800,       0,     203,  1914906),
 (8, 'ワンラブ ペットプラザ岐阜店',          '2026年1月に撤退',
      1842692,  1093462,   -97371, 0,  583202,        0,       0,       0,  1755945),
 (None, 'セキュリティ事業',                  '－',
      2944810,        0,        0, 0,       0,        0,       0,       0,        0),
 (None, '大垣管理センター',                  '2026年2月に閉鎖',
            0,   348658,  -251815, 0,       0,        0,       0,       0,   113163),
 (None, '共通部門（本社）',                  '－',
     13650127,  2450000,  9752797, 6300000, 0,  2345000, 4469888,  848821, 58643080),
]
SUB_AT = 6          # 稼働6店舗 小計をこの行数のあとに差し込む

HEAD = ['№', '部門', '状況', '売上高', '従業員給与', '法定福利費', '人件費計', '人件費率',
        '役員報酬', 'ロイヤリティ', '地代家賃', '広告宣伝費', '支払手数料', 'その他',
        '販管費計', '販管費率']
WIDTH = {'A': 13, 'B': 4, 'C': 32, 'D': 15, 'E': 12, 'F': 12, 'G': 12, 'H': 11, 'I': 9,
         'J': 11, 'K': 12, 'L': 12, 'M': 12, 'N': 12, 'O': 11, 'P': 12, 'Q': 9}
NOTE = (
 '※出所：店舗別（部門別）勘定科目残高推移表 2025年12月〜2026年6月（2026年9月18日出力・11部門）。'
 '従業員給与・法定福利費・役員報酬・ロイヤリティ・地代家賃・広告宣伝費・支払手数料の各科目について、'
 '11部門の合計が全社PL（64,179,726／2,778,916／6,300,000／14,193,011／19,879,129／15,163,208／'
 '6,927,603円）と一致することを確認済。\n'
 '※「その他」は販管費計から上記各科目を控除した差額（旅費交通費・厚生費・リース料・通信費・'
 '水道光熱費・租税公課・接待交際費・保険料・備品消耗品費・管理諸費・雑費等）。\n'
 '※法定福利費は本社（共通部門）に一括計上したうえで各店舗へマイナス計上（振替）する処理が'
 'とられており、店舗側は各店ともマイナス残高となる。したがって店舗別の販管費は社会保険料負担分だけ'
 '過少に表示されている点に留意を要する（店舗計△6,449千円）。\n'
 '※ロイヤリティはワンラブFC加盟店のうち明和・伊勢・久居・大垣・岐阜の5店のみが負担し、'
 'ドン・キホーテ緑店は負担がない一方で地代家賃5,407千円（売上高比14.6%）を負担している。'
 '緑店のみ契約形態が異なる理由は要確認。\n'
 '※広告宣伝費はFC加盟店で月額220千円（緑店は110千円）の定額計上。'
 '路面店2店は月額300〜360千円程度で、こちらも概ね定額である。')


def style(cell, *, font=JP, size=9, bold=False, color=None, fmt=None,
          fill=None, halign=None, valign='bottom', wrap=False):
    cell.font = Font(name=font, size=size, bold=bold, color=color)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    cell.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    cell.border = BOX


def build():
    wb = openpyxl.load_workbook(MASTER)
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, wb.sheetnames.index('店舗別損益') + 1)
    ws.sheet_view.showGridLines = False
    for col, w in WIDTH.items():
        ws.column_dimensions[col].width = w

    c = ws.cell(row=2, column=2, value='店舗別（部門別）販管費の内訳　進行期7か月累計（2025年12月〜2026年6月）')
    c.font = Font(name=JPM, size=11, bold=True, color=NAVY)
    c.border = Border()
    c = ws.cell(row=2, column=16, value='(単位：千円)')
    c.font = Font(name=JPM, size=9)
    c.alignment = Alignment(horizontal='right')
    c.border = Border()

    for i, h in enumerate(HEAD):
        style(ws.cell(row=4, column=2 + i, value=h), font=JPM, bold=True,
              color='FFFFFFFF', fill=HDRFIL, halign='center', valign='center', wrap=True)
    ws.row_dimensions[4].height = 28

    def emit(r, no, name, state, sales, pay, ins, exe, roy, rent, ad, fee, sga,
             bold=False, fill=None):
        """1行書き出す。列は B=№ C=部門 D=状況 E=売上高 F=従業員給与 G=法定福利費
           H=人件費計 I=人件費率 J=役員報酬 K=ロイヤリティ L=地代家賃 M=広告宣伝費
           N=支払手数料 O=その他 P=販管費計 Q=販管費率。"""
        vals = [no, name, state, sales, pay, ins,
                f'=F{r}+G{r}', f'=IFERROR(H{r}/$E{r},"")',
                exe, roy, rent, ad, fee,
                f'=P{r}-(F{r}+G{r}+J{r}+K{r}+L{r}+M{r}+N{r})', sga,
                f'=IFERROR(P{r}/$E{r},"")']
        for i, v in enumerate(vals):
            col = 2 + i
            cell = ws.cell(row=r, column=col, value=v)
            if col == 2:
                style(cell, font=NUM, bold=bold, halign='center', fill=fill)
            elif col in (3, 4):
                style(cell, font=JP, bold=bold, fill=fill)
            elif col in (9, 17):
                style(cell, font=NUM, bold=bold, fmt=FMT_P, fill=fill)
            else:
                style(cell, font=NUM, bold=bold, fmt=FMT_K, fill=fill)
        ws.row_dimensions[r].height = 15

    SUMCOLS = ('E', 'F', 'G', 'J', 'K', 'L', 'M', 'N', 'P')

    r, store_rows, body_rows = 5, [], []
    for i, (no, name, state, *nums) in enumerate(ROWS):
        if i == SUB_AT:                                   # 稼働6店舗 小計
            emit(r, None, '【稼働6店舗 小計】', '－', *[None] * 9, bold=True, fill=SUBFIL)
            for col in SUMCOLS:
                ws[f'{col}{r}'] = f'=SUM({col}{store_rows[0]}:{col}{store_rows[-1]})'
            r += 1
        emit(r, no, name, state, *nums)
        if i < SUB_AT:
            store_rows.append(r)
        body_rows.append(r)
        r += 1

    total = r
    emit(total, None, '【合計】', '－', *[None] * 9, bold=True, fill=SUBFIL)
    for col in SUMCOLS:
        ws[f'{col}{total}'] = '=' + '+'.join(f'{col}{x}' for x in body_rows)

    nt = ws.cell(row=total + 2, column=2, value=NOTE)
    nt.font = Font(name=JP, size=9)
    nt.alignment = Alignment(wrap_text=True, vertical='top')
    nt.border = Border()
    ws.freeze_panes = 'D5'
    wb.save(MASTER)
    print('saved:', MASTER, '/ sheet:', SHEET, '/ rows 5-%d' % total)


if __name__ == '__main__':
    build()
