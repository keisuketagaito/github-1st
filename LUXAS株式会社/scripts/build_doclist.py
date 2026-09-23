# -*- coding: utf-8 -*-
"""必要資料一覧（LUXAS株式会社／売主様ご提出用）を当社テンプレートから作成する。
   テンプレート：85fc53f6（当社標準の必要資料一覧）
   ブロック構成 Ⅰ3件/Ⅱ5件/Ⅲ5件/Ⅳ4件/Ⅴ2件 の各ブロック内の行を削って
   Ⅰ2件/Ⅱ3件/Ⅲ4件/Ⅳ2件/Ⅴ2件（計13件）にする。
   ブロック最終行（罫線の下線を持つ行）は残すため、罫線の作り直しは不要。"""
import openpyxl
from openpyxl.styles import Alignment

SRC = '/root/.claude/uploads/fe5370cf-c5f7-5843-870c-82796c774b08/85fc53f6-____________________.xlsx'
OUT = '/home/user/github-1st/LUXAS株式会社/必要資料一覧_LUXAS株式会社_20260923.xlsx'

wb = openpyxl.load_workbook(SRC)
ws = wb['Sheet1']

# ブロック内の行を削除（降順）。各ブロックの最終行は残す。
for r in (18, 17, 12, 8, 7, 4):
    ws.delete_rows(r)

ws['B1'] = '必要資料一覧　　LUXAS株式会社 御中'
ws['B2'] = ('いずれもコピーまたはPDF等のデータにて、ご用意ください。コピーやPDF化が困難な場合は、原本にてご準備いただければ弊社で対応致しますのでご相談ください。\n'
            '※企業概要書の更新と買手候補へのご提示に必要な最低限の資料に絞っています。'
            '※2026年9月23日時点。「受領」欄に✔のある資料は受領済です。')

ITEMS = [
 # (行, カテゴリ記号, カテゴリ名, 資料名, 受領済)
 (3,  'Ⅰ', '概要', '定款（最新のもの）', False),
 (4,  None, None,  '株主名簿（株主名・持株数・持株比率がわかるもの）', False),
 (5,  'Ⅱ', '財務', '決算書・勘定科目内訳明細書・法人税／消費税申告書（直近3期分）', True),
 (6,  None, None,  '月次試算表（2026年7月分以降）　※2026年6月分までは受領済', False),
 (7,  None, None,  '借入金の返済予定表（金融機関別・直近のもの）', False),
 (8,  'Ⅲ', '事業', '店舗別の売上・粗利がわかる資料（2025年12月〜2026年6月）', True),
 (9,  None, None,  '撤退2店舗（アクアウォーク大垣店・ペットプラザ岐阜店）の原状回復費用・違約金がわかる資料', False),
 (10, None, None,  '棚卸資産の内訳（2025年11月期末）　※2026年3月計上の廃棄損44,000千円の処分方法・証憑もあわせてご教示ください', True),
 (11, None, None,  'ペット保険の代理店手数料がわかる資料（直近1期分の精算書等）', False),
 (12, 'Ⅳ', '人事', '従業員名簿（氏名・生年月日・入社年月日・雇用形態・勤務店舗がわかるもの）', True),
 (13, None, None,  '給与台帳（直近1期分｜直近の年末調整の資料等）', True),
 (14, 'Ⅴ', '契約', 'フランチャイズ契約書（有限会社ワンラブ）', False),
 (15, None, None,  '店舗の賃貸借契約書（稼働6店舗分。FC本部からの転借契約を含む）', False),
]
for r, mark, cat, name, got in ITEMS:
    ws.cell(row=r, column=2).value = mark          # B列：ローマ数字
    ws.cell(row=r, column=3).value = cat           # C列：カテゴリ
    ws.cell(row=r, column=5).value = name          # E列：資料名
    # F列：チェックボックス機能は openpyxl で保持できないため「✔／空欄」に置き換える
    ws.cell(row=r, column=6).value = '✔' if got else None
    ws.cell(row=r, column=6).alignment = Alignment(horizontal='center', vertical='center')

# テンプレートに残っていた不規則な塗りを一旦すべて消し、受領済の行だけ淡いグレーを掛ける
from openpyxl.styles import PatternFill
NOFILL = PatternFill(fill_type=None)
GOTFILL = PatternFill('solid', fgColor='F2F2F2')
for r, _m, _c, _n, got in ITEMS:
    for col in (2, 3, 4, 5, 6):
        ws.cell(row=r, column=col).fill = GOTFILL if got else NOFILL

last = ITEMS[-1][0]
ws.print_area = f'B1:F{last}'
ws.row_dimensions[2].height = 51           # 前書き3行分
for _r in (8, 10):                          # 文字数の多い行だけ高さを確保
    ws.row_dimensions[_r].height = 30.75
wb.save(OUT)
print('saved:', OUT)
