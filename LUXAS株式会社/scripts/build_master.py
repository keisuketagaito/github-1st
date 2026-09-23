# -*- coding: utf-8 -*-
"""案件マスター（LUXAS株式会社）を受領資料から作成する。
   数値は円単位で入力（表示形式は千円単位に設定済み）。"""
import openpyxl, copy, datetime
import luxas_prog as LP
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

SRC = 'master.xlsx'
OUT = '/home/user/github-1st/LUXAS株式会社/案件マスター_LUXAS株式会社_20260923.xlsx'
wb = openpyxl.load_workbook(SRC)

YG   = '游ゴシック'
YGM  = '游ゴシック Medium'
ARIAL= 'Arial'

def put(ws, addr, val):
    ws[addr] = val

def clear(ws, cells):
    for c in cells:
        ws[c] = None


def sset(ws, row, col, val):
    """MergedCell を避けて値をセットする"""
    try:
        ws.cell(row=row, column=col).value = val
    except AttributeError:
        pass

def copy_style(ws, src_row, dst_row, cols):
    for col in cols:
        s = ws.cell(row=src_row, column=col)
        d = ws.cell(row=dst_row, column=col)
        d._style = copy.copy(s._style)

# ---------------------------------------------------------------- 期ヘッダー
bs = wb['BS(借方)']
bs['C4'] = '－'          # 第2期（決算書未受領）
bs['D4'] = '23年11月期'
bs['E4'] = '24年11月期'
bs['F4'] = '25年11月期'
bs['G4'] = '26年11月期\n（7か月TB）'

# ---------------------------------------------------------------- 進行期の差替え
# 2026年9月18日出力の試算表（2025年12月〜2026年6月）で進行期を7か月分に更新する。
# 旧試算表（2026年6月26日出力・〜4月）から月次が一部修正されているため、新試算表を正とする。
def _rowof(ws, label):
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == label:
            return r
    raise KeyError(label)

_bsd, _bsc = wb['BS(借方)'], wb['BS (貸方)']
for _lab, _v in LP.BS_PROG_D.items():
    _bsd.cell(row=_rowof(_bsd, _lab), column=7).value = _v
for _lab, _v in LP.BS_PROG_C.items():
    _bsc.cell(row=_rowof(_bsc, _lab), column=7).value = _v

_pl = wb['PL']
for _lab in ('生体売上高', 'サービス売上高', '物販売上高', '期首棚卸高', '商品仕入高', '外注費',
             '他勘定振替高', '期末棚卸高', '受取利息', '受取配当金', '雑収入', '支払利息',
             '雑損失', '棚卸資産廃棄損', '法人税等'):
    _pl.cell(row=_rowof(_pl, _lab), column=11).value = LP.PL_PROG.get(_lab)
_pl['K59'] = LP.NI_PROG          # 当期純利益の整合チェック用

_sga = wb['SGA']
for _lab, _v in LP.SGA_PROG.items():
    _sga.cell(row=_rowof(_sga, _lab), column=11).value = _v or None


# ---------------------------------------------------------------- 会社概要
ws = wb['会社概要']
put(ws,'C3','LUXAS株式会社［ラクサス］（法人番号 4200001038172）\nhttps://luxas.co.jp/　※HPの有無・URLは要確認')
put(ws,'C4','岐阜県大垣市青柳町3丁目227番地2\n※代表者自宅を法人が賃借（貸主：個人／年額1,000千円）')
put(ws,'C5','直営店舗8箇所（岐阜県3・三重県3・愛知県1・大阪府1）、\n大垣管理センター、松阪管理センター\n※2025年11月期「売上高等の事業所別内訳書」ベース')
put(ws,'C6','2019年8月 個人事業として創業／2020年12月 法人設立（第5期）')
put(ws,'C7','1,000千円')
put(ws,'C8','代表取締役　淺野　ゆう子（1985年11月生）')
put(ws,'C9','淺野　ゆう子（100株／100%）\n※株券不発行会社、譲渡制限規定あり（株主総会承認）、設立時から変更なし')
put(ws,'C10','ペットショップ（生体販売・ペット用品販売）、トリミング、ペットホテル、\nペット保険代理店、セキュリティ機器販売')
put(ws,'C11','73名（役員1名、正社員20名、アルバイト53名）（2026年3月時点）')
put(ws,'C12','第一種動物取扱業許可（種別：販売／保管）　8事業所で取得')
put(ws,'C13','558,390千円（2025年11月期）')
put(ws,'C14','6,681千円（2025年11月期）／過去3期平均 18,185千円')
put(ws,'C15','成長戦略（資本力・経営基盤のある企業グループ下での多店舗展開の加速）')

# ---------------------------------------------------------------- PLハイライト（調整項目）
ws = wb['PLハイライト']
ws['C3'] = '－'
ws['D3'] = '2023年11月期'
ws['E3'] = '2024年11月期'
ws['F3'] = '2025年11月期'
ws['G3'] = '進行期(7か月TB)'
ws['I4'] = None

# 調整方針：M&A後に継続的に発生しない費用を戻し入れ、継続的に必要な費用は控除する。
# 経営者人件費の前提：代表者(淺野ゆう子氏)は引継完了後に退任し、現場責任者(淺野崇氏)が
# 役員に就任して現行と同水準の役員報酬を受ける想定。よって①代表者報酬を戻し入れ、
# ②同額の後任役員報酬を控除し、③淺野崇氏の従業員給与・社会保険料を戻し入れる。
adj = {
    18: ('役員報酬（退任予定）＿淺野ゆう子氏',            None,  9600000,  9600000, 14400000, None),
    19: ('役員報酬（派遣想定）＿後任代表者（現行と同水準）', None, -9600000, -9600000,-14400000, None),
    21: ('支払保険料＿事業関連性の低い生命保険料',          None,   696000,  1599000,  1224000, None),
    22: ('接待交際費＿適正水準4,000千円との差額',           None,  4453000,  6230000,  4218000, None),
    24: ('非経常的販管費＿非事業用車両リース料',            None,  2386000,  3154000,  5604000, None),
    25: ('経常的営業外収益＿雑収入のうち社員割引',          None,  1321000,   635000,   654868, None),
    26: ('現場責任者人件費の役員報酬への振替',              None,        0,  8831000,  8183000, None),
}
for row, (label, c, d, e, f, g) in adj.items():
    ws.cell(row=row, column=2).value = label
    for col, v in zip(('C','D','E','F','G'), (c, d, e, f, g)):
        if v is not None:
            ws[f'{col}{row}'] = v
# 26行目の書式を25行目から引き継ぐ
copy_style(ws, 25, 26, range(2, 8))
# 調整合計は SUM(18:27) のため26行目も自動的に含まれる
ws['B39'] = '調整後EBITDA（過去3期平均）'
ws['D39'] = '=AVERAGE(D37:F37)'
ws['B40'] = '調整後EBITDA（直近2期平均）'
ws['D40'] = '=AVERAGE(E37:F37)'
for r in (39, 40):
    copy_style(ws, 37, r, range(2, 8))
ws['B42'] = ('【調整の考え方】M&A実行後に継続的に発生しない費用を戻し入れ、継続的に必要となる費用を控除。'
             '代表者（淺野ゆう子氏）は引継完了後に退任、現場責任者（淺野崇氏）が役員に就任し現行と同水準の'
             '役員報酬を受ける前提とした。進行期は月次で減価償却費が未計上のため調整後EBITDAは算定していない（要確認）。')
ws['B42'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 取引先（販売先）
ws = wb['取引先']
ws['G3'] = '2023年11月期'; ws['I3'] = '2024年11月期'; ws['K3'] = '2025年11月期'
rows = [
    (5, '一般個人顧客（店舗販売）', '生体・ペット用品・トリミング・ペットホテル・ペット保険',
        '－', '個人', 461431841, 530126419, 552881515,
        '店舗での個人向け販売。得意先別の売上集計は行っていないため、事業所別内訳書の合計から下記セキュリティ事業を控除して算定'),
    (6, 'セキュリティ機器販売先', 'セキュリティ機器の販売・保守', '－', '法人・個人',
        12996480, 8908600, 5508219,
        '代表者配偶者（淺野崇氏）が個人で開始した事業を法人が承継。現在は一部顧客にのみ対応で実質休眠'),
]
for r, name, naiyou, addr, gyoshu, v23, v24, v25, biko in rows:
    ws[f'C{r}'] = name; ws[f'D{r}'] = naiyou; ws[f'E{r}'] = addr; ws[f'F{r}'] = gyoshu
    ws[f'G{r}'] = v23; ws[f'I{r}'] = v24; ws[f'K{r}'] = v25; ws[f'M{r}'] = biko
ws['C15'] = '－'
ws['G16'] = '=SUM(G5:G15)'; ws['I16'] = '=SUM(I5:I15)'; ws['K16'] = '=SUM(K5:K15)'
ws['B18'] = ('※本件は個人向け販売が売上の大宗を占め、得意先別売上高は把握していない（要確認）。'
             '2025年11月期末の売上債権先は ㈲ワンラブ 17,373千円／りそなカード㈱ 5,910千円／'
             'シティックスカード㈱ 3,109千円／PayPay㈱ 821千円／ポケットカード㈱ 415千円（計27,628千円）。')
ws['B18'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 仕入先
ws = wb['仕入先']
ws['F3'] = '2023年11月期'; ws['H3'] = '2024年11月期'; ws['J3'] = '2025年11月期'
shiire = [
    (5, 'ジャペル株式会社', '生体・飼育用品等', '愛知県春日井市', 91166000, 109451000, 107140000,
        '肥料・飼料卸売業。仕入高の約6割を占める最大の仕入先'),
    (6, '有限会社ワンラブ', '生体・飼育用品等', '愛知県名古屋市', 44927000, 50158000, 29244000,
        'フランチャイズ本部。仕入に関する取引契約を締結。生体はワンラブがオークション会場で仕入れ、マイクロチップ埋込・ワクチン等の処理後に当社へ供給'),
    (7, '個人ブリーダー（複数）', '生体', '－', None, 9184000, 17561000,
        '大型犬専門店の展開に伴い直接仕入を拡大'),
    (8, 'デザイナーズワン株式会社', '生体', '愛知県一宮市', 10500000, 12280000, 9400000,
        '代表者の親族が経営する会社。当社が営業権を保有（簿価1,500千円）。大垣・松阪の両管理センターで同社の生体管理を受託'),
]
for r, name, naiyou, addr, v23, v24, v25, biko in shiire:
    ws[f'C{r}'] = name; ws[f'D{r}'] = naiyou; ws[f'E{r}'] = addr
    if v23 is not None: ws[f'F{r}'] = v23
    ws[f'H{r}'] = v24; ws[f'J{r}'] = v25; ws[f'L{r}'] = biko
ws['C15'] = 'その他約50社（個人含む）'
ws['F15'] = 37261000; ws['H15'] = 39728000; ws['J15'] = 18899000
ws['F16'] = '=SUM(F5:F15)'; ws['H16'] = '=SUM(H5:H15)'; ws['J16'] = '=SUM(J5:J15)'
ws['B18'] = ('※出所：他社仲介作成の企業概要書（対象会社提供の総勘定元帳より「商品仕入高」を集計）。'
             'PL上の商品仕入高（183,854／220,801／182,243千円）と一致することを確認済。')
ws['B18'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 外注先
ws = wb['外注先']
ws['G3'] = '2023年11月期'; ws['I3'] = '2024年11月期'; ws['K3'] = '2025年11月期'
ws['C5'] = '要確認（内訳未受領）'
ws['D5'] = 'トリミング・生体管理等の外部委託と推察'
ws['E5'] = '－'; ws['F5'] = '－'
ws['G5'] = 29627732; ws['I5'] = 31465971; ws['K5'] = 25058151
ws['M5'] = '売上原価の「外注費」総額。取引先別の内訳は未受領のため要確認'
ws['C15'] = '－'
ws['G16'] = '=SUM(G5:G15)'; ws['I16'] = '=SUM(I5:I15)'; ws['K16'] = '=SUM(K5:K15)'

# ---------------------------------------------------------------- 株主
ws = wb['株主']
ws['E3'] = '役員報酬\n（2025年11月期）'
ws['C5'] = '淺野　ゆう子'; ws['D5'] = '代表取締役'; ws['E5'] = 14400000
ws['F5'] = 100; ws['H5'] = '本件依頼者。引継完了後、退任予定'
clear(ws, ['C6','D6','E6','F6','H6'])
ws['B7'] = 7  # ダミー行番号はそのまま
ws['B17'] = ('※株券不発行会社。譲渡制限規定あり（株主総会承認）。設立（2020年12月）以降、株主の変動なし。')
ws['B17'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 取締役
ws = wb['取締役']
ws['F3'] = '役員報酬\n（2025年11月期）'
ws['C5'] = '淺野　ゆう子'; ws['D5'] = '代表取締役'; ws['E5'] = '本人'
ws['F5'] = 14400000; ws['G5'] = '経理全般、金融機関対応'; ws['H5'] = '引継完了後、退任予定'
for r in (6, 7, 8):
    clear(ws, [f'C{r}', f'D{r}', f'E{r}', f'F{r}', f'G{r}', f'H{r}'])
ws['B17'] = ('※取締役会・監査役はいずれも非設置。役員は代表取締役1名のみ。\n'
             '※現場責任者の淺野　崇氏（代表者の配偶者・1980年10月生）は役員ではなく従業員（2025年11月期 給与7,000千円）。'
             '経営業務・人材育成・店舗管理を担い、ドン・キホーテ緑店および大垣管理センターの動物取扱責任者を兼務するキーパーソン。'
             '本件は既知で、譲渡後の継続関与を前提としている（要確認）。')
ws['B17'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 従業員
ws = wb['従業員']
ws['F2'] = '基準日→適宜変更'
ws['G2'] = datetime.datetime(2026, 9, 23)
hdr = ['№', '部署・店舗', '正社員等\n（名）', 'アルバイト\n（名）', '合計\n（名）',
       '平均年齢\n（歳）', '平均勤続\n年数（年）', '備考', '']
for i, h in enumerate(hdr):
    ws.cell(row=4, column=2 + i).value = h
for col in range(11, 18):
    sset(ws, 4, col, None)
EMPNOTE = {
 '本社': '代表取締役 淺野ゆう子氏および現場責任者 淺野崇氏。いずれも給与体系は「役員」',
 'ワンラブ カインズホーム明和店': '松阪管理センターと一体で運営されていると推察',
 'ワンラブ バローミタス伊勢店': None,
 'ワンラブ ホームセンターバロー久居店': None,
 'ワンラブ ドン・キホーテ緑店': None,
 'おっきなもふもふ応援隊 岐阜総本店': None,
 'おっきなもふもふ応援隊 大阪総本店': None,
 '松阪管理センター': 'バックヤード管理／デザイナーズワン㈱の生体管理',
}
for i, (dep, sei, arb, tot_, age, ten) in enumerate(LP.EMPLOYEES):
    r = 5 + i
    ws.cell(row=r, column=2).value = i + 1
    ws.cell(row=r, column=3).value = dep
    ws.cell(row=r, column=4).value = sei
    ws.cell(row=r, column=5).value = arb
    ws.cell(row=r, column=6).value = tot_
    ws.cell(row=r, column=7).value = age
    ws.cell(row=r, column=8).value = ten
    ws.cell(row=r, column=9).value = EMPNOTE.get(dep)
    for col in range(10, 18):
        sset(ws, r, col, None)
r = 5 + len(LP.EMPLOYEES)
ws.cell(row=r, column=3).value = '合計'
ws.cell(row=r, column=4).value = LP.EMP_TOTAL[0]
ws.cell(row=r, column=5).value = LP.EMP_TOTAL[1]
ws.cell(row=r, column=6).value = LP.EMP_TOTAL[2]
ws.cell(row=r + 2, column=2).value = (
 '※出所：対象会社提供「従業員名簿（全従業員195名）」（CSV）。'
 '退職年月日が入力されていない49名を2026年9月時点の在籍者として集計した。\n'
 '※本名簿は2026年8月27日入社・同年8月31日退職までを収録しており、進行期の異動を反映している。'
 '2025年12月以降の入社者44名（うち在籍14名）が含まれるため、在籍49名は足元の実態と整合する。\n'
 '※撤退したアクアウォーク大垣店・ペットプラザ岐阜店および大垣管理センターには在籍者がいない。'
 'これは店舗別試算表で両店の売上が2026年1月・2月で終了していることと整合する。\n'
 '※平均年齢・平均勤続年数は2026年9月23日時点。役員2名は入社日の登録がないため勤続年数の集計対象外。\n'
 '※195名のうち146名が退職済。2025年12月以降の退職者は57名（うち大垣店6名・岐阜店4名・大垣管理センター1名）。\n'
 '※氏名・生年月日・住所を含む個票を受領済。買手候補への開示資料にはイニシャル等の匿名化を行ったうえで記載する。')
ws.cell(row=r + 2, column=2).font = Font(name=YGM, size=9)
for r2 in range(5, 35):
    for col in range(2, 18):
        cell = ws.cell(row=r2, column=col)
        if r2 > 5 + len(LP.EMPLOYEES) or (isinstance(cell.value, str) and cell.value.startswith('=ROUNDDOWN')):
            if r2 != r and r2 != r + 2:
                sset(ws, r2, col, None)
        if col in (7, 8):
            cell.number_format = '0.0;;"－"'
        elif col in (4, 5, 6):
            cell.number_format = '0;;"－"'

# ---------------------------------------------------------------- 拠点(本社等)
ws = wb['拠点(本社等)']
kyoten = [
 ('岐阜県大垣市青柳町3丁目227番地2\n（本社／登記上の本店・代表者自宅）','賃貸借\n（貸主：個人）',None,83333,None,
  'JR東海道本線「大垣駅」より車で約8分（要確認）'),
 ('岐阜県大垣市上面4丁目17-1\n（ペット出産用施設／大垣管理センターと推察）','賃貸借\n（大東建託パートナーズ㈱）',None,47831,None,'要確認'),
 ('三重県松阪市（松阪管理センター）\n※明和店倉庫（貸主：ジェイリース㈱）と同一施設の可能性','賃貸借',None,201667,None,'要確認'),
]
for i, (addr, keitai, menseki, chin, shiki, access) in enumerate(kyoten):
    r = 3 + i
    if r > 3:
        copy_style(ws, 3, r, range(2, 8))
    ws.cell(row=r, column=2).value = addr
    ws.cell(row=r, column=3).value = keitai
    ws.cell(row=r, column=4).value = menseki
    ws.cell(row=r, column=5).value = chin
    ws.cell(row=r, column=6).value = shiki
    ws.cell(row=r, column=7).value = access
ws.cell(row=7, column=2).value = (
 '※月間賃料は2025年11月期の年間支払賃借料（地代家賃等の内訳書）を12で除した参考値。延床面積・敷金・契約期間は未受領（要確認）。\n'
 '※本社は代表者の自宅であり、譲渡後の取扱い（賃貸借契約の継続／移転）は要確認。')
ws.cell(row=7, column=2).font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 不動産(賃借)
ws = wb['不動産(賃借) ']
tenpo = [
 ('ワンラブ カインズホーム明和店','建物）店舗（インショップ）','三重県多気郡明和町中村1236\nカインズ明和 ペッツワンコーナー',None,
  'FC本部（㈲ワンラブ）\nからの転借','有限会社ワンラブ',None,None,'月額賃料は売上歩合（約15%）。COC条項あり。ロイヤリティに含めて支払'),
 ('ワンラブ バローミタス伊勢店','建物）店舗（インショップ）','三重県伊勢市船江1-10-190\nバローホームセンター ミタス伊勢店内',None,
  '2020年4月〜2030年4月\n※2年毎の更新','有限会社ワンラブ',None,None,'月額賃料は売上歩合（売上の11%）。COC条項あり'),
 ('ワンラブ ホームセンターバロー久居店','建物）店舗（インショップ）','三重県津市戸木町7869-1\nバローメガストア久居インター店内',None,
  '2020年12月〜2030年11月\n※2年毎の更新','有限会社ワンラブ',None,None,'月額賃料は売上歩合（売上の12%）。COC条項あり'),
 ('ワンラブ アクアウォーク大垣店','建物）店舗（インショップ）','岐阜県大垣市林町6丁目80番21',None,
  '要確認','有限会社ワンラブ',None,None,'決算書上の事業所名は「ペッツビレッジ大垣店」（要確認）。別途、駐車場地代をユニー㈱へ年184.8千円支払'),
 ('ワンラブ ペットプラザ岐阜店','建物）店舗（インショップ）','岐阜県岐阜市鶉町3丁目18',None,
  '要確認','有限会社ワンラブ',None,None,'月額賃料は売上歩合。COC条項の有無は要確認'),
 ('ワンラブ ドン・キホーテ緑店','建物）店舗（インショップ）','愛知県名古屋市緑区鳥澄1-527\nドン・キホーテ緑店ペット館',None,
  '要確認','日本商業施設株式会社',749000,2700000,'2025年11月期の年間家賃9,285千円。COC条項は確認中'),
 ('おっきなもふもふ応援隊 岐阜総本店','建物）店舗（路面店）','岐阜県羽島郡岐南町徳田3-182-1\nザ・ビッグ岐南店1階 区画No.1',395.38,
  '2023年4月〜2033年1月\n定期建物賃貸借（更新なし）','株式会社カムテイ',640000,2800000,'2025年11月期の年間家賃8,153千円。COC条項あり。再契約可否は要確認'),
 ('おっきなもふもふ応援隊 大阪総本店','建物）店舗（路面店）','大阪府茨木市東太田1-4-48\nドン・キホーテ茨木店1階',421.10,
  '2024年9月〜2027年11月\n定期建物賃貸借（更新なし）','株式会社ドン・キホーテ\n（サブリース／転貸人：日本商業施設㈱）',920000,3600000,
  '駐車場3台分を含む。2025年11月期の年間家賃10,756千円＋駐車場363千円。COC条項あり。再契約は協議のうえ可'),
 ('本社','建物）事務所（代表者自宅）','岐阜県大垣市青柳町3丁目227番地2',None,'要確認','清水　満（個人）',83333,None,'2025年11月期の年間家賃1,000千円'),
 ('ペット出産用施設','建物）その他','岐阜県大垣市上面4丁目17-1',None,'要確認','大東建託パートナーズ株式会社',47831,None,'2025年11月期の年間家賃574千円'),
 ('明和店倉庫','建物）倉庫','要確認',None,'要確認','ジェイリース株式会社',201667,None,'2025年11月期の年間家賃2,420千円'),
 ('従業員社宅','建物）社宅','要確認',None,'要確認','株式会社エポスカード',42725,None,'2025年11月期の年間家賃513千円'),
]
for i, row in enumerate(tenpo):
    r = 4 + i
    if r > 4:
        copy_style(ws, 4, r, range(2, 11))
    for j, v in enumerate(row):
        ws.cell(row=r, column=2 + j).value = v
for r in range(4 + len(tenpo), 22):
    for col in range(2, 11):
        sset(ws, r, col, None)
ws.cell(row=18, column=2).value = (
 '※出所：2025年11月期「地代家賃等の内訳書」、他社仲介作成の企業概要書「店舗の状況1・2」。'
 'ワンラブFC店（明和・伊勢・久居・大垣・岐阜）の店舗使用料は売上歩合であり、販管費上は「ロイヤリティ」に計上されている。\n'
 '※面積・敷金・契約期間の空欄および「要確認」欄は資料未受領。COC条項（会社組織の変更・資本構成の重大な変更時の届出義務、'
 '合併時の契約解除可能）が①〜③および岐阜総本店・大阪総本店の契約書に存在するため、本件実行時の対応は要確認。')
ws.cell(row=18, column=2).font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 不動産(所有)
ws = wb['不動産(所有)']
ws['B4'] = '大垣（非店舗）'
ws['C4'] = '土地）駐車場'
ws['D4'] = '岐阜県大垣市南若森町字柳原227番3'
ws['E4'] = 89.4
ws['F4'] = '宅地'
ws['G4'] = '2025年11月期中に取得（要確認）'
ws['H4'] = 2974400
ws['I4'] = '要確認'
ws['J4'] = '要確認'
ws['K4'] = '固定資産税評価額が未受領のため時価は要確認。簿価＝2,974,400円'
for r in range(5, 22):
    for col in range(2, 12):
        sset(ws, r, col, None)
ws['B7'] = ('※法人所有の不動産は上記土地1筆（89.4㎡）のみ。建物の所有はなく、店舗はすべて賃借（インショップ／路面店）。\n'
            '※代表者個人が所有する事業用不動産は無し（他社仲介作成の企業概要書「希望条件」欄にて「該当無し」）。')
ws['B7'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 設備
ws = wb['設備']
setsubi = [
 ('建物附属設備','おっきなもふもふ大阪総本店 新装工事','定額法','2024年','10月','15年',15950000,14347025,'大阪総本店の出店に伴う内装一式'),
 ('建物附属設備','おっきなもふもふ大阪総本店 内装工事','定額法','2024年','10月','15年',13000000,12927417,None),
 ('建物附属設備','空調・照明設備','定額法','2024年','7月','15年',8091127,7368321,None),
 ('建物附属設備','ドン・キホーテ緑店 改装工事（給排水・電気配管・照明・エアコン）','定額法','2024年','7月','15年',4319200,4270969,None),
 ('建物附属設備','水回り工事・外装工事','定額法','2024年','9月','15年',2158650,2050179,None),
 ('建物附属設備','壁面サイン（金属製）','定額法','2023年','7月','18年',1945900,1800607,None),
 ('建物附属設備','おっきなもふもふ大阪総本店 サイン改修工事（金属製）','定額法','2024年','6月','18年',1760000,1751787,None),
 ('建物附属設備','ドン・キホーテ緑店 改装工事（電気工事）','定額法','2024年','8月','15年',1892000,1839182,None),
 ('建物附属設備','電気設備・その他','定額法','2024年','8月','15年',1040200,1011162,None),
 ('建物附属設備','おっきなもふもふ大阪総本店 外装工事（消防設備）','定額法','2024年','8月','8年',286000,241313,None),
 ('車両運搬具','ハイエースバン（トヨタ自動車㈱）','200%定率法','2022年','4月','6年',3168768,731557,None),
 ('車両運搬具','ピクシスバン（トヨタ自動車㈱）','200%定率法','2023年','7月','4年',1647234,326016,None),
 ('車両運搬具','プリウス（トヨタ自動車㈱）','200%定率法','2020年','12月','6年',1163000,1,'備忘価額1円'),
 ('車両運搬具','BMW X7 xDrive40d M Sport（BMW AG）','200%定率法','2023年','10月','6年',13540645,0,
  '2025年11月期中に㈱Luck Autoへ7,562千円で売却（固定資産売却損968千円）。他社仲介資料の「進行期に売却」との記載は要確認'),
 ('工具器具備品','おっきなもふもふ応援隊 展示用什器・陳列用ゲージ','200%定率法','2024年','7月','8年',2400000,1612500,None),
 ('工具器具備品','おっきなもふもふ大阪総本店 展示用什器・冷蔵庫','200%定率法','2024年','9月','8年',3750000,2695313,None),
 ('工具器具備品','複合機・通信機器（bizhub C287ほか）等','200%定率法','2022〜2024年','－','5年',11769229,6044761-1612500-2695313,
  '複数台。事務機器・通信機器・Airレジ等の合計'),
]
for i, row in enumerate(setsubi):
    r = 3 + i
    if r > 21:
        break
    if r > 3:
        copy_style(ws, 3, r, range(2, 12))
    ws.cell(row=r, column=2).value = i + 1
    for j, v in enumerate(row):
        ws.cell(row=r, column=3 + j).value = v
last = 3 + len(setsubi)
ws.cell(row=last, column=3).value = '合計（上記のうち2025年11月期末簿価）'
ws.cell(row=last, column=9).value = f'=SUM(I3:I{last-1})'
ws.cell(row=last, column=10).value = f'=SUM(J3:J{last-1})'
ws.cell(row=last + 2, column=2).value = (
 '※出所：2025年11月期「旧定率法・定額法による固定資産減価償却内訳明細書」。'
 '2025年11月期末の有形固定資産簿価は 建物附属設備44,397千円／車両運搬具1,058千円／工具器具備品6,045千円／土地2,974千円／その他184千円。\n'
 '※上記のほかリース車両3台（メルセデス・ベンツ GLE Coupe 月195千円・2024年9月〜60ヶ月、GLS 月212千円・2024年12月〜60ヶ月、'
 '日産オーラ 月60千円・2024年9月〜36ヶ月）。ベンツ2台は非事業用であり、本件実行後に解約予定。\n'
 '※一括償却資産：ドン・キホーテ緑店 ゲージ改修工事146千円、同 カーテン修繕工事130千円（いずれも2025年6月供用）。')
ws.cell(row=last + 2, column=2).font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 店舗別売上（新規シート）
if '店舗別売上' in wb.sheetnames:
    del wb['店舗別売上']
ws = wb.create_sheet('店舗別売上', wb.sheetnames.index('取引先'))
NAVY = '0B3041'; BLUE = '0889C9'; PALE = 'E1F3FB'
thin = Side(style='thin', color='FFFFFF')
box  = Side(style='thin', color='BFBFBF')
ws.sheet_view.showGridLines = False
ws['B2'] = '店舗別売上高（決算書「売上高等の事業所別内訳書」ベース）'
ws['B2'].font = Font(name=YG, size=11, bold=True, color=NAVY)
ws['L2'] = '(単位：千円)'
ws['L2'].font = Font(name=YGM, size=9)
ws['L2'].alignment = Alignment(horizontal='right')
head = ['№','事業所名','業態','所在地','2023年11月期','構成比','2024年11月期','構成比',
        '2025年11月期','構成比','進行期7か月\n(25/12〜26/6)','構成比',
        '期末棚卸高\n(25/11期)','期末従事\n員数(名)']
for i, h in enumerate(head):
    c = ws.cell(row=4, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
tenpo_uri = [
 ('本社（セキュリティ事業）','非店舗','岐阜県大垣市', 12996480,  8908600,  5508219, 2944810, None, 3),
 ('本社（サービス・物販）','非店舗','岐阜県大垣市', 15610820, 18397207, 22125782, 13650127, 44000000, None),
 ('ワンラブ カインズホーム明和店','FC／インショップ','三重県多気郡明和町', 65416318, 67267220, 37216460, 25108810, 21818576, 9),
 ('ワンラブ バローミタス伊勢店','FC／インショップ','三重県伊勢市', 88527370, 75571000, 66609459, 40221277, 2985224, 8),
 ('ワンラブ ホームセンターバロー久居店','FC／インショップ','三重県津市', 66886986, 62981034, 57326532, 33840444, 2814741, 7),
 ('ワンラブ アクアウォーク大垣店 ※2026年2月撤退','FC／インショップ','岐阜県大垣市', 50138278, 43305102, 38021413, 6112430, 3045850, 6),
 ('ワンラブ ペットプラザ岐阜店 ※2026年1月撤退','FC／インショップ','岐阜県岐阜市', 32770375, 20636067, 18944680, 1842692, 2459389, 6),
 ('ワンラブ ドン・キホーテ緑店','FC／インショップ','愛知県名古屋市緑区', None, 19793315, 51761466, 36963203, 3747447, 6),
 ('ワンラブ コメリパワー中志段味店 ※2024年6月譲渡','FC／インショップ','愛知県名古屋市守山区', 43554994, 29765016, None, None, None, None),
 ('おっきなもふもふ応援隊 岐阜総本店','大型犬専門／路面店','岐阜県羽島郡岐南町', 98526700, 157755564, 118491175, 60022946, 6566330, 11),
 ('おっきなもふもふ応援隊 大阪総本店','大型犬専門／路面店','大阪府茨木市', None, 34654894, 142384548, 58446704, 8579727, 10),
 ('デザイナーズワン株式会社（仕入先預け在庫）','－','愛知県一宮市', None, None, None, None, 10000000, None),
]
r0 = 5
for i, (name, gyotai, addr, v23, v24, v25, vpg, tana, nin) in enumerate(tenpo_uri):
    r = r0 + i
    vals = [i + 1, name, gyotai, addr, v23, None, v24, None, v25, None, vpg, None, tana, nin]
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=2 + j)
        c.value = v
        c.border = Border(left=box, right=box, top=box, bottom=box)
    for col in (7, 9, 11, 13):
        src = ws.cell(row=r, column=col - 1).coordinate[0]
        ws.cell(row=r, column=col).value = f'=IFERROR({src}{r}/{src}${r0+len(tenpo_uri)},"")'
    for col in (2, 14, 15):
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9)
        ws.cell(row=r, column=col).alignment = Alignment(horizontal='center')
    for col in (3, 4, 5):
        ws.cell(row=r, column=col).font = Font(name=YG, size=9)
    for col in (6, 8, 10, 12, 14):
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9)
        ws.cell(row=r, column=col).number_format = '#,##0,;[Red]\\-#,##0,;"－"'
    for col in (7, 9, 11, 13):
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9, italic=True)
        ws.cell(row=r, column=col).number_format = '0.0%'
    ws.cell(row=r, column=15).number_format = '0;;"－"'
rt = r0 + len(tenpo_uri)
ws.cell(row=rt, column=3).value = '合計'
for col in (6, 8, 10, 12, 14, 15):
    L = ws.cell(row=rt, column=col).coordinate[0]
    ws.cell(row=rt, column=col).value = f'=SUM({L}{r0}:{L}{rt-1})'
for col in range(2, 16):
    c = ws.cell(row=rt, column=col)
    c.fill = PatternFill('solid', fgColor=PALE)
    c.border = Border(left=box, right=box, top=box, bottom=box)
    if col in (6, 8, 10, 12, 14):
        c.font = Font(name=ARIAL, size=9, bold=True)
        c.number_format = '#,##0,;[Red]\\-#,##0,;"－"'
    elif col == 15:
        c.font = Font(name=ARIAL, size=9, bold=True); c.number_format = '0;;"－"'
        c.alignment = Alignment(horizontal='center')
    else:
        c.font = Font(name=YG, size=9, bold=True)
for col in (7, 9, 11, 13):
    ws.cell(row=rt, column=col).value = '－'
    ws.cell(row=rt, column=col).font = Font(name=ARIAL, size=9, bold=True)
    ws.cell(row=rt, column=col).alignment = Alignment(horizontal='center')
note = ws.cell(row=rt + 2, column=2)
note.value = ('※出所：第3期〜第5期 決算報告書「売上高等の事業所別内訳書」、および店舗別（部門別）勘定科目残高推移表 2025年12月〜2026年6月。'
 '各期のPL売上高（474,428／539,035／558,390／279,153千円）と一致することを確認済。\n'
 '※コメリパワー中志段味店は2024年6月に譲渡。ドン・キホーテ緑店は2024年7月、おっきなもふもふ応援隊 大阪総本店は2024年10月に開店。\n'
 '※ペットプラザ岐阜店は2026年1月、アクアウォーク大垣店は2026年2月をもって売上計上が終了しており、両店が撤退した2店舗である'
 '（店舗別試算表および従業員名簿の在籍状況により確認）。2026年3月以降の稼働店舗は6店舗。\n'
 '※本社（サービス・物販）の進行期13,650千円は月額約195万円で計上されており、ペット保険の代理店手数料と推察される（仮説・要確認）。')
note.font = Font(name=YGM, size=9)
note.alignment = Alignment(wrap_text=False)
for col, w in zip('BCDEFGHIJKLMNOP', (4, 36, 18, 20, 13, 8, 13, 8, 13, 8, 14, 8, 13, 10, 4)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C5'

# ---------------------------------------------------------------- 店舗別損益（進行期）
# 出所：店舗別（部門別）勘定科目残高推移表 2025年12月〜2026年6月
if '店舗別損益' in wb.sheetnames:
    del wb['店舗別損益']
ws = wb.create_sheet('店舗別損益', wb.sheetnames.index('店舗別売上') + 1)
ws.sheet_view.showGridLines = False
NAVY = '0B3041'; PALE = 'E1F3FB'
thin = Side(style='thin', color='FFFFFF')
box = Side(style='thin', color='BFBFBF')
BOX = Border(left=box, right=box, top=box, bottom=box)
ws['B2'] = '店舗別（部門別）損益　進行期7か月累計（2025年12月〜2026年6月）'
ws['B2'].font = Font(name=YG, size=11, bold=True, color=NAVY)
ws['K2'] = '(単位：千円)'
ws['K2'].font = Font(name=YGM, size=9)
ws['K2'].alignment = Alignment(horizontal='right')
head = ['№', '部門', '業態', '状況', '売上高', '構成比', '売上総利益', '粗利率',
        '販管費', '営業利益', '営業利益率']
for i, h in enumerate(head):
    c = ws.cell(row=4, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)

def _put_store(r, no, name, gyotai, jokyo, sales, gp, sga, op, emph=False):
    vals = [no, name, gyotai, jokyo, sales, None, gp, None, sga, op, None]
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=2 + j)
        c.value = v
        c.border = BOX
        if emph:
            c.fill = PatternFill('solid', fgColor=PALE)
    for col, src in ((7, 6), (9, 8), (12, 11)):   # 構成比・粗利率・営業利益率
        L = ws.cell(row=r, column=src).coordinate[0]
        ws.cell(row=r, column=col).value = (
            f'=IFERROR({L}{r}/$F${rt},"")' if col == 7 else f'=IFERROR({L}{r}/$F{r},"")')
        ws.cell(row=r, column=col).number_format = '0.0%;-0.0%;"－"'
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9, italic=True, bold=emph)
    for col in (2, 6, 8, 10, 11):
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9, bold=emph)
        if col != 2:
            ws.cell(row=r, column=col).number_format = '#,##0,;[Red]△ #,##0,;"－"'
        else:
            ws.cell(row=r, column=col).alignment = Alignment(horizontal='center')
    for col in (3, 4, 5):
        ws.cell(row=r, column=col).font = Font(name=YG, size=9, bold=emph)

ROWS = []
for i, (name, gyotai, jokyo, sales, gp, sga, op, _m) in enumerate(LP.STORE_PL):
    ROWS.append((i + 1, name, gyotai, jokyo, sales, gp, sga, op, False))
ROWS.append((None, '【稼働6店舗 小計】', '－', '－',
             sum(x[3] for x in LP.STORE_PL), sum(x[4] for x in LP.STORE_PL),
             sum(x[5] for x in LP.STORE_PL), sum(x[6] for x in LP.STORE_PL), True))
for i, (name, gyotai, jokyo, sales, gp, sga, op, _m) in enumerate(LP.STORE_CLOSED):
    ROWS.append((7 + i, name, gyotai, jokyo, sales, gp, sga, op, False))
for name, gyotai, jokyo, sales, gp, sga, op, _m in LP.STORE_OTHER:
    ROWS.append((None, name, gyotai, jokyo, sales, gp, sga, op, False))
ROWS.append((None, '【合計】', '－', '－', *LP.COMPANY_TOTAL, True))

r0 = 5
rt = r0 + len(ROWS) - 1          # 合計行（構成比の分母）
for i, (no, name, gyotai, jokyo, sales, gp, sga, op, emph) in enumerate(ROWS):
    _put_store(r0 + i, no, name, gyotai, jokyo, sales, gp, sga, op, emph)
note = ws.cell(row=rt + 2, column=2)
note.value = (
 '※出所：店舗別（部門別）勘定科目残高推移表 2025年12月〜2026年6月（2026年9月18日出力）。'
 '部門別の合計は全社の試算表（売上高279,153千円・営業利益6,637千円）と一致することを確認済。\n'
 '※店舗別の営業利益は本社費を配賦する前の「店舗貢献利益」であり、本社費等は共通部門に一括計上されている。\n'
 '※共通部門の売上13,650千円は、月額約195万円で計上されているサービス売上高であり、'
 'ペット保険の代理店手数料と推察される（仮説・要確認）。\n'
 '※共通部門の売上原価46,604千円は、本社で仕入れた商品（92,274千円のうち40,994千円）および外注費5,610千円が'
 '各店舗へ振り替えられずに共通部門に残っているものと推察される（仮説・要確認）。'
 'このため店舗別の粗利・営業利益は実力値より過大に表示されている可能性があり、解釈には留意を要する。\n'
 '※明和店の販管費には松阪管理センターの人件費が含まれていると推察される'
 '（棚卸表でも松阪管理センターの在庫19,669千円が明和店に含めて計上されている）。\n'
 '※大垣管理センターは2026年2月以降、費用のマイナス計上（振替）のみとなっている。')
note.font = Font(name=YGM, size=9)
note.alignment = Alignment(vertical='top')
for col, w in zip('BCDEFGHIJKLM', (4, 34, 17, 16, 13, 9, 13, 9, 13, 13, 10, 4)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C5'

# ---------------------------------------------------------------- 棚卸明細
if '棚卸明細' in wb.sheetnames:
    del wb['棚卸明細']
ws = wb.create_sheet('棚卸明細', wb.sheetnames.index('店舗別損益') + 1)
ws.sheet_view.showGridLines = False
ws['B2'] = '棚卸資産の内訳（2025年12月1日時点＝2025年11月期末）'
ws['B2'].font = Font(name=YG, size=11, bold=True, color=NAVY)
ws['J2'] = '(単位：千円／税抜)'
ws['J2'].font = Font(name=YGM, size=9)
ws['J2'].alignment = Alignment(horizontal='right')
head = ['拠点', '生体：アクア', '生体：小動物', '生体：犬猫', '生体 小計', '物販',
        '合計', '決算書の計上額', '備考']
for i, h in enumerate(head):
    c = ws.cell(row=4, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
for i, row in enumerate(LP.INVENTORY):
    r = 5 + i
    emph = row[0].startswith('【')
    for j, v in enumerate(row):
        c = ws.cell(row=r, column=2 + j)
        c.value = v
        c.border = BOX
        if emph:
            c.fill = PatternFill('solid', fgColor=PALE)
        if j == 0:
            c.font = Font(name=YG, size=9, bold=emph)
        elif j == 8:
            c.font = Font(name=YGM, size=9)
        else:
            c.font = Font(name=ARIAL, size=9, bold=emph)
            c.number_format = '#,##0,;[Red]△ #,##0,;"－"'
ws.cell(row=5 + len(LP.INVENTORY) + 2, column=2).value = (
 '※出所：対象会社作成「棚卸表 2025年12月1日時点（最終版）」。全社合計106,017,284円は第5期決算書の商品残高と一致する。\n'
 '※本社32,000千円＋バックヤード12,000千円＝44,000千円は、進行期の2026年3月に棚卸資産廃棄損として全額を特別損失に計上している。\n'
 '※外部保管分10,000千円は仕入先デザイナーズワン株式会社（代表者の親族が経営）への預け在庫。'
 '2026年6月末の商品残高58,446千円のうち同額が含まれているかは要確認。\n'
 '※「決算書の計上額」は第5期決算報告書「売上高等の事業所別内訳書」の期末棚卸高。'
 '松阪管理センターは明和店に、中志段味店は緑店に含めて計上されている。')
ws.cell(row=5 + len(LP.INVENTORY) + 2, column=2).font = Font(name=YGM, size=9)
ws.cell(row=5 + len(LP.INVENTORY) + 2, column=2).alignment = Alignment(vertical='top')
for col, w in zip('BCDEFGHIJK', (36, 12, 12, 12, 12, 12, 13, 13, 46, 4)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C5'


# ---------------------------------------------------------------- その他（要確認事項）
ws = wb['その他']
ws.sheet_view.showGridLines = False
ws['B2'] = '要確認事項・ネクストステップ'
ws['B2'].font = Font(name=YG, size=12, bold=True, color=NAVY)
hdr = ['№', '区分', '論点', '現時点で確認できている事実／当社の仮説', '確認方法・依頼先']
for i, h in enumerate(hdr):
    c = ws.cell(row=4, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
items = [
 ('【解決済】店舗','2026年1月末に撤退した赤字2店舗の特定',
  '店舗別試算表により、ペットプラザ岐阜店が2026年1月、アクアウォーク大垣店が2026年2月をもって'
  '売上計上を終了していることを確認。従業員名簿でも両店および大垣管理センターに在籍者がおらず整合する。'
  '進行期7か月の両店合計は売上7,955千円・営業損失4,086千円',
  '解決（2026年9月受領資料）。撤退に伴う原状回復費用・違約金の有無は引き続き要確認'),
 ('【解決済】資産','進行期に計上された棚卸資産廃棄損44,000千円の内容',
  '棚卸表（2025年12月1日時点）により、本社32,000千円＋バックヤード12,000千円＝44,000千円と一致することを確認',
  '解決（2026年9月受領資料）。廃棄に至った理由・処分方法・証憑は引き続き要確認'),
 ('財務','2026年7月以降の月次試算表',
  '2025年12月〜2026年6月の7か月分を受領済。進行期7か月の売上279,153千円（前年同期比80.9%）、'
  '営業利益6,637千円（同48.2%）。通期の着地見込みを判断するには下期の実績が必要',
  '対象会社・顧問税理士へ依頼'),
 ('財務','進行期に未計上の減価償却費・法定福利費',
  '進行期の月次には減価償却費の科目が立っていない（2025年11月期の通期実績6,349千円）。'
  '法定福利費も7か月累計2,778千円と通期実績15,295千円に比して低く、期中の未計上分が下期に乗る可能性がある',
  '顧問税理士へ確認'),
 ('財務','共通部門に残る仕入原価の配賦',
  '店舗別試算表では、本社で仕入れた商品40,994千円および外注費5,610千円が各店舗へ振り替えられず'
  '共通部門の売上原価に残っている（共通部門の売上原価46,604千円＝本社仕入40,994千円＋外注費5,610千円）。'
  'このため店舗別の粗利・営業利益は実力値より過大に表示されている可能性がある（仮説）',
  '対象会社ヒアリング（本社仕入分の店舗への配分方法）'),
 ('財務','2026年5月の租税公課4,969千円の内容',
  '5月単月で4,969千円を計上（7か月累計5,281千円の94%）。自動車税・固定資産税等の年次課税と推察（仮説）。'
  '同月の営業損失5,152千円の主因',
  '対象会社ヒアリング／総勘定元帳'),
 ('店舗収益','明和店の販管費が売上に対して突出している要因',
  '進行期7か月で売上25,109千円に対し販管費18,977千円（75.6%）。'
  '他のFC加盟店（伊勢35.2%・久居39.4%・緑37.6%）と比べて著しく高く、'
  '松阪管理センターの人件費が明和店に含まれていると推察（棚卸表でも松阪管理センターの在庫19,669千円が'
  '明和店に含めて計上されている）（仮説）',
  '対象会社ヒアリング（部門設定の考え方）'),
 ('収益モデル','ペット保険の代理店手数料',
  '共通部門のサービス売上高が月額約195万円で計上されており、7か月累計13,650千円（年換算23,400千円）。'
  '他社仲介資料の「年間約19,000千円の保険収入」との聴取と概ね整合し、同手数料と推察（仮説）',
  '対象会社ヒアリング／アイペット損保との代理店手数料明細'),
 ('【解決済】人事','進行期を含む最新の従業員名簿',
  '名簿195名（2026年8月27日入社・同年8月31日退職まで収録）を受領。退職日の入力がない49名を在籍者として集計。'
  '2025年12月以降の入社44名・退職57名を織り込んでおり、足元の在籍実態を反映している',
  '2026年9月受領の従業員名簿で確認済'),
 ('人事','現場責任者 淺野崇氏の譲渡後の継続関与',
  '従業員名簿上、同氏の給与体系は「役員」・区分は「社員」。正常収益力の算定において'
  '同氏が役員に就任し現行と同水準の役員報酬を受ける前提を置いている。'
  '同氏はドン・キホーテ緑店および大垣管理センターの動物取扱責任者を兼務するキーパーソン',
  '対象会社ヒアリング（意向確認）'),
 ('契約','フランチャイズ契約書および店舗賃貸借契約書',
  'FC本部（㈲ワンラブ）との契約内容、店舗使用料の料率、COC条項の詳細が未確認。'
  '定期建物賃貸借（岐阜総本店2033年1月／大阪総本店2027年11月満了）の再契約見通しも要確認',
  '対象会社へ依頼（必要資料No.12・13）'),
 ('資産','外部保管在庫10,000千円の実在性',
  '2025年11月期末の棚卸表に「外部保管分10,000千円（生体）」として計上。'
  '仕入先デザイナーズワン株式会社（代表者の親族が経営）への預け在庫と推察。'
  '2026年6月末の商品残高58,446千円に同額が含まれているかは未確認',
  '対象会社ヒアリング／預け在庫の残高証明'),
 ('資産','2026年6月に取得した車両運搬具438千円',
  '車両運搬具が2026年5月末1,058千円から6月末1,495千円へ増加。事業用・非事業用の別は未確認',
  '対象会社ヒアリング／固定資産台帳'),
 ('資産','営業権1,500千円の取得経緯・算定根拠',
  '仕入先であるデザイナーズワン株式会社（代表者の親族が経営）に係るもの',
  '取引契約書／営業権の取得経緯・算定根拠'),
 ('許認可','動物取扱責任者の配置状況と許可の更新',
  '第一種動物取扱業許可は2025年11月期末時点で8事業所。'
  '撤退したアクアウォーク大垣店（2026年11月29日満了）・ペットプラザ岐阜店（2027年1月31日満了）の'
  '廃業届の提出状況、および明和店・久居店（2026年5月20日満了）の更新状況が未確認',
  '許可証の写し／動物取扱責任者名簿'),
 ('財務','借入金の返済予定と月次での据置き',
  '長期借入金は月次試算表では期首残高191,858千円のまま据え置かれており、返済額は決算時に'
  '一括計上されていると推察（仮説）。役員借入金は7,566千円（2025年11月期末）から'
  '3,403千円（2026年6月末）へ減少している',
  '金融機関別の返済予定表（必要資料No.5）'),
 ('財務','本社所在地（代表者自宅）の賃貸借の譲渡後の取扱い',
  '貸主は個人。本社は登記上の本店かつ代表者自宅であり、譲渡後の本店移転・契約継続の要否を整理する必要',
  '対象会社ヒアリング'),
 ('プロセス関連（社内確認事項）','他社仲介の併走状況',
  '2025年11月期末の仮払金1,650千円は他社仲介への支払（着手金等）と認められる。'
  '本項は社内確認事項であり、買手候補への開示資料には記載しない',
  '社内'),
]
for i, (kubun, ronten, jijitsu, houhou) in enumerate(items):
    r = 5 + i
    vals = [i + 1, kubun, ronten, jijitsu, houhou]
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=2 + j)
        c.value = v
        c.border = Border(left=box, right=box, top=box, bottom=box)
        c.alignment = Alignment(vertical='top', wrap_text=True)
        c.font = Font(name=ARIAL if j == 0 else YGM, size=9)
    ws.cell(row=r, column=3).alignment = Alignment(vertical='center', wrap_text=True, horizontal='center')
    ws.cell(row=r, column=2).alignment = Alignment(vertical='center', horizontal='center')
    ws.row_dimensions[r].height = 46
for col, w in zip('BCDEF', (4, 16, 40, 66, 32)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C5'

# ---------------------------------------------------------------- 月次推移（進行期）
# 出所：勘定科目残高推移表 2025年12月〜2026年6月（2026年9月18日出力）
if '月次推移' in wb.sheetnames:
    del wb['月次推移']
ws = wb.create_sheet('月次推移', wb.sheetnames.index('PLハイライト'))
ws.sheet_view.showGridLines = False
NAVY = '0B3041'; PALE = 'E1F3FB'; GREY = 'E3E7E9'
thin = Side(style='thin', color='FFFFFF')
box = Side(style='thin', color='BFBFBF')
BOX = Border(left=box, right=box, top=box, bottom=box)

ws['B2'] = '進行期（2026年11月期）の月次推移　※2025年12月〜2026年6月の7か月'
ws['B2'].font = Font(name=YG, size=11, bold=True, color=NAVY)
ws['L2'] = '(単位：千円)'
ws['L2'].font = Font(name=YGM, size=9)
ws['L2'].alignment = Alignment(horizontal='right')

hdr = ['科目'] + LP.MONTHS + ['7か月累計', '前年同期比']
for i, h in enumerate(hdr):
    c = ws.cell(row=4, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
r0 = 5
for i, (label, vals, tot, yoy, emph) in enumerate(LP.PL_MONTHLY):
    r = r0 + i
    ws.cell(row=r, column=2).value = label
    ws.cell(row=r, column=2).font = Font(name=YG, size=9, bold=emph)
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=3 + j)
        c.value = v if v else None
        c.number_format = '#,##0,;[Red]△ #,##0,;"－"'
        c.font = Font(name=ARIAL, size=9, bold=emph)
    c = ws.cell(row=r, column=10)
    c.value = tot
    c.number_format = '#,##0,;[Red]△ #,##0,;"－"'
    c.font = Font(name=ARIAL, size=9, bold=emph)
    c = ws.cell(row=r, column=11)
    c.value = yoy
    c.number_format = '0.0%;;"－"'
    c.font = Font(name=ARIAL, size=9, italic=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).border = BOX
        if emph:
            ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=PALE)
rEnd = r0 + len(LP.PL_MONTHLY)

ws.cell(row=rEnd + 2, column=2).value = '【月末の主要B/S科目】'
ws.cell(row=rEnd + 2, column=2).font = Font(name=YG, size=10, bold=True, color=NAVY)
bh = rEnd + 3
hdr2 = ['科目', '2025年11月末（決算）'] + [m + '末' for m in LP.MONTHS]
for i, h in enumerate(hdr2):
    c = ws.cell(row=bh, column=2 + i)
    c.value = h
    c.font = Font(name=YG, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
for i, (label, base, vals, emph) in enumerate(LP.BS_MONTHLY):
    r = bh + 1 + i
    ws.cell(row=r, column=2).value = label
    ws.cell(row=r, column=2).font = Font(name=YG, size=9, bold=emph)
    for j, v in enumerate([base] + vals):
        c = ws.cell(row=r, column=3 + j)
        c.value = v
        c.number_format = '#,##0,;[Red]△ #,##0,;"－"'
        c.font = Font(name=ARIAL, size=9, bold=emph)
    for col in range(2, 11):
        ws.cell(row=r, column=col).border = BOX
        if emph:
            ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=PALE)
note = ws.cell(row=bh + len(LP.BS_MONTHLY) + 3, column=2)
note.value = (
 '※出所：勘定科目残高推移表 2025年12月〜2026年6月（2026年9月18日出力）。「2025年11月末（決算）」列のみ第5期決算報告書。\n'
 '※2026年6月26日出力の旧試算表（2026年4月まで）から、2026年1月〜4月の月次が一部修正されている（販管費・営業外費用・役員借入金等）。本シートは新試算表を正としている。\n'
 '※進行期の月次には減価償却費が計上されていない（販管費に「減価償却費」の科目が立っていない）。'
 '2025年11月期の通期実績は6,349千円であり、通期ではその分だけ営業利益が減少する見込み（要確認）。\n'
 '※法定福利費は2025年12月・2026年1月がマイナス計上となっており、年末調整・社会保険料の精算処理の影響と推察（仮説・要確認）。'
 '7か月累計2,779千円は2025年11月期の通期15,295千円に比して低く、期中の未計上分が下期に乗る可能性がある。\n'
 '※2026年5月の租税公課4,969千円は、自動車税・固定資産税等の年次課税と推察（仮説）。同月の営業損失の主因。\n'
 '※2026年3月に棚卸資産廃棄損44,000千円を計上（他勘定振替高△44,000千円）。'
 '棚卸表の「本社32,000千円＋バックヤード12,000千円」に一致する（棚卸明細シート参照）。\n'
 '※長期借入金は月次では期首残高のまま据え置かれており、返済額は決算時に一括計上されていると推察（要確認）。')
note.font = Font(name=YGM, size=9)
note.alignment = Alignment(vertical='top')
for col, w in zip('BCDEFGHIJKL', (26, 12, 12, 12, 12, 12, 12, 12, 13, 11, 4)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'C5'

# ---------------------------------------------------------------- BSハイライト（進行期）
ws = wb['BSハイライト']
ws['H3'] = '進行期（2026年6月末TB）'
ws['H3'].font = Font(name=YG, size=10, bold=True, color=NAVY)
prog_l = [('流動資産', None), ('現預金', 'G6'), ('売掛金', 'G7'), ('在庫', None),
          ('前払費用', 'G11'), ('未収入金', None), ('未収還付法人税等', None), ('仮払金', 'G14'),
          ('固定資産', None), ('建物', 'G23'), ('車両運搬具', 'G25'), ('工具器具及び備品', 'G26'),
          ('土地', 'G27'), ('長期前払費用', 'G48'), ('差入保証金', 'G49'), ('営業権', 'G50'),
          ('その他', None), ('資産合計', None)]
ws['H4'] = '流動資産'; ws['I4'] = "='BS(借方)'!G5"
rows_l = [('現預金', "='BS(借方)'!G6"), ('売掛金', "='BS(借方)'!G7"),
          ('在庫', "='BS(借方)'!G8+'BS(借方)'!G9"), ('前払費用', "='BS(借方)'!G11"),
          ('仮払金', "='BS(借方)'!G14"), ('その他', "='BS(借方)'!G15")]
for i, (lab, f) in enumerate(rows_l):
    ws.cell(row=5 + i, column=8).value = lab
    ws.cell(row=5 + i, column=9).value = f
ws['H11'] = '固定資産'; ws['I11'] = "='BS(借方)'!G20"
rows_f = [('建物附属設備', "='BS(借方)'!G23"), ('構築物', "='BS(借方)'!G24"),
          ('車両運搬具', "='BS(借方)'!G25"), ('工具器具及び備品', "='BS(借方)'!G26"),
          ('土地', "='BS(借方)'!G27"), ('長期前払費用', "='BS(借方)'!G48"),
          ('差入保証金', "='BS(借方)'!G49"), ('営業権', "='BS(借方)'!G50"),
          ('その他', "='BS(借方)'!G28+'BS(借方)'!G47")]
for i, (lab, f) in enumerate(rows_f):
    ws.cell(row=12 + i, column=8).value = lab
    ws.cell(row=12 + i, column=9).value = f
ws['H21'] = '資産合計'; ws['I21'] = "='BS(借方)'!G61"
ws['K3'] = '進行期（2026年6月末TB）'
ws['K3'].font = Font(name=YG, size=10, bold=True, color=NAVY)
ws['K4'] = '流動負債'; ws['L4'] = "='BS (貸方)'!G5"
rows_d = [('買掛金', "='BS (貸方)'!G6"), ('短期借入金', "='BS (貸方)'!G7"),
          ('未払金', "='BS (貸方)'!G8"), ('未払費用', "='BS (貸方)'!G9"),
          ('未払消費税等', "='BS (貸方)'!G11"), ('預り金', "='BS (貸方)'!G12"),
          ('カード未払金', "='BS (貸方)'!G13")]
for i, (lab, f) in enumerate(rows_d):
    ws.cell(row=5 + i, column=11).value = lab
    ws.cell(row=5 + i, column=12).value = f
ws['K12'] = '固定負債'; ws['L12'] = "='BS (貸方)'!G18"
rows_g = [('長期借入金', "='BS (貸方)'!G19"), ('役員借入金', "='BS (貸方)'!G20"),
          ('長期未払金', "='BS (貸方)'!G21")]
for i, (lab, f) in enumerate(rows_g):
    ws.cell(row=13 + i, column=11).value = lab
    ws.cell(row=13 + i, column=12).value = f
ws['K16'] = '負債合計'; ws['L16'] = "='BS (貸方)'!G25"
ws['K18'] = '資本金'; ws['L18'] = "='BS (貸方)'!G36"
ws['K19'] = '利益剰余金'; ws['L19'] = "='BS (貸方)'!G43"
ws['K20'] = '純資産合計'; ws['L20'] = "='BS (貸方)'!G50"
ws['K21'] = '負債・純資産合計'; ws['L21'] = "='BS (貸方)'!G51"
for r in range(3, 22):
    for col in (8, 11):
        ws.cell(row=r, column=col).font = Font(name=YG, size=9,
                                               bold=ws.cell(row=r, column=col).value in
                                               ('流動資産', '固定資産', '資産合計', '流動負債',
                                                '固定負債', '負債合計', '純資産合計', '負債・純資産合計'))
    for col in (9, 12):
        ws.cell(row=r, column=col).font = Font(name=ARIAL, size=9)
        ws.cell(row=r, column=col).number_format = '#,##0,;[Red]△ #,##0,;"－"'
for col, w in zip('HIJKL', (17, 12, 1.2, 17, 12)):
    ws.column_dimensions[col].width = w
ws['H23'] = '※進行期は2026年6月末（7か月）の試算表ベース。決算整理（減価償却・棚卸評価等）は未了。'
ws['H23'].font = Font(name=YGM, size=9)

# ---------------------------------------------------------------- 仕上げ（数式エラーの解消）
# 第2期（C列）はデータ未受領のため、C列を分母とする比率が #DIV/0! になる。IFERROR で包む。
import re as _re
_div = _re.compile(r'^=(?!IFERROR)(.+/.+)$')
for _name in ('SGA', 'CR', 'PL', 'PLハイライト', 'BSハイライト'):
    _ws = wb[_name]
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and _c.value.startswith('='):
                _m = _div.match(_c.value)
                if _m:
                    _c.value = '=IFERROR({},"")'.format(_m.group(1))

# 従業員シート：テンプレートの年齢・勤続年数の数式と日付書式を除去する
_ws = wb['従業員']
_last = 5 + len(LP.EMPLOYEES)
for _r in range(5, 35):
    for _col in range(2, 18):
        _cell = _ws.cell(row=_r, column=_col)
        if (_r > _last and _r != _last + 2) or (isinstance(_cell.value, str)
                                                and _cell.value.startswith('=ROUNDDOWN')):
            sset(_ws, _r, _col, None)
        if _col in (7, 8):
            _cell.number_format = '0.0;;"－"'
        elif _col in (4, 5, 6):
            _cell.number_format = '0;;"－"'

wb.save(OUT)
print('saved:', OUT)
