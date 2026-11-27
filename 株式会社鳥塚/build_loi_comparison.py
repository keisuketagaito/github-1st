"""株式会社鳥塚 各社意向表明書 比較表の作成スクリプト。

過去案件（タビュラ株式会社）の比較表を書式の型として読み込み、
タイトル・列見出し・項目・スコア欄の書式を引き継いで本件の比較表を組む。

使い方:
    python3 build_loi_comparison.py <タビュラ比較表.xlsx> <出力先.xlsx>

スコアは各項目で売手（鳥塚社長・兼子副社長）にとって有利な順に 1〜3（同順位あり、
タビュラ表と同じく密な順位付け）。合計が小さいほど売手に有利。「-」は採点対象外。
"""
import copy
import math
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side

SRC, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------- 内容
TITLE = "【比較表】株式会社鳥塚御中_各社意向表明書"
DISCLAIMER = ("※本比較表の項目及びスコアは、当社にて設定した比較項目に基づくものです。"
              "最終的なご判断・意思決定は、売主様の経営陣にてお決めください。")
# 受領（書面日付）順
BIDDERS = ["株式会社創桐", "株式会社ストラウト", "リージョナルフィッシュ株式会社"]

# (項目名, scored, [(本文, スコア) x 3])  scored=False の行はスコア「-」
ROWS = [
    ("買手概要\n（書面日付）", False, [
        ("兵庫県三田市／1987年設立／従業員83名\n"
         "住宅用基礎ユニット鉄筋の製造及び販売\n"
         "代表取締役会長 山本新九郎、代表取締役社長 山本英輝\n"
         "（2026年11月23日付）", "-"),
        ("沖縄県沖縄市／2020年設立／社員10名\n"
         "AI・IoTを活用した陸上養殖、種苗生産、養殖場の運営受託及び経営改善支援\n"
         "グループ売上210百万円（2025年10月期）\n"
         "関係会社：合同会社白水養殖（クルマエビ）、有限会社柴崎養鱒場（ニジマス）\n"
         "代表取締役社長 平林馨（過去に貴社の経営改善に関与）\n"
         "（2026年11月27日付）", "-"),
        ("京都府京都市／2019年設立／社員135名（役員含む）\n"
         "水産物の品種改良（ゲノム編集・選抜育種）、陸上養殖システムの構築・運営、種苗・水産物の生産販売\n"
         "資金調達総額154億円（補助金含む）\n"
         "代表取締役 梅川忠典\n"
         "（2026年11月30日付）", "-"),
    ]),
    ("株式譲受の形態", False, [
        ("200株（発行済株式全株）\n"
         "鳥塚社長150株・兼子副社長50株\n"
         "弊社又は弊社が指定する取得主体が譲受（最終契約締結までに確定）\n"
         "※ご希望があれば、実行後に株式の一部再譲渡（20〜50％程度）を協議可（確約ではない）", "-"),
        ("200株（発行済株式全株）\n"
         "鳥塚社長150株・兼子副社長50株\n"
         "弊社が譲受", "-"),
        ("200株（発行済株式全株）\n"
         "鳥塚社長150株・兼子副社長50株\n"
         "弊社が譲受", "-"),
    ]),
    ("株式譲受対価", True, [
        ("20,000,000円\n"
         "（初期的な希望価格。DDによる借入金・役員借入金・運転資本・簿外債務等の精査結果を踏まえ、最終価格及び調整方法を協議）", 1),
        ("20,000,000円\n"
         "（事業用不動産の対価及び役員借入金の取扱いとは別建て）", 1),
        ("15,000,000円", 2),
    ]),
    ("事業用不動産の取扱い\n（鳥塚社長個人名義：\n各池及び本社周辺の\n土地25筆・建物3棟）", True, [
        ("総額5,000万円程度を想定\n"
         "買収後に貴社による資金調達を活用した取得を含め検討\n"
         "売買価格・取得主体・時期・方法は、不動産評価・権利関係・金融機関との協議結果を踏まえ協議", 2),
        ("50,000,000円〜60,000,000円程度を想定\n"
         "本件実行と同時に、貴社・弊社又は弊社指定先が不動産鑑定評価等に基づく適正価格で譲受\n"
         "譲受主体・資金手当て（弊社による貴社への出資を含む）は取引金融機関と協議\n"
         "本件不動産に設定された貴社借入金の担保は譲受後も維持", 1),
        ("金額の提示なし（要確認）\n"
         "本件実行と同時又は実行後速やかに、貴社への売却その他適切な方法、又は弊社への直接譲渡により移転\n"
         "価格は合理的な方法で時価との整合性を確認し、最終契約締結までに協議\n"
         "上記以外でも事業用の社長個人名義不動産は対象に含む（ご自宅等の私用不動産は対象外）", 3),
    ]),
    ("役員借入金の取扱い\n（31,300千円：\n鳥塚社長14,000千円・\n兼子副社長17,300千円）", True, [
        ("本件実行後、うち30,000千円を返済する方針（可能な限り早期）\n"
         "具体的な返済時期・方法は協議", 2),
        ("本件実行時に返済せず、全額を貴社の借入金として存続\n"
         "返済時期・方法は財務基盤の再構築の内容と整合する形で別途協議", 3),
        ("本件実行時に貴社をして返済させる\n"
         "（又は一部債権放棄・株式譲渡対価との調整等による整理）\n"
         "方法・金額・時期は資金繰り及び金融機関借入金の返済状況を踏まえ、最終契約締結までに協議", 1),
    ]),
    ("株式対価・不動産対価・\n役員借入金返済の合計\n（概算・税引前）", True, [
        ("約100,000,000円\n"
         "（株式20,000千円＋不動産約50,000千円＋借入金返済30,000千円）\n"
         "※不動産取得・借入金返済はいずれも実行後で、時期は未確定", 1),
        ("70,000,000円〜80,000,000円\n"
         "（株式20,000千円＋不動産50,000〜60,000千円）\n"
         "※役員借入金31,300千円は貴社に存続（返済は財務基盤再構築の内容次第）", 2),
        ("46,300,000円＋不動産対価（金額未提示）\n"
         "（株式15,000千円＋借入金返済31,300千円＋不動産：要確認）\n"
         "※借入金は一部債権放棄・対価調整による整理の可能性あり", 3),
    ]),
    ("本件後の経営体制並びに\n鳥塚社長の関わり方", True, [
        ("現行の経営体制（鳥塚社長・兼子副社長中心）が基本\n"
         "鳥塚社長：取締役社長として継続\n"
         "兼子副社長：取締役（副社長の肩書の記載なし）\n"
         "成長目標：売上約3億円→8億円、営業利益1億円", 2),
        ("現行の経営体制（鳥塚社長・兼子副社長中心）が基本\n"
         "鳥塚社長：取締役社長として継続\n"
         "兼子副社長：取締役副社長として継続", 1),
        ("現行の経営体制（鳥塚社長・兼子副社長中心）が基本\n"
         "鳥塚社長：取締役社長として継続\n"
         "兼子副社長：取締役副社長として継続", 1),
    ]),
    ("役員派遣", False, [
        ("有\n"
         "弊社代表取締役会長 山本新九郎が貴社の取締役CFO（最高財務責任者）兼CSO（最高戦略責任者）として就任し、経営管理本部を管掌\n"
         "担当：金融機関との折衝、財務基盤の再構築、経営戦略の立案・推進\n"
         "上記以外の派遣は記載なし", "-"),
        ("協議事項\n"
         "弊社からの役員派遣の要否及びその時期は今後双方で協議", "-"),
        ("協議事項\n"
         "弊社からの役員派遣の要否及びその時期は今後双方で協議\n"
         "※役員とは別に、開発・生産現場への技術人材の派遣、採用面の支援を検討（シナジーとして記載）", "-"),
    ]),
    ("報酬内容\n（役員報酬・退職慰労金）", True, [
        ("役員報酬：現行水準の維持を基本とし、事業の成長・収益状況を踏まえ段階的な増額を目指す\n"
         "退職慰労金：記載なし", 1),
        ("役員報酬：原則として現行水準を維持\n"
         "退職慰労金：継続勤務が前提のため、本件実行時の支給は想定せず（監査役も同様）", 2),
        ("役員報酬：原則として現行水準を維持し、提携後の業績に応じて毎年協議\n"
         "退職慰労金：継続勤務が前提のため、本件実行時の支給は想定せず（監査役も同様）", 2),
    ]),
    ("その他役員・\n従業員の処遇", True, [
        ("監査役 鳥塚あゆみ様：本件実行時の退任が基本案（処遇・必要な決議等は協議）\n"
         "その他の役員：今後の面談でご意向等を伺い協議\n"
         "従業員：既存の雇用関係及び現行の雇用条件を維持", 2),
        ("監査役 鳥塚あゆみ様：本件実行時に退任\n"
         "取締役統括部長 兼子雅和様：継続勤務、現行の役職・役員報酬水準を維持\n"
         "従業員：全員の雇用契約を引継ぎ、現状水準以上の労働条件を維持", 1),
        ("監査役 鳥塚あゆみ様：本件実行時に退任\n"
         "取締役統括部長 兼子雅和様：継続勤務、現行の役職・役員報酬水準を維持\n"
         "従業員：全員の雇用契約を引継ぎ、現状水準以上の労働条件を維持", 1),
    ]),
    ("連帯保証及び担保の解除", True, [
        ("個人保証・個人資産への担保のご負担軽減を重要事項と認識\n"
         "解除又は変更に向け金融機関等と誠実に協議（方法・時期は債権者の承諾・借換条件等を踏まえ協議）", 2),
        ("本件実行後、弊社の責任で取引金融機関等と協議し、「経営者保証に関するガイドライン」等に沿って脱退・解除・抹消の手続を最大限進める\n"
         "※本件不動産の担保は維持（不動産の項参照）", 1),
        ("本件実行後、弊社の責任により脱退等ができるよう最大限努力", 1),
    ]),
    ("財務基盤・金融機関対応\n（実行の前提条件）", False, [
        ("山本会長（取締役CFO就任想定）が金融機関との折衝・財務基盤の再構築を担う\n"
         "前提条件の明示なし（調査状況・資金調達・債権者承諾・許認可等の確認状況により条件・時期は変更あり得る）", "-"),
        ("中小企業活性化協議会等を活用し、返済条件の再構築・資本性借入金の活用等により財務基盤を再構築（実行後6か月以内の計画合意を目途）\n"
         "【実行の前提条件】取引金融機関の了解：①株主の異動 ②計画合意までの元本返済猶予等 ③再構築の方向性（資本性借入金を含む）及び連帯保証の取扱い\n"
         "前提条件が整わない場合は条件・進め方を再協議", "-"),
        ("金融機関対応の明示なし\n"
         "資金面での協力、政府認定実績を活かした補助金・助成金の活用支援を検討（シナジーとして記載）", "-"),
    ]),
    ("許認可及び契約の継続", False, [
        ("漁業免許（区画漁業権）、河川占用許可等、愛知分場・尾上池の土地賃貸借契約の継続が前提\n"
         "名義・株主変更に伴う届出・承認・更新要件等をDDで確認し手続を協議", "-"),
        ("漁業免許（区画漁業権）、河川占用許可等、愛知分場・尾上池の土地賃貸借契約の継続が前提\n"
         "行政機関への届出・事前説明、賃貸人への通知・承諾の要否は別途相談", "-"),
        ("漁業免許（区画漁業権）、河川占用許可等、愛知分場・尾上池の土地賃貸借契約の継続が前提\n"
         "行政機関への届出・事前説明、賃貸人への通知・承諾の要否は別途相談", "-"),
    ]),
    ("独占交渉権", False, [
        ("2027年3月31日まで\n"
         "発生時点・対象範囲等は別途書面で定める（本書の提出のみでは義務は生じない）\n"
         "延長は別途協議", "-"),
        ("2027年3月末日まで\n"
         "期間満了時に交渉継続中の場合は協議の上延長", "-"),
        ("2027年3月末日まで\n"
         "期間満了時に交渉継続中の場合は協議の上延長", "-"),
    ]),
    ("デューデリジェンスについて", False, [
        ("・財務\n・税務\n・法務\n・事業\n"
         "（12月の稚鮎仕入・4月の放流時期に配慮／費用は原則弊社負担）", "-"),
        ("・財務（公認会計士）\n・法務（弁護士）\n・不動産（不動産鑑定士）\n"
         "・事業（事業計画の検証、実地調査、インタビュー）\n"
         "（繁忙期に配慮／費用は帳簿コピー代を除き弊社負担）\n"
         "※並行して取引金融機関へ事前相談", "-"),
        ("・財務（公認会計士）\n・法務（弁護士）\n・不動産（不動産鑑定士）\n"
         "・事業（事業計画の検証、実地調査、インタビュー）\n"
         "（繁忙期に配慮／費用は帳簿コピー代を除き弊社負担）", "-"),
    ]),
    ("想定スケジュール\n(決済予定日)", True, [
        ("2026年12月中旬〜2027年1月中旬：企業調査、金融機関等との協議\n"
         "2027年1月下旬：株式譲渡契約の締結\n"
         "2027年2月中旬：クロージング", 1),
        ("2026年12月中旬〜2027年2月中旬：企業調査、金融機関への事前相談\n"
         "2027年2月中旬〜3月上旬：金融機関の了解取得\n"
         "2027年3月上旬：株式譲渡契約の締結\n"
         "2027年3月下旬：決済\n"
         "実行後6か月以内：財務基盤再構築計画の策定・合意", 3),
        ("2027年1月下旬〜2月中旬：企業調査\n"
         "2027年2月下旬：調査結果のご報告・具体的なご提案\n"
         "2027年3月上旬〜中旬：株式譲渡契約の締結\n"
         "2027年3月中旬：決済", 2),
    ]),
]

NOTES_LABEL = "留意点・要確認事項"
NOTES = [
    "・対価は初期的な希望価格で、DD結果による調整が前提\n"
    "・不動産取得は買収後の資金調達を前提とし、時期・金額は未確定\n"
    "・役員借入金は30,000千円の返済方針（残1,300千円の扱い、鳥塚社長・兼子副社長の按分は要確認）\n"
    "・兼子副社長の役職は「取締役」と記載（副社長の肩書の扱いは要確認）\n"
    "・異業種（住宅用基礎鉄筋）。山本会長が取締役CFO兼CSOとして経営に参画（役員派遣の項参照）\n"
    "・将来の創業家への株式一部再譲渡（20〜50％）の協議を他社にない提案として記載",
    "・役員借入金31,300千円はクロージング時に返済されない（貴社に存続）\n"
    "・取引金融機関の了解が実行の前提条件。了解が得られない場合は条件再協議となり、実行の確度・時期に影響\n"
    "・中小企業活性化協議会等を活用した財務基盤再構築の進め方を具体的に記載\n"
    "・同業（陸上養殖・種苗生産、養殖場の経営改善）。代表は過去に貴社の経営改善に関与",
    "・株式対価は3社中最低（15,000千円）。不動産価格は未提示のため金額の確認が必要\n"
    "・役員借入金は実行時返済とするが、一部債権放棄・対価調整による整理も併記されており、手取額の確認が必要\n"
    "・シナジー（人工種苗の生産効率向上、琵琶湖産鮎の資源回復、設備更新・補助金活用、新魚種・遊休池活用、技術人材の派遣）を具体的に記載\n"
    "・書面日付は2026年11月30日付",
]
FOOTNOTE = ("※スコアは各項目について売手（鳥塚社長・兼子副社長）にとって有利な順に1〜3（同順位あり）。"
            "合計が小さいほど売手に有利。「-」は採点対象外。金額は意向表明書記載額による概算（税引前）。"
            "意向表明書は独占交渉権・秘密保持を除き法的拘束力なし。")

# ---------------------------------------------------------------- 書式
tpl = openpyxl.load_workbook(SRC).active
wb = openpyxl.load_workbook(SRC)
ws = wb.active

for rng in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(rng))
for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 40), max_col=12):
    for c in row:
        c.value = None
        c.style = "Normal"


def put(dst, src_ref, value=None, **align):
    s = tpl[src_ref]
    dst.font = copy.copy(s.font)
    dst.fill = copy.copy(s.fill)
    dst.border = copy.copy(s.border)
    dst.alignment = copy.copy(s.alignment)
    dst.number_format = s.number_format
    if align:
        a = copy.copy(dst.alignment)
        dst.alignment = Alignment(
            horizontal=align.get("horizontal", a.horizontal),
            vertical=align.get("vertical", a.vertical),
            wrap_text=align.get("wrap_text", a.wrap_text),
            indent=align.get("indent", a.indent),
        )
    if value is not None:
        dst.value = value


def set_right(cell, style):
    b = cell.border
    cell.border = Border(left=b.left, right=Side(style=style), top=b.top, bottom=b.bottom)


def set_font_size(cell, sz):
    f = copy.copy(cell.font)
    cell.font = Font(name=f.name, sz=sz, b=f.b, i=f.i, color=f.color, scheme=f.scheme,
                     family=f.family, charset=f.charset)


TXT = ["C", "E", "G"]
SCR = ["D", "F", "H"]
W_LABEL, W_TXT, W_SCR = 30, 60, 8.6
BODY_PT, LABEL_PT = 12, 14

for col, w in {"A": 5.25, "B": W_LABEL, "C": W_TXT, "D": W_SCR, "E": W_TXT,
               "F": W_SCR, "G": W_TXT, "H": W_SCR, "I": 3}.items():
    ws.column_dimensions[col].width = w
for col in ("J", "K"):
    ws.column_dimensions[col].width = 8.43


def disp_w(s):
    return sum(2 if unicodedata.east_asian_width(ch) in "WFA" else 1 for ch in s)


def lines(text, width_chars, pt):
    cap = width_chars * 11 / pt * 0.92  # 半角換算の1行あたり容量（余白控除）
    return sum(max(1, math.ceil(disp_w(p) / cap)) for p in text.split("\n"))


def row_height(texts_with_width, pt):
    n = max(lines(t, w, pt) for t, w in texts_with_width)
    return max(45, n * pt * 1.45 + 14)


# タイトル
put(ws["B2"], "B2", TITLE, horizontal="left")
ws.row_dimensions[2].height = 30

# 注意書き（タイトル直下）
ws["B3"] = DISCLAIMER
ws["B3"].font = Font(name=tpl["C5"].font.name, sz=12, b=True)
ws["B3"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws.merge_cells("B3:H3")
ws.row_dimensions[3].height = 26

# 見出し行
HR = 4
put(ws[f"B{HR}"], "B4")
for t, s, name in zip(TXT, SCR, BIDDERS):
    put(ws[f"{t}{HR}"], "C4", name)
    put(ws[f"{s}{HR}"], "D4")
set_right(ws[f"H{HR}"], "medium")
ws.row_dimensions[HR].height = 30

first = HR + 1
last = first + len(ROWS) - 1
for i, (label, scored, vals) in enumerate(ROWS):
    r = first + i
    tr = 5 if r == first else (14 if r == last else 6)  # タビュラ表の上端・中段・下端の罫線
    put(ws[f"B{r}"], f"B{tr}", label, wrap_text=True, horizontal="center")
    set_font_size(ws[f"B{r}"], LABEL_PT)
    for t, s, (text, score) in zip(TXT, SCR, vals):
        put(ws[f"{t}{r}"], f"C{tr}", text, horizontal="left", wrap_text=True, indent=1)
        set_font_size(ws[f"{t}{r}"], BODY_PT)
        put(ws[f"{s}{r}"], f"D{tr}", score)
    set_right(ws[f"H{r}"], "medium")
    ws.row_dimensions[r].height = row_height(
        [(label, W_LABEL)] + [(v[0], W_TXT) for v in vals], BODY_PT)

# 合計スコア
TR = last + 1
put(ws[f"B{TR}"], "B15", "合計スコア")
for t, s in zip(TXT, SCR):
    put(ws[f"{t}{TR}"], "C15", f"=SUM({s}{first}:{s}{last})")
    put(ws[f"{s}{TR}"], "D15")
    ws.merge_cells(f"{t}{TR}:{s}{TR}")
set_right(ws[f"H{TR}"], "medium")
ws.row_dimensions[TR].height = 45

# 留意点・要確認事項
NR = TR + 1
put(ws[f"B{NR}"], "B15", NOTES_LABEL)
set_font_size(ws[f"B{NR}"], LABEL_PT)
for t, s, note in zip(TXT, SCR, NOTES):
    put(ws[f"{t}{NR}"], "I15", note, horizontal="left", vertical="top", wrap_text=True, indent=1)
    set_font_size(ws[f"{t}{NR}"], BODY_PT)
    put(ws[f"{s}{NR}"], "D15")
    ws.merge_cells(f"{t}{NR}:{s}{NR}")
set_right(ws[f"H{NR}"], "medium")
ws.row_dimensions[NR].height = row_height([(n, W_TXT + W_SCR) for n in NOTES], BODY_PT)

# 注記
FR = NR + 1
ws[f"B{FR}"] = FOOTNOTE
ws[f"B{FR}"].font = Font(name=tpl["C5"].font.name, sz=11)
ws[f"B{FR}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws.merge_cells(f"B{FR}:H{FR}")
ws.row_dimensions[FR].height = 36

for r in range(FR + 1, 41):
    ws.row_dimensions[r].height = None

# 印刷設定：A3横・幅1ページ、見出し行を各ページに
ws.print_area = f"A1:I{FR + 1}"
ws.print_title_rows = f"{HR}:{HR}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.freeze_panes = f"C{HR + 1}"
ws.title = "比較表"

wb.save(OUT)
print("saved", OUT, "rows", first, "-", FR)
