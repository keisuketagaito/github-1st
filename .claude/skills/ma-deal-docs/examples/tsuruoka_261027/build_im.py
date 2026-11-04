import os
"""鶴岡中央青果 企業概要書：公式マスター（master.xlsx）の貼付範囲から表を貼り直す（鳥塚準拠）。
ベース：リポジトリの 企業概要書_鶴岡中央青果_261019.pptx（34ページ）
"""
import sys, copy
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts'))
from im_paste import *
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor
from lxml import etree

S = os.environ.get('MA_WORK', '.')   # 作業フォルダ（up5/=受領ファイル、master/・im/=出力）
BASE = os.environ.get('REPO', '.') + '/株式会社鶴岡中央青果/企業概要書_鶴岡中央青果_261019.pptx'
wb = load_workbook(f'{S}/master/master.xlsx', data_only=True)
prs = Presentation(BASE)
SL = list(prs.slides)
v = lambda sh, a: wb[sh][a].value
k = lambda x: f'{x / 1000:,.0f}'          # 円→千円表記
pct = lambda x: f'{x * 100:.1f}%'


def sld(n):
    return SL[n - 1]


def drop_tables(s):
    for sh in list(s.shapes):
        if sh.has_table and sh.name.startswith('表（'):
            sh._element.getparent().remove(sh._element)


def drop_notes(s, ids=()):
    for sh in list(s.shapes):
        if sh.shape_id in ids:
            sh._element.getparent().remove(sh._element)


def fit(ref, target, hmax=1.0):
    """表の高さが target（インチ）に収まる行高倍率"""
    return min(hmax, target / range_height(wb, ref))


def paste(s, ref, x, y, w, fs, target=None, hmax=1.0):
    hs = fit(ref, target, hmax) if target else hmax
    t = paste_range(s, wb, ref, x, y, w, hscale=hs, fscale=fs)
    return t, y + range_height(wb, ref, hscale=hs)


def body_text(s, sid):
    return shape_by_id(s, sid)


def move_slide(prs, old_index, new_index):
    lst = prs.slides._sldIdLst
    el = list(lst)[old_index]
    lst.remove(el)
    lst.insert(new_index, el)


def dup_slide(prs, src, keep=lambda sh: True):
    """同じレイアウトで新規スライドを作り、src の図形（表以外）を複製する。"""
    new = prs.slides.add_slide(src.slide_layout)
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    for sh in src.shapes:
        if keep(sh):
            new.shapes._spTree.append(copy.deepcopy(sh._element))
    return new


# ---------------------------------------------------------------- 表紙
for sh in sld(1).shapes:
    if sh.has_text_frame and '企業概要書' in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace('青果卸売業', '野菜卸売業')

# ---------------------------------------------------------------- 基本情報
s = sld(5)
drop_tables(s)
t, yb = paste(s, '会社概要!B2:C15', 0.34, 2.06, 6.41, 0.83, target=4.75)
place(shape_by_id(s, 16), 6.95, 1.62, 4.40, yb - 1.62 + 0.05)
place(shape_by_id(s, 17), 6.95, 2.20, 4.40, yb - 2.20)
ph = 'PLハイライト'
fill_comment(shape_by_id(s, 17), [
    '≪鶴岡市内唯一の青果市場≫',
    '・市内の青果市場は当社のみ（2026年1月に2社体制から1社に）。山形県知事の認定を受けた地方卸売市場の開設者兼卸売業者。',
    '≪農家の持込を全量受け入れる集荷力≫',
    '・持込は数量が多くても原則すべて受け入れる。地元優先で販売し、過剰分は豊洲市場・他県へ出荷。',
    '≪安定した売上規模と分散した品揃え≫',
    f'・売上高は約12億円で安定。野菜約6割・果実約3割で、だだちゃ豆・庄内柿・メロン等の地場産品を扱う。',
    '≪足元の収益改善≫',
    f'・2026年4〜9月の営業利益は{k(v(ph, "G8"))}千円（前年同期{k(v(ph, "G38"))}千円）。人件費・運賃の削減による。',
    '≪事業用不動産を自社保有≫',
    '・本社市場とパッケージ場は隣接する自社保有地。賃借物件はなく地代負担がない。',
    '≪確認中の論点≫',
    '・(株)主婦の店への依存度、社内勘定の性質、2026年9月の新規借入の使途、不動産の時価。',
], body=8.5, head=9, after=1)

# ---------------------------------------------------------------- 財務ハイライト（P/L）
s = sld(7)
drop_tables(s); drop_notes(s, (18,))
t, yb = paste(s, 'PLハイライト!B2:G25', 0.34, 1.55, 6.62, 0.875, target=5.0)
note(s, 0.34, yb + 0.1, 6.62, [
    f'※EBITDA＝営業利益＋減価償却費。調整後EBITDAの3期平均は{k(v(ph, "C29"))}千円、直近2期平均は{k(v(ph, "C30"))}千円。',
    '※進行期は2026年4〜9月の試算表（6か月累計、年換算しない）。前期比は前年同期比。調整項目は同じ方針で6か月分を算定（賞与は年額の6/12）。'], size=7.5)
g = lambda a: v(ph, a)
set_lead(s, f'売上高は約12億円で安定。2026/3期は営業黒字へ回復し、調整後EBITDAは3期平均{k(g("C29"))}千円。'
            f'進行期（2026年4〜9月）は人件費・運賃の削減で営業利益{k(g("G8"))}千円と前年同期の{g("G8") / g("G38"):.1f}倍。')
dpers = v('PLハイライト', 'G45') - g('G43'); dfre = v('SGA', 'K17') - g('G44'); ddep = g('G15') - g('G41')
fill_comment(shape_by_id(s, 9), [
    '≪売上高≫',
    f'・{k(g("D4"))}→{k(g("E4"))}→{k(g("F4"))}千円と約12億円で安定。繁忙期は7〜12月、閑散期は1〜5月。',
    '≪売上総利益≫',
    f'・粗利率は{pct(g("D7"))}→{pct(g("E7"))}→{pct(g("F7"))}と低下（要因は要確認）。',
    '≪営業利益≫',
    '・2025/3期は粗利減少で営業赤字。2026/3期は役員報酬・賞与・償却費の減少で黒字に回復。',
    '≪進行期（2026年4〜9月）≫',
    (f'・売上高{k(g("G4"))}千円（前年同期比{pct(g("G5"))}）。販管費は前年同期比{pct(v("SGA", "K40") / g("G37") - 1)}'
     f'（人件費{k(dpers)}千円、運賃{k(dfre)}千円、償却費{k(ddep)}千円）。').replace('-', '△'),
    f'・営業利益{k(g("G8"))}千円（前年同期{k(g("G38"))}千円）、調整後EBITDA{k(g("G25"))}千円。',
], body=8.5, head=9.5, after=2)
fill_comment(shape_by_id(s, 10), [
    '【調整項目】',
    '≪役員報酬・顧問料・社会保険料≫',
    '・退任役員（代表者・配偶者・監査役・退任済みの常務）の報酬を全額控除（社保15%連動）し、代表者の顧問料（月18万円）を計上',
    '≪賞与≫',
    '・賞与規定はなく業績連動。2026/3期は支給なし。継続支給を前提に前2期平均（社保込）を控除（保守的に想定）',
    '≪寄付金・受取奨励金≫',
    '・譲渡後に発生しない寄付金と、経常的に発生する受取奨励金（純額）を加算',
], body=8, head=9, after=1)

# ---------------------------------------------------------------- 財務ハイライト（B/S）
s = sld(8)
drop_tables(s); drop_notes(s, (15,))
t, yb = paste(s, 'BSハイライト!B2:F21', 0.37, 1.55, 5.47, 0.875, target=5.0)
note(s, 0.37, yb + 0.1, 5.47, ['※2026年3月期末（評価基準日）。3期推移・2026年9月末は「財務情報詳細｜貸借対照表」に記載。'], size=7.5)
bs = lambda a: v('BS (貸方)', a)
fill_comment(shape_by_id(s, 7), [
    '≪現金及び預金≫',
    f'・{k(v("BS(借方)", "F6"))}千円：鶴岡信用金庫84,940千円、山形銀行63,122千円、荘内銀行59,634千円ほか。',
    '≪売上債権・棚卸資産≫',
    '・売掛金は(株)主婦の店7,194千円が最大。棚卸資産は商品・資材・地物で、廃棄を要する滞留在庫はないとのこと。',
    '≪固定資産≫',
    '・土地93,137千円：本社・市場70,757千円、パッケージ場21,380千円、旧倉庫敷地1,000千円。',
    '・保証金33,303千円：仕入代金のデポジット（約10日分）で取引終了時に返金。',
    '≪負債≫',
    f'・長期借入金{k(bs("F19"))}千円：鶴岡信用金庫182,840千円、山形銀行70,250千円、日本政策金融公庫107,968千円。',
    f'・2026年9月末は長期借入金{k(bs("G19"))}千円（3月末比+{k(bs("G19") - bs("F19"))}千円）、現預金{k(v("BS(借方)", "G6"))}千円（試算表）。新規借入の使途・条件は要確認。',
], body=8.5, head=9.5, after=2)

# ---------------------------------------------------------------- 事業の仕組みと収益源
s = sld(12)
drop_tables(s); drop_notes(s, (30,))
t, yb = paste(s, '売上構成!B4:E11', 0.49, 3.95, 10.71, 1.0, target=2.1)
note(s, 0.49, yb + 0.08, 10.71, [
    '※決算書では売上高を「商品売上高」1本で表示しているため、収益源別の内訳は試算表（2026年4〜9月）で代替。資材売上には自社利用分の振替を含む（同額を資材仕入に計上）。',
    f'※手数料率8.5%から逆算した受託販売の取扱高は6か月で約{v("売上構成", "C17") / 1e6:,.0f}百万円（推計）。センター費は月約{v("売上構成", "C8") / 6 / 1e6:.1f}百万円。',
    '※上記のほか、受取奨励金（内容は要確認）、仲卸業者等からのゴミ処理料・冷蔵庫保管料を営業外収益に計上。仲卸業者数・売買参加者数は要確認。'], size=7.5)

# ---------------------------------------------------------------- 主要得意先・仕入先（年間取引高に差替え）
s = sld(15)
drop_tables(s); drop_notes(s, (15,))
t, yb = paste(s, '取引先!B2:M17', 0.34, 1.55, 11.0, 0.72, target=4.4)
note(s, 0.34, yb + 0.2, 11.0, [
    '※買参人別順位表（販売管理システム、各期上位20位）の年間売上金額。上位20位圏外の年度は「－」。構成比は決算書の売上高に対する比率（集計基準が異なるため参考値）。業種等は公開情報からの推察を含む。',
    '※社内勘定は近郷部の集荷担当別（KK）・パッケージ部門別の勘定で、仕入側にも同名の勘定がある（性質は要確認）。長井中央青果(株)の2024/3期は旧コード分を合算。'], size=7.5)
tb = lambda a: v('取引先', a)
set_lead(s, f'(株)主婦の店（鶴岡店）が売上高の{pct(tb("L5"))}を占める最大の販売先。生協・近隣の青果卸・食品加工業者へも販売し、'
            '上位先は3期を通じて概ね同じ顔ぶれで長期・安定的な取引関係にある。')
s = sld(16)
drop_tables(s); drop_notes(s, (17,))
t, yb = paste(s, '仕入先!B2:L17', 0.34, 1.55, 11.0, 0.72, target=4.4)
note(s, 0.34, yb + 0.2, 11.0, [
    '※荷主別順位表（販売管理システム、各期上位20件）の年間仕入金額。上位20件圏外の年度は「－」。構成比は決算書の仕入高に対する比率（集計基準が異なるため参考値）。所在地・業種等は推察を含む。',
    '※個人生産者は匿名で記載。社内勘定は近郷部の集荷担当別（KK）・パッケージ部門別の勘定（性質は要確認）。'], size=7.5)
sp = lambda a: v('仕入先', a)
set_lead(s, f'公設庄内青果物地方卸売市場（三川町）の卸売業者2社からの仕入が{pct(sp("K5") + sp("K6"))}を占める。'
            '長井中央青果(株)等の県内市場、JA・地元生産者からも集荷し、品揃えを確保している。')

# ---------------------------------------------------------------- 株主及び役員
s = sld(20)
drop_tables(s)
t, yb = paste(s, '株主!B2:H15', 0.85, 2.00, 9.6, 0.78, target=2.55)
place(shape_by_id(s, 3), 0.34, yb + 0.12)
t, yb2 = paste(s, '取締役!B2:H15', 0.85, yb + 0.55, 9.6, 0.78, target=1.7)
place(shape_by_id(s, 16), 0.34, max(yb2 + 0.15, 7.25))

# ---------------------------------------------------------------- 従業員一覧
s = sld(21)
drop_tables(s); drop_notes(s, (15,))
t, yb = paste(s, '従業員!B4:Q28', 0.34, 1.5, 11.0, 0.75, target=5.35)
note(s, 0.34, yb + 0.08, 11.0, [
    '※職員名簿（2026年9月1日現在）。役員は含まない。賞与規定はなく業績に応じて支給。退職金規程あり（中退共・オリックス生命で外部積立。会社の追加負担は基本的にないとの認識）。',
    '※法人事業概況説明書（2026/3期末）では従業員30名（常勤役員3名含む）と記載があり、人数の定義差は要確認。就業規則・定年（60歳／65歳）は要確認。'], size=7.5)
em = lambda a: v('従業員', a)
set_lead(s, f'従業員{em("E37")}名（平均年齢{em("E38"):.1f}歳・平均勤続{em("E39"):.1f}年）。勤続20年以上のベテランが多い一方、60歳以上が{em("E40")}名在籍。'
            f'セリ人資格者{em("E42")}名が市場運営を担う。')

# 組織図の人員構成ボックス（平均年齢はマスターの従業員シートから）
for sh in sld(19).shapes:
    if sh.has_text_frame and '平均年齢' in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if '平均年齢' in r.text:
                    import re as _re
                    r.text = _re.sub(r'平均年齢[\d.]+歳', f'平均年齢{em("E38"):.1f}歳', r.text)
                    r.text = _re.sub(r'平均勤続[\d.]+年', f'平均勤続{em("E39"):.1f}年', r.text)

# ---------------------------------------------------------------- 拠点・不動産・設備
s = sld(22)
drop_tables(s)
t, yb = paste(s, '拠点(本社等)!B2:G6', 0.34, 2.05, 11.0, 0.83, target=2.1)
place(shape_by_id(s, 15), 0.34, yb + 0.25, 11.0, 7.25 - yb - 0.25)
place(shape_by_id(s, 16), 0.34, yb + 0.8, 11.0, 7.2 - yb - 0.8)

s = sld(23)
drop_tables(s)
t, yb = paste(s, '不動産概況!B4:J9', 0.34, 1.6, 11.0, 1.0)

s = sld(24)
drop_tables(s)
t, yb = paste(s, '不動産(所有)!B3:K15', 0.34, 2.02, 11.0, 0.72, target=3.3)
place(shape_by_id(s, 3), 0.34, yb + 0.4); place(shape_by_id(s, 15), 2.15, yb + 0.4)
place(shape_by_id(s, 16), 0.34, yb + 0.95)
set_text(shape_by_id(s, 16), ['※建物の簿価は棟別ではなく改修工事を含めた合計（BS建物・建物付属設備）。棟別の帳簿内訳は「保有設備一覧」に記載。庄内町長沼の土地は建築不可（代表者談）。',
                              '※1981年築の建物は新耐震基準（1981年6月施行）の適用有無を建築確認日で要確認。土地の時価は鑑定・路線価で要確認。担保設定は登記簿謄本で要確認。'], size=8)

s = sld(25)
drop_tables(s)
import json
EQ = json.load(open(f'{S}/master/paste_index.json'))['EQ']
t, yb = paste(s, EQ[0], 0.34, 1.5, 6.2, 0.75, target=5.6)
t, yb2 = paste(s, EQ[1], 6.75, 1.5, 4.6, 0.75, target=2.4)
place(shape_by_id(s, 16), 6.75, yb2 + 0.12)

# ---------------------------------------------------------------- 財務情報詳細（財務入力シートをそのまま貼る：公式マスターの運用）
s = sld(27)
drop_tables(s)
t, yb = paste(s, "'BS(借方)'!B3:G61", 0.34, 1.52, 5.42, 0.7, target=5.8)
t, yb2 = paste(s, "'BS (貸方)'!B3:G25", 5.88, 1.52, 5.42, 0.7, target=2.4, hmax=0.55)
ref_na = "'BS (貸方)'!B34:G51"
hna = min(0.55, 2.4 / range_height(wb, ref_na))
y_na = yb - range_height(wb, ref_na, hscale=hna)
paste_range(s, wb, ref_na, 5.88, y_na, 5.42, hscale=hna, fscale=0.7)
set_lead(s, f'直近3期と進行期（2026年9月末）の貸借対照表。総資産は借入金の返済で{v("BS(借方)", "D61") / 1e6:,.0f}→{v("BS(借方)", "F61") / 1e6:,.0f}百万円に圧縮したが、'
            f'9月末は新規借入で長期借入金が{bs("G19") / 1e6:,.0f}百万円に増加。')

for n, ref, cid in ((28, 'PL!B2:L56', 6), (29, 'SGA!B2:L40', 3)):
    s = sld(n)
    drop_tables(s)
    t, yb = paste(s, ref, 0.34, 1.5, 7.3, 0.62, target=5.75)
    place(shape_by_id(s, 2), 7.8, 1.45, 3.55, 6.1)
    place(shape_by_id(s, cid), 7.8, 2.05, 3.53, 5.4)
s = sld(28)
fill_comment(shape_by_id(s, 6), [
    '≪売上高≫',
    '・決算書では「商品売上高」1本で表示。受託販売は2019年10月以降、手数料（8.5%）のみの純額表示。売上割戻は期限内に代金を支払った八百屋等への完納奨励金（市場の商慣習）。',
    '≪売上原価≫',
    '・粗利率は18.6%→17.9%→16.6%。棚卸資産は最終仕入原価法で評価。',
    '≪営業外損益≫',
    '・受取奨励金（毎期約2.3〜2.5百万円、内容は要確認）、雑収入（ゴミ処理料・冷蔵庫保管料等）を計上。支払利息は年約4百万円。',
    '≪特別損益≫',
    '・2026/3期の前期損益修正損20,900千円は過年度の保険解約に係る訂正額（非経常）。',
    '≪進行期（2026年4〜9月）≫',
    f'・売上高{k(v("PL", "K11"))}千円、営業利益{k(v("PL", "K19"))}千円、経常利益{k(v("PL", "K36"))}千円（6か月累計、試算表）。繁忙期を含むため年換算しない。',
], body=8, head=9, after=2)
set_lead(s, '売上高は約12億円で推移。2026/3期は保険積立金の取崩しで当期純損失。進行期（2026年4〜9月）は経常利益'
            f'{k(v("PL", "K36"))}千円と前年同期（{k(v("PLハイライト", "G39"))}千円）を大きく上回る。')
s = sld(29)
fill_comment(shape_by_id(s, 3), [
    '≪人件費≫',
    '・役員報酬は22,200→15,360千円に減少。賞与規定はなく業績に応じて支給（2026/3期は支給なし）。',
    f'・進行期は役員報酬・給料手当等が前年同期比{k(dpers)}千円。2026年10月の最低賃金改定に伴い給与を改定（影響額は要確認）。'.replace('比-', '比△'),
    '・退職金は中退共・オリックス生命で外部積立。',
    '≪運賃・車両関連費≫',
    '・自社配送に伴う費用（東京向け出荷等の一部のみ外部運送）。車両は自己所有（償却済み）とリース（年約6百万円）を併用。',
    '≪租税公課・消費税引当金≫',
    '・消費税は税込経理。決算では納付額を租税公課、試算表では消費税引当金に計上。',
    '≪減価償却費≫',
    '・13,664→8,879千円へ減少（車両等の償却終了）。進行期は月次で概算計上。',
], body=8, head=9, after=2)

# ---------------------------------------------------------------- Appendix
s = sld(31)
drop_tables(s); drop_notes(s, (15,))
t, yb = paste(s, '沿革!B4:C16', 0.49, 1.6, 10.71, 1.0, target=4.9)
note(s, 0.49, yb + 0.1, 10.71, ['※出所：履歴事項全部証明書、地方卸売市場の認定証、固定資産税課税明細書、個別注記表、固定資産内訳書。創業の経緯・代表者の経歴は売主様へ確認中。'])
s = sld(32)
drop_tables(s)
t, yb = paste(s, 'その他補足!B4:C15', 0.49, 1.65, 10.71, 1.0, target=5.2)
set_lead(s, '許認可・競合環境・取引慣行・人事制度等の補足情報。地方卸売市場は2020年6月に山形県知事の認定を受け、2026年2月に変更の認定を受けている。')

# ---------------------------------------------------------------- 追加ページ①：品目別売上高の3期推移（仕入・販売の仕組みの後）
src = sld(15)
new = dup_slide(prs, src, keep=lambda sh: not sh.has_table and sh.shape_id not in (15,) and not (sh.has_text_frame and sh.text_frame.text.startswith('※')))
for sh in new.shapes:
    if sh.has_text_frame and sh.text_frame.text == '主要得意先の3期推移':
        set_text(sh, '品目別売上高の3期推移')
lead_new = None
for sh in new.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith('(株)主婦の店'):
        lead_new = sh
pj = lambda a: v('品目別売上', a)
veg = pj('H25') / pj('H23'); fru = pj('H26') / pj('H23')
t, yb = paste(new, '品目別売上!B4:J23', 0.34, 1.5, 8.2, 0.92, target=5.0)
note(new, 0.34, yb + 0.2, 8.2, [
    '※分類別商品売上実績表（販売管理システム、各期3月末現在の年間累計）。決算書の商品売上高とは集計範囲が異なる（受託販売分を販売額の総額で集計していると推察：要確認）。'], size=7.5)
# 区分別構成（2026/3期）の円グラフ（数値はマスターの集計）
cd = CategoryChartData()
cats = [('野菜', pj('H25')), ('果実', pj('H26')), ('花き', pj('H21')), ('加工品', pj('H22'))]
tot = sum(x for _, x in cats)
cd.categories = [f'{c} {x / tot * 100:.1f}%' for c, x in cats]
cd.add_series('2026/3期', [round(x / 1000) for _, x in cats])
gf = new.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(8.75), Inches(1.6), Inches(2.6), Inches(3.4), cd)
ch = gf.chart
ch.has_title = True
ch.chart_title.text_frame.text = '区分別構成（2026/3期）'
for p in ch.chart_title.text_frame.paragraphs:
    for r in p.runs:
        r.font.size = Pt(10); r.font.bold = True; r.font.name = JP
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(8.5); ch.legend.font.name = JP
plot = ch.plots[0]
plot.has_data_labels = False
for i, col in enumerate(['0B3041', '2E75B6', '9DC3E6', 'BFBFBF']):
    pt = plot.series[0].points[i]
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(col)
comment(new, shape_by_id(sld(7), 8), shape_by_id(sld(7), 9), 8.75, 5.2, 2.6, 2.1, [
    '≪品揃え≫',
    f'・野菜{pct(veg)}・果実{pct(fru)}。地場の葉物・だだちゃ豆・庄内柿・メロンが柱。',
    '≪季節性≫',
    '・7〜8月の枝豆・メロン、10〜11月の庄内柿・梨が繁忙期を形成。',
], head='Comment', body_dy=0.55)
set_lead(new, f'年間取扱高は約{pj("H23") / 1e8:.1f}億円で3期とも横ばい。葉茎菜類・果菜類・メロン類が上位を占め、だだちゃ豆・庄内柿・メロン等の'
              '庄内の特産品を扱う。野菜と果実をバランスよく揃える総合青果市場。')
# 位置：仕入・販売の仕組み（13）の直後
move_slide(prs, len(prs.slides) - 1, 13)

# ---------------------------------------------------------------- 追加ページ②：許認可一覧（保有設備一覧の後）
src = SL[31]   # その他補足情報（元32ページ）
new2 = dup_slide(prs, src, keep=lambda sh: not sh.has_table)
for sh in new2.shapes:
    if sh.has_text_frame and sh.text_frame.text == 'その他補足情報':
        set_text(sh, '許認可一覧')
    if sh.has_text_frame and sh.text_frame.text == 'Appendix':
        set_text(sh, '組織、拠点、設備等')
t, yb = paste(new2, '許認可!B5:F8', 0.49, 1.65, 10.71, 1.0)
note(new2, 0.49, yb + 0.1, 10.71, ['※出所：地方卸売市場の認定証・変更の認定証（山形県知事）、履歴事項全部証明書。認定の条件（取扱品目・開場日等）は業務規程で要確認。'])
set_lead(new2, '卸売市場法に基づき、山形県知事から地方卸売市場の認定を受けている（2020年6月認定、2026年2月変更認定）。市場の開設・運営は事業継続の前提となる許認可。')
# 位置：保有設備一覧（元25 → 追加ページ①の挿入で26）の直後
move_slide(prs, len(prs.slides) - 1, 26)

for s in prs.slides:
    remove_null_pics  # noqa（テンプレート由来の不要画像はベースで除去済み）
out = f'{S}/im/企業概要書_鶴岡中央青果_261027.pptx'
import os; os.makedirs(f'{S}/im', exist_ok=True)
prs.save(out)
print('saved', out, len(prs.slides))
