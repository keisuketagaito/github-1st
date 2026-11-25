"""ジャパニアス株式会社 化学・バイオ事業（事業譲渡）企業概要書を当社テンプレートから組む。

表はすべて案件マスター（MA_WORK/work/master.xlsx：再計算・inject_cache済み）の貼付範囲を paste_range で貼る。
リード文・Comment・グラフの数字もマスターの値から生成する（パワポに数字を打たない）。
"""
import os, sys, copy, json
REPO = os.environ.get('REPO', os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
sys.path.insert(0, os.path.join(REPO, '.claude', 'skills', 'ma-deal-docs', 'scripts'))
import im_builder as ib
from im_paste import (paste_range, range_height, fill_comment, set_lead, note, clone, set_box, place, shape_by_id, remove,
                      remove_yellow, remove_null_pics, remove_watermark, comment)
from im_paste import set_text as _st


def set_text(sh, t, **kw):
    return _st(sh, t.split('\n') if isinstance(t, str) else t, **kw)
from openpyxl import load_workbook
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

W = os.environ.get('MA_WORK', '.')
OUT = os.environ.get('OUT', f'{W}/work/im.pptx')
wb = load_workbook(f'{W}/work/master.xlsx', data_only=True)
META = json.load(open(f'{W}/work/meta.json'))
REFE = json.load(open(f'{W}/work/ref_emp.json'))
v = lambda sh, a: wb[sh][a].value
k = lambda x: f'{x / 1000:,.0f}'                 # 円→千円
mm = lambda x: f'{x / 1e6:,.1f}'                # 円→百万円（1桁）
pc = lambda x, d=1: f'{x * 100:.{d}f}%'
ph = lambda a: v('PLハイライト', a)
ec = lambda lab: v('人員構成', REFE[lab])

prs = ib.open_template()
T = list(prs.slides)                              # テンプレート36頁（0始まりで T[i-1]）


def dup(n, section=None, title=None):
    """テンプレートn頁（1始まり）を末尾に複製"""
    s = ib.duplicate_slide(prs, prs.slides.index(T[n - 1]), after=len(prs.slides) - 1)
    if title:
        ib.set_title(s, title)
    if section:
        ib.set_section(s, section)
    return s


def drop_pics(s):
    for sh in list(s.shapes):
        if sh.shape_type == 13:
            sh._element.getparent().remove(sh._element)


def fit(ref, target, hmax=1.0):
    return min(hmax, target / range_height(wb, ref))


def paste(s, ref, x, y, w, fs, target=None, hmax=1.0):
    hs = fit(ref, target, hmax) if target else hmax
    paste_range(s, wb, ref, x, y, w, hscale=hs, fscale=fs)
    return y + range_height(wb, ref, hscale=hs)


SRC_HEAD = shape_by_id(T[5], 8)                   # P/LページのComment枠（複製元）
SRC_BODY = shape_by_id(T[5], 9)
HEAD_EL = copy.deepcopy(SRC_HEAD._element); BODY_EL = copy.deepcopy(SRC_BODY._element)


def cbox(s, x, y, w, h, lines, head='Comment', body=8.5, hd=9.5, after=2, body_dy=0.66):
    """テンプレートのComment枠を複製して置く"""
    s.shapes._spTree.append(copy.deepcopy(HEAD_EL)); bx = s.shapes[-1]
    s.shapes._spTree.append(copy.deepcopy(BODY_EL)); bd = s.shapes[-1]
    ids = [x_.shape_id for x_ in s.shapes]
    for i, el in enumerate((bx._element, bd._element)):
        el.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr').set('id', str(max(ids) + 1 + i))
    place(bx, x, y, w, h); place(bd, x, y + body_dy, w, h - body_dy - 0.05)
    set_text(bx, head)
    fill_comment(bd, lines, body=body, head=hd, after=after)
    return bx, bd


def fc(sh, lines, body=9, hd=10, after=3):
    return fill_comment(sh, lines, body=body, head=hd, after=after)


def skip_labels(ch, n=3):
    from lxml import etree
    from pptx.oxml.ns import qn
    cat = ch._chartSpace.find('.//' + qn('c:catAx'))
    for tag in ('c:tickLblSkip', 'c:tickMarkSkip'):
        e = etree.SubElement(cat, qn(tag)); e.set('val', str(n))
    # スキーマ順（noMultiLvlLblの前）に並べ替え
    nm = cat.find(qn('c:noMultiLvlLbl'))
    if nm is not None:
        cat.remove(nm); cat.append(nm)


def chart_fmt(ch, size=8):
    ch.font.size = Pt(size); ch.font.name = ib.JP
    ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(size)


# ============================================================== 表紙・目次・中扉
s = T[0]
ib.set_cover_title(s, '化学・バイオ人材派遣', '首都圏')
s = T[2]
for sh in s.shapes:
    if sh.has_text_frame and '組織、拠点、設備等' in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.text = r.text.replace('組織、拠点、設備等', '組織・人員')
ib.set_title(T[15], '組織・人員')
for n in (17, 19):
    ib.set_section(T[n - 1], '組織・人員')

# ============================================================== P5 基本情報
s = T[4]
drop_pics(s)
remove(s, 3, 4, 6, 7, 8)
yb = paste(s, '会社概要!B2:C16', 0.34, 2.06, 6.41, 0.83, target=5.1)
set_text(shape_by_id(s, 2), '対象事業の概要')
set_lead(s, f'東証グロース上場のジャパニアスが展開する化学・バイオ領域の技術者派遣事業を事業譲渡で承継。理系技術者33名と営業1名、'
            f'首都圏の顧客{v("顧客分析", "E19")}社との取引基盤を有し、売上高は3期で約{ph("C21"):.1f}倍に拡大。')
cbox(s, 6.95, 1.62, 4.40, yb - 1.62 + 0.05, [
    '≪理系技術者33名を一括で承継可能≫',
    f'・技術者33名中32名が正社員（無期雇用）。平均年齢{ec("平均年齢（技術者）"):.1f}歳・20代が{pc(ec("20代比率"), 0)}で、33名中{ec("理系の学歴（専攻系統が理系以外以外）")}名が理系の学歴。',
    f'≪首都圏の顧客{v("顧客分析", "E19")}社と派遣取引≫',
    '・計測機器・化学・食品メーカー、試験・分析機関、大手メーカーグループ会社等。2026/11期の売上は全額が前期からの継続顧客。',
    '≪立上げ3期目で黒字化の見込み≫',
    f'・売上高{mm(ph("C4"))}→{mm(ph("D4"))}→{mm(ph("E4"))}百万円（見込）。採用費が先行した立上げ期を経て、'
    f'2026/11期は事業所利益{mm(ph("E8"))}百万円、共通費配賦後の営業利益{mm(ph("E11"))}百万円の見込み。',
    '≪選択と集中による譲渡≫',
    '・事業は成長しているが、売主の主力事業とのシナジーが限定的なため譲渡を検討（売主談）。',
    '≪確認中の論点≫',
    '・本社共通費の内訳と譲渡後の管理コスト、従業員の転籍同意・顧客の承諾、譲渡後の運営体制（事業部長等の関与）。',
], head='本件のポイント', body=8.5, hd=9.5, after=2)
note(s, 0.34, yb + 0.1, 6.41, ['※売上高等は売主の事業部PL（管理会計）。2026年11月期は10・11月の見込を含む。従業員数は2026年10月末。'], size=7.5)

# ============================================================== P6 インベストメント・ハイライト（テンプレ12）
s = T[11]
ib.clear_body(s)
ib.set_title(s, 'インベストメント・ハイライト')
ib.set_section(s, 'エグゼクティブサマリー')
pe = lambda lab: ec(lab)
rs = lambda a: v('人員構成', a)
r1 = META['r1']
items = [
    ('人材基盤', '採用難度の高い化学・バイオ系の理系技術者33名を一括で承継',
     f'技術者33名のうち32名が正社員（無期雇用）、33名中{pe("理系の学歴（専攻系統が理系以外以外）")}名が理系の学歴（大学院卒{rs("D16")}名を含む）。'
     f'平均年齢{pe("平均年齢（技術者）"):.1f}歳と若く、研究開発{rs(f"D{r1 + 3}")}名・分析補助{rs(f"D{r1 + 4}")}名・検査評価{rs(f"D{r1 + 5}")}名・'
     f'検査{rs(f"D{r1 + 6}")}名が首都圏の顧客先で就業。作業主任者等の資格保有者{rs(f"D{r1 + 15}")}名。'),
    ('顧客基盤', f'計測機器・化学・食品・試験分析等の顧客{v("顧客分析", "E19")}社と派遣取引',
     f'2026/11期（12〜10月）の顧客は{v("顧客分析", "E19")}社。最大顧客のガス検知・計測機器メーカーには9名、大手メーカーグループ会社・メーカーに各5名が就業。'
     f'2026/11期の売上は全額が前期からの継続顧客で、既存顧客への増員により拡大している。'),
    ('成長・採算', f'売上高は3期で約{ph("C21"):.1f}倍、2026/11期に黒字化の見込み',
     f'売上高{mm(ph("C4"))}→{mm(ph("D4"))}→{mm(ph("E4"))}百万円（見込）、売上総利益{mm(ph("E6"))}百万円（見込、粗利率{pc(ph("E7"))}）。'
     f'稼働率は各期94〜96%台。立上げ期の採用関連費（2025/11期{mm(ph("C29"))}百万円）が一巡し、2026/11期は事業所利益{mm(ph("E8"))}百万円、'
     f'共通費配賦後の営業利益{mm(ph("E11"))}百万円の見込み。'),
    ('統合余地', '技術者派遣の既存基盤との統合による拡大余地（可能性）',
     '採用・営業・管理機能を有する買手であれば、営業担当1名の体制を補強し、自社の採用力・顧客網を活かした増員や新規顧客の開拓が期待できる。'
     '効果は買手の体制により異なるため、本資料では定量化していない。'),
]
ib.feature_bands(s, items, x=0.49, y=0.98, w=10.71, h=1.38, gap=0.14)
note(s, 0.49, 7.08, 10.71, ['※数値は事業部PL（管理会計）・従業員一覧（2026年10月末）・顧客別売上による。「統合余地」は当社の見立てであり、実績ではない。'], size=7.5)

# ============================================================== P7 売主会社の概要（テンプレ34を複製）
s7 = dup(34, 'エグゼクティブサマリー', '売主会社の概要｜ジャパニアス株式会社')
set_lead(s7, '売主は1999年設立、東証グロース上場のエンジニアリング人材企業（2025/11期の全社売上高約120.8億円）。'
             '対象事業はその一事業部門で、開示セグメントは単一のため対象事業単独の数値は公表されていない。')
yb = paste(s7, '売主概要!B4:C15', 0.34, 1.6, 6.55, 0.83, target=5.4)
note(s7, 0.34, yb + 0.08, 6.55, ['※出所：会社案内（2025年4月）、公表決算。全社の情報であり、対象事業（化学・バイオ事業）の情報ではない。'], size=7.5)
labels = [str(v('売主概要', f'B{21 + i}')).replace('/11期', '') for i in range(11)]
vals = [v('売主概要', f'C{21 + i}') for i in range(11)]
ib.bar_chart(s7, 7.15, 1.85, 4.2, 2.9, [x[2:] for x in labels], vals, title='売主全社の売上高推移（百万円、11月期）', highlight=10, size=7, value_fmt='{:,}')
cbox(s7, 7.15, 4.95, 4.2, 2.2, [
    '≪売主全体と対象事業の区別≫',
    '・左表・グラフは売主全社の情報。社員数・拠点・主要取引先は全社のもので、対象事業の実績ではない。',
    '・対象事業の数値は売主の事業部PL（管理会計）による（財務ハイライト参照）。',
    '≪売主の拠点と対象事業≫',
    '・対象事業の技術者は全員が大宮事業所（2019年開設）に所属。',
], body=8.5, hd=9.5)

# ============================================================== P8 財務ハイライト（P/L）
s = T[5]
drop_pics(s)
yb = paste(s, 'PLハイライト!B2:G17', 0.34, 1.55, 6.62, 0.875, target=4.9)
note(s, 0.34, yb + 0.08, 6.62, [
    '※売主の事業部PL（管理会計）。2026/11期見込は2025/12〜2026/9の実績に10・11月の見込を加えたもの（年換算ではない）。',
    f'※在籍技術者数の2026/11期欄は2026年9月末（2026年10月末は従業員一覧で33名）。延べ稼働人月は人員欄が9月までのため見込欄は「－」。',
    '※減価償却費の計上はなく、EBITDAは営業利益と同額。正常化調整は行っていない。'], size=7.5)
set_lead(s, f'売上高は2024/11期{mm(ph("C4"))}百万円から2026/11期見込{mm(ph("E4"))}百万円へ拡大。採用費が先行した立上げ期を経て、'
            f'2026/11期は事業所利益{mm(ph("E8"))}百万円、本社共通費配賦後の営業利益{mm(ph("E11"))}百万円と黒字化の見込み。')
place(shape_by_id(s, 8), 7.06, 1.51, 4.29, 3.75); place(shape_by_id(s, 9), 7.06, 2.1, 4.28, 3.1)
fc(shape_by_id(s, 9), [
    '≪売上高≫',
    f'・2023年12月に技術者3名で計上開始。在籍は{ph("C14")}→{ph("D14")}→{ph("E14")}名（2026/9末）に増え、売上高は前期比{pc(ph("D5"))}・{pc(ph("E5"))}（見込）。'
    f'2025/12〜2026/9は前年同期比{pc(ph("F5"))}。',
    '≪売上総利益≫',
    f'・原価は技術者の人件費と通勤交通費。粗利率は{pc(ph("C7"))}→{pc(ph("D7"))}→{pc(ph("E7"))}（見込）。'
    f'稼働1人月当たり売上総利益は{k(ph("C17"))}→{k(ph("D17"))}→{k(ph("F17"))}千円（12〜9月）。',
    '≪事業所利益≫',
    f'・採用紹介料・広告宣伝費が{mm(ph("C28"))}→{mm(ph("C29"))}→{mm(ph("C30"))}百万円（見込）と一巡し、事業所利益は2026/11期に黒字化。',
    '≪見込値≫',
    f'・2026/11期は10・11月の見込{mm(ph("C32"))}百万円を含む。顧客別売上の10月実績は事業部PLを{mm(-ph("C35"))}百万円下回る（実績化後に更新）。',
], body=8, hd=9, after=1.5)
place(shape_by_id(s, 2), 7.06, 5.36, 4.29, 1.95); place(shape_by_id(s, 10), 7.07, 5.42, 4.28, 1.85)
fc(shape_by_id(s, 10), [
    '【共通費配賦と収益性の見方】',
    f'≪共通販管費≫・売主本社の共通費を売上高比（2026/11期{pc(ph("C25"))}）で配賦。営業利益は配賦後の数値。',
    f'≪譲渡後≫・売主の配賦はなくなる一方、給与計算・労務・経理等の管理コストは買手の体制で発生。PL外に記載の所長人件費（年{mm(ph("C26"))}百万円）の計上範囲は要確認。',
    f'✓ 独立運営時の収益性は、事業所利益{mm(ph("E8"))}百万円から買手側の管理コスト等を差し引いて検討いただく必要がある',
], body=7.8, hd=8.5, after=1)

# ============================================================== P9 収益構造（テンプレ7）
s = T[6]
drop_pics(s)
ib.set_title(s, '収益構造｜費用区分と共通費配賦')
yb = paste(s, '収益構造!B3:H16', 0.34, 1.5, 11.0, 0.8, target=4.25)
set_lead(s, f'原価の大半は技術者人件費（売上高比{pc(v("収益構造", "G6"))}）。本社共通費の配賦（同{pc(v("収益構造", "G13"))}）は譲渡後に売主からは発生しない一方、'
            '管理機能のコストは買手の体制で発生するため、事業所利益を起点に検討いただきたい。')
yc = yb + 0.15
place(shape_by_id(s, 2), 0.34, yc, 5.42, 7.3 - yc); place(shape_by_id(s, 7), 0.34, yc + 0.55, 5.42, 7.3 - yc - 0.6)
fc(shape_by_id(s, 7), [
    '≪費用構造≫',
    f'・売上原価の約95%が技術者の労務費。稼働1人月当たり売上高{k(ph("F16"))}千円に対し売上原価{k(v("人員推移", "E15"))}千円（2025/12〜2026/9）。',
    '≪譲渡後の考え方≫',
    '・売上と技術者人件費は転籍・契約承継により継続。共通費は買手の管理コストに置き換わる。',
], body=8, hd=9, after=1.5)
place(shape_by_id(s, 6), 5.92, yc, 5.42, 7.3 - yc); place(shape_by_id(s, 8), 5.92, yc + 0.07, 5.42, 7.3 - yc - 0.1)
fc(shape_by_id(s, 8), [
    '【要確認事項】',
    '≪共通販管費の内訳≫・配賦対象の本社機能と配賦基準（売上高比）の妥当性。',
    f'≪所長人件費≫・PL外記載の年{mm(ph("C26"))}百万円と、販管費の本社給与手当等（年{mm(ph("C27"))}百万円）との関係（重複の有無）。',
    '≪譲渡後に必要な機能≫・派遣元責任者、給与計算・社会保険手続、勤怠管理、請求・回収、採用。買手の既存組織で吸収できる範囲により追加コストは異なる。',
], body=7.8, hd=8.5, after=1)
set_text(shape_by_id(s, 2), 'Comment')

# ============================================================== P10 譲渡希望条件
s = T[7]
set_text(shape_by_id(s, 19), '事業譲渡（化学・バイオ事業）')
set_text(shape_by_id(s, 56), '2,000万円')
set_text(shape_by_id(s, 256), ['従業員34名の転籍による承継（従業員の個別同意が前提）', '顧客との派遣契約の承継（顧客の同意が前提）', '運転資金（売掛金等）は譲渡対象外の想定'], size=13)
place(shape_by_id(s, 256), 3.45, 4.68, 7.74, 1.1)
set_text(shape_by_id(s, 264), ['経営資源の選択と集中', '（事業は成長しており、不採算撤退ではない旨を売主より聴取）'], size=13)
place(shape_by_id(s, 264), 3.45, 6.4, 7.74, 0.8)
note(s, 0.49, 7.55, 10.71, ['※譲渡対象の範囲（人員・契約・資産・権利義務）は今後確定（Appendix「譲渡対象の範囲」）。譲渡対価に係る消費税の取扱いは要確認。'], size=7.5)

# ============================================================== P11 参考｜譲渡希望対価の位置づけ
s11 = dup(34, 'エグゼクティブサマリー', '参考｜譲渡希望対価の位置づけ')
tv = lambda a: v('対価対比', a)
set_lead(s11, f'希望譲渡対価2,000万円は承継予定人員1名当たり約{tv("C7") / 1e4:,.0f}万円。対象事業の入社1名当たり採用関連費（約{tv("C10") / 1e4:,.0f}万円）を下回り、'
              f'2026/11期見込の売上総利益の{tv("C12"):.2f}倍に相当（当社算定の参考指標）。')
yb = paste(s11, '対価対比!B4:D14', 0.34, 1.6, 9.8, 0.9, target=3.6)
note(s11, 0.34, yb + 0.08, 9.8, ['※当社算定の参考指標であり、価格の妥当性を示すものではない。'], size=7.5)
cbox(s11, 0.34, yb + 0.45, 11.0, min(2.3, 7.25 - yb - 0.45), [
    '≪指標の見方≫',
    '・採用関連費は採用活動に要した費用であり、技術者が既に顧客先で就業している状態（稼働中の派遣契約・就業実績）の価値は含まない。',
    '・転籍は従業員の個別同意が前提であり、承継人数は変動しうる。',
    '≪対価以外に見込むべきもの≫',
    f'・譲受会社側の管理コスト（給与計算・労務・経理・派遣元責任者等）と運転資金（労務費の約1か月分、2026/11期見込の月平均{mm(ph("C36"))}百万円が目安）。',
], body=8.5, hd=9.5)

# ============================================================== P13 事業内容（テンプレ10）
s = T[9]
ib.set_title(s, '事業内容（ビジネスモデル）')
set_lead(s, '理系人材を正社員（無期雇用）で採用し、首都圏の化学・計測機器・食品・試験検査等の顧客へ派遣する労働者派遣事業。'
            '売上高の全額が派遣料（時間給制）で、技術者の人件費を差し引いた額が売上総利益の源泉。')
set_text(shape_by_id(s, 5), '人材（採用）')
set_text(shape_by_id(s, 17), '顧客（派遣先）')
set_text(shape_by_id(s, 21), '対象事業（化学・バイオ事業）')
set_text(shape_by_id(s, 12), '売主本社（共通機能）※譲渡対象外')
set_text(shape_by_id(s, 8), ['理系の大学・大学院・専門学校等の卒業者', 'スカウトサービスで個別にスカウト', '人材紹介会社・求人広告', '営業担当が面接に入り選考'], size=10)
place(shape_by_id(s, 8), 0.55, 2.4, 2.25, 4.2)
set_text(shape_by_id(s, 13), [f'{v("顧客分析", "E19")}社（2026/11期）', '計測・光学機器メーカー', '化学・素材メーカー', '試験・分析受託機関', '食品メーカー', 'バイオ関連企業', '総合建設会社 ほか'], size=10)
place(shape_by_id(s, 13), 8.9, 2.4, 2.25, 4.2)
set_text(shape_by_id(s, 39), '営業（所長）1名')
set_text(shape_by_id(s, 51), '顧客開拓・契約\n顧客窓口')
set_text(shape_by_id(s, 52), '技術者の面接・選考\n就業フォロー（要確認）')
set_text(shape_by_id(s, 46), f'技術者33名（正社員32名・有期1名）')
set_text(shape_by_id(s, 49), f'研究開発{rs(f"D{r1 + 3}")}名・分析補助{rs(f"D{r1 + 4}")}名')
set_text(shape_by_id(s, 50), f'検査評価{rs(f"D{r1 + 5}")}名・検査{rs(f"D{r1 + 6}")}名')
set_text(shape_by_id(s, 47), '取引条件')
set_text(shape_by_id(s, 55), '労働者派遣契約\n時間給制')
set_text(shape_by_id(s, 56), '月末締め\n翌月末払い')
set_text(shape_by_id(s, 25), ['給与計算・労務・経理等', '（想定）'], size=9)
set_text(shape_by_id(s, 26), ['共通費として', '売上高比で配賦'], size=9)
for sid in (2,):
    remove(s, sid)
ib.textbox(s, 2.86, 3.55, 0.6, 0.5, '雇用\n(正社員)', size=7.5, color=ib.NAVY, align=PP_ALIGN.CENTER)
ib.textbox(s, 8.26, 3.55, 0.6, 0.5, '派遣\n(時間給)', size=7.5, color=ib.NAVY, align=PP_ALIGN.CENTER)

# ============================================================== P14 業務領域（テンプレ34複製）
s14 = dup(34, '事業概要', '業務領域｜売主の対応分野と対象事業の実績')
set_lead(s14, f'対象事業の技術者は研究開発{rs(f"D{r1 + 3}")}名・分析補助{rs(f"D{r1 + 4}")}名・検査評価{rs(f"D{r1 + 5}")}名・検査{rs(f"D{r1 + 6}")}名。'
              '売主が示す4領域のうち開発・分析は実績を確認でき、動物実験・細胞実験は対象事業での実績を確認中。')
yb = paste(s14, '業務領域!B4:E8', 0.34, 1.55, 11.0, 0.85, target=2.55)
note(s14, 0.34, yb + 0.06, 11.0, ['※売主の対応分野は公式サイト・会社案内による売主全体の記載であり、対象事業の実績を示すものではない。'], size=7.5)
ib.section_bar(s14, 0.34, yb + 0.38, 11.0, '参考｜化学・バイオ領域の営業資料に掲載の業務例', h=0.3, size=10.5)
yb2 = paste(s14, '業務領域!B13:E17', 0.34, yb + 0.72, 11.0, 0.85, target=7.0 - (yb + 0.72))
note(s14, 0.34, yb2 + 0.06, 11.0, ['※営業資料に「あくまで一例」として掲載された業務例。対象事業の現在の顧客・技術者の業務との対応は要確認。'], size=7.5)

# ============================================================== P15 取引条件と人材モデル
s15 = dup(34, '事業概要', '取引条件と人材モデル')
set_lead(s15, '取引は労働者派遣契約・時間給制（月末締め翌月末払い）が標準。技術者は正社員（無期雇用）が基本で派遣期間の制限がなく、'
              '顧客先での長期就業が可能。営業担当が採用面接に入り、人材を見極めている。')
yb = paste(s15, '取引条件!B4:D13', 0.34, 1.55, 11.0, 0.85, target=4.3)
note(s15, 0.34, yb + 0.06, 11.0, ['※会社案内（化学・バイオ領域の営業資料、2025年4月）の記載。個別の契約内容（単価・期間・更新）は契約書で要確認。'], size=7.5)
cbox(s15, 0.34, yb + 0.4, 11.0, min(2.0, 7.25 - yb - 0.4), [
    '≪無期雇用派遣のメリット≫',
    '・派遣可能期間の制限を受けないため、顧客は同じ技術者に長期で業務を任せられる。会社案内でも「長期的に労働力を供給でき、離職リスクも抑止できる」点を訴求。',
    '≪譲受会社での扱い≫',
    '・譲受会社が無期雇用で受け入れれば、顧客先での就業を継続しやすい（雇用形態・労働条件の設計は譲受会社による）。',
], body=8.5, hd=9.5)

# ============================================================== P16 主要顧客（テンプレ13）
s = T[12]
drop_pics(s)
ib.set_title(s, '主要顧客の3期推移')
yb = paste(s, '取引先!B2:M21', 0.34, 1.5, 11.0, 0.78, target=4.95)
note(s, 0.34, yb + 0.08, 11.0, [
    '※顧客別売上高（売主資料）。2026/11期は2025/12〜2026/10の実績（11か月）。顧客名は匿名（記号は2026/11期の売上順）。業種は公開情報等に基づく当社の推察（要確認）。',
    f'※配属人数は2026年10月末の就業先別人数（従業員一覧）。このほか休職1名、B社のグループ会社への配属1名（有期契約社員）。勤務地は就業先の所在都府県。',
    f'※2024/11期の顧客別合計は事業部PLの売上高を{mm(-v("取引先", "F32"))}百万円、2026/11期（12〜10月）は同{mm(-v("取引先", "J32"))}百万円下回る（集計範囲・10月の見込値の差。要確認）。'], size=7.5)
cd_ = lambda a: v('顧客分析', a)
set_lead(s, f'2026/11期（12〜10月）は{cd_("E19")}社と取引し、上位3社で{pc(cd_("E21"))}。最大顧客のガス検知・計測機器メーカー（A社）は9名が就業。'
            '上位5社のうち3社は事業開始初年度からの取引で、取引を深耕している。')

# ============================================================== P17 顧客基盤の分析（テンプレ14）
s = T[13]
drop_pics(s)
ib.set_title(s, '顧客基盤の分析（業種構成・集中度・取引継続）')
yb = paste(s, '顧客分析!B3:J14', 0.34, 1.55, 6.7, 0.8, target=3.3)
note(s, 0.34, yb + 0.05, 6.7, ['※業種区分は当社の推察（要確認）。メーカー系グループ会社はB社。'], size=7.5)
yb2 = paste(s, '顧客分析!B18:E26', 7.25, 1.55, 4.1, 0.8, target=2.5)
note(s, 7.25, yb2 + 0.05, 4.1, ['※構成比の分母は顧客別売上の合計。'], size=7.5)
yb3 = paste(s, '顧客分析!L30:O35', 7.25, yb2 + 0.4, 4.1, 0.8, target=1.6)
set_lead(s, f'業種は精密・計測機器、化学・素材、試験・分析受託、食品等に分散。2026/11期の売上は全額が前期からの継続顧客で、'
            '新規開拓よりも既存顧客への増員で成長している。')
cbox(s, 0.34, yb + 0.35, 6.7, 7.3 - yb - 0.35, [
    '≪顧客基盤の特徴≫',
    f'・2025/11期に新規{cd_("D24")}社を獲得し、2026/11期は{cd_("E19")}社すべてが前期からの継続顧客。A社（9名）・C社（5名）等で配属人数が増加。',
    '≪取引終了先≫',
    f'・2024/11期の最大顧客（構成比{pc(cd_("C42"))}）を含む{META["ENDED"]}社とは2026/11期の取引がない。終了理由は要確認。',
    f'・2026年10月末に配属のない顧客が3社（M・N・O社、2026/11期の売上は計{mm(v("取引先", "J17") + v("取引先", "J18") + v("取引先", "J19"))}百万円）。',
], body=8, hd=9, after=1.5)

# ============================================================== P18 月次業績の推移（テンプレ15）
s = T[14]
drop_pics(s)
ib.set_title(s, '月次業績の推移')
mv = lambda a: v('月次推移', a)
cats, sal, gpv = [], [], []
for i in range(36):
    r = 5 + i
    lab = str(mv(f'B{r}'))[2:]
    cats.append(lab + ('*' if mv(f'C{r}') == '見込' else ''))
    sal.append(round(mv(f'D{r}') / 1000)); gpv.append(round(mv(f'E{r}') / 1000))
cd = CategoryChartData(); cd.categories = cats
cd.add_series('売上高', sal); cd.add_series('売上総利益', gpv)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.34), Inches(1.55), Inches(7.2), Inches(3.9), cd)
ch = gf.chart; chart_fmt(ch)
ch.has_title = True; ch.chart_title.text_frame.text = '月次の売上高・売上総利益（千円）　*は見込'
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9); ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
pl_ = ch.plots[0]; pl_.gap_width = 40; pl_.overlap = 0
for i, col in enumerate(['0B3041', '0889C9']):
    ser = pl_.series[i]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = RGBColor.from_string(col)
    for j in (34, 35):
        pt = ser.points[j]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string('9DB3C2' if i == 0 else 'A9D6EE')
ca = ch.category_axis; ca.tick_labels.font.size = Pt(7); ca.has_major_gridlines = False; skip_labels(ch, 3)
va = ch.value_axis; va.tick_labels.font.size = Pt(7); va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
yq = paste(s, '月次推移!B43:G56', 7.7, 1.5, 3.65, 0.8, target=3.9)
note(s, 7.7, yq + 0.05, 3.65, ['※*は10・11月の見込を含む。前年同期比は倍率。'], size=7)
set_lead(s, f'月次売上高は2023年12月の{mm(mv("D5"))}百万円から2026年9月には{mm(mv("D38"))}百万円へ増加。四半期売上高は各期とも前年同期を上回り、'
            f'直近3か月（7〜9月）平均の年換算は約{ph("C24") / 1e6:,.0f}百万円（試算値）。')
cbox(s, 0.34, 5.6, 7.2, 1.7, [
    '≪推移の特徴≫',
    '・在籍技術者の増加に合わせて売上高・売上総利益が段階的に拡大。12〜2月の四半期は前四半期をやや下回る傾向（要因は要確認）。',
    f'・2026年10・11月の見込（計{mm(ph("C32"))}百万円）は実績化後に更新する。年換算値は試算であり、見込・計画ではない。',
], body=8, hd=9, after=1.5)

# ============================================================== P20 運営体制（テンプレ17 組織図）
s = T[16]
ib.set_title(s, '運営体制（組織図）')
box_dir = shape_by_id(s, 30); box_dep = shape_by_id(s, 31); box_mem = shape_by_id(s, 32); box_top = shape_by_id(s, 72)
box_hr = shape_by_id(s, 82)
keep_ids = {13, 5, 85}
src_el = {n: copy.deepcopy(x._element) for n, x in (('dir', box_dir), ('dep', box_dep), ('mem', box_mem), ('top', box_top), ('hr', box_hr))}
for sh in list(s.shapes):
    if sh.shape_id not in keep_ids:
        sh._element.getparent().remove(sh._element)
nid = [200]


def mk(kind, x, y, w, h, lines, size=None, size1=None):
    el = copy.deepcopy(src_el[kind]); s.shapes._spTree.append(el); sh = s.shapes[-1]
    nid[0] = max([x_.shape_id for x_ in s.shapes if x_._element is not el] + [nid[0]]) + 1
    el.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr').set('id', str(nid[0]))
    place(sh, x, y, w, h); set_box(sh, lines, size=size, size1=size1)
    return sh


def line(x, y, w, h):
    ib.rect(s, x, y, max(w, 0.012), max(h, 0.012), fill=RGBColor(0x80, 0x80, 0x80))


# 売主側（譲渡対象外の想定）
mk('top', 4.6, 1.6, 2.5, 0.55, ['売主：ジャパニアス株式会社', '（本社機能・共通費）'], size=8, size1=9)
mk('dir', 3.3, 2.45, 2.2, 0.6, ['事業部長', '運営に関与／譲渡後の関与は要確認'], size=7.5, size1=9)
mk('dir', 6.2, 2.45, 2.2, 0.6, ['専務取締役', '運営に関与／譲渡後の関与は要確認'], size=7.5, size1=9)
line(5.85, 2.15, 0.012, 0.3); line(4.4, 2.3, 2.9, 0.012); line(4.4, 2.3, 0.012, 0.15); line(7.3, 2.3, 0.012, 0.15)
# 譲渡対象（34名）
ib.rect(s, 0.34, 3.3, 11.0, 3.95, fill=None, line=ib.NAVY, lw=1.0)
ib.textbox(s, 0.45, 3.36, 4.0, 0.25, '譲渡対象（想定）：34名', size=10, bold=True, color=ib.NAVY)
mk('dep', 4.6, 3.45, 2.5, 0.62, ['大宮事業所 所長（営業）1名', '顧客窓口・契約・技術者の面接'], size=7.5, size1=9)
line(5.85, 3.05, 0.012, 0.4)
line(5.85, 4.07, 0.012, 0.25)
groups = [('A社', 9), ('B社', 5), ('C社', 5), ('D社', 3), ('E社', 2)]
xs = [0.5, 2.05, 3.6, 5.15, 6.7, 8.25, 9.8]
line(1.2, 4.32, 9.3, 0.012)
for i, (c_, n_) in enumerate(groups):
    x = xs[i]; line(x + 0.7, 4.32, 0.012, 0.2)
    mk('mem', x, 4.52, 1.38, 0.6, [f'{c_}（{n_}名）', '技術者'], size=7.5, size1=9)
x = xs[5]; line(x + 0.7, 4.32, 0.012, 0.2)
mk('mem', x, 4.52, 1.38, 0.6, ['F〜L社（各1名）', '技術者 計7名'], size=7.5, size1=9)
x = xs[6]; line(x + 0.7, 4.32, 0.012, 0.2)
mk('mem', x, 4.52, 1.38, 0.6, ['その他（2名）', 'グループ会社1・休職1'], size=7, size1=9)
mk('hr', 0.5, 5.35, 3.4, 1.75, ['【人員構成】（2026年10月末）', '営業（所長）：1名', '技術者：33名（正社員32・有期契約1）',
                                  '合計：34名', f'技術者の平均年齢：{ec("平均年齢（技術者）"):.1f}歳', '所属：全員が大宮事業所'], size=8.5, size1=9)
cbox(s, 4.15, 5.3, 7.05, 1.85, [
    '≪譲渡後の運営体制（要確認）≫',
    '・顧客窓口・契約・採用面接は営業担当（所長）が担う（会社案内では営業担当が面接に入ると記載）。事業部長・専務取締役の関与内容（顧客対応・採用・技術者フォロー）と、譲渡後の引継ぎ期間は要確認。',
    '・技術者は全員が大宮事業所所属で、就業先はA〜L社等の顧客先。休職1名（2026年10月末）。',
], body=8, hd=9, after=1.5)
set_lead(s, '譲渡対象は大宮事業所の営業担当（所長）1名と技術者33名。現在は売主の事業部長・専務取締役が運営に関与しており、'
            '譲渡後の運営体制と引継ぎの方法は重要な確認事項。')

# ============================================================== P21 人員構成（テンプレ34複製）
s21 = dup(34, '組織・人員', '人員構成（技術者33名）')
y1 = paste(s21, f'人員構成!B4:E{META["endL"]}', 0.34, 1.55, 3.75, 0.8, target=5.6)
y2 = paste(s21, f'人員構成!B{r1}:E{META["endR"]}', 4.25, 1.55, 3.75, 0.8, target=5.6)
note(s21, 0.34, max(y1, y2) + 0.05, 7.66, ['※技術者33名（営業1名を除く）。2026年10月末の従業員一覧による。専攻系統は学科名からの当社分類。勤務地は就業先の所在都府県。'], size=7)
set_lead(s21, f'技術者33名は平均年齢{ec("平均年齢（技術者）"):.1f}歳・20代が{pc(ec("20代比率"), 0)}と若く、正社員（無期雇用）が32名。'
              f'研究開発・分析・検査に従事し、化学系・生命科学系等の理系の学歴を持つ人材が{ec("理系の学歴（専攻系統が理系以外以外）")}名。')
cbox(s21, 8.2, 1.55, 3.15, 5.7, [
    '≪若い理系人材≫',
    f'・大学院卒{rs("D16")}名・大卒{rs("D17")}名。専攻は化学系{rs("D19")}名、生物・生命科学系{rs("D20")}名、医・薬・栄養系{rs("D21")}名。',
    '≪勤続≫',
    f'・事業開始が2023年11月頃のため平均勤続は{ec("平均勤続（技術者、年）"):.1f}年。2年以上は{rs("D15")}名。',
    '≪資格≫',
    f'・資格保有者{rs(f"D{r1 + 15}")}名（有機溶剤作業主任者{rs(f"D{r1 + 16}")}名、特定化学物質作業主任者{rs(f"D{r1 + 17}")}名、危険物乙4 {rs(f"D{r1 + 18}")}名ほか）。',
    '≪給与水準≫',
    f'・正社員技術者の基本給は月{ec("正社員（技術者）基本給 最低") / 1000:,.0f}〜{ec("正社員（技術者）基本給 最高") / 1000:,.0f}千円（平均{ec("正社員（技術者）平均基本給（円/月）") / 1000:,.1f}千円）。'
    f'2026年5月賞与は{ec("2026年5月賞与 支給者数")}名に支給（平均{ec("2026年5月賞与 支給者平均（円）") / 1000:,.0f}千円）。',
    '≪居住地≫',
    f'・神奈川県{ec("居住地：神奈川県")}名、東京都{ec("居住地：東京都")}名、埼玉県{ec("居住地：埼玉県")}名ほか。',
], body=8, hd=9, after=1.5)

# ============================================================== P22・P23 従業員一覧
s = T[18]
drop_pics(s)
ib.set_title(s, '従業員一覧①')
EL = META['EMP_LAST']
yb = paste(s, '従業員!B4:Q21', 0.34, 1.5, 11.0, 0.8, target=5.4)
note(s, 0.34, yb + 0.06, 11.0, ['※従業員一覧（2026年10月末在籍）。氏名・学校名・個人別の給与は記載していない。勤続は2026年10月末までの年数。就業先は顧客の記号（主要顧客の3期推移と同じ）。'], size=7.5)
set_lead(s, '№1は営業担当（大宮事業所所長）。技術者は就業先別に記載し、A社（9名）・B社（5名）・C社（5名）に配属が集中。'
            '全員が正社員で、20代の理系人材が中心。')
s23 = ib.duplicate_slide(prs, prs.slides.index(T[18]), after=len(prs.slides) - 1)
for sh in list(s23.shapes):
    if sh.has_table or (sh.has_text_frame and sh.text_frame.text.startswith('※')):
        sh._element.getparent().remove(sh._element)
ib.set_title(s23, '従業員一覧②')
yh = paste(s23, '従業員!B4:Q4', 0.34, 1.5, 11.0, 0.8)
yb = paste(s23, f'従業員!B22:Q{EL}', 0.34, yh, 11.0, 0.8, hmax=fit('従業員!B4:Q21', 5.4) )
note(s23, 0.34, yb + 0.06, 11.0, ['※有期契約社員1名は時給制。休職中1名は2026年10月末時点。B社グループ会社への配属分の売上計上先は要確認。'], size=7.5)
set_lead(s23, f'D社以降の就業先と、休職者・B社グループ会社への配属者を記載。有機溶剤・特定化学物質の作業主任者等、'
              f'資格保有者{rs(f"D{r1 + 15}")}名が在籍。')

# ============================================================== P24 人員の推移と1人当たり採算（テンプレ34複製）
s24 = dup(34, '組織・人員', '人員の推移と1人当たり採算')
pv = lambda a: v('人員推移', a)
yb = paste(s24, '人員推移!B3:E18', 0.34, 1.55, 5.75, 0.85, target=4.6)
note(s24, 0.34, yb + 0.05, 5.75, ['※事業部PLの人員欄（月末在籍・稼働等）。2026/11期は人員欄が9月までのため2025/12〜2026/9で表示。'], size=7)
cats, hc, op = [], [], []
for i in range(34):
    r = 5 + i
    cats.append(str(mv(f'B{r}'))[2:]); hc.append(mv(f'H{r}')); op.append(mv(f'I{r}'))
cd = CategoryChartData(); cd.categories = cats
cd.add_series('在籍', hc); cd.add_series('稼働', op)
gf = s24.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(6.3), Inches(1.55), Inches(5.05), Inches(3.3), cd)
ch = gf.chart; chart_fmt(ch)
ch.has_title = True; ch.chart_title.text_frame.text = '月末の在籍・稼働技術者数（名）'
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9); ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
for i, col in enumerate(['0B3041', '0889C9']):
    ser = ch.plots[0].series[i]; ser.format.line.color.rgb = RGBColor.from_string(col); ser.smooth = False
    ser.marker.size = 4; ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = RGBColor.from_string(col)
ch.category_axis.tick_labels.font.size = Pt(7); ch.value_axis.tick_labels.font.size = Pt(7); skip_labels(ch, 3)
ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
set_lead(s24, f'稼働率は各期{pc(min(pv("C8"), pv("D8"), pv("E8")))}〜{pc(max(pv("C8"), pv("D8"), pv("E8")))}で待機人員は少ない。退職率は2025/11期の{pc(pv("D13"))}から'
              f'2026/11期（12〜9月）は{pc(pv("E13"))}に低下し、入社1名当たり採用関連費も{k(pv("C18"))}千円→{k(pv("E18"))}千円に低減。')
cbox(s24, 6.3, 5.0, 5.05, 2.25, [
    '≪稼働≫',
    f'・未稼働（スタンバイ等）は延べ{pv("C9")}・{pv("D9")}・{pv("E9")}人月と少なく、採用後は短期間で顧客先へ配属。2026/11期は休職者が延べ{pv("E10")}人月。',
    '≪採用・定着≫',
    f'・入社は{pv("C11")}→{pv("D11")}→{pv("E11")}名（12〜9月）。2025/11期は退職{pv("D12")}名と多かったが、2026/11期は{pv("E12")}名に減少（退職理由は要確認）。',
    '≪1人当たり採算≫',
    f'・稼働1人月当たり売上総利益は{k(pv("C16"))}→{k(pv("D16"))}→{k(pv("E16"))}千円。',
], body=8, hd=9, after=1.5)

# ============================================================== P26 損益計算書（テンプレ26）
s = T[25]
ib.set_title(s, '損益計算書（事業部PL）')
yb = paste(s, '事業部PL!B3:J19', 0.34, 1.5, 6.6, 0.75, target=4.6)
note(s, 0.34, yb + 0.06, 6.6, ['※売主の事業部PL（管理会計）。2026/11期見込は10・11月の見込を含む。対象事業単独の貸借対照表はない。'], size=7)
dp = lambda a: v('事業部PL', a)
set_lead(s, f'売上高の全額が派遣売上高で、売上原価の約95%が技術者の労務費。売上原価率は{pc(dp("D14"))}→{pc(dp("F14"))}→{pc(dp("H14"))}（見込）と低下し、'
            '粗利率が改善している。')
place(shape_by_id(s, 2), 7.06, 1.51, 4.29, 5.75); place(shape_by_id(s, 6), 7.06, 2.05, 4.26, 5.1)
fc(shape_by_id(s, 6), [
    '≪売上高≫',
    '・全額が派遣売上高（請負売上高の計上なし）。2026/11期は10・11月の見込を含む。',
    '≪売上原価≫',
    '・技術者の給与手当・法定福利費・賞与引当金・退職給付費用（概算）と旅費交通費（通勤交通費等と推察）。賞与・退職給付は毎月の引当計上。',
    '≪販売費及び一般管理費≫',
    '・事業部に直課された費用（採用関連費・管理人件費・事務所家賃等）。内訳は次頁。',
    '≪共通販管費≫',
    f'・売主本社の共通費を売上高比で配賦（{pc(dp("D18"))}→{pc(dp("F18"))}→{pc(dp("H18"))}）。内訳は要確認。',
    '≪その他≫',
    '・減価償却費・営業外損益・特別損益・法人税等の計上はない（事業部PLのため）。',
], body=8, hd=9, after=1.5)

# ============================================================== P27 販管費（テンプレ28）
s = T[27]
ib.set_title(s, '販売費及び一般管理費（事業部直接費）')
ST = META['SGA_TOT_ROW']
yb = paste(s, f'事業部販管費!B3:J{ST}', 0.34, 1.5, 6.6, 0.75, target=5.3)
note(s, 0.34, yb + 0.06, 6.6, ['※3期とも計上のない科目は非表示。共通販管費（本社共通費の配賦）は含まない。'], size=7)
sg = lambda a: v('事業部販管費', a)
set_lead(s, f'事業部直接の販管費は採用関連費と管理人件費が中心。採用関連費は2025/11期{mm(ph("C29"))}百万円から2026/11期見込{mm(ph("C30"))}百万円に減少し、'
            f'販管費の売上高比は{pc(sg(f"D{ST}"))}→{pc(sg(f"H{ST}"))}に低下。')
place(shape_by_id(s, 2), 7.06, 1.51, 4.29, 5.75); place(shape_by_id(s, 3), 7.06, 2.05, 4.26, 5.1)
fc(shape_by_id(s, 3), [
    '≪本社給与手当・法定福利費・賞与引当金≫',
    f'・事業部に直課された管理人件費。PL外に記載の所長人件費（年{mm(ph("C26"))}百万円）との関係は要確認。',
    '≪採用紹介料・広告宣伝費≫',
    '・人材紹介会社への紹介料と、求人媒体等の費用と推察（広告宣伝費は月額定額で計上。内容は要確認）。立上げ期に先行し、2026/11期は減少。',
    '≪事務所家賃・その他≫',
    '・大宮事業所の家賃等の配分と推察。在籍人数の増加に応じて増加。',
], body=8, hd=9, after=1.5)

# ============================================================== Appendix
s = T[29]
yb = paste(s, '沿革!B4:D15', 0.49, 1.6, 10.71, 0.9, target=4.5)
note(s, 0.49, yb + 0.08, 10.71, ['※出所：会社案内（経営状況）、事業部PL、従業員一覧、顧客別売上、公表決算。顧客数は顧客別売上資料の取引社数。'], size=7.5)
set_lead(s, '売主は1999年設立、2019年に大宮営業所を開設し2022年に上場。対象事業は2023年11月頃に立ち上がり、約3年で技術者33名・'
            f'顧客{v("顧客分析", "E19")}社の規模に成長。')
s = T[31]
ib.set_title(s, '譲渡対象の範囲')
yb = paste(s, '譲渡対象!B4:E15', 0.34, 1.55, 11.0, 0.85, target=5.4)
note(s, 0.34, yb + 0.08, 11.0, ['※現時点の当社の想定。譲渡対象の資産・負債・契約・権利義務の正確な範囲は、売主からの回収資料・協議により確定する。'], size=7.5)
set_lead(s, '承継対象は人員（34名）と顧客との派遣契約を想定。派遣事業の許可は承継されず、運転資金は対象外の想定。'
            '拠点・採用基盤・データ等の範囲は今後確認する。')
s = T[32]
ib.set_title(s, '事業譲渡に伴う確認事項')
yb = paste(s, '確認事項!B4:D13', 0.34, 1.55, 11.0, 0.85, target=5.0)
note(s, 0.34, yb + 0.08, 11.0, ['※一般的な実務上の論点の整理。法令上の取扱いは譲受会社の顧問社会保険労務士・弁護士等の確認を前提とする。'], size=7.5)
set_lead(s, '事業譲渡では従業員の転籍同意と顧客の承諾が前提。無期雇用の技術者は派遣期間制限の対象外であり、'
            '譲受会社でも無期雇用で受け入れれば顧客先での就業を継続しやすい。')
s = T[33]
ib.set_title(s, '要確認事項・ネクストステップ')
yb = paste(s, '要確認事項!B4:D16', 0.34, 1.55, 11.0, 0.85, target=5.2)
note(s, 0.34, yb + 0.08, 11.0, ['※回答を受領次第、本資料を更新する。'], size=7.5)
set_lead(s, '収益性（共通費の内訳・所長人件費）、人員（転籍条件・運営体制）、顧客（契約条件・承諾）を優先して確認する。')

# ============================================================== 頁構成
order = [T[0], T[1], T[2], T[3], T[4], T[11], s7, T[5], T[6], T[7], s11,
         T[8], T[9], s14, s15, T[12], T[13], T[14],
         T[15], T[16], s21, T[18], s23, s24,
         T[23], T[25], T[27],
         T[28], T[29], T[31], T[32], T[33], T[34], T[35]]
lst = prs.slides._sldIdLst
ids = {id(sl): el for sl, el in zip(prs.slides, list(lst))}
keep_el = [ids[id(sl)] for sl in order]
for el in list(lst):
    if el not in keep_el:
        rid = el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        prs.part.drop_rel(rid); lst.remove(el)
for el in keep_el:
    lst.remove(el)
for el in keep_el:
    lst.append(el)

ib.strip_red_banners(prs)
ib.fix_broken_pictures(prs)
remove_yellow(prs); remove_null_pics(prs); remove_watermark(prs)
ib.force_white_headers(prs)
for sh in prs.slides[1].shapes:
    if sh.has_text_frame and 'ご照会先' in sh.text_frame.text:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
prs.save(OUT)
print('saved', OUT, len(prs.slides))
