"""株式会社STOP 初回面談資料（ご面談資料）のビルドスクリプト。

    python build_meeting_deck.py <LogiTechnoService最終版.pptx> <IT事例の入った参考資料.pptx> <出力.pptx>

LogiTechnoService案件の最終版（依頼者修正済み・31頁）を複製し、
  - 表紙・会社概要・事業概要・主な歩み・強みと確認事項を本件の内容に差し替え
  - 会社概要（13頁）の直後に「財務内容」（PL・BS 3期）を1頁追加
  - 成約事例2頁は、物流の事例画像を IT業界の事例画像（参考資料の19・20頁）に差し替え
財務数値は東京商工リサーチ財務情報（2026/11/21取得）を FIN に一度だけ入力し、比率・合計はここで計算する
（パワポに数字を手で打たない）。
"""
import copy
import io
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                '..', '.claude', 'skills', 'ma-deal-docs', 'scripts'))
from im_builder import duplicate_slide, house_table  # noqa: E402
from im_paste import scrub_doc_props  # noqa: E402

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
NAVY = RGBColor(0x0B, 0x30, 0x41)
GRAY = RGBColor(0x59, 0x59, 0x59)
BODY = RGBColor(0x1A, 0x1A, 0x1A)

BASE, CASES, OUT = sys.argv[1:4]
prs = Presentation(BASE)

# ---------------------------------------------------------------- 財務データ（TSR・千円）
PERIODS = ['2024年1月期', '2024年9月期', '2025年9月期']
MONTHS = [12, 8, 12]                     # 2024年9月期は決算期変更による8か月決算
FIN = {
    '売上高': [23_976, 45_619, 204_877],
    '売上原価': [17_025, 5_087, 25_596],
    '販管費': [6_462, 38_955, 172_692],
    '営業利益': [489, 1_576, 6_588],
    '経常利益': [489, 1_704, 6_256],
    '当期純利益': [-675, 1_471, 4_628],
    '従業員数': [1, 15, 31],
    # BS
    '現金及び預金': [4_401, 770, 125],
    '売掛金': [22_596, 21_333, 73_194],
    '仕掛品': [1_400, 76_671, 79_061],
    '流動資産合計': [40_706, 129_907, 200_135],
    '固定資産合計': [35_146, 31_100, 77_529],
    '無形固定資産': [4_113, 3_038, 50_306],
    '前払費用': [317, 16_855, 37_011],
    '長期前払費用': [31_033, 26_895, 25_799],
    '資産合計': [75_853, 161_007, 277_665],
    '買掛金': [9_066, 1_540, 11_229],
    '未払金': [4_737, 19_318, 76_086],
    '短期借入金': [0, 5_250, 5_000],
    '1年内長期借入金': [0, 0, 2_004],
    '長期借入金': [0, 20_000, 27_662],
    '負債合計': [14_410, 46_593, 133_873],
    '純資産合計': [61_442, 114_413, 143_791],
    '払込資本': [42_000 + 10_000, 72_750 + 30_750, 87_750 + 40_500],  # 資本金＋資本準備金（＋申込証拠金）
}
F = FIN
gp = [6_951, 40_531, 179_281]            # 売上総利益（TSR記載値。千円未満の端数で差額±1あり）
assert all(abs(s - c - g) <= 1 for s, c, g in zip(F['売上高'], F['売上原価'], gp))
assert all(abs(g - x - o) <= 1 for g, x, o in zip(gp, F['販管費'], F['営業利益']))
debt = [a + b + c for a, b, c in zip(F['短期借入金'], F['1年内長期借入金'], F['長期借入金'])]
payable = [a + b for a, b in zip(F['買掛金'], F['未払金'])]
other_ca = [t - a - b - c for t, a, b, c in zip(F['流動資産合計'], F['現金及び預金'], F['売掛金'], F['仕掛品'])]
other_li = [t - a - b for t, a, b in zip(F['負債合計'], payable, debt)]
eq_ratio = [n / t for n, t in zip(F['純資産合計'], F['資産合計'])]
annual = [s * 12 / m for s, m in zip(F['売上高'], MONTHS)]
growth = annual[2] / annual[1]
soft = [a + b + c + d for a, b, c, d in zip(F['仕掛品'], F['前払費用'], F['長期前払費用'], F['無形固定資産'])]
soft_ratio = soft[2] / F['資産合計'][2]


def k(v):
    return f'△{-v:,.0f}' if v < 0 else f'{v:,.0f}'


def pc(v):
    return f'{v * 100:.1f}%'


# ---------------------------------------------------------------- helpers
def _fix_lang(r_el):
    rpr = r_el.find(A + 'rPr')
    if rpr is not None:
        rpr.set('lang', 'ja-JP')
        rpr.set('altLang', 'en-US')


def set_lines(tf_owner, lines):
    """テキスト枠の書式（段落・最初のrun）を保ったまま行を差し替える。"""
    tf = tf_owner.text_frame
    if isinstance(lines, str):
        lines = [lines]
    protos = [p._p for p in tf.paragraphs]
    txBody = protos[0].getparent()
    for p in protos:
        txBody.remove(p)
    for i, text in enumerate(lines):
        p = copy.deepcopy(protos[min(i, len(protos) - 1)])
        for extra in p.findall(A + 'br') + p.findall(A + 'fld'):
            p.remove(extra)
        runs = p.findall(A + 'r')
        if not runs:
            end = p.find(A + 'endParaRPr')
            r = etree.SubElement(p, A + 'r')
            rpr = copy.deepcopy(end) if end is not None else etree.Element(A + 'rPr')
            rpr.tag = A + 'rPr'
            r.append(rpr)
            etree.SubElement(r, A + 't')
            if end is not None:
                p.remove(end)
                p.append(end)
            runs = [r]
        for r in runs[1:]:
            p.remove(r)
        runs[0].find(A + 't').text = text
        _fix_lang(runs[0])
        txBody.append(p)


def shape(slide, name):
    hits = [s for s in slide.shapes if s.name == name]
    assert len(hits) == 1, (name, len(hits))
    return hits[0]


def put(slide, mapping):
    for name, lines in mapping.items():
        set_lines(shape(slide, name), lines)


def title_ph(slide):
    return [s for s in slide.placeholders
            if s.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)][0]


def lead_ph(slide):
    return [s for s in slide.placeholders if s.name in ('字幕 3', 'Subtitle 1')][0]


def add_text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    """paras: [(text, size, bold, color)]"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(72000)
    for i, (txt, size, bold, rgb) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = txt
        r.font.size, r.font.bold = Pt(size), bold
        r.font.color.rgb = rgb
        _fix_lang(r._r)
    return tb


def replace_picture(slide, old_name, blob, src_pic, cover_on_top):
    """old_name の画像を blob に差し替える（位置・大きさは参考資料の画像に合わせ、重なり順は保つ）。"""
    old = shape(slide, old_name)
    new = slide.shapes.add_picture(io.BytesIO(blob), src_pic.left, src_pic.top, src_pic.width, src_pic.height)
    for side in ('left', 'right', 'top', 'bottom'):      # 参考資料のトリミング（画像内のフッター隠し）を引き継ぐ
        setattr(new, f'crop_{side}', getattr(src_pic, f'crop_{side}'))
    cover = [x for x in slide.shapes if x.name == '正方形/長方形 62']
    # 参考資料と同じく、画像の右上ラベルを隠す白地の矩形を画像の上に重ねる
    (cover[0]._element if (cover and cover_on_top) else old._element).addprevious(new._element)
    old._element.getparent().remove(old._element)


S = prs.slides

# ---------------------------------------------------------------- 頁構成：13頁の直後に財務内容を追加
fin_slide = duplicate_slide(prs, 14, after=12)   # 「主な歩み」頁の型（見出し帯・ご確認事項帯）を流用
S = prs.slides
# 以後（0始まり）12=会社概要,13=財務内容,14=事業概要,15=主な歩み,16=強み,17=中扉,18・19=成約事例

# ---------------------------------------------------------------- 1 表紙・3 目次
cover = S[0]
set_lines(shape(cover, 'テキスト ボックス 1'), '株式会社STOP　御中')
[x for x in cover.shapes if x.name.endswith('1') and '御中' in x.text_frame.text][0].text_frame.word_wrap = True  # 短い宛名でも左寄せ
date_runs = shape(cover, 'テキスト ボックス 11').text_frame.paragraphs[0].runs
date_runs[2].text, date_runs[4].text = '11', '25'
agenda = shape(S[2], 'Google Shape;259;p2')
set_lines(agenda, [p.text if p.text != '候補先企業について' else '成約事例に関して'
                   for p in agenda.text_frame.paragraphs])

# ---------------------------------------------------------------- 13 会社概要
s = S[12]
set_lines(lead_ph(s),
          '2010年設立。2023年にIT事業を再始動し、2024年5月に連続起業家の藤方裕伸氏が代表に就任。生成AI・受託開発・'
          'SES・IT人材支援を展開し、2025年9月期は売上高約2億円・従業員31名へ拡大。')
tbl = shape(s, '表 12').table
rows = [
    ('会社名', '株式会社STOP'),
    ('本社所在地', ['東京都港区六本木七丁目15番7号', '（2026年2月 江東区有明より移転）']),
    ('資本金', '92,250千円（発行済株式総数 544株）'),
    ('設立', '2010年4月30日'),
    ('役員', None),
    ('株主', ['高垣 佳典 様（30％）、藤方 雅志 様（18.8％）', '電子公園（15％）ほか ※TSR情報・現状は要確認']),
    ('事業内容', '生成AIプラットフォーム、受託開発、SES、IT人材採用支援 等'),
    ('従業員数', '31名（2025年9月期末）'),
    ('許認可', ['労働者派遣事業 派13-316741', '有料職業紹介事業 13-ユ-316241']),
    ('取引銀行', 'みずほ銀行（京橋）、日本政策金融公庫、りそな銀行（新都心営業部）、三井住友銀行（京橋）'),
    ('販売先', 'グリーエックス（30％）、サイオステクノロジー（20％）、電通（15％）、地方自治体'),
    ('仕入先', 'アマゾンウェブサービスジャパン'),
]
for r, (label, val) in enumerate(rows):
    set_lines(tbl.cell(r, 0), label)
    if val is not None:
        set_lines(tbl.cell(r, 1), val)
set_lines(tbl.cell(4, 1), ['代表取締役', '取締役', '取締役'])
set_lines(tbl.cell(4, 2), ['藤方 裕伸 様', '高垣 佳典 様', '藤方 雅志 様'])
for r, hh in enumerate([340, 430, 340, 340, 640, 430, 400, 340, 430, 430, 430, 340]):
    tbl.rows[r].height = Emu(hh * 1000)
# ロゴ：公式サイトから取得できないため文字ロゴに置き換える
logo = shape(s, '図 6')
x, y, w, h = logo.left, logo.top, logo.width, logo.height
logo._element.getparent().remove(logo._element)
add_text(s, x, y, w, h, [('STOP', 28, True, RGBColor(0x0B, 0x3C, 0x6E)),
                         ('生成AI × 受託開発 × IT人材支援', 11, False, RGBColor(0x40, 0x40, 0x40))])
ch = shape(s, 'グラフ 15').chart
cd = CategoryChartData()
cd.categories = ['2024年1月期', '2024年9月期（8か月）', '2025年9月期']
cd.add_series('年間売上', [v * 1000 for v in F['売上高']])
cd.add_series('当期純利益', [v * 1000 for v in F['当期純利益']])
ch.replace_data(cd)
cs = ch._chartSpace
for dl in list(cs.iter(C + 'dLbl')):
    dl.getparent().remove(dl)
for ax, (mx, mn) in zip(cs.findall('.//' + C + 'valAx'), [(250_000_000, 0), (10_000_000, -10_000_000)]):
    sc = ax.find(C + 'scaling')
    for tag in ('max', 'min'):
        el = sc.find(C + tag)
        if el is not None:
            sc.remove(el)
    for tag, v in (('max', mx), ('min', mn)):
        el = etree.SubElement(sc, C + tag)
        el.set('val', str(v))

# ---------------------------------------------------------------- 14 財務内容（新規）
s = fin_slide
keep = {'Subtitle 1', 'Title 2', 'タイトル 1', 'Rounded Rectangle 8', 'TextBox 9', 'Rounded Rectangle 37',
        'TextBox 38', 'Rounded Rectangle 48', 'Rounded Rectangle 49', 'Rectangle 50', 'TextBox 51', 'TextBox 52'}
for sh in list(s.shapes):
    if sh.name not in keep:
        sh._element.getparent().remove(sh._element)
set_lines(title_ph(s), '財務内容')
set_lines(lead_ph(s),
          f'2024年の代表交代以降、受託開発・SESの拡大で急成長し、2025年9月期は売上高{k(F["売上高"][2])}千円・'
          f'営業利益{k(F["営業利益"][2])}千円の黒字。増資で調達した資金を人員増強と自社AI基盤へ投じていると推察。')
L, R_EDGE = 448056, 10259568
LW = 4535424
RX = L + LW + 137160
RW = R_EDGE - RX
for bar, tb, x, w in (('Rounded Rectangle 8', 'TextBox 9', L, LW), ('Rounded Rectangle 37', 'TextBox 38', RX, RW)):
    b, t = shape(s, bar), shape(s, tb)
    b.left, b.width = x, w
    t.left, t.width = x + 155448, w - 2 * 155448
set_lines(shape(s, 'TextBox 9'), '損益計算書（PL）')
set_lines(shape(s, 'TextBox 38'), '貸借対照表（BS）')
TY = Emu(1673352 + 283464 + 64008).inches
head = ['（単位：千円）'] + [f'{p}\n（{m}か月）' for p, m in zip(PERIODS, MONTHS)]
pl_rows = [head,
           ['売上高'] + [k(v) for v in F['売上高']],
           ['売上原価'] + [k(v) for v in F['売上原価']],
           ['売上総利益'] + [k(v) for v in gp],
           ['　売上総利益率'] + [pc(g / s_) for g, s_ in zip(gp, F['売上高'])],
           ['販売費及び一般管理費'] + [k(v) for v in F['販管費']],
           ['営業利益'] + [k(v) for v in F['営業利益']],
           ['　営業利益率'] + [pc(o / s_) for o, s_ in zip(F['営業利益'], F['売上高'])],
           ['経常利益'] + [k(v) for v in F['経常利益']],
           ['当期純利益'] + [k(v) for v in F['当期純利益']],
           ['期末従業員数（名）'] + [str(v) for v in F['従業員数']]]
house_table(s, Emu(L).inches, TY, Emu(LW).inches, [1.75, 1.07, 1.07, 1.07], pl_rows,
            row_h=0.30, head_h=0.42, size=8.5, bold_rows=(1, 3, 6, 9))
bs_rows = [['（単位：千円）'] + [p.replace('期', '末') for p in PERIODS],
           ['現金及び預金'] + [k(v) for v in F['現金及び預金']],
           ['売掛金'] + [k(v) for v in F['売掛金']],
           ['仕掛品'] + [k(v) for v in F['仕掛品']],
           ['その他流動資産'] + [k(v) for v in other_ca],
           ['固定資産'] + [k(v) for v in F['固定資産合計']],
           ['　うち無形固定資産'] + [k(v) for v in F['無形固定資産']],
           ['資産合計'] + [k(v) for v in F['資産合計']],
           ['買掛金・未払金'] + [k(v) for v in payable],
           ['借入金'] + [k(v) for v in debt],
           ['その他負債'] + [k(v) for v in other_li],
           ['純資産'] + [k(v) for v in F['純資産合計']],
           ['　自己資本比率'] + [pc(v) for v in eq_ratio]]
house_table(s, Emu(RX).inches, TY, Emu(RW).inches, [1.85, 1.1, 1.1, 1.1], bs_rows,
            row_h=0.25, head_h=0.42, size=8.5, bold_rows=(7, 11))
# 財務面のポイント
PY, PH = Inches(5.70), Inches(0.98)
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, L, PY, R_EDGE - L, PH)
box.adjustments[0] = 0.08
box.fill.solid()
box.fill.fore_color.rgb = RGBColor(0xF2, 0xF6, 0xFA)
box.line.color.rgb = RGBColor(0xB9, 0xC2, 0xC7)
box.line.width = Pt(0.75)
box.shadow.inherit = False
add_text(s, L + Emu(90000), PY + Emu(30000), R_EDGE - L - Emu(180000), PH - Emu(60000), [
    ('財務面のポイント', 9.5, True, NAVY),
    (f'・2025年9月期は売上高{k(F["売上高"][2])}千円（8か月決算の年換算比 約{growth:.1f}倍）、'
     f'営業利益{k(F["営業利益"][2])}千円と黒字を確保。売上総利益率は{pc(gp[2] / F["売上高"][2])}と高水準', 8.5, False, BODY),
    (f'・増資により払込資本は{F["払込資本"][0] / 1000:,.0f}万円→{F["払込資本"][2] / 1000:,.0f}万円へ'
     f'（資本金は現在9,225万円）。借入金は{k(debt[2])}千円、自己資本比率は{pc(eq_ratio[2])}', 8.5, False, BODY),
    (f'・総資産の約{soft_ratio * 100:.0f}％を仕掛品・前払費用・無形固定資産が占め、期末の現預金は'
     f'{k(F["現金及び預金"][2])}千円', 8.5, False, BODY),
], anchor=MSO_ANCHOR.MIDDLE)
set_lines(shape(s, 'TextBox 52'), '未払金・前払費用・仕掛品の内容／売上原価と販管費の計上区分／月次の資金繰りと今期（2026年9月期）の見通し')
add_text(s, L, Inches(7.40), R_EDGE - L, Inches(0.2), [
    ('出典：東京商工リサーチ 財務情報（2026年11月取得）。2024年9月期は決算期変更による8か月決算。'
     '千円未満の端数処理により合計が一致しない場合があります', 7, False, GRAY)])

# ---------------------------------------------------------------- 15 事業概要・提供サービス
s = S[14]
set_lines(lead_ph(s),
          '受託開発・SESで顧客基盤と開発体制を築きつつ、生成AIプラットフォーム等の自社サービスを展開する'
          '「受託＋自社プロダクト」の二層モデル。人材紹介・派遣の許可も保有する。')
put(s, {
    'TextBox 9': '主な事業領域（公式サイト・TSR情報）',
    'TextBox 13': '生成AI基盤', 'TextBox 14': ['50超のAIモデルを統合した', '「OrionNaut」を提供'],
    'TextBox 18': '受託開発', 'TextBox 19': ['AI・基幹システム・スマホ', 'アプリ等（2024年4月開始）'],
    'TextBox 23': 'SES', 'TextBox 24': ['顧客先へのエンジニア常駐', '（2023年9月開始）'],
    'TextBox 28': 'IT人材支援', 'TextBox 29': ['IT・Web専門の人材紹介・', '派遣（許可取得済）'],
    'TextBox 33': 'データ解析・監視', 'TextBox 34': ['データ解析プラットフォーム、', 'サーバ監視運用'],
    'TextBox 38': '動画配信基盤', 'TextBox 39': ['VOD／ライブ配信基盤、', '動画マーケティング支援'],
    'TextBox 41': '事業フロー（公開情報に基づく整理）',
    'TextBox 43': '顧客企業（大手Web・IT・広告・自治体）',
    'TextBox 46': 'STOP（開発・運用・人材）',
    'TextBox 47': '受託開発 ／ SES ／ 生成AI・データ基盤 ／ 人材紹介',
    'TextBox 48': '↕ 連携',
    'TextBox 50': 'クラウド基盤・開発パートナー',
    'TextBox 51': ['AWS（主要仕入先）上でサービスを構築・運用', '外部パートナーとの協業の有無は要確認'],
    'TextBox 53': '※事業別の売上構成、元請／二次請の比率は要確認',
    'TextBox 58': '事業別（受託・SES・自社サービス・人材）の売上・粗利構成／元請比率／稼働エンジニア数',
})

# ---------------------------------------------------------------- 16 主な歩み・自社サービス
s = S[15]
set_lines(lead_ph(s),
          '2023年のIT事業再始動、2024年の代表交代を経て、SES→受託開発→自社AIプラットフォームへと'
          '短期間で事業領域を拡大している。')
put(s, {
    'TextBox 9': '主な歩み（公式サイト・登記・公開情報）',
    'TextBox 15': '2010年4月', 'TextBox 16': ['会社設立', 'Webメディア「NeoL」の運営等'],
    'TextBox 19': '2023年8月', 'TextBox 20': ['労働者派遣事業の許可を取得', '9月 SES事業を開始（IT事業の再始動）'],
    'TextBox 23': '2024年4月', 'TextBox 24': ['受託開発事業を開始', '5月 藤方 裕伸 様が代表取締役に就任'],
    'TextBox 27': '2024年6月', 'TextBox 28': ['六本木オフィスを開設', '9月 決算期を1月から9月へ変更'],
    'TextBox 31': '2026年2月', 'TextBox 32': '本店を東京都港区六本木へ移転',
    'TextBox 35': '2026年7月', 'TextBox 36': ['統合AIプラットフォーム「OrionNaut」', 'のβ版提供を開始'],
    'TextBox 38': '自社サービス・経営体制',
    'TextBox 40': '統合AIプラットフォーム「OrionNaut」',
    'TextBox 41': ['50以上の最新AIモデルを1つの基盤に統合', '（2026年7月β版提供開始）'],
    'TextBox 43': '受託開発・SES',
    'TextBox 44': ['AI・基幹システム・スマホアプリの受託開発と', 'エンジニア常駐。求人媒体では「大手直請メイン」', 'と紹介'],
    'TextBox 46': '代表 藤方 裕伸 様',
    'TextBox 47': ['ソフトウェア開発会社を創業し、2019年に', '上場企業へ売却した連続起業家（報道ベース）'],
    'TextBox 52': 'OrionNautの収益化の見通しと投資額（無形資産計上の有無）／受託・SESの継続契約の状況',
})

# ---------------------------------------------------------------- 17 事業の強み・ご確認事項
s = S[16]
set_lines(lead_ph(s),
          '大手企業との取引基盤、生成AIの自社プラットフォーム、急成長を実現した経営チームを主な特徴と認識。'
          '現状の事業構成と今後のご意向を伺い、貴社にとって最適な選択肢を取れるよう、ご支援いたします。')
put(s, {
    'TextBox 13': '急成長の実績',
    'TextBox 15': ['代表交代後、売上高は', '約2億円・従業員31名へ', f'拡大（年換算 約{growth:.0f}倍）'],
    'TextBox 16': f'2025年9月期 売上高{k(F["売上高"][2])}千円',
    'TextBox 20': '大手との取引基盤',
    'TextBox 22': ['グリーエックス、サイオス', 'テクノロジー、電通等との', '取引基盤'],
    'TextBox 23': 'TSR様の情報を参考',
    'TextBox 27': '生成AIの自社基盤',
    'TextBox 29': ['50以上のAIモデルを統合', 'した「OrionNaut」を', '自社で開発・提供'],
    'TextBox 30': '2026年7月β版提供開始',
    'TextBox 34': '経営者の実行力',
    'TextBox 36': ['開発会社の創業・売却を', '経験した代表のもと、', '採用・組織を急拡大'],
    'TextBox 37': '派遣・職業紹介の許可も保有',
    'TextBox 43': '事業構成・収益性',
    'TextBox 44': ['受託・SES・自社サービス・人材紹介の', '売上・粗利構成と2026年9月期の見通し'],
    'TextBox 48': '主要顧客・受注形態',
    'TextBox 49': ['上位顧客の取引比率と継続性、', '元請／二次請の比率、稼働エンジニア数'],
    'TextBox 53': '資本政策・資金繰り',
    'TextBox 54': ['増資の引受先と現在の株主構成、', '運転資金・資金繰りの状況'],
    'TextBox 58': 'M&Aのご意向・時期',
    'TextBox 59': ['株式譲渡・資本提携（成長資金）等の', 'お考えと、藤方様の今後のご関与'],
})

# ---------------------------------------------------------------- 19・20 成約事例（IT業界の事例画像へ差し替え）
ref = Presentation(CASES)
for dst_idx, src_idx in ((18, 18), (19, 19)):       # 参考資料の19・20頁（0始まり18・19）
    src_pic = [x for x in ref.slides[src_idx].shapes if x.shape_type == 13 and x.width > Inches(5)][0]
    dst = S[dst_idx]
    old = [x for x in dst.shapes if x.shape_type == 13 and x.width > Inches(5)][0]
    src_shapes = list(ref.slides[src_idx].shapes)
    rect = [i for i, x in enumerate(src_shapes) if x.name == '正方形/長方形 62']
    on_top = bool(rect) and rect[0] > src_shapes.index(src_pic)   # 参考資料と同じ重なり順にする
    replace_picture(dst, old.name, src_pic.image.blob, src_pic, on_top)

prs.save(OUT)
scrub_doc_props(OUT, title='ご面談資料')
print('saved', OUT, len(prs.slides), 'slides')
print(f'年換算成長率 {growth:.3f} / 資産に占める仕掛品等 {soft_ratio:.3f} / 自己資本比率 {eq_ratio}')
