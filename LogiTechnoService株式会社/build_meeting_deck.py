"""LogiTechnoService株式会社 初回面談資料（ご面談資料）のビルドスクリプト。

    python build_meeting_deck.py <参考にした初回面談資料.pptx> <出力.pptx>

参考資料（当社の初回面談資料フォーマット・32頁）を複製し、
  - 当社について／今後のステップ／Appendix の定型頁はそのまま残す
  - 「貴社への理解」の4頁と「候補先企業について」の1頁を本件の内容に差し替える
  - 前案件固有の頁（画像の事業概要、類似業種のM&A事例2頁）は削除する
数値の出所：調査会社データ（2025/10/31調査）、履歴事項（2026/11/20取得）、公開記事（各頁の注記）。
"""
import copy
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.util import Pt, Emu

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'

SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)


# ---------------------------------------------------------------- helpers
def _fix_lang(r_el):
    rpr = r_el.find(A + 'rPr')
    if rpr is not None:
        rpr.set('lang', 'ja-JP')
        rpr.set('altLang', 'en-US')


def set_lines(tf_owner, lines):
    """テキスト枠の書式（段落・最初のrun）を保ったまま行を差し替える。lines は文字列か list。"""
    tf = tf_owner.text_frame
    if isinstance(lines, str):
        lines = [lines]
    paras = list(tf.paragraphs)
    protos = [p._p for p in paras]
    txBody = protos[0].getparent()
    for p in protos:
        txBody.remove(p)
    for i, text in enumerate(lines):
        proto = protos[min(i, len(protos) - 1)]
        p = copy.deepcopy(proto)
        runs = p.findall(A + 'r')
        for extra in p.findall(A + 'br') + p.findall(A + 'fld'):
            p.remove(extra)
        if not runs:  # 空段落の場合は endParaRPr から run を作る
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
    from pptx.enum.shapes import PP_PLACEHOLDER
    return [s for s in slide.placeholders
            if s.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)][0]


sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                '..', '.claude', 'skills', 'ma-deal-docs', 'scripts'))
from im_builder import delete_slide, duplicate_slide as _dup  # noqa: E402


def duplicate_slide(index, after):
    return _dup(prs, index, after)


# ---------------------------------------------------------------- 頁構成
# 参考資料：14=事業概要（画像1枚）、19・20=類似業種（IT）のM&A事例（画像）→ 本件では使わない
for i in sorted([14, 19, 20], reverse=True):
    delete_slide(prs, i - 1)
# ここで（0始まり）12=会社概要,13=事業概要・提供サービス,14=主要実績,15=強み・確認事項,16=中扉
cand = duplicate_slide(15, after=16)  # 17頁（強み・確認事項）の型で「候補先企業のイメージ」を作る

S = prs.slides

# ---------------------------------------------------------------- 1 表紙
cover = S[0]
set_lines(shape(cover, 'テキスト ボックス 1'), 'LogiTechnoService株式会社　御中')

# ---------------------------------------------------------------- 13 会社概要
s = S[12]
set_lines([x for x in s.placeholders if x.name == '字幕 3'][0],
          '2023年7月設立。海上コンテナ中心の貨物利用運送を主力に、物流コンサルティング・物流システム企画販売を展開。'
          '設立2期目の2025年6月期に売上高約3.5億円（前期比+78.9％）へ拡大。')
tbl = [x for x in s.shapes if x.has_table][0].table
rows = [
    ('会社名', 'LogiTechnoService株式会社（ロジテクノサービス）'),
    ('本社所在地', ['東京都武蔵村山市神明二丁目51番地の15', '東京事務所：東京都大田区東海4-10-8 10F']),
    ('資本金', '4,000千円（発行済株式総数 400株）'),
    ('設立', '2023年7月25日'),
    ('役員', None),
    ('株主', ['清島 陽介 様（50％）、渡邉 裕 様（50％）', '※調査会社情報（2025年10月時点）・現状は要確認']),
    ('事業内容', '貨物利用運送（90％）、物流コンサルティング（9％）、物流システム企画販売（1％）'),
    ('従業員数', '1名（調査会社情報・2025年10月時点）'),
    ('許認可', '貨物利用運送事業 関自貨第900号'),
    ('取引銀行', '山梨中央銀行（立川）、日本政策金融公庫、GMOあおぞらネット銀行、楽天銀行'),
    ('販売先', '近鉄エクスプレス、鴻池運輸、伊藤忠ロジスティクス、エスワイプロモーション'),
    ('仕入先', 'アスカ、OK物流サービス、エムカンパニー、日軽物流、庄野運輸、ラップタイマー'),
]
for r, (label, val) in enumerate(rows):
    set_lines(tbl.cell(r, 0), label)
    if val is not None:
        set_lines(tbl.cell(r, 1), val)
set_lines(tbl.cell(4, 1), ['代表取締役', '（辞任）'])
set_lines(tbl.cell(4, 2), ['清島 陽介 様', '渡邉 裕 様（2026年9月30日）'])
tbl.rows[4].height = Emu(560000)
# 前案件のロゴ画像を外し、社名の文字ロゴに置き換える
logo = shape(s, '図 8')
x, y, w, h = logo.left, logo.top, logo.width, logo.height
logo._element.getparent().remove(logo._element)
tb = s.shapes.add_textbox(x, y, w, h)
tf = tb.text_frame
tf.word_wrap = True
for i, (txt, size, bold, rgb) in enumerate([
        ('LogiTechnoService', 26, True, RGBColor(0x0B, 0x3C, 0x6E)),
        ('海上コンテナ輸送 × ICDコンサルティング × 物流DX', 11, False, RGBColor(0x40, 0x40, 0x40))]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    run = p.add_run()
    run.text = txt
    run.font.size, run.font.bold = Pt(size), bold
    run.font.color.rgb = rgb
    run.font.name = 'Arial' if i == 0 else None
    _fix_lang(run._r)
set_lines(shape(s, '正方形/長方形 13'), '財務データ（単位：千円）')
# グラフ：調査会社データ（決算期は6月。第1期は2023年7月〜2024年6月）
ch = shape(s, 'グラフ 4').chart
cd = CategoryChartData()
cd.categories = ['2024年6月期（第1期）', '2025年6月期（第2期）']
cd.add_series('年間売上', (194_522_000, 347_950_000))
cd.add_series('利益', (7_810_000, 6_189_000))
ch.replace_data(cd)
cs = ch._chartSpace
for dl in cs.iter(C + 'dLbl'):  # 前案件の3期目用の手動ラベル位置を外す
    dl.getparent().remove(dl)
vals = cs.findall('.//' + C + 'valAx')
for ax, (mx, mn) in zip(vals, [(400_000_000, 0), (20_000_000, 0)]):
    sc = ax.find(C + 'scaling')
    for tag, v in (('max', mx), ('min', mn)):
        el = sc.find(C + tag)
        if el is None:
            el = etree.SubElement(sc, C + tag)
        el.set('val', str(v))
    # scaling の子要素は orientation → max → min の順
    orient = sc.find(C + 'orientation')
    for tag in ('max', 'min'):
        el = sc.find(C + tag)
        sc.remove(el)
        sc.append(el)
    if orient is not None:
        sc.remove(orient)
        sc.insert(0, orient)
note = s.shapes.add_textbox(Emu(5490550), Emu(6960000), Emu(4800000), Emu(200000))
note.text_frame.word_wrap = True
r = note.text_frame.paragraphs[0].add_run()
r.text = '出典：調査会社データ（2025年10月調査）。利益の段階（経常/当期）は要確認。2026年6月期は未入手'
r.font.size = Pt(7)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
_fix_lang(r._r)

# ---------------------------------------------------------------- 14 事業概要・提供サービス
s = S[13]
set_lines([x for x in s.placeholders if x.name == 'Subtitle 1'][0],
          '荷主・フォワーダーから海上コンテナ輸送を受注し、協力運送会社に実運送を委託する利用運送モデル（アセットライト型）。'
          'ICDコンサルティングとシステム提供で、2024年問題下のドレージ不足に対応。')
put(s, {
    'TextBox 9': '主な事業領域（調査会社情報・公開情報）',
    'TextBox 13': '貨物利用運送', 'TextBox 14': ['海上コンテナ輸送が中心', '売上構成比 約90％'],
    'TextBox 18': '物流コンサル', 'TextBox 19': ['ICD・デポの活用による', '輸送効率化の助言（約9％）'],
    'TextBox 23': '物流システム販売', 'TextBox 24': ['業務管理システム「ロジマネ」', 'の共同開発・販売（約1％）'],
    'TextBox 28': '海コンマッチング', 'TextBox 29': ['荷主・フォワーダーと', 'ドレージ会社を結ぶ会員制'],
    'TextBox 33': 'ICD拠点の支援', 'TextBox 34': ['山梨県のICD中継ターミナル', 'を顧問として支援'],
    'TextBox 38': '物流研究会の開催', 'TextBox 39': ['荷主・運送会社等が業界課題', 'を議論する場を継続開催'],
    'TextBox 41': '事業フロー（公開情報に基づく整理）',
    'TextBox 43': '荷主・フォワーダー・物流企業',
    'TextBox 46': 'LogiTechnoService（利用運送事業者）',
    'TextBox 47': '受注・配車手配 ／ ICD活用の提案 ／ 運行・請求管理',
    'TextBox 48': '↕ 委託',
    'TextBox 50': '協力運送会社（海上コンテナのドレージ）',
    'TextBox 51': ['トレーラー・ドライバーを保有する実運送会社へ委託', '必要に応じICDで中継・保管・空コンテナ再利用（CRU）'],
    'TextBox 53': '※事業別の売上総利益、協力会社への委託条件・支払サイトは要確認',
    'TextBox 58': '事業別の売上・粗利構成／主要顧客の取引比率と契約形態／協力運送会社の確保状況',
})

# ---------------------------------------------------------------- 15 主な歩み・自社サービス
s = S[14]
set_lines(title_ph(s), '主な歩み・自社サービス')
set_lines([x for x in s.placeholders if x.name == 'Subtitle 1'][0],
          '設立以来、ICD中継ターミナルの立ち上げ支援、業務管理システム「ロジマネ」、マッチングサービス「海コンマッチング」と、'
          '海上コンテナ物流の課題解決型サービスを相次いで展開している。')
put(s, {
    'TextBox 9': '設立以来の主な歩み（公開情報）',
    'TextBox 11': '時期', 'TextBox 12': '主な出来事',
    'TextBox 15': '2023年7月', 'TextBox 16': ['会社設立（本社：東京都武蔵村山市）', '海上コンテナ中心の貨物利用運送事業を開始'],
    'TextBox 19': '2024年3月', 'TextBox 20': ['顧問先エムズ物流（山梨県南アルプス市）の', '「ICD中継ターミナル」が本格稼働'],
    'TextBox 23': '2025年5月', 'TextBox 24': '業務管理システム「ロジマネ」の提供開始',
    'TextBox 27': '2026年1月', 'TextBox 28': ['「海コンマッチング」の提供開始を発表', '（エックスグラビティ社と共同、2月リリース）'],
    'TextBox 31': '2026年7月', 'TextBox 32': '「物流研究会」第8回を開催',
    'TextBox 35': '2026年9月', 'TextBox 36': ['取締役 渡邉 裕 様が辞任（10月登記）', '清島 陽介 様の単独代表体制へ'],
    'TextBox 38': '自社サービスの概要',
    'TextBox 40': '業務管理システム「ロジマネ」',
    'TextBox 41': ['海コンドレージ向けクラウド型。受注・帳票', '（依頼書・配車台帳）・請求を一元管理'],
    'TextBox 43': '海コンマッチング',
    'TextBox 44': ['空車情報・案件情報を登録・検索し、荷主・', 'フォワーダーとドレージ会社をマッチング', '成立まで企業名は非公開、月額制の会員サービス'],
    'TextBox 46': 'ICDコンサルティング',
    'TextBox 47': ['「日本初のICDコンサルタント」を標榜。内陸', 'デポでの中継・保管・CRUで輸送を効率化'],
    'TextBox 52': 'ロジマネ・海コンマッチングの売上・会員数／エックスグラビティ社との契約関係（権利帰属・収益配分）',
})

# ---------------------------------------------------------------- 16 事業の強み・ご確認事項
s = S[15]
set_lines([x for x in s.placeholders if x.name == 'Subtitle 1'][0],
          '大手物流企業との取引基盤、海コン×ICDの専門性、協力運送会社ネットワーク、DXサービスを主な特徴と認識。'
          '現状の事業構成と今後のご意向を伺い、貴社にとって最適な選択肢を検討したい。')
put(s, {
    'TextBox 13': '海コン×ICDの専門性',
    'TextBox 15': ['海上コンテナ輸送とICD活用', 'に特化した提案力。「日本初', 'のICDコンサルタント」を標榜'],
    'TextBox 16': 'ICD中継ターミナル立上げを支援',
    'TextBox 20': '大手物流企業との取引',
    'TextBox 22': ['近鉄エクスプレス、鴻池運輸、', '伊藤忠ロジスティクス等との', '取引基盤'],
    'TextBox 23': '販売先：調査会社情報',
    'TextBox 27': 'アセットライトな成長',
    'TextBox 29': ['車両を持たない利用運送モデル', 'で、設立2期目に売上高', '約3.5億円（前期比+78.9％）'],
    'TextBox 30': '2025年6月期 売上高347,950千円',
    'TextBox 34': 'DX・業界ネットワーク',
    'TextBox 36': ['「ロジマネ」等のサービスと', '「物流研究会」による', '業界横断のネットワーク'],
    'TextBox 37': 'エックスグラビティ社と共同展開',
    'TextBox 43': '事業構成・収益性',
    'TextBox 44': ['事業別の売上・粗利構成、', '2026年6月期の着地と今期の見通し'],
    'TextBox 48': '主要顧客・協力会社',
    'TextBox 49': ['上位顧客の取引比率・契約形態、', '協力運送会社の確保と代表者への依存度'],
    'TextBox 53': '株主・経営体制',
    'TextBox 54': ['2026年9月の役員変更後の株主構成、', '清島様の今後のご関与についてのお考え'],
    'TextBox 58': 'M&Aのご意向・時期',
    'TextBox 59': ['株式譲渡・資本提携・事業拡大（譲受）', 'の中でのお考えと、ご希望の時期・条件'],
})

# ---------------------------------------------------------------- 17 中扉
s = S[16]
set_lines(title_ph(s), '候補先企業について')

# ---------------------------------------------------------------- 18 候補先企業のイメージ
s = cand
set_lines(title_ph(s), '想定される候補先企業のイメージ')
set_lines(shape(s, 'タイトル 1'), '　候補先企業について')
set_lines([x for x in s.placeholders if x.name == 'Subtitle 1'][0],
          '貴社の取引基盤・協力会社ネットワーク・ICDとDXのノウハウは、物流事業者に加え物流DX企業にもシナジーが'
          '見込めると想定。具体的な候補先は、本日伺うご意向を踏まえてご提案いたします。')
put(s, {
    'TextBox 9': '想定される候補先のタイプ（仮説）',
    'TextBox 13': '国際物流・フォワーダー',
    'TextBox 15': ['海コン陸送の手配力を取り込み、', '輸出入顧客へのサービスを', '強化したい企業'],
    'TextBox 16': '例：大手・中堅フォワーダー、商社系物流',
    'TextBox 20': '運送・ドレージ会社',
    'TextBox 22': ['利用運送の受注基盤を取り込み、', '車両稼働率・荷主開拓を', '高めたい実運送会社'],
    'TextBox 23': '例：海コン・トレーラー輸送事業者',
    'TextBox 27': '港湾・倉庫・ICD事業者',
    'TextBox 29': ['ICD・デポのノウハウを活かし、', '内陸拠点ビジネスを', '拡大したい事業者'],
    'TextBox 30': '例：港湾運送・倉庫・バンプール運営',
    'TextBox 34': '物流DX・IT企業',
    'TextBox 36': ['「ロジマネ」等のサービスと', '物流会社との接点を得て、', '物流SaaSを拡大したい企業'],
    'TextBox 37': '例：物流SaaS・システム開発会社',
    'TextBox 39': '候補先から見た貴社の魅力（仮説）',
    'TextBox 43': '大手物流企業との取引口座',
    'TextBox 44': ['近鉄エクスプレス等との取引基盤と、', '設立2期で売上高約3.5億円の成長性'],
    'TextBox 48': '協力運送会社ネットワーク',
    'TextBox 49': ['ドライバー不足の中で希少な', '海上コンテナ陸送の手配力'],
    'TextBox 53': 'ICDコンサルティングの知見',
    'TextBox 54': ['2024年問題への解決策となる', '中継輸送・CRUの提案力'],
    'TextBox 58': 'サービス・業界ネットワーク',
    'TextBox 59': ['ロジマネ・海コンマッチングと、', '物流研究会を通じた業界との接点'],
    'TextBox 61': '※上記は公開情報を基にした初期仮説です。貴社のご意向を伺ったうえで、具体的な候補先をご提案いたします。',
})

prs.save(OUT)
from im_paste import scrub_doc_props  # noqa: E402
scrub_doc_props(OUT, title='ご面談資料')
print('saved', OUT, len(prs.slides), 'slides')
