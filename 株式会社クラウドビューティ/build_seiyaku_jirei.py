"""成約事例発表資料 <売 クラウドビューティ社 × 買 財成社> のビルドスクリプト。

成約事例テンプレート（成約事例テンプレート.pptx）を複製し、各頁を埋める。
数値の出所：
  - 企業概要書 260526（部門PL・譲渡対象資産・契約状況・不動産）
  - 論点整理シート 260623（価格2案・賃貸借・承継）
  - 初期DD資料・QAリスト（2026/9/12〜9/15、財成社DD）
未確認の事項は「要確認」と明記している。

実行：python3 build_seiyaku_jirei.py
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, '.claude/skills/ma-deal-docs/scripts'))

from pptx.util import Pt, Emu, Inches  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
import im_builder as B  # noqa: E402

TEMPLATE = os.path.join(HERE, '成約事例テンプレート.pptx')
OUT = os.path.join(HERE, '成約事例_クラウドビューティ×財成_261123.pptx')
IMG = os.path.join(HERE, 'img')

# ---------------------------------------------------------------- 数値（出所は冒頭）
PRICE_A = 52_500          # 基準額（運営委託なし）千円
PRICE_B = 37_500          # 基準額（運営委託あり）千円
OUTSOURCE_Y = 5_000       # 運営委託費（3年で約15,000千円を均等按分）千円/年
EBITDA = 16_436           # 2026年2月期 部門EBITDA 千円
OP_FCST = 17_111          # 2027年2月期 営業利益見通し 千円
EBITDA_3 = [6_494, 9_910, 16_436]
FULL_REV = 53_866         # 満室時年額
CUR_REV = 42_346          # 2026年1月末 契約ベース年換算

cf_a, cf_b = EBITDA, EBITDA - OUTSOURCE_Y
y_a, y_b = cf_a / PRICE_A, cf_b / PRICE_B
pb_a, pb_b = PRICE_A / cf_a, PRICE_B / cf_b
yf_a, yf_b = OP_FCST / PRICE_A, (OP_FCST - OUTSOURCE_Y) / PRICE_B
avg3 = round(sum(EBITDA_3) / 3)


def pct(v):
    return f'{v * 100:.1f}%'


def by_name(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError(name)


def fill(sh, text, size=None, bold=None, align=None):
    """テンプレート図形の文字を差し替える。最初の段落・ランの書式（色・フォント・箇条書き）を複製して使う。

    im_builder.set_text は色を持たないランを作るため、保存前の force_white_headers で白文字にされる。
    """
    import copy
    from pptx.oxml.ns import qn
    txBody = sh.text_frame._txBody
    ps = txBody.findall(qn('a:p'))
    p_tpl = copy.deepcopy(ps[0])
    r0 = p_tpl.find(qn('a:r'))
    rpr = copy.deepcopy(r0.find(qn('a:rPr'))) if r0 is not None and r0.find(qn('a:rPr')) is not None else None
    for p in ps:
        txBody.remove(p)
    for ln in str(text).split('\n'):
        p = copy.deepcopy(p_tpl)
        for ch in list(p):
            if ch.tag != qn('a:pPr'):
                p.remove(ch)
        r = p.makeelement(qn('a:r'), {})
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        tt = r.makeelement(qn('a:t'), {})
        tt.text = ln
        r.append(tt)
        p.append(r)
        txBody.append(p)
    bp = txBody.find(qn('a:bodyPr'))
    for tag in ('a:normAutofit', 'a:spAutoFit'):
        for el in bp.findall(qn(tag)):
            bp.remove(el)
    for para in sh.text_frame.paragraphs:
        if align is not None:
            para.alignment = align
        for run in para.runs:
            if size is not None:
                run.font.size = Pt(size)
            if bold is not None:
                run.font.bold = bold
    return sh


def put(slide, name, text, size=None, bold=None, align=None):
    return fill(by_name(slide, name), text, size, bold, align)


def parts(shape, items, size=10):
    """≪見出し≫（太字）と本文を交互に書く。items=[(text, bold), ...]、'' は空行。"""
    tf = shape.text_frame
    tf.word_wrap = True
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    first = True
    for text, bold in items:
        p = p0 if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(1)
        if text:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = B.BODY


def picture_into(slide, shape, path, remove=True):
    """図形の位置に画像を縦横比を保って収める（中央寄せ）。"""
    from PIL import Image
    x, y, w, h = shape.left, shape.top, shape.width, shape.height
    iw, ih = Image.open(path).size
    s = min(w / iw, h / ih)
    pw, ph = int(iw * s), int(ih * s)
    pic = slide.shapes.add_picture(path, x + (w - pw) // 2, y + (h - ph) // 2, pw, ph)
    if remove:
        shape._element.getparent().remove(shape._element)
    return pic


def photo_fill(slide, shape, path):
    """正方形枠いっぱいに写真を置く（写真は正方形に切り出し済み）。"""
    pic = slide.shapes.add_picture(path, shape.left, shape.top, shape.width, shape.height)
    shape._element.getparent().remove(shape._element)
    return pic


def replace_runs(shape, mapping):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.text in mapping:
                r.text = mapping[r.text]


# ================================================================ build
prs = B.open_template(TEMPLATE)
B.fix_broken_pictures(prs)
B.strip_red_banners(prs)
S = prs.slides

# ---- 1. 表紙
s = S[0]
title = [sh for sh in s.shapes if sh.is_placeholder][0]
replace_runs(title, {'売 ○○○社': '売 クラウドビューティ社', '買 ○○○社': '買 財成社'})
for r in title.text_frame.paragraphs[0].runs:
    if r.font.size is not None:
        r.font.size = Pt(20)
sub = [sh for sh in s.shapes if not sh.is_placeholder and sh.has_text_frame][0]
replace_runs(sub, {'売り主担当：○○ ○○／買い主担当：○○ ○○':
                   '売り主担当：田垣内 啓介／買い主担当：○○ ○○（要確認）'})

# ---- 4. 譲渡企業
s = S[3]
B.set_lead(s, '美容業界特化型プラットフォーム企業が、直営シェアサロン事業（渋谷表参道・恵比寿の2店舗／全19室・'
              'いずれも賃借）を事業譲渡で切り出した案件。')
put(s, '正方形/長方形 19', '1999年6月（第27期）', 10)
put(s, '正方形/長方形 23', '311,763千円（資本準備金等含む）', 10)
put(s, '正方形/長方形 11', '望月 雅彦 様', 10)
put(s, '正方形/長方形 26', '美容プラットフォーム・POS等（譲渡対象：シェアサロン2店舗）', 9)
put(s, '正方形/長方形 17', '望月 雅彦 様（100％）', 10)
put(s, '正方形/長方形 29', '全社149,488千円／対象事業38,179千円（2026年2月期）', 9)
put(s, '正方形/長方形 32', f'{EBITDA:,}千円（2026年2月期・対象事業の部門EBITDA）', 9)
put(s, '正方形/長方形 35', '簿価9円（内装・設備9点）／再計算簿価12,503千円', 9)
put(s, '正方形/長方形 38', '事業の選択と集中（経営資源の最適配分）\n'
                          '本業の美容プラットフォーム「BEAUTY SHARE」・店舗管理クラウド「BEAUTY POS」へ経営資源を集中', 9)
picture_into(s, by_name(s, '正方形/長方形 2'), os.path.join(IMG, 'logo_bss.png'))
for nm, f in [('正方形/長方形 3', 'ph1.jpg'), ('正方形/長方形 5', 'ph2.jpg'),
              ('正方形/長方形 6', 'ph3.jpg'), ('正方形/長方形 7', 'ph4.jpg')]:
    photo_fill(s, by_name(s, nm), os.path.join(IMG, f))

# ---- 5. 譲渡希望条件
s = S[4]
put(s, 'テキスト ボックス 18', '事業譲渡', 22, True)
put(s, 'テキスト ボックス 23', '事業譲渡（2店舗）', 20, True)
put(s, 'テキスト ボックス 55', '5,750万円', 22, True)
pb = put(s, 'テキスト ボックス 25', '5,250万円（運営委託なし）\n3,750万円＋運営委託（B案）', 13, True)
pb.top, pb.height = Inches(3.50), Inches(0.52)
put(s, 'テキスト ボックス 255', '従業員の転籍なし（専任者不在、兼務1名が週12〜15時間で運営）\n'
                              '引継期間は応相談（売手担当者が2〜3ヶ月程度を想定）\n'
                              '譲渡後、売手への運営委託も選択可能', 14)
put(s, 'テキスト ボックス 263', '事業の選択と集中（経営資源の最適配分）', 18, True)
B.textbox(s, 6.53, 4.07, 4.6, 0.20, '※2案で合意（260623時点）。B案は別途3年で約1,500万円の運営委託費。最終採用案は要確認',
          size=7.5, color=B.NOTE)

# ---- 6. 財務ハイライト（P/L）
s = S[5]
B.set_title(s, '財務ハイライト（P/L：対象事業の部門PL）')
B.set_lead(s, f'2026年2月期は売上高38,179千円・EBITDA{EBITDA:,}千円（利益率43.1%）。減価償却ゼロ・固定費中心で、'
              '増収分がほぼそのまま利益になる構造。')
picture_into(s, by_name(s, '図 17'), os.path.join(IMG, 'pl_table.png'))
parts(by_name(s, 'テキスト ボックス 8'), [
    ('≪売上≫', True),
    ('・恵比寿店の4部屋増床（2024年12月）と稼働率改善で、2026年2月期は前期比127.7%', False),
    ('≪費用≫', True),
    ('・主な費用は地代家賃（売上比約53%）。リネン・清掃・消耗品は軽微', False),
    ('≪営業利益≫', True),
    ('・6,494→9,910→16,436千円と2期連続の増益', False),
    ('≪見通し≫', True),
    ('・2027年2月期は稼働率80%前提で売上高42,260千円、営業利益17,111千円程度の見込み'
     '（2026年1月末の契約ベース年換算売上は42,346千円）', False),
], size=9.5)
parts(by_name(s, 'テキスト ボックス 9'), [
    ('【調整項目】', True),
    ('≪減価償却費≫', True),
    ('・内装・設備は全額特損処理済のため0（EBITDA＝営業利益）', False),
    ('≪人件費≫', True),
    ('・専任者なし・本部兼務のため部門PLに未計上。買手側の運営コスト（または運営委託費）を別途考慮', False),
    ('', False),
    (f'調整後EBITDA（直近期）は{EBITDA:,}千円（過去3期平均{avg3:,}千円）', True),
], size=9.5)

# ---- 7. 譲渡対象資産（B/S に代えて）
s = S[6]
B.set_title(s, '譲渡対象資産（B/Sに代えて）')
B.set_lead(s, '事業譲渡のためB/Sは承継しない。承継資産は内装・設備9点（簿価合計9円）のみで、不動産・敷金・借入金の'
              '承継もない「資産性ゼロ」の案件。')
by_name(s, '図 16')._element.getparent().remove(by_name(s, '図 16')._element)
y = B.section_bar(s, 0.37, 1.51, 5.55, '譲渡対象の整理（帳簿上）', h=0.30, size=11)
B.house_table(s, 0.37, y + 0.04, 5.55, [1.55, 2.55, 1.45], [
    ['項目', '内容', '金額'],
    ['内装・設備等（9点）', '恵比寿店内装・間仕切・造作・什器等', '簿価 9円'],
    ['　取得原価', '2013年〜2022年取得', '17,190千円'],
    ['　参考：再計算簿価', '通常の耐用年数で償却した場合', '12,503千円'],
    ['店舗不動産', '所有なし（2店舗とも賃借）', '－'],
    ['敷金・保証金', '差入れなし', '－'],
    ['借入金・その他債務', '承継なし（引継債務は支払家賃のみ）', '－'],
    ['従業員', '転籍なし', '－'],
], row_h=0.27, head_h=0.27, size=8.5)
y2 = 4.15
y2 = B.section_bar(s, 0.37, y2, 5.55, '実質的に引き継ぐもの（＝買手が買うもの）', h=0.30, size=11)
B.house_table(s, 0.37, y2 + 0.04, 5.55, [1.35, 4.2], [
    ['項目', '内容'],
    ['利用者契約', '22契約・15.0室分（月額3,529千円／年換算42,346千円）'],
    ['転貸借契約', 'BG社・売手・財成社の三者で転借人の地位を承継'],
    ['集客導線', '公式サイト（301リダイレクトで順位維持）・Instagram・meta広告'],
    ['運営ノウハウ', '業務フロー、清掃・リネン・スマートロック等の外注先、公式LINE運用'],
    ['引継ぎ支援', '売手担当者が2〜3ヶ月程度を想定（問合せ〜契約・入退去）'],
], row_h=0.36, head_h=0.27, size=8.5)
parts(by_name(s, 'テキスト ボックス 6'), [
    ('≪資産≫', True),
    ('・内装・設備は全額特別損失処理済で簿価9円。帳簿上の純資産で価格を説明することはできない', False),
    ('≪負債≫', True),
    ('・引き継ぐ債務は支払家賃のみ（DDで確認）。借入金・保証債務なし', False),
    ('≪不動産≫', True),
    ('・2店舗とも賃借で敷金の差入れなし。買手の初期投資は譲渡対価のみ', False),
], size=9.5)
parts(by_name(s, 'テキスト ボックス 7'), [
    ('【時価修正】', True),
    ('≪内装・設備≫', True),
    ('・再計算簿価12,503千円は参考値。賃借物件の内装で換金性は乏しく、時価は簿価同様ほぼゼロと見るのが妥当', False),
    ('≪敷金・保証金≫', True),
    ('・なし', False),
    ('', False),
    ('想定時価純資産　　　≒0千円', True),
    ('想定ネットキャッシュ　対象外（現預金・借入金は承継せず）', True),
], size=9.5)

# ---- 8. 譲受企業
s = S[7]
B.set_lead(s, '要確認：財成社の特徴を一文で（業種・規模・本件を買った理由）')
for nm in ['正方形/長方形 82', '正方形/長方形 88', '正方形/長方形 76', '正方形/長方形 92', '正方形/長方形 80']:
    put(s, nm, '要確認', 10)
put(s, '正方形/長方形 96', '要確認\n（DDでは賃貸借の継続・区画別収支・設備更新・承継手続を重点確認しており、'
                          '運営事業というより「収益を生む資産」として利回り目線で検討していたと推察）', 9)
put(s, '正方形/長方形 2', '財成社ロゴ（要差替）', 10)
for nm in ['正方形/長方形 3', '正方形/長方形 5', '正方形/長方形 6', '正方形/長方形 7']:
    put(s, nm, '写真（要差替）', 10)

# ---- 10. スケジュール
s = S[9]
B.set_lead(s, '1社目（同業のサロン運営会社）はトップ面談・意向表明書ドラフトまで進むも成約に至らず（経緯要確認）。'
              '利回りで判断する財成社と意向表明→初期DD→成約。')
events = {  # group name: (date, label)
    'グループ化 38': ('26/4/2', 'アドバイ\nザリー契約'),
    'グループ化 39': ('26/5/26', '企業概要書完成\n（打診開始）'),
    'グループ化 40': ('26/6/24', '1社目\nトップ面談'),
    'グループ化 41': ('26/9上旬', '財成社\n意向表明（要確認）'),
    'グループ化 42': ('26/9/12', '初期DD\n（資料・Q&A）'),
    'グループ化 43': ('要確認', 'クロージング'),
}
for g in s.shapes:
    if g.name in events:
        d, lab = events[g.name]
        for ch in g.shapes:
            if not ch.has_text_frame or not ch.text_frame.text.strip():
                continue
            t = ch.text_frame.text
            if t[:2] in ('25', '26') or '/' in t[:6]:
                fill(ch, d)
            else:
                fill(ch, lab, size=9)

# ---- 12. 論点と具体施策（行を高くして3段）
s = S[11]
B.set_lead(s, '資産性ゼロ・賃借2店舗の事業譲渡を、「契約キャッシュフローを買う投資案件」として組み立て直したことが成約の決め手。')
issues = [
    ('【論点①：案件として成立するか】\n不動産は所有せず2店舗とも賃借、内装は簿価9円、従業員の転籍もなし。'
     '買手が「何を買うのか」が見えにくく、事業譲渡として成立するかが最初の論点。',
     '「モノ」ではなく「継続する契約キャッシュフロー」を売る整理に転換。22契約・月額353万円の契約一覧、'
     '平均利用22.9ヶ月・最低保証15ヶ月、満室シミュレーションを開示し、収益が続くことを可視化。'),
    ('【論点②：価格の見せ方】\n時価純資産≒0のため、通常の「時価純資産＋営業権」の評価では価格のほぼ全額がのれんになり、'
     '5,000万円超の根拠を説明しにくい。',
     f'評価額ではなく「投資利回り」で提示。EBITDA{EBITDA:,}千円÷譲渡価格5,250万円＝利回り約{y_a*100:.0f}%・'
     f'回収約{pb_a:.1f}年。運営委託ありの3,750万円案でも利回り約{y_b*100:.0f}%と、2案とも同水準に設計（次頁）。'),
    ('【論点③：賃借物件・利用者・集客導線の承継】\n店舗はBG社からの転借、利用者22契約は個別契約。'
     '承継できなければ収益そのものが消える。',
     'BG社・売手・財成社の三者で転借人の地位承継契約を締結。利用者とは地位承継の覚書、公式サイトは301リダイレクトで'
     '検索順位を維持。売手担当者が2〜3ヶ月の引継ぎを支援。'),
]
ROW_TOP, ROW_H, GAP = 1.92, 1.58, 0.10
left = ['正方形/長方形 7', '正方形/長方形 9', '正方形/長方形 10']
right = ['正方形/長方形 16', '正方形/長方形 18', '正方形/長方形 20']
tris = ['二等辺三角形 22', '二等辺三角形 24', '二等辺三角形 26']
for i, (a, b) in enumerate(issues):
    top = Inches(ROW_TOP + i * (ROW_H + GAP))
    for nm, txt in ((left[i], a), (right[i], b)):
        sh = by_name(s, nm)
        sh.top, sh.height = top, Inches(ROW_H)
        fill(sh, txt, size=11.5)
        sh.text_frame.word_wrap = True
        sh.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    # 論点の見出し行だけ太字
    p0 = by_name(s, left[i]).text_frame.paragraphs[0]
    for r in p0.runs:
        r.font.bold = True
    t = by_name(s, tris[i])
    t.top = top + Inches(ROW_H / 2) - t.height // 2

# ---- 13. 新規頁：価格の見せ方（利回り）
s = B.duplicate_slide(prs, 11, after=11)
B.clear_body(s)
B.set_title(s, '価格の見せ方：「評価額」ではなく「投資利回り」で提示')
B.set_lead(s, '時価純資産≒0の事業を「いくらで買えば年何％・何年で回収できるか」で提示。2案とも利回り約30%・回収3年強に'
              '揃え、買手が初期投資額を選べる設計にした。')
# 上段：通常の見せ方 → 本件の見せ方
bx = B.rect(s, 0.37, 1.62, 4.95, 1.18, fill=B.GREY)
B.rich(s, 0.50, 1.70, 4.70, 1.05, [
    ('通常の見せ方：株価 ＝ 時価純資産 ＋ 営業権', {'size': 11, 'bold': True, 'color': B.NAVY}),
    ('本件は時価純資産≒0（簿価9円・不動産なし・敷金なし）。価格のほぼ全額が営業権となり、'
     '「EBITDA×年数」の説明だけでは割高に見えやすい。', {'size': 10.5, 'newpara': True}),
])
arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.42), Inches(2.00), Inches(0.42), Inches(0.42))
arr.fill.solid(); arr.fill.fore_color.rgb = B.BLUE; arr.line.fill.background()
B.rect(s, 5.94, 1.62, 5.40, 1.18, fill=B.PALE, line=B.BLUE, lw=1.0)
B.rich(s, 6.07, 1.70, 5.15, 1.05, [
    ('本件の見せ方：投資額に対する年間リターン', {'size': 11, 'bold': True, 'color': B.NAVY}),
    ('利回り＝年間CF÷譲渡価格、回収年数＝譲渡価格÷年間CF。不動産・金融商品と同じ物差しで比べられ、'
     '投資目線の買手に価格の妥当性が一目で伝わる。', {'size': 10.5, 'newpara': True}),
])
# 中段：2案の比較表
y = B.section_bar(s, 0.37, 3.00, 6.75, '提示した2案の利回り試算', h=0.30, size=11)
B.house_table(s, 0.37, y + 0.04, 6.75, [2.95, 1.9, 1.9], [
    ['（単位：千円）', 'A案：運営委託なし', 'B案：運営委託あり'],
    ['譲渡価格（基準額）', f'{PRICE_A:,}', f'{PRICE_B:,}'],
    ['年間EBITDA（2026年2月期実績）', f'{EBITDA:,}', f'{EBITDA:,}'],
    ['運営委託費（年・試算）', '－', f'△{OUTSOURCE_Y:,}'],
    ['買手の年間CF（実績ベース）', f'{cf_a:,}', f'{cf_b:,}'],
    ['投資利回り', pct(y_a), pct(y_b)],
    ['投資回収年数', f'約{pb_a:.1f}年', f'約{pb_b:.1f}年'],
    [f'参考：進行期見通し（営業利益{OP_FCST:,}）の利回り', pct(yf_a), pct(yf_b)],
    ['売手の受取総額（手残り＋委託収入）', '37,500', '37,500'],
], row_h=0.30, head_h=0.30, size=9, bold_rows=(5, 6),
    aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT],
    fills={('row', 5): B.PALE, ('row', 6): B.PALE})
B.textbox(s, 0.37, 6.38, 6.75, 0.80,
          '※利回り＝年間CF÷譲渡価格。仲介手数料・税金・A案の買手側運営人件費は考慮前の簡易試算。運営委託費は「3年で約1,500万円」を'
          '年500万円に均等按分。売手の受取総額はA案＝手残り37,500、B案＝手残り22,500＋委託収入15,000（基準額＝手残り下限＋当社フィー15,000）。'
          '実際に買手へ提示した数値と最終採用案は要確認。出所：企業概要書260526・論点整理シート260623。',
          size=7.5, color=B.NOTE)
# 右：Comment
B.comment_box(s, 7.32, 3.00, 4.02, 4.15, [
    ('≪設計のポイント≫', {'size': 10, 'bold': True}),
    ('・2案とも利回り約30%に揃え、「初期投資を抑えて委託するか、自ら運営して手取りを厚くするか」を選ぶだけの比較にした', {'size': 9.5, 'newpara': True}),
    ('・売手の受取総額は両案とも3,750万円で同額', {'size': 9.5, 'newpara': True}),
    ('≪利回りを支える材料≫', {'size': 10, 'bold': True, 'newpara': True}),
    ('・最低保証15ヶ月・平均利用22.9ヶ月で、回収期間中の解約リスクは限定的', {'size': 9.5, 'newpara': True}),
    (f'・満室時の年間売上{FULL_REV:,}千円（現状契約ベース比＋{FULL_REV - CUR_REV:,}千円）がアップサイド', {'size': 9.5, 'newpara': True}),
    ('・運営工数は週12〜15時間程度と軽く、本業の傍らで保有できる', {'size': 9.5, 'newpara': True}),
    ('≪税務（補足）≫', {'size': 10, 'bold': True, 'newpara': True}),
    ('・事業譲渡のため、法人の買手は対価の大半を資産調整勘定（5年均等償却）として損金算入できると想定（要税理士確認）', {'size': 9.5, 'newpara': True}),
], title='Comment', title_size=14)

# ---- 15. 学び（行を高くして3段）
s = S[14]
B.set_lead(s, '「資産がない事業は売れない」ではなく、「続く収益」と「買手の物差し」で見せれば成約できる。')
learn = [
    '資産性がなくても事業は売れる。判断軸は「資産」ではなく、収益の継続性・再現性・承継可能性。賃借・簿価ゼロの事業でも、'
    '契約一覧と稼働データで「続く収益」を示せば案件として成立する。',
    '価格は「評価額」ではなく「利回り・回収年数」で語る。純資産で説明できない案件ほど、買手の投資判断の物差しに合わせた'
    '見せ方が効く。同業だけでなく投資目線の買手まで候補を広げられる。',
    '承継の選択肢で買手の不安と負担を先回りして解く。運営委託あり／なしの2案（売手の受取総額は同額）、転貸借の三者地位承継、'
    '売手担当者の引継ぎ支援で、「運営できるか」「契約が続くか」の不安を解消した。',
]
heads = ['正方形/長方形 8', '正方形/長方形 13', '正方形/長方形 6']
bodies = ['正方形/長方形 10', '正方形/長方形 15', '正方形/長方形 17']
for i, txt in enumerate(learn):
    top = Inches(ROW_TOP + i * (ROW_H + GAP))
    for nm in (heads[i], bodies[i]):
        sh = by_name(s, nm)
        sh.top, sh.height = top, Inches(ROW_H)
    b = by_name(s, bodies[i])
    fill(b, txt, size=12)
    b.text_frame.word_wrap = True
    b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill(by_name(s, heads[i]), f'学び{"①②③"[i]}\n' + ['続く収益を売る', '利回りで語る', '承継の選択肢'][i],
         size=11, bold=True, align=PP_ALIGN.CENTER)
    by_name(s, heads[i]).text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

B.force_white_headers(prs)
prs.save(OUT)
print('saved', OUT, len(prs.slides), 'slides')
print(f'利回り A {pct(y_a)} / B {pct(y_b)}  回収 A {pb_a:.2f} / B {pb_b:.2f}  3期平均 {avg3}')
