"""鶴岡中央青果 株式価値評価書（買手候補先様向け）を当社テンプレートから作る。

目的：売手向け（3月末基準・保守的。株価を握るため）とは逆に、希望譲渡価格（1.2億円）が割安であることを
買手に示す。前提は検証可能なものに限り、すべて頁内に明記する（依頼者指示 2026/11/17）。
  ①基準日は進行期末（2026/9末試算表）。上期の法人税等は見込計上
  ②土地は事業用不動産として簿価評価（固定資産税評価額÷0.7の推計は比較頁で参考表示）
  ③正常収益力は直近12か月（LTM）。調整項目は概要書と同じ（調整後EBITDAは概要書と一致）
  ④営業権1〜3年、基準は税考慮前の修正後営業利益（LTM）。評価差額の税効果は認識しない
  ⑤マーケットアプローチはネットデットが大きく株価が出ないため不採用（依頼者修正 2026/11/20）
数値はすべて案件マスターの「買手評価」シート（と既存のVR→・財務シート）を貼る。
売手向けの「株価評価にあたっての注意点」（売主の期待値調整用）は買手向けには載せない。

使い方: python build_vr_buyer.py <案件マスター.xlsx> <出力.pptx>
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '../../.claude/skills/ma-deal-docs/scripts'))
from openpyxl import load_workbook  # noqa: E402
import im_paste as ip  # noqa: E402
from im_builder import delete_slide  # noqa: E402
from vr_builder import (open_vr_template, strip_red_banners, fix_broken_pictures,  # noqa: E402
                        vr_paste, vr_text, set_range_chart, insert_detail_pages, drop_tables)

MASTER, OUT = sys.argv[1], sys.argv[2]
DATE = '2026年11月●日'          # 依頼者指定（提出日に確定）
wb = load_workbook(MASTER, data_only=True)
B = '買手評価'
v = lambda sh, a: wb[sh][a].value
b = lambda a: v(B, a)
m = lambda x: f'{x / 1e6:,.1f}'
K = lambda x: f'{x / 1000:,.0f}' if x >= 0 else f'△{-x / 1000:,.0f}'
import copy  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

prs = open_vr_template()
strip_red_banners(prs)
fix_broken_pictures(prs)
insert_detail_pages(prs, ['（参考）正常収益力［直近12か月｜2025/10~2026/9］'], after_title='修正PLの算定［正常収益力の集計］')
insert_detail_pages(prs, ['希望譲渡価格の水準'], after_title='年買法（コストアプローチ）による算出')


def title_of(s):
    return next((sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.name.startswith('タイトル')
                 and sh.top < 228600 and sh.text_frame.text.strip()), '')   # 頁タイトル（y<0.25in）。右上の章名は除く


# 売手向けの期待値調整頁・3月末基準の修正BS頁・不採用としたマルチプル法の頁は使わない
for t in ('株価評価にあたっての注意点', 'BS補足［資産］', 'BS補足［負債・純資産］', '修正BS［負債・純資産］',
          'EV/EBITDAマルチプル法（マーケットアプローチ）による算出'):
    i = next(i for i, s in enumerate(prs.slides) if title_of(s) == t)
    delete_slide(prs, i)
SL = list(prs.slides)
T = {title_of(s): s for s in SL}
sid = ip.shape_by_id
paste = lambda s, ref, x, y, w, target, fs=None: vr_paste(s, wb, ref, x, y, w, target, fs)


def remove(s, *ids):
    for i in ids:
        sh = sid(s, i)
        sh._element.getparent().remove(sh._element)


# ------------------------------------------------------------------ 表紙
for sh in SL[0].shapes:
    if sh.has_text_frame:
        t = sh.text_frame.text
        if '御中' in t:
            ip.set_text(sh, '株式会社鶴岡中央青果')
        elif '年' in t and '月' in t and len(t) < 20:
            ip.set_text(sh, DATE)
for sh in SL[1].shapes:
    if sh.has_table and len(sh.table.rows) == 2:
        for c, t in enumerate(['株式会社M&A承継機構', '田垣内　啓介', '080-4737-4326', 'tagaito_keisuke@ma-sg.co.jp']):
            ip.set_text(sh.table.cell(1, c), t)

# ------------------------------------------------------------------ 採用するアプローチの選定（マーケットを不採用に）
s = T['採用するアプローチの選定']
remove(s, 62)                                   # 類似会社比準法の□に付いたチェック
grp = sid(s, 84)._element                       # ○採用／✕不採用の図形グループ
kid = lambda i: next(g for g in grp.iter(qn('p:grpSp')) if g.find('.//' + qn('p:cNvPr')).get('id') == str(i))
mkt, ng = kid(81), kid(21)                      # 81＝マーケットの「採用」、21＝インカムの「不採用」
new = copy.deepcopy(ng)
y_mkt = mkt.find(qn('p:grpSpPr')).find(qn('a:xfrm')).find(qn('a:off')).get('y')
new.find(qn('p:grpSpPr')).find(qn('a:xfrm')).find(qn('a:off')).set('y', y_mkt)
for i, el in enumerate(new.iter(qn('p:cNvPr'))):
    el.set('id', str(900 + i))
mkt.getparent().replace(mkt, new)
cell = next(sh for sh in s.shapes if sh.has_table).table.cell(2, 4)
paras = cell.text_frame.paragraphs
for p_ in paras[2:]:
    p_._p.getparent().remove(p_._p)
runs = paras[1].runs
runs[0].text = (f'本件では、市場用地・建物等の事業用資産に対応する長期借入金（ネットデット{m(-b("C21"))}百万円）を'
                '控除すると株式価値が算定されないため、不採用とした。')
for r_ in runs[1:]:
    r_._r.getparent().remove(r_._r)

# ------------------------------------------------------------------ 株式価値の試算結果
s = T['株式価値の試算結果']
t = sid(s, 15).table
for r, txt in enumerate(['株式会社鶴岡中央青果　普通株式4,000株（発行済株式の100%）', DATE,
                         '2026年9月30日（進行期の試算表。上期の法人税等は見込計上）',
                         '2025年10月1日～2026年9月30日（直近12か月。2026年3月期決算と2026年4〜9月試算表から算定）',
                         '買手候補先様における譲渡価額検討の参考資料とすることを目的に本株式価値試算報告書は作成されております。']):
    ip.set_text(t.cell(r, 1), txt)
ip.set_text(sid(s, 30), f'{K(b("L6"))}　～　{K(b("N6"))}（千円）')
band = copy.deepcopy(sid(s, 6)._element)        # 「※※※不採用※※※」帯をマーケットの行に複製
s.shapes._spTree.append(band)
remove(s, 4, 3)                                 # マーケットの想定レンジ枠
nb = s.shapes[-1]
nb.element.find('.//' + qn('p:cNvPr')).set('id', '950')
ip.place(nb, 0.49, 5.45, 10.83, 0.56)
set_range_chart(s, ranges={'コスト': (b('L6'), b('N6'))})
gf = next(sh for sh in s.shapes if sh.has_chart)
ip.place(gf, 6.9, 3.95, 4.4, 1.45)
ip.place(sid(s, 34), 10.32, 3.78)
ip.note(s, 0.49, 7.07, 10.71, [
    f'※希望譲渡価格{m(b("C25"))}百万円は、年買法の下限（営業権1年：{m(b("L6"))}百万円）を下回り、時価純資産{m(b("C26"))}百万円に'
    f'営業権（直近12か月の修正後営業利益）約{b("C29"):.1f}年分を加えた水準に相当する。',
    f'※マーケットアプローチ（EV/EBITDAマルチプル法）は、市場用地・建物等の事業用資産に対応する長期借入金（ネットデット{m(-b("C21"))}百万円）を'
    '控除すると株式価値が算定されないため不採用とし、事業用不動産を自社保有する対象会社の価値は時価純資産を基礎とする年買法で評価した。'], size=7.5)

# ------------------------------------------------------------------ 過去BSの推移（3期＋2026年9月末試算表）
s = T['過去BSの推移']
drop_tables(s)
paste(s, "'BS(借方)'!B3:G61", 0.37, 1.0, 5.47, 6.35)
_, yb = paste(s, "'BS (貸方)'!B3:G25", 5.85, 1.0, 5.47, 2.75)
paste(s, "'BS (貸方)'!B34:G51", 5.85, yb + 0.25, 5.47, 7.35 - yb - 0.25)

# ------------------------------------------------------------------ 時価純資産の算定［2026年9月末］（修正BS［資産］の枠）
s = T['修正BS［資産］']
ip.set_text(sid(s, 4), '時価純資産の算定［2026年9月末］')
drop_tables(s)
paste(s, f'{B}!B3:C10', 0.37, 1.2, 5.85, 3.2, fs=0.85)
vr_text(sid(s, 9), [
    f'基準日：進行期末（2026年9月末）の試算表。上期の利益計上により簿価純資産は{m(b("C4"))}百万円（2026年3月末{m(v("BS (貸方)", "F50"))}百万円）。'
    f'上期の法人税等は税引前利益×34%で見込計上（{K(b("C9"))}千円）。',
    '土地：本社・市場、パッケージ場の事業用不動産で売却予定がないため簿価評価。',
    '投資有価証券：上場株式（フィデアHD 3,000株）を時価評価（9月末時点の保有株式数は要確認）。',
    '貸倒引当金は個別に回収可能性を判断し戻入。電話加入権は換金価値なしとしてゼロ評価。',
    f'✓ 時価純資産 {K(b("C10"))}千円'], size=9.5)
ip.place(sid(s, 9), 6.6, 1.75, 4.5, 4.9)

# ------------------------------------------------------------------ ネットキャッシュの算定（2026年9月末）
s = T['ネットキャッシュの算定']
drop_tables(s)
paste(s, f'{B}!B16:F21', 0.37, 1.1, 10.34, 2.2, fs=0.85)
vr_text(sid(s, 8), [
    '基準日（2026年9月末）の試算表を参照。長期借入金は8〜9月の新規借入を含む（使途は繁忙期の運転資金と想定）。',
    '保証金（仕入代金のデポジット33,303千円。取引終了時に返金）と預り保証金8,980千円は事業上の預託としてネットキャッシュに含めていない。'], size=9.5)

# ------------------------------------------------------------------ 過去PL・販管費の修正・修正PL（概要書と同じ調整）
s = T['過去PLの推移']
drop_tables(s)
paste(s, 'PL!B2:L56', 0.37, 0.92, 10.83, 6.45)
s = T['販管費及び一般管理費の修正']
drop_tables(s)
paste(s, '修正SGA!B3:N39', 0.37, 0.78, 11.0, 5.45)
vr_text(sid(s, 8), ['役員報酬・法定福利費：譲渡時に退任する役員（代表者・配偶者・監査役・退任済みの常務）の報酬を控除（社保15%連動）し、'
                    '代表者の退任後の顧問料（年2,160千円）を計上。新体制での役員報酬は買手の方針により変動する。'], size=8.5)
vr_text(sid(s, 7), ['賞与：規定はない。業績連動で支給義務がないため修正しない。',
                    '寄付金：譲渡後に発生しない費用として控除。',
                    '受取奨励金：経常的な事業関連収益として営業利益に取込み。'], size=8.5)
s = T['修正PLの算定［正常収益力の集計］']
drop_tables(s)
paste(s, '修正PL!B3:N56', 0.37, 0.82, 10.95, 5.55)
vr_text(sid(s, 13), [
    '調整項目は企業概要書の財務ハイライト（P/L）と同一（調整後EBITDA 2026/3期27,842千円）。',
    '営業外損益は受取利息・受取配当金のみ残し、受取奨励金・出荷奨励金は販管費へ振替。受取延滞料・雑収入・雑損失は全額控除、支払利息は資金調達構造に影響されない収益力とするため控除。'
    '雑収入には仲卸業者等からのゴミ処理料・冷蔵庫保管料等が含まれ、経常部分は上振れ余地となる。',
    '特別損益（前期損益修正損20,900千円・固定資産売却損益・貸倒引当金繰入／戻入）は非経常的なため全額控除。'], size=8.5)

# ------------------------------------------------------------------ （参考）正常収益力［LTM］（追加頁）
s = T['（参考）正常収益力［直近12か月｜2025/10~2026/9］']
_, yb = paste(s, f'{B}!E3:I8', 0.49, 1.05, 10.71, 2.5, fs=0.85)
_, yb = paste(s, f'{B}!E13:I13', 0.49, yb + 0.12, 10.71, 0.36, fs=0.85)
ip.place(sid(s, 19), 0.49, yb + 0.35, 10.71, 1.9)
ip.place(sid(s, 8), 1.95, yb + 0.47, 9.1, 1.7)
vr_text(sid(s, 8), [
    f'・進行期（2026年4〜9月）は人件費・運賃の削減により営業利益{m(b("H4"))}百万円と前年同期（{m(b("G4"))}百万円）の約{b("H4") / b("G4"):.1f}倍に改善しており、'
    '現在の収益力は直近12か月（LTM）に最もよく表れる。',
    '・LTM＝2026年3月期＋2026年4〜9月（試算表）−2025年4〜9月（前年同期）。前年同期の修正額は2026/3期の修正額の6/12で推計（役員報酬等は月額一定と想定）。',
    f'・LTMの修正後営業利益{K(b("I10"))}千円を年買法の基準営業利益とする（修正後EBITDA {K(b("I13"))}千円）。'], size=9.5)

# ------------------------------------------------------------------ 年買法
s = T['年買法（コストアプローチ）による算出']
remove(s, 6, 24, 26, 20, 21)
paste(s, f'{B}!B3:C10', 0.49, 1.25, 3.48, 2.9, fs=0.8)
paste(s, f'{B}!E3:I11', 4.59, 1.25, 6.72, 3.2, fs=0.8)
paste(s, f'{B}!K3:N6', 0.49, 5.0, 6.78, 1.6, fs=0.85)
ip.set_text(sid(s, 4), '時価純資産（2026年9月末）')
ip.set_text(sid(s, 11), '基準営業利益（LTM）')
ip.set_text(sid(s, 8), '・営業権は直近12か月の修正後営業利益（LTM）を基準値として算出。')

# ------------------------------------------------------------------ 希望譲渡価格の水準（追加頁）
s = T['希望譲渡価格の水準']
ip.set_text(sid(s, 3), '3. Valuation Details')
_, yb = paste(s, f'{B}!B24:C33', 0.49, 1.05, 10.71, 3.5, fs=0.9)
ip.place(sid(s, 19), 0.49, yb + 0.3, 10.71, 2.1)
ip.place(sid(s, 8), 1.95, yb + 0.42, 9.1, 1.9)
vr_text(sid(s, 8), [
    f'希望譲渡価格{m(b("C25"))}百万円のうち、時価純資産（{m(b("C26"))}百万円）を超える部分（営業権相当）は{m(b("C27"))}百万円で、'
    f'直近12か月の修正後営業利益（{m(b("C28"))}百万円）の約{b("C29"):.1f}年分にとどまる。',
    f'年買法の中央値（営業権2年：{m(b("C30"))}百万円）に対し、希望譲渡価格は約{m(b("C31"))}百万円低い水準にある。',
    f'参考：土地を固定資産税評価額÷0.7で評価した場合でも時価純資産は{m(b("C32"))}百万円で、'
    f'希望譲渡価格は営業権約{b("C33"):.1f}年分に相当し、国内中小企業M&Aのボリュームゾーン（1〜3年）の範囲内となる。'], size=10)

from vr_builder import force_white_headers  # noqa: E402
force_white_headers(prs)
prs.save(OUT)
ip.scrub_doc_props(OUT, title='株式価値評価書（買手候補先様向け）')
print('saved', OUT, len(prs.slides))
