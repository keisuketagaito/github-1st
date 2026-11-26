# -*- coding: utf-8 -*-
"""企業概要書（簡易版）：ティーザーと企業概要書の中間に位置づける13頁の資料。

当社テンプレート im_template.pptx を複製し、表は案件マスターの貼付範囲を paste_range で貼る。
リード文・Commentの数字もマスター（PLハイライト・BSハイライト・PL・CR・SGA・参考数値）から読む。

出典の扱い
  - 決算書（2022〜2025年12月期）は事実として記載
  - 売主面談（2026年11月）は「売主説明」と明示。面談メモにはつばめグループ全体の話が混在しうるため、
    対象会社単体と確認できないもの（車両台数・進行期見込み・関係会社等）は「確認中」と書く
  - 外部調査報告書は社外極秘のため直接引用しない（登記・公表情報・売主説明で裏付く事実のみ）
  - 内部限（希望株価・他社仲介・契約方針・焦げ付き先の事情・解雇経緯）は載せない（check_confidential で検査）
  - 時価純資産・含み損益（未収入金の回収不能見込み額、保険の解約返戻金額）は評価書で扱い、ここでは金額を出さない
"""
import copy
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.path.join(HERE, '..', '.claude', 'skills', 'ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
import im_builder as ib  # noqa: E402
from im_paste import (comment, fill_comment, note, paste_range, place, range_height, remove,  # noqa: E402
                      remove_null_pics, remove_watermark, remove_yellow, scrub_doc_props, set_lead,
                      set_text, shape_by_id)
from openpyxl import load_workbook  # noqa: E402
from pptx.util import Inches  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

OUT = os.path.join(HERE, '企業概要書_簡易版_ウェルカム・バスケット_261126.pptx')
wb = load_workbook(os.path.join(HERE, '案件マスター_ウェルカム・バスケット.xlsx'), data_only=True)
v = lambda sh, a: wb[sh][a].value
M = lambda x: f'{x / 1e6:,.0f}'                 # 百万円
K = lambda x: f'{x / 1e3:,.0f}'                 # 千円
P = lambda x: f'{x * 100:.1f}%'
PH = lambda a: v('PLハイライト', a)
BS = lambda a: v('BSハイライト', a)
SAN = lambda a: v('参考数値', a)


def fit(ref, target, hmax=1.0):
    return min(hmax, target / range_height(wb, ref))


def paste(s, ref, x, y, w, fs, target=None):
    hs = fit(ref, target) if target else 1.0
    paste_range(s, wb, ref, x, y, w, hscale=hs, fscale=fs)
    return y + range_height(wb, ref, hscale=hs)


def title(s, text):
    for sh in s.shapes:
        if sh.is_placeholder and sh.placeholder_format.idx == 0:
            set_text(sh, text)
            return


def section(s, text):
    for sh in s.shapes:
        if sh.name == 'タイトル 1' and not sh.is_placeholder:
            set_text(sh, text)


def drop_pics(s):
    for sh in list(s.shapes):
        if sh.shape_type == 13:
            sh._element.getparent().remove(sh._element)


# ================================================================ テンプレート
prs = ib.open_template()
ib.keep_slides(prs, [1, 2, 5, 6, 7, 8, 10, 12, 18, 20, 34, 35, 36])
ib.strip_red_banners(prs)
ib.fix_broken_pictures(prs)
remove_yellow(prs); remove_null_pics(prs); remove_watermark(prs)
S = list(prs.slides)
(S_COV, S_INTRO, S_BASIC, S_PL, S_BS, S_COND, S_BIZ, S_FEAT, S_SH, S_SITE, S_TODO, S_FEE, S_END) = S
CM_HEAD, CM_BODY = shape_by_id(S_PL, 8), shape_by_id(S_PL, 9)     # Comment枠の見本（他頁へ複製する）

# 期の表記
Y = [v('BS(借方)', c + '4') for c in 'CDEF']                       # 2022年12月期 …
sales = [PH(c + '4') for c in 'CDEF']
gp = [PH(c + '6') for c in 'CDEF']
op = [PH(c + '8') for c in 'CDEF']
dep = [PH(c + '15') for c in 'CDEF']
ebitda = [PH(c + '37') for c in 'CDEF']
temp_staff = [v('CR', c + '18') for c in 'CEGI']                   # 臨時傭員費
outsrc = [v('CR', c + '27') for c in 'CEGI']                       # その他外注費
referral = [v('CR', c + '40') for c in 'CEGI']                     # 紹介手数料
sga = [v('PL', c + '18') for c in 'CEGI']
officer = [v('SGA', c + '6') for c in 'CEGI']
ad = [v('SGA', c + '21') for c in 'CEGI']                          # 宣伝広告費
op_fc = SAN('C5')                                                  # 進行期の営業利益見込み（売主説明）
staff, trucks = SAN('C6'), SAN('C7')
debt = v('BS (貸方)', 'F6') + v('BS (貸方)', 'F7') + v('BS (貸方)', 'F19')
debt_prev = v('BS (貸方)', 'E6') + v('BS (貸方)', 'E7') + v('BS (貸方)', 'E19')
assert v('SGA', 'B21') == '宣伝広告費' and v('CR', 'B40') == '紹介手数料' and v('CR', 'B18') == '臨時傭員費'

# ================================================================ 1 表紙
ib.set_cover_title(S_COV, '引越業', '東海地区')           # 「企業概要書 <業種／地区>」の2ラン構成（1行に収める）
for sh in S_COV.shapes:
    if not sh.is_placeholder and sh.has_text_frame and 'Information' in sh.text_frame.text:
        set_text(sh, 'Information Memorandum（簡易版）')

# ================================================================ 2 はじめに（免責文はテンプレートのまま。照会先の担当者は●で仮置き）
for sh in S_INTRO.shapes:
    if sh.has_table:
        t = sh.table
        for c, txt in zip((1, 2, 3), ('●●　●●', '0●0-●●●●-●●●●', '●●_●●@ma-sg.co.jp')):
            ib_cell = t.cell(1, c)
            ib_cell.text_frame.paragraphs[0].runs[0].text = txt
            for r in ib_cell.text_frame.paragraphs[0].runs[1:]:
                r._r.getparent().remove(r._r)

# ================================================================ 3 基本情報
s = S_BASIC
drop_pics(s)
remove(s, 3, 4, 6, 7, 8)                                          # ロゴ・写真の仮枠（写真は受領後に掲載）
set_lead(s, f'東海地方で業歴20年超の引越専業会社。個人向けを主力に法人・医療機関の移転、電気工事等の付帯サービスまで提供し、'
            f'{Y[-1]}は売上高{M(sales[-1])}百万円・営業黒字へ回復。')
yb = paste(s, '会社概要!B2:C15', 0.34, 2.06, 6.41, 0.83, target=5.0)
comment(s, CM_HEAD, CM_BODY, 6.95, 1.66, 4.40, yb - 1.66, [
    '≪地域で選ばれる引越会社≫',
    '・オリコン顧客満足度調査（引越し会社）の地域別ランキングで第1位（8年連続と売主説明。部門・年数は公表資料で確認中）。',
    '≪紹介と直接受注の両輪≫',
    f'・不動産仲介・住宅関連企業等からの紹介受注に加え、Web等の直接受注。紹介手数料は年{M(referral[-1])}百万円（{Y[-1]}）。',
    '≪引越事業への集中と収益改善≫',
    f'・家電FC店・一般貨物運送から撤退し引越事業に一本化。外注・臨時人員と販管費を圧縮し、{Y[-1]}は営業利益{M(op[-1])}百万円。',
    '≪承継の受け皿≫',
    '・現場運営を担う取締役1名が譲渡後も継続勤務の意向。従業員は約70名（売主説明）。',
    '≪論点（確認中）≫',
    '・未収入金の一部（取引先の倒産分）の処理、車両の台数・保有形態、関係会社（福祉事業）との関係、役員報酬の内訳。',
])
place(shape_by_id(s, 2), 0.34, 1.66)

# ================================================================ 4 財務ハイライト（P/L）
s = S_PL
drop_pics(s)
set_lead(s, f'売上高はピーク時の{M(sales[0])}百万円から{M(sales[-1])}百万円へ縮小したが、{Y[-1]}は粗利率{P(gp[-1] / sales[-1])}へ改善し'
            f'EBITDA{M(ebitda[-1])}百万円。進行期は営業利益約{M(op_fc)}百万円の見込み（売主説明）。')
yb = paste(s, 'PLハイライト!B2:G37', 0.34, 1.6, 6.62, 0.875, target=4.9)
note(s, 0.34, yb + 0.12, 6.62, [
    '※EBITDA＝営業利益＋減価償却費（運送原価・販管費の合計）で、正常化調整前。役員報酬等の調整は内訳の受領後に反映（確認中）。',
    '※進行期（2026年12月期）は試算表未受領。営業利益の見込みは売主説明で、対象会社単体の数値かを含め確認中。'], size=7.5)
place(shape_by_id(s, 8), 7.06, 1.6, 4.29, 3.95)
place(shape_by_id(s, 9), 7.06, 2.18, 4.28, 3.3)
place(shape_by_id(s, 2), 7.06, 5.65, 4.29, 1.55)
place(shape_by_id(s, 10), 7.07, 5.72, 4.28, 1.45)
fill_comment(shape_by_id(s, 9), [
    '≪売上高≫',
    f'・{M(sales[0])}→{M(sales[1])}→{M(sales[2])}→{M(sales[3])}百万円。家電FC店・一般貨物運送からの撤退と、'
    '法人の個別訪問営業を縮小しWeb受注へ移行したことによる減少（移行後の受注は伸び悩み）。',
    '≪売上総利益≫',
    f'・{Y[-1]}は減収ながら{M(gp[-1])}百万円と前期並み。臨時傭員費{M(temp_staff[2])}→{M(temp_staff[3])}百万円、'
    f'外注費{M(outsrc[2])}→{M(outsrc[3])}百万円の抑制で粗利率は{P(gp[2] / sales[2])}→{P(gp[3] / sales[3])}。',
    '≪営業利益≫',
    f'・販管費は{M(sga[2])}→{M(sga[3])}百万円（役員報酬{M(officer[2])}→{M(officer[3])}、宣伝広告費{M(ad[2])}→{M(ad[3])}百万円等）。'
    '2期連続の営業赤字から黒字へ回復。',
    '≪進行期の見通し≫',
    f'・コスト改善の効果が通年で寄与し、営業利益約{M(op_fc)}百万円の着地見込み（売主説明。試算表で確認中）。',
], body=8.5, head=9.5, after=2)
fill_comment(shape_by_id(s, 10), [
    '【調整項目（確認中）】',
    '≪役員報酬≫',
    '・譲渡後は取締役1名を除き退任予定。退任役員の報酬（社保込）の加算と、後任体制の費用は個人別内訳の受領後に算定',
    '≪営業外損益≫',
    '・雑収入（補助金等）、固定資産の賃貸収支は経常性を確認のうえ扱いを決める',
    '✓ 調整後EBITDAは内訳の受領後に提示。表は調整前のEBITDA',
], body=8, head=9, after=1)

# ================================================================ 5 財務ハイライト（B/S）
s = S_BS
drop_pics(s)
set_lead(s, f'{Y[-1]}末の純資産は{M(BS("F20"))}百万円（自己資本比率{P(BS("F20") / BS("C21"))}）。'
            f'有利子負債は前期末比{M(debt_prev - debt)}百万円減の{M(debt)}百万円で、現預金は{M(BS("C5"))}百万円。')
yb = paste(s, 'BSハイライト!B2:F21', 0.37, 1.55, 5.47, 0.875, target=5.2)
note(s, 0.37, yb + 0.1, 5.47, [f'※{Y[-1]}末（決算書）。進行期の試算表は未受領。'], size=7.5)
fill_comment(shape_by_id(s, 7), [
    '≪未収入金≫',
    f'・{M(BS("C7"))}百万円。売掛金の科目はなく、営業債権を未収入金で計上。大半は通常の債権との説明。'
    '一部（取引先の倒産分）は回収困難の見込みで、処理方針を確認中。',
    '≪立替金・仮払金・前払金≫',
    f'・計{M(BS("C8") + BS("C9"))}百万円。内容・相手先は勘定科目内訳書で確認中。',
    '≪未払費用≫',
    f'・{M(BS("F7"))}百万円。コロナ期に猶予を受けた社会保険料を含み、既に完済との説明（確認中）。',
    '≪借入金≫',
    f'・短期{M(BS("F5"))}百万円・長期{M(BS("F12"))}百万円。借入先別残高・保証は確認中。',
], body=8.5, head=9.5, after=2)
fill_comment(shape_by_id(s, 8), [
    '【着眼点】',
    '≪保険積立金≫',
    f'・簿価{M(BS("C17"))}百万円。解約返戻金は簿価を上回る見込み（売主説明。証憑は確認中）',
    '≪未収入金≫',
    '・回収困難分の処理により純資産は減少する。保険積立金の含み益と合わせて株式価値評価で反映',
    '≪土地・関係会社≫',
    f'・土地{M(BS("C14"))}百万円の所在・用途、関係会社との取引・簿外債務の有無は確認中',
    '✓ 時価純資産は勘定科目内訳書の受領後に株式価値評価書で提示',
], body=8, head=9, after=1)

# ================================================================ 6 譲渡希望条件（テンプレートの4バンドの文字だけ差し替え）
s = S_COND
set_text(shape_by_id(s, 19), '株式100％譲渡')
set_text(shape_by_id(s, 56), '応相談')
set_text(shape_by_id(s, 256), ['従業員の継続雇用', '取締役1名（現場運営）が継続勤務。その他の役員は退任予定',
                               '福祉事業を営む関係会社は譲渡対象外（資本関係は確認中）', '代表者保証等の解除（保証の状況は確認中）'],
         size=13)
set_text(shape_by_id(s, 264), '確認中')

# ================================================================ 7 事業内容（商流図の枠を流用）
s = S_BIZ
set_lead(s, '個人向け引越を主力に、法人・医療機関の移転と付帯サービスを提供。受注は紹介ルートと直接受注の両輪で、'
            '専門作業や繁忙期の応援は協力会社を活用する。')
for sid, txt in {5: '協力会社・仕入先', 17: '顧客・受注経路', 21: '対象会社', 39: '引越事業（主力）',
                 46: '付帯サービス', 47: '撤退済みの事業', 12: '協力会社（外注）'}.items():
    set_text(shape_by_id(s, sid), txt)
set_text(shape_by_id(s, 8), ['梱包資材', '燃料', '車両（リース等）', '車両整備（外注）'], size=11)
set_text(shape_by_id(s, 13), ['個人顧客（直接受注）', '紹介元（不動産仲介・', '　住宅関連企業等）', '法人顧客（オフィス・', '　医療機関）'], size=10)
set_text(shape_by_id(s, 51), ['個人向け引越', '（単身〜家族）'])
set_text(shape_by_id(s, 52), ['法人移転', '（オフィス・医療機関）'])
set_text(shape_by_id(s, 49), ['電気工事（エアコン等）', 'ハウスクリーニング'])
set_text(shape_by_id(s, 50), ['ピアノ輸送', '不用品回収'])
set_text(shape_by_id(s, 55), ['家電FC店'])
set_text(shape_by_id(s, 56), ['一般貨物運送'])
set_text(shape_by_id(s, 25), ['電気工事・ピアノ輸送等'], size=9)
set_text(shape_by_id(s, 26), ['繁忙期の作業応援'], size=9)
place(shape_by_id(s, 12), 3.21, 6.24, 5.28, 1.12)
place(shape_by_id(s, 14), 3.21, 6.24)
place(shape_by_id(s, 25), 3.44, 6.86, 2.07, 0.42)
place(shape_by_id(s, 26), 5.80, 6.86, 2.07, 0.42)
place(shape_by_id(s, 22), 5.51, 5.92); place(shape_by_id(s, 31), 5.51, 6.07)
for sid in (5, 17):
    sh_ = shape_by_id(s, sid); sh_.height = Inches(5.76)
note(s, 0.49, 7.42, 10.71, [
    f'※{Y[-1]}の運送原価のうち、紹介手数料{M(referral[-1])}百万円、その他外注費{M(outsrc[-1])}百万円、臨時傭員費{M(temp_staff[-1])}百万円。'
    '撤退済みの事業の売上・撤退時期は確認中。'], size=7.5)

# ================================================================ 8 特徴、強み、成長性等
s = S_FEAT


def band(sh, head, body):
    txb = sh.text_frame._txBody
    ps = txb.findall(qn('a:p'))
    hp, bp = copy.deepcopy(ps[0]), copy.deepcopy(ps[1])
    for p in ps:
        txb.remove(p)
    for tmpl, text in ((hp, head), (bp, body)):
        p = copy.deepcopy(tmpl)
        r = p.find(qn('a:r'))
        r.find(qn('a:t')).text = text
        txb.append(p)


band(shape_by_id(s, 14), '地域密着の引越専業として20年超',
     '東海地方で個人の引越を主力に、オフィス・医療機関の移転まで対応。電気工事・ハウスクリーニング等の付帯サービスを'
     'ワンストップで提供。家電FC店・一般貨物運送から撤退し、経営資源を引越事業に集中した。')
band(shape_by_id(s, 20), '顧客満足度の高さと紹介ルート',
     'オリコン顧客満足度調査（引越し会社）の地域別ランキングで第1位（8年連続と売主説明。部門・年数は確認中）。'
     '社内研修で作業品質を標準化し、不動産仲介・住宅関連企業等からの紹介受注とリピートにつなげている。')
band(shape_by_id(s, 29), '受注チャネルの再構築と内製化による伸び代',
     '法人の個別訪問営業を縮小しWeb受注へ移行したが、移行後は受注が伸び悩んでいる。法人・紹介ルートの再強化、'
     '外注している電気工事等の内製化、買手の拠点・車両との相互活用で収益を伸ばす余地がある（弊社の評価）。')
set_lead(s, '顧客満足度の高さと紹介ルートが受注の基盤。コスト構造の見直しを終え、営業の再強化が次の打ち手になる。')

# ================================================================ 9 株主及び役員
s = S_SH
drop_pics(s)
set_lead(s, '株式は創業家（天野家）が80%、役員と取引先法人が10%ずつ保有。譲渡後は現場運営を担う取締役1名が継続勤務し、'
            'その他の役員は退任予定。')
yb = paste(s, '株主!B2:H15', 0.85, 2.05, 9.6, 0.78, target=1.6)
place(shape_by_id(s, 3), 0.34, yb + 0.15)
yb2 = paste(s, '取締役!B2:H15', 0.85, yb + 0.60, 9.6, 0.78, target=2.75)
note(s, 0.85, yb2 + 0.1, 9.6, [
    '※株主・役員は2024年11月時点の情報（最新の株主名簿・登記で確認中）。上記のほか常務1名が在籍（登記の有無は確認中）。',
    '※役員報酬の個人別内訳、各役員の勤務実態は確認中。'], size=7.5)

# ================================================================ 10 拠点・車両
s = S_SITE
drop_pics(s)
remove(s, 11)
title(s, '拠点・車両')
set_text(shape_by_id(s, 2), '拠点')
set_lead(s, '本部・支店を愛知県内と岐阜県に置き、東海地方をカバー。拠点はすべて賃借で、本店事務所は関係会社所有の建物の一部を'
            '借りている。車両は主にリースとの説明。')
yb = paste(s, '拠点一覧!B4:F9', 0.34, 2.10, 11.0, 0.8)
note(s, 0.34, yb + 0.08, 11.0, ['※2024年時点の情報。現況・賃貸借契約の条件（賃料・期間・譲渡後の継続可否）は確認中。'
                                 '撤退した事業の拠点は現在の用途を確認中。'], size=7.5)
car_lease = v('CR', 'I36')
comment(s, CM_HEAD, CM_BODY, 0.34, yb + 0.55, 11.0, 2.3, [
    '≪車両≫',
    f'・トラック約{trucks:,.0f}台で、主にリースとの売主説明（面談では50台との説明もあり、営業車を含むかを含め車両一覧で確認中）。'
    f'{Y[-1]}の自動車リース料は{M(car_lease)}百万円で、契約形態・計上科目を確認中。',
    '≪繁忙期≫',
    '・3月〜ゴールデンウィーク前が繁忙期。臨時人員・協力会社で波動に対応。月別売上は確認中。',
    '≪本店事務所≫',
    '・関係会社所有の建物（本店所在地）の一部を賃借。譲渡後の継続の可否・条件は確認中。',
], head='Comment')

# ================================================================ 11 確認中の事項・今後の進め方
s = S_TODO
title(s, '確認中の事項・今後の進め方')
set_lead(s, '勘定科目内訳書・従業員一覧・車両一覧を受領次第、本資料を更新し、企業概要書・株式価値評価書として開示する予定。')
yb = paste(s, '確認事項_開示!B4:D13', 0.34, 1.62, 11.0, 0.8)
note(s, 0.34, yb + 0.15, 11.0, [
    '今後の進め方（想定）：秘密保持契約の締結 → 本資料・追加資料の開示 → トップ面談・現地見学 → 意向表明 → 基本合意 → '
    'デューデリジェンス → 最終契約・クロージング'], size=9)

# ================================================================ 保存・後処理
ib.force_white_headers(prs)
prs.save(OUT)
scrub_doc_props(OUT, title='企業概要書（簡易版）')
print('saved', OUT)
