# -*- coding: utf-8 -*-
"""株式価値評価書（越ヶ谷製氷冷蔵株式会社）。数値は全て案件マスターのVRセクションから読み込む。"""
import sys
sys.path.insert(0, '.')
from deck import *

TEMPLATE, MASTER, OUT = sys.argv[1:4]
M = Master(MASTER)
D = Deck(TEMPLATE)
S1, S2, S3, S4 = 'Valuation Summary', '財務分析（B/S）', '財務分析（P/L）', 'Valuation Details'
X0, W0 = 0.37, 10.95
PER = [M.v('BS(借方)', f'{c}4') for c in 'CDEF']
PS = [p.replace('年', '/').replace('月期', '期')[2:] for p in PER]
C4 = 'CDEF'
YR3 = PS[1:]

yb = M.wb['年買法']
ev = M.wb['EV EBITDAマルチプル']
sv = M.wb['株式価値 ']
nc = M.wb['ネットキャッシュ']
mb = M.wb['修正BS(借方)']
mc = M.wb['修正BS(貸方) ']
mp = M.wb['修正PL']
RB, RC, RP, RN = M.ref['修正BS(借方)'], M.ref['修正BS(貸方) '], M.ref['修正PL'], M.ref['ネットキャッシュ']
na_fair = mc[f'E{RC["純資産"]}'].value
na_book = mc[f'C{RC["純資産"]}'].value
netcash = nc[f'C{RN["nc"]}'].value
base_op = yb['F10'].value
base_ebitda = ev['F8'].value
cost_vals = [yb.cell(row=16, column=c).value for c in (3, 4, 5)]
mkt_vals = [ev.cell(row=17, column=c).value for c in (3, 4, 5)]
ev_d = [ev.cell(row=15, column=c).value for c in (3, 4, 5)]
sens = [yb.cell(row=22, column=c).value for c in (3, 4, 5)]
cap_case = [ev.cell(row=20, column=c).value for c in (3, 4, 5)]
debt = M.rv('BS (貸方)', '有利子負債', 'F')

# ============================================================ 表紙
s = D.use(1)
for sh in s.shapes:
    if sh.has_text_frame and '企業概要書' in sh.text_frame.text:
        rs = sh.text_frame.paragraphs[0].runs
        rs[0].text = '株式価値評価書  '
        rs[1].text = '<製氷業／埼玉県>'
    elif sh.has_text_frame and 'Information' in sh.text_frame.text:
        rs = sh.text_frame.paragraphs[0].runs
        rs[0].text = 'Valuation'
        rs[2].text = 'Report'
text_box(s, 1.03, 5.2, 8, 0.3, '評価基準日：2026年3月31日（第67期末）　／　越ヶ谷製氷冷蔵株式会社', size=12, color=NAVY)

# ============================================================ はじめに
s = D.use(2)
for sh in s.shapes:
    if sh.has_text_frame and sh.name == 'テキスト ボックス 22':
        texts = ['本評価書は、株式会社M&A承継機構（以下「弊社」）が、越ヶ谷製氷冷蔵株式会社（以下「対象会社」）の株式譲渡に関する仲介者として、対象会社の株式価値の目安を把握するために作成した簡易的な試算資料です。',
                 '本評価書は、対象会社から受領した決算書（第67期）・勘定科目内訳書・減価償却計算書・法人税申告書、および他の仲介会社が作成した企業概要書に記載された第64〜66期の財務数値に基づいて作成しています。会計監査・財務デューデリジェンスは実施しておらず、受領資料の正確性・完全性を前提としています。',
                 '本評価書に記載の株式価値は、公正価値の算定や、取引価格の妥当性に関する意見を表明するものではありません。最終的な譲渡価格は、譲受企業によるデューデリジェンスの結果および当事者間の交渉により決定されます。',
                 '資料未受領等により評価できなかった科目は簿価で評価し、その旨を本評価書に記載しています。これらの科目は、追加資料の受領により評価額が変動する可能性があります。',
                 '本評価書は内部検討用の資料であり、弊社の書面による承諾なく第三者へ開示することを禁じます。']
        ps = sh.text_frame.paragraphs
        for i, p in enumerate(ps):
            if i < len(texts):
                p.runs[0].text = texts[i]
                for r_ in p.runs[1:]:
                    r_._r.getparent().remove(r_._r)
            else:
                p._p.getparent().remove(p._p)

# ============================================================ 目次
s = D.use(3)
for sh in s.shapes:
    if sh.has_text_frame and sh.name == 'Google Shape;259;p2':
        for p, t in zip(sh.text_frame.paragraphs, ['Valuation Summary', 'Financial Analysis（B/S）', 'Financial Analysis（P/L）', 'Valuation Details', '要確認事項']):
            p.runs[0].text = t
            for r_ in p.runs[1:]:
                r_._r.getparent().remove(r_._r)

# ============================================================ Section1
D.divider('Valuation Summary')
s = D.new('採用するアプローチの選定', S1, 'コストアプローチ（年買法）とマーケットアプローチ（EV/EBITDAマルチプル法）を採用。詳細な事業計画が存在しないためインカムアプローチは採用しない。')
rows = [['アプローチ', '評価手法', '概要', '採否', '理由'],
        ['コストアプローチ', '時価純資産＋営業権法\n（年買法）', '修正BSの時価純資産に、税引後の正常営業利益×1〜3年分を営業権として加算', '採用', '中小企業M&Aで実務上広く用いられ、譲渡当事者にとって理解しやすい'],
        ['マーケットアプローチ', '類似会社比準法\n（EV/EBITDAマルチプル法）', '正常EBITDAに倍率3〜5倍を乗じた事業価値から、非流動性ディスカウントとネットデットを控除', '採用', '装置産業であり償却前の収益力（EBITDA）で評価するのが実態に即する'],
        ['インカムアプローチ', 'DCF法・収益還元法', '将来キャッシュフローの現在価値', '不採用', '詳細な事業計画が存在せず、合理的な算定が困難'],
        ['インカムアプローチ', '配当還元法', '将来の配当期待額の現在価値', '不採用', '少数株主視点の手法であり、経営権の移転を伴う本件の性質に不適']]
table(s, X0, 1.65, [1.6, 2.0, 3.6, 0.7, 3.05], rows, row_h=0.62, size=9.5, align_cols={3: PP_ALIGN.CENTER},
      fill_rows={1: LBLUE, 2: LBLUE})
rect(s, X0, 5.0, W0, 1.9, fill=GRAY)
text_box(s, X0 + 0.2, 5.1, W0 - 0.4, 1.75, [
    [('評価の前提', {'bold': True, 'size': 11, 'color': NAVY})],
    [('・評価基準日：2026年3月31日（第67期末）。正常収益力は直近3期（24/3期〜26/3期）を対象とし、加重は直近期50%・前期30%・前々期20%', {'size': 9.5})],
    [('・実効税率34%。非流動性ディスカウント30%（マーケットアプローチのみ）', {'size': 9.5})],
    [('・概要書の「調整後EBITDA」と本評価書の「修正後EBITDA」は同一の調整方針（案件マスター「修正PL」）で算定', {'size': 9.5})],
    [('・数値は全て案件マスター（Excel）から転記。千円未満四捨五入のため合計が一致しない場合がある', {'size': 9.5})]])

s = D.new('株式価値の試算結果', S1,
          f'基本ケースでは両手法とも株式価値はマイナス。ネットデット{fk(-netcash)}千円が事業価値（{fk(ev_d[0])}〜{fk(ev_d[2])}千円）を上回り、株式価値は備忘価額程度。工場建物の評価が結論を左右する。')
unit(s, 9.3, 1.55, 2.0)
rows = [['評価手法', '下限', '中央', '上限', '主な前提'],
        ['コストアプローチ（年買法）', *[fk(v) for v in cost_vals], f'時価純資産{fk(na_fair)}千円＋営業権（基準営業利益{fk(base_op)}千円×1〜3年）'],
        ['マーケットアプローチ（EV/EBITDA法）', *[fk(v) for v in mkt_vals], f'基準EBITDA{fk(base_ebitda)}千円×3〜5倍×(1-30%)＋ネットキャッシュ{fk(netcash)}千円'],
        ['算定上の株式価値', '0', '0', '0', '両手法ともマイナスのため0円（備忘価額）']]
table(s, X0, 1.78, [2.9, 1.15, 1.15, 1.15, 4.6], rows, row_h=0.45, size=9.5, num_cols={1, 2, 3}, bold_rows=[3], fill_rows={3: LBLUE})
# 価値の構成（中央値）
sub_head(s, X0, 3.75, 6, '事業価値と有利子負債の関係（中央値ベース）')
bar_line_chart(s, X0, 4.05, 6.2, 3.0, ['事業価値EV\n（ディスカウント後）', '現預金＋\n保険積立金', '有利子負債等', '株式価値\n（試算）'],
               [('千円', [rk(ev_d[1]), rk(nc[f'C{RN["cl"]}'].value), rk(nc[f'C{RN["dl"]}'].value), rk(mkt_vals[1])])],
               number_format='#,##0;△#,##0', legend=False, bar_colors=(NAVY,), font_size=8)
comment_box(s, 6.8, 3.75, 4.5, 3.35, [
    ('結論', ['基本ケースの算定上の株式価値は0円（備忘価額）。']),
    ('マイナスとなる要因', [f'有利子負債{fk(debt)}千円が、事業規模（基準EBITDA{fk(base_ebitda)}千円）の約{debt / base_ebitda:.1f}倍と大きい。',
                         f'工場建物は会計上の償却を停止しており、固定資産税評価額への評価替えで{fk(-mb[f"D{RB["建物"]}"].value)}千円の評価減。']),
    ('感応度', [f'建物を簿価で評価すると年買法は{fk(sens[0])}〜{fk(sens[2])}千円とプラスに転じる。建物時価の検証（不動産鑑定・再調達原価）が最重要の確認事項。']),
    ('交渉上の論点', ['代表者の連帯保証・担保提供の解除、代表者借入金・資本性ローンの取扱いが実質的な対価となる。']),
], size=8)

# ============================================================ Section2 B/S
D.divider('Financial Analysis（B/S）')
s = D.new('過去B/Sの推移', S2, '借入金の約定返済と利益の積上げにより純資産は4期で増加したが、借入金は依然として総資産の8割超を占める。')
bh = M.wb['BSハイライト']
RG = M.ref['BSハイライト']
keys = ['現預金（保険積立除く）', '保険積立金', '売掛金', '棚卸資産', 'その他流動資産', '流動資産合計', '建物・構築物', '機械・車両・器具', '土地', '無形固定資産・投資等',
        '固定資産合計', '資産合計', '支払手形・買掛金', '借入金（短期・長期）', 'その他負債', '負債合計', '純資産合計', '負債・純資産合計']
rows = [['科目', *PS]]
bold = []
for k in keys:
    rows.append([k, *[fk(bh[f'{c}{RG[k]}'].value) for c in C4]])
    if '合計' in k:
        bold.append(len(rows) - 1)
rows.append(['自己資本比率', *[fp(bh[f'{c}{RG["自己資本比率"]}'].value) for c in C4]])
table(s, X0, 1.62, [2.2, 1.1, 1.1, 1.1, 1.1], rows, row_h=0.27, size=9, num_cols={1, 2, 3, 4}, bold_rows=bold, fill_rows={b: LBLUE for b in bold})
unit(s, 5.0, 1.42, 1.6)
comment_box(s, 7.15, 1.62, 4.15, 5.45, [
    ('資産', ['工場の土地・建物と機械で資産の約7割を占める装置型の資産構成。',
            '建物は4期とも259,242千円で、会計上の減価償却を行っていない。',
            '現預金は25/3期に定期積金の解約等で減少。保険積立金（諸掛金）は現預金に含めて表示されている。']),
    ('負債', [f'借入金は23/3期{fk(bh[f"C{RG["借入金（短期・長期）"]}"].value)}千円から26/3期{fk(bh[f"F{RG["借入金（短期・長期）"]}"].value)}千円へ減少。',
            '短期借入金80,000千円は継続的に借換えされている（借換条件は要確認）。']),
    ('純資産', [f'26/3期の簿価純資産{fk(na_book)}千円。自己資本比率は{fp(bh[f"F{RG["自己資本比率"]}"].value)}。']),
], size=8.5)

s = D.new('評価科目別の検討結果', S2, '工場建物を固定資産税評価額へ、土地を同評価額÷0.7へ評価替え。資料未受領の科目は簿価で評価し、要確認事項として整理した。')
rows = [['科目', '帳簿価額', '評価差額', '評価額', '調整事由']]
for k in ['現金及び預金（保険積立除く）', '売掛金', '棚卸資産', 'その他流動資産', '保険積立金', '建物', '構築物', '機械', '車両運搬具・器具備品', '土地', '特許商標権', '出資金']:
    r_ = RB[k]
    rows.append([k, fk(mb[f'C{r_}'].value), fk(mb[f'D{r_}'].value, dash='0'), fk(mb[f'E{r_}'].value), mb[f'F{r_}'].value])
for k in ['賞与引当金', '退職給付引当金']:
    r_ = RC[k]
    rows.append([k + '（負債）', '-', '0', '-', mc[f'F{r_}'].value])
r_ = RB['資産合計']
rows.append(['資産の評価差額 合計', fk(mb[f'C{r_}'].value), fk(mb[f'D{r_}'].value), fk(mb[f'E{r_}'].value), ''])
hl = {i: LBLUE for i, r in enumerate(rows) if i and r[2] not in ('0', '-', '')}
hl[len(rows) - 1] = LBLUE
table(s, X0, 1.62, [2.3, 1.2, 1.2, 1.2, 5.05], rows, row_h=0.3, size=9, num_cols={1, 2, 3}, bold_rows=[len(rows) - 1], fill_rows=hl)
unit(s, 5.6, 1.42, 1.3)
land = M.wb['土地']
footnote(s, [f'※土地：固定資産税評価額{fk(land["G5"].value)}千円÷0.7＝{fk(land["I5"].value)}千円。建物：固定資産税評価額{fk(land["G6"].value)}千円（再調達原価法による検証は未実施）。',
             '※構築物・機械：減価償却計算書の当期償却不足額を減額。過年度の累積不足額は資料不足のため未反映。',
             '※評価差額が純額で評価損のため、税効果（繰延税金資産）は回収可能性が不確実として認識しない。'], y=6.55)

s = D.new('修正B/S（評価基準日：2026年3月31日）', S2, f'評価替えにより資産は{fk(-mb[f"D{RB["資産合計"]}"].value)}千円減少し、時価純資産は{fk(na_fair)}千円（簿価純資産{fk(na_book)}千円）となる。')
rows = [['資産の部', '簿価', '修正額', '修正後']]
for k in ['現金及び預金（保険積立除く）', '売掛金', '棚卸資産', 'その他流動資産', '保険積立金', '建物', '構築物', '機械', '車両運搬具・器具備品', '土地', '特許商標権', '出資金', '資産合計']:
    r_ = RB[k]
    rows.append([k, fk(mb[f'C{r_}'].value), fk(mb[f'D{r_}'].value), fk(mb[f'E{r_}'].value)])
table(s, X0, 1.62, [2.3, 1.0, 1.0, 1.0], rows, row_h=0.34, size=9.5, num_cols={1, 2, 3}, bold_rows=[13], fill_rows={13: LBLUE})
rows = [['負債・純資産の部', '簿価', '修正額', '修正後']]
for k in ['支払手形・買掛金', '借入金（短期・長期）', '未払金', '未払法人税等', 'その他流動負債', '賞与引当金', '退職給付引当金', '負債合計', '税効果', '純資産']:
    r_ = RC[k]
    lab = {'税効果': '評価差額に係る税効果', '純資産': '純資産（時価純資産）'}.get(k, k)
    rows.append([lab, fk(mc[f'C{r_}'].value), fk(mc[f'D{r_}'].value), fk(mc[f'E{r_}'].value)])
table(s, 5.95, 1.62, [2.35, 1.0, 1.0, 1.0], rows, row_h=0.34, size=9.5, num_cols={1, 2, 3}, bold_rows=[8, 10], fill_rows={8: GRAY, 10: LBLUE})
unit(s, 9.5, 1.42, 1.8)
footnote(s, '※借入金には代表者借入14,000千円・日本政策金融公庫の資本性ローン105,420千円を含む。評価上はいずれも有利子負債として扱う。', y=6.55)

s = D.new('ネットキャッシュの算定', S2, f'キャッシュライク{fk(nc[f"C{RN["cl"]}"].value)}千円に対し、デットライク{fk(nc[f"C{RN["dl"]}"].value)}千円。ネットデットは{fk(-netcash)}千円。')
rows = [['区分', '項目', '金額（千円）', '判断根拠']]
for k, grp in [('cash', 'キャッシュライク'), ('ins', ''), ('cl', ''), ('bank', 'デットライク'), ('owner', ''), ('tax', ''), ('dl', ''), ('nc', '')]:
    r_ = RN[k]
    rows.append([grp, nc[f'B{r_}'].value, fk(nc[f'C{r_}'].value), nc[f'D{r_}'].value or ''])
table(s, X0, 1.62, [1.5, 3.2, 1.4, 4.85], rows, row_h=0.4, size=9.5, num_cols={2}, bold_rows=[3, 7, 8], fill_rows={3: GRAY, 7: GRAY, 8: LBLUE})
footnote(s, [f'※（参考）日本政策金融公庫の資本性ローン105,420千円を資本とみなした場合のネットキャッシュは{fk(nc[f"C{RN["nc_cap"]}"].value)}千円。',
             '※保険積立金は解約返戻金の資料未受領のため簿価。代表者借入金はクロージング時の返済・放棄の取扱いにより変動する。'], y=5.5)

# ============================================================ Section3 P/L
D.divider('Financial Analysis（P/L）')
s = D.new('過去P/Lの推移', S3, '売上高は4期連続増収。23/3期は電力料の高騰で営業赤字となったが、24/3期以降はEBITDAで30〜37百万円を確保している。')
ph = M.wb['PLハイライト']
RH = M.ref['PLハイライト']
bar_line_chart(s, X0, 1.6, 5.6, 3.0, PS, [('売上高', [rk(ph[f'{c}{RH["売上高"]}'].value) for c in C4])],
               [('EBITDA（右軸）', [rk(ph[f'{c}{RH["EBITDA"]}'].value) for c in C4], '0889C9', '#,##0;△#,##0')],
               number_format='#,##0', bar_colors=(NAVY,))
rows = [['科目', *PS]]
for k, lab in [('売上高', '売上高'), ('売上総利益', '売上総利益'), ('粗利率', '　売上総利益率'), ('営業利益', '営業利益'), ('営業利益率', '　営業利益率'),
               ('経常利益', '経常利益'), ('当期純利益', '当期純利益'), ('減価償却費', '減価償却費'), ('EBITDA', 'EBITDA')]:
    vals = [ph[f'{c}{RH[k]}'].value for c in C4]
    rows.append([lab, *[fp(v) if '率' in k else fk(v) for v in vals]])
table(s, X0, 4.75, [2.0, 0.9, 0.9, 0.9, 0.9], rows, row_h=0.24, size=8.5, num_cols={1, 2, 3, 4}, bold_rows=[1, 4, 9], small_rows=[3, 5])
comment_box(s, 6.3, 1.6, 5.0, 5.5, [
    ('売上高', ['製氷関係が約9割。7〜8月の盛夏期に年間売上の約3割が集中し、天候の影響を受けやすい。']),
    ('売上原価', [f'電力料（動力費）は23/3期{fk(M.rv("CR", "製造経費/動力費", "C"))}千円→24/3期{fk(M.rv("CR", "製造経費/動力費", "D"))}千円へ低下し粗利率が回復。',
              '減価償却費は23/3期にほぼ計上なし（19,763円）。24/3期以降は11〜22百万円を計上しており、期間比較では償却前のEBITDAが実態に近い。']),
    ('販管費', ['配送運賃が24/3期以降増加（26/3期27,570千円）。',
             '支払手数料（25/3期3,156千円・26/3期1,765千円）の内容は要確認。']),
    ('営業外', ['支払利息は年6〜8百万円。65期の雑収入8,482千円・雑損5,192千円の内容は要確認。']),
], size=8.5)

s = D.new('正常収益力の算定（修正P/L）', S3, f'臨時的な退職給与を加算した修正後EBITDAは3期とも34〜37百万円で安定。加重平均の基準EBITDAは{fk(base_ebitda)}千円、基準営業利益（税引後）は{fk(base_op)}千円。')
rows = [['科目'] + sum([[f'{y}\n帳簿', '修正', '修正後'] for y in YR3], [])]
for k, lab in [('売上高', '売上高'), ('売上原価', '売上原価'), ('売上総利益', '売上総利益'), ('販管費', '販管費'), ('営業利益', '営業利益'), ('減価償却費', '減価償却費'), ('EBITDA', 'EBITDA')]:
    r_ = RP[k]
    rows.append([lab] + [fk(mp.cell(row=r_, column=c).value) for c in range(3, 12)])
table(s, X0, 1.62, [1.5] + [1.05, 0.9, 1.05] * 3, rows, row_h=0.33, size=9, num_cols=set(range(1, 10)), bold_rows=[3, 5, 7], fill_rows={5: LBLUE, 7: LBLUE}, hdr_size=8.5)
unit(s, 9.5, 1.42, 1.8)
rows = [['調整項目', *YR3, '調整理由']]
for k in ['adj_退職給与', 'adj_役員報酬']:
    r_ = RP[k]
    rows.append([mp[f'B{r_}'].value, *[fk(mp.cell(row=r_, column=c).value, dash='0') for c in (4, 7, 10)], mp[f'L{r_}'].value])
r_ = RP['adj_total']
rows.append(['調整額合計', *[fk(mp.cell(row=r_, column=c).value, dash='0') for c in (4, 7, 10)], ''])
table(s, X0, 4.6, [2.7, 0.9, 0.9, 0.9, 5.55], rows, row_h=0.42, size=9, num_cols={1, 2, 3}, bold_rows=[3], fill_rows={3: LBLUE})
footnote(s, ['※調整しなかった項目：建物の償却停止（EBITDAベースのマーケットアプローチを重視することで対応）、支払手数料（内容要確認）、代表者からの駐車場賃借（継続前提）。',
             '※本調整方針は企業概要書の「調整後EBITDA」と同一（案件マスター「修正PL」シート）。'], y=6.55)

# ============================================================ Section4
D.divider('Valuation Details')
s = D.new('年買法による算出', S4, f'基準営業利益{fk(base_op)}千円に1〜3年を乗じた営業権を時価純資産{fk(na_fair)}千円に加算。株式価値は{fk(cost_vals[0])}〜{fk(cost_vals[2])}千円とマイナス。')
rows = [['項目', *YR3, '合計']]
for r_, lab in [(5, '修正後営業利益'), (6, '税考慮額（△34%）'), (7, '税考慮後営業利益'), (8, '加重'), (9, '加重後')]:
    vals = [yb.cell(row=r_, column=c).value for c in (3, 4, 5)]
    rows.append([lab, *[f'{int(v * 100)}%' if r_ == 8 else fk(v) for v in vals], ''])
rows.append(['基準営業利益（加重平均）', '', '', '', fk(base_op)])
table(s, X0, 1.62, [3.0, 1.3, 1.3, 1.3, 1.5], rows, row_h=0.34, size=9.5, num_cols={1, 2, 3, 4}, bold_rows=[6], fill_rows={6: LBLUE})
rows = [['', '下限（1年）', '中央（2年）', '上限（3年）']]
for r_, lab in [(14, '営業権'), (15, '時価純資産'), (16, '想定株式価値')]:
    rows.append([lab, *[fk(yb.cell(row=r_, column=c).value) for c in (3, 4, 5)]])
rows.append(['（感応度）建物を簿価評価した場合', *[fk(v) for v in sens]])
table(s, X0, 4.25, [3.0, 1.3, 1.3, 1.3], rows, row_h=0.38, size=9.5, num_cols={1, 2, 3}, bold_rows=[3], fill_rows={3: LBLUE}, small_rows=[4])
unit(s, 8.5, 1.42, 1.3)
comment_box(s, 8.4, 4.25, 2.9, 2.4, [('ポイント', ['基本ケースでは時価純資産のマイナスを営業権で補えない。', f'建物を簿価で評価すると{fk(sens[0])}〜{fk(sens[2])}千円とプラスに転じ、建物の時価が結論を左右する。'])], size=8.5, title='Point', title_size=11)
footnote(s, '※営業権倍率1〜3年は国内中小企業M&Aのボリュームゾーン。税考慮は営業利益がマイナスの期は0。', y=6.85)

s = D.new('EV/EBITDAマルチプル法による算出', S4, f'基準EBITDA{fk(base_ebitda)}千円×3〜5倍の事業価値から非流動性ディスカウント30%を控除。ネットデット{fk(-netcash)}千円を差し引くと株式価値は{fk(mkt_vals[0])}〜{fk(mkt_vals[2])}千円。')
rows = [['項目', *YR3, '合計']]
for r_, lab in [(5, '修正後EBITDA'), (6, '加重'), (7, '加重後')]:
    vals = [ev.cell(row=r_, column=c).value for c in (3, 4, 5)]
    rows.append([lab, *[f'{int(v * 100)}%' if r_ == 6 else fk(v) for v in vals], ''])
rows.append(['基準EBITDA（加重平均）', '', '', '', fk(base_ebitda)])
table(s, X0, 1.62, [3.0, 1.3, 1.3, 1.3, 1.5], rows, row_h=0.34, size=9.5, num_cols={1, 2, 3, 4}, bold_rows=[4], fill_rows={4: LBLUE})
rows = [['', '下限（3倍）', '中央（4倍）', '上限（5倍）']]
for r_, lab in [(13, '事業価値（EV）'), (14, '非流動性ディスカウント（△30%）'), (15, 'ディスカウント後事業価値'), (16, 'ネットキャッシュ'), (17, '想定株式価値')]:
    rows.append([lab, *[fk(ev.cell(row=r_, column=c).value) for c in (3, 4, 5)]])
rows.append(['（参考）資本性ローンを資本とみなした場合', *[fk(v) for v in cap_case]])
table(s, X0, 3.65, [3.0, 1.3, 1.3, 1.3], rows, row_h=0.36, size=9.5, num_cols={1, 2, 3}, bold_rows=[5], fill_rows={5: LBLUE}, small_rows=[6])
unit(s, 8.5, 1.42, 1.3)
comment_box(s, 8.4, 3.65, 2.9, 2.9, [('ポイント', ['倍率3〜5倍は中小企業M&Aの実務レンジ。', f'株式価値がプラスとなるにはEV/EBITDA約{(-netcash) / base_ebitda / 0.7:.1f}倍（ディスカウント前）が必要。'])], size=8.5, title='Point', title_size=11)
footnote(s, ['※非流動性ディスカウント：上場会社の株価倍率には市場性（流動性）が織り込まれているため、非上場株式への適用にあたり控除。',
             '※製氷専業の上場会社は存在しないため、類似上場会社（低温物流・食品）の倍率は算定日時点で別途更新する（案件マスター「参照企業」）。'], y=6.75)

# ============================================================ 要確認事項
D.divider('要確認事項')
s = D.new('評価に影響する要確認事項', '要確認事項', '以下の事項は資料未受領等により簿価または仮定で評価しており、追加資料の受領により評価額が変動する。')
rows = [['No', '項目', '確認内容', '評価への影響']]
items = [('建物', '会計上の償却停止の経緯、建物簿価（BS259,242千円／内訳書280,673千円）の差異、最新の固定資産税評価額、不動産鑑定等による時価の検証', '時価純資産（△168百万円の評価減。簿価評価なら年買法はプラス）'),
         ('定期預金', '預貯金内訳（19,601千円）とBS残高（26,401千円）の差異6,800千円、担保差入の有無', 'ネットキャッシュ'),
         ('保険積立金', '大同生命・第一生命の解約返戻金額', '時価純資産・ネットキャッシュ'),
         ('売掛金', '年齢表・滞留債権、(株)小野田商店の計上区分と関係', '時価純資産'),
         ('退職給与', '67期6,178千円の対象者・退職金規程・今後の退職予定者', '正常収益力（臨時性の判断）・退職給付債務'),
         ('賞与', '支給時期・算定期間', '賞与引当金（負債）'),
         ('支払手数料', '25/3期3,156千円・26/3期1,765千円の内容（一過性か）', '正常収益力'),
         ('借入金', '代表者借入14,000千円・資本性ローン105,420千円の取扱い、保証・担保の解除条件', 'ネットキャッシュ・譲渡条件'),
         ('発行済株式数', '登記22,000株／決算注記200株／別表二220株の相違', '1株当たり価値・株式譲渡手続')]
for i, (a, b, c) in enumerate(items):
    rows.append([str(i + 1), a, b, c])
table(s, X0, 1.62, [0.45, 1.5, 5.9, 3.1], rows, row_h=0.46, size=9, align_cols={0: PP_ALIGN.CENTER})
D.use(64)
D.finalize(OUT)
print('saved', OUT, len(D.order))
