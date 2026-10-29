# -*- coding: utf-8 -*-
"""企業概要書（越ヶ谷製氷冷蔵株式会社）。数値は全て案件マスターから読み込む。"""
import sys
sys.path.insert(0, '.')
from deck import *
from data import EQUIPMENT_MAIN, LICENSES, SITE_SALES_67, BUSINESS_MIX_67, MONTHLY_67

TEMPLATE, MASTER, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
IMG = sys.argv[4] if len(sys.argv) > 4 else None
M = Master(MASTER)
D = Deck(TEMPLATE)
ES, BZ, ORG, FIN, APX = 'エグゼクティブサマリー', '事業概要', '組織、拠点、設備等', '財務情報詳細', 'Appendix'
PER = [M.v('BS(借方)', f'{c}4') for c in 'CDEF']           # 2023年3月期…
PS = [p.replace('年', '/').replace('月期', '期')[2:] for p in PER]  # 23/3期…
X0, W0 = 0.37, 10.95


def pl(key, col):
    return M.rv('PLハイライト', key, col)


def bs(key, col='F'):
    return M.rv('BSハイライト', key, col)


C4 = 'CDEF'

# ============================================================ 1 表紙
s = D.use(1)
for sh in s.shapes:
    if sh.name == 'タイトル 2' and sh.has_text_frame and '企業概要書' in sh.text_frame.text:
        sh.text_frame.paragraphs[0].runs[1].text = '<製氷業／埼玉県>'
# 2 はじめに / 3 目次 / 4 中扉
D.use(2)
D.use(3)
D.use(4)

# ============================================================ 5 基本情報
s = D.new('基本情報', ES,
          '1960年の創業以来65年超、埼玉県越谷市の自社工場で純水を凍結させた「純氷」を製造。角氷・加工氷・袋氷を首都圏の食品卸・氷卸等へ供給し、自社ブランド「キングアイス」を展開する老舗製氷メーカー。')
section_bar(s, X0, 1.62, 6.55, '会社概要')
ws = M.wb['案件サマリー']
rows = [['項目', '内容']]
for k in ['会社名', '本店所在地', '工場・営業所', '設立', '資本金', '代表者', '株主', '従業員数', '事業内容', '主要取引先', '主要設備', '売上高', '譲渡理由']:
    rows.append([k, ws[f'I{M.row("案件サマリー", "co_" + k)}'].value])
rows[1][1] = '越ヶ谷製氷冷蔵株式会社（ブランド：純氷「キングアイス」）'
rows[3][1] = '越谷工場：越谷市大字平方999-4（平方工業団地内）\n松戸営業所：千葉県松戸市稔台5-14-3'
rows[7][1] = '小野田 洋（100%）'
table(s, X0, 2.0, [1.15, 5.4], rows, row_h=0.36, size=9, header_rows=1)
photo(s, 'product_ice', 7.15, 1.62, 3.95, 1.75, '角氷・純氷（出典：king-ice.com/product.html）', IMG)
photo(s, 'product_bag', 7.15, 3.52, 1.9, 1.55, '袋氷「キングアイス」（king-ice.com/product_04.html）', IMG)
photo(s, 'product_dia', 9.2, 3.52, 1.9, 1.55, 'ダイヤアイス（king-ice.com/product_09.html）', IMG)
photo(s, 'factory', 7.15, 5.22, 3.95, 1.75, '越谷工場（外観）（king-ice.com/corporate.html）', IMG)
footnote(s, '※資本金は2026年5月15日に11,000千円から10,000千円へ減資（登記）。売上高・従業員数は2026年3月期。', y=7.22)

# ============================================================ 6 事業概要（サマリー）
s = D.new('事業概要', ES,
          '純水製造から凍結・加工・保管・配送までを自社で一貫して行う製氷事業が売上の約9割。ドライアイス等の商品販売と冷蔵保管を組み合わせ、夏季の需要ピークに対応する供給体制を持つ。')
sub_head(s, X0, 1.6, 5, '事業構成（2026年3月期・兼業割合）')
doughnut(s, X0, 1.95, 2.3, 2.3, [b[0] for b in BUSINESS_MIX_67], [b[1] / 100 for b in BUSINESS_MIX_67], [NAVY, BLUE, RGBColor(0x9D, 0xC3, 0xE6)])
yb = 2.0
for i, (nm, pct) in enumerate(BUSINESS_MIX_67):
    col = [NAVY, BLUE, RGBColor(0x9D, 0xC3, 0xE6)][i]
    rect(s, 2.8, yb + i * 0.72, 0.12, 0.55, fill=col)
    text_box(s, 3.0, yb + i * 0.72, 2.7, 0.6, [[(f'{pct:.1f}%', {'bold': True, 'size': 14, 'color': col if i < 2 else BLUE, 'font': NUMF})], [(nm, {'size': 8.5})]])
text_box(s, X0, 4.35, 5.4, 0.2, '出所：法人事業概況説明書（第67期）の兼業割合', size=7.5, color=DGRAY)
# KPI タイル
sub_head(s, 5.95, 1.6, 5, '主要指標')
kpi = [('売上高（26/3期）', f'{fk(pl("売上高", "F"))}千円'),
       ('調整後EBITDA（直近3期平均）', f'{fk(M.rv("PLハイライト", "avg3", "F"))}千円'),
       ('製氷能力', '日産40トン'),
       ('従業員数', '25名')]
for i, (k, v) in enumerate(kpi):
    xx, yy = 5.95 + (i % 2) * 2.7, 1.98 + (i // 2) * 1.15
    rect(s, xx, yy, 2.55, 1.0, fill=LBLUE)
    text_box(s, xx, yy + 0.1, 2.55, 0.25, k, size=8.5, color=NAVY, align=PP_ALIGN.CENTER)
    text_box(s, xx, yy + 0.38, 2.55, 0.45, [[(v, {'bold': True, 'size': 16, 'color': NAVY, 'font': NUMF})]], align=PP_ALIGN.CENTER)
# 事業フロー
sub_head(s, X0, 4.6, 6, '事業フロー')
steps = [('純水製造', 'イオン交換・RO膜\nろ過装置'), ('凍結', 'アイス缶で\n約48時間'), ('加工', 'カット・砕氷\n袋詰'), ('保管', '冷蔵設備\nF級1,000トン'), ('配送・販売', '食品卸・氷卸\n業務用顧客')]
for i, (h, b) in enumerate(steps):
    xx = X0 + i * 2.2
    rect(s, xx, 4.95, 2.05, 0.42, fill=NAVY, shape=MSO_SHAPE.PENTAGON if i < 4 else MSO_SHAPE.RECTANGLE, text=h, size=11, color=WHITE, bold=True)
    rect(s, xx, 5.42, 1.9, 0.62, fill=WHITE, line=LINE, text=b, size=9)
# 強み
rect(s, X0, 6.2, W0, 0.9, fill=GRAY)
rect(s, X0, 6.2, 1.3, 0.9, fill=NAVY, text='強み', size=14, color=WHITE, bold=True)
pts = ['65年超の業歴と伝統製法による\n高品質な純氷（自社ブランド）', '大手食品卸を含む首都圏の\n安定した販路（掛売99%）', '製氷・加工・冷蔵・配送の\n一貫体制と自社保有の工場']
for i, t in enumerate(pts):
    xx = 1.85 + i * 3.15
    rect(s, xx, 6.47, 0.32, 0.32, fill=BLUE, shape=MSO_SHAPE.OVAL, text='✓', size=11, color=WHITE, bold=True)
    text_box(s, xx + 0.4, 6.36, 2.7, 0.55, t, size=9.5, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
footnote(s, '※製法・設備の記載は同社公式サイト等の公開情報および法人事業概況説明書による。', y=7.2)

# ============================================================ 7 財務ハイライト（P/L）
avg3 = M.rv('PLハイライト', 'avg3', 'F')
avg2 = M.rv('PLハイライト', 'avg2', 'F')
s = D.new('財務ハイライト（P/L）', ES,
          f'売上高は4期連続で増加し2026年3月期に{fk(pl("売上高", "F"))}千円。電力費高騰で営業赤字となった2023年3月期から回復し、臨時的支出を調整した調整後EBITDAは直近3期平均{fk(avg3)}千円で安定。')
unit(s, 4.9, 1.55, 1.8)
rows = [['', *PS]]
spec = [('売上高', '売上高', 'v'), ('前期比', '前期比', 'p'), ('売上総利益', '売上総利益', 'v'), ('粗利率', '利益率', 'p'),
        ('営業利益', '営業利益', 'v'), ('営業利益率', '利益率', 'p'), ('経常利益', '経常利益', 'v'), ('経常利益率', '利益率', 'p'),
        ('当期純利益', '当期純利益', 'v'), ('減価償却費', '減価償却費', 'v'), ('EBITDA', 'EBITDA', 'v'),
        ('adj_退職', '　退職給与（臨時）', 'v'), ('adj_役員', '　役員報酬', 'v'), ('調整後EBITDA', '調整後EBITDA', 'v')]
small, bold_r, fills = [], [], {}
for key, nm, kind in spec:
    vals = [M.rv('PLハイライト', key, c) for c in C4]
    if kind == 'p':
        cells = ['-' if (key == '前期比' and i == 0) else fp(v) for i, v in enumerate(vals)]
        small.append(len(rows))
        rows.append([nm, *cells])
    else:
        if key.startswith('adj'):
            cells = ['-' if i == 0 else (fk(v, dash='0') if v == 0 else ('+' + fk(v) if v > 0 else fk(v))) for i, v in enumerate(vals)]
        else:
            cells = [fk(v) for v in vals]
        rows.append([nm, *cells])
    if key in ('売上高', '営業利益', '調整後EBITDA'):
        bold_r.append(len(rows) - 1)
fills[len(rows) - 1] = LBLUE
t = table(s, X0, 1.78, [1.6, 1.25, 1.25, 1.25, 1.25], rows, row_h=0.31, size=10, num_cols={1, 2, 3, 4},
          bold_rows=bold_r, fill_rows=fills, small_rows=small, align_cols=None)
for r_ in small:
    for j in range(1, 5):
        t.table.cell(r_, j).text_frame.paragraphs[0].runs[0].font.size = Pt(8)
footnote(s, [f'※EBITDA＝営業利益＋減価償却費（製造原価・販管費計上分）。調整後EBITDAの直近2期平均は{fk(avg2)}千円。',
             '※23/3期は評価対象期間外のため調整なし。'], y=6.5, w=6.6)
comment_box(s, 7.05, 1.62, 4.25, 4.15, [
    ('売上高', [f'23/3期{fk(pl("売上高", "C"))}千円から24/3期{fk(pl("売上高", "D"))}千円（前期比{fp(pl("前期比", "D"))}）へ大きく伸長し、以降も増収を継続。',
              '製氷関係が約9割を占め、7〜8月の盛夏期に年間売上の約3割が集中する。']),
    ('売上総利益', [f'23/3期は電力料（動力費）が{fk(M.rv("CR", "製造経費/動力費", "C"))}千円に膨らみ粗利率{fp(pl("粗利率", "C"))}に低下。',
                f'24/3期以降は電力料が32〜34百万円台に落ち着き、粗利率は22〜26%で推移。']),
    ('営業利益', [f'配送委託の運賃増加（24/3期{fk(M.rv("SGA", "輸送関係費/運賃", "D"))}千円→26/3期{fk(M.rv("SGA", "輸送関係費/運賃", "F"))}千円）等により、25/3期以降は減益。',
              f'26/3期は臨時的な退職給与{fk(M.rv("CR", "労務費/退職給与", "F"))}千円を含み{fk(pl("営業利益", "F"))}千円。']),
    ('EBITDA', [f'減価償却費は24/3期以降11〜22百万円で推移（23/3期はほぼ未計上）。',
                f'臨時的な退職給与を調整した調整後EBITDAは34〜37百万円で安定し、直近3期平均{fk(avg3)}千円。']),
], size=8.5)
rect(s, 7.05, 5.88, 4.25, 1.22, fill=LBLUE, line=BLUE)
text_box(s, 7.18, 5.93, 4.0, 1.15, [
    [('【調整項目】', {'bold': True, 'size': 9.5})],
    [('≪退職給与≫ 26/3期・24/3期の退職給与は臨時的な支出と推察し加算（要確認）', {'size': 8.5})],
    [('≪役員報酬≫ 現行7,200千円／年は標準的な水準と想定し調整なし', {'size': 8.5})],
    [(f'✓ 調整後EBITDA（直近3期平均）は{fk(avg3)}千円', {'bold': True, 'size': 9.5, 'color': NAVY})]])

# ============================================================ 8 財務ハイライト（B/S）
s = D.new('財務ハイライト（B/S）', ES,
          f'越谷工場の土地・建物と製氷設備を自社保有する装置型の資産構成。一方で借入金{fk(bs("借入金（短期・長期）"))}千円を抱え、工場建物は会計上の償却が行われていないため簿価が時価を大きく上回る点に留意を要する。')
unit(s, 3.9, 1.55, 1.8)
L = [('現預金', fk(bs('現預金（保険積立除く）'))), ('保険積立金', fk(bs('保険積立金'))), ('売掛金', fk(bs('売掛金'))), ('棚卸資産', fk(bs('棚卸資産'))),
     ('その他', fk(bs('その他流動資産'))), ('流動資産', fk(bs('流動資産合計'))), ('建物・構築物', fk(bs('建物・構築物'))),
     ('機械・車両・器具', fk(bs('機械・車両・器具'))), ('土地', fk(bs('土地'))), ('無形・投資等', fk(bs('無形固定資産・投資等'))),
     ('固定資産', fk(bs('固定資産合計'))), ('資産合計', fk(bs('資産合計')))]
cl_other = M.v('BS (貸方)', f'F{M.row("BS (貸方)", "sub_cl")}') - M.v('BS (貸方)', f'F{M.row("BS (貸方)", "支払手形")}') - M.v('BS (貸方)', f'F{M.row("BS (貸方)", "買掛金")}') - M.v('BS (貸方)', f'F{M.row("BS (貸方)", "短期借入金")}')
Rr = [('支払手形・買掛金', fk(bs('支払手形・買掛金'))), ('短期借入金', fk(M.rv('BS (貸方)', '短期借入金', 'F'))), ('その他', fk(cl_other)),
      ('流動負債', fk(M.rv('BS (貸方)', 'sub_cl', 'F'))), ('長期借入金', fk(M.rv('BS (貸方)', '長期借入金', 'F'))), ('固定負債', fk(M.rv('BS (貸方)', 'sub_fl', 'F'))),
      ('負債合計', fk(bs('負債合計'))), ('資本金', fk(M.rv('BS (貸方)', '資本金', 'F'))),
      ('利益剰余金', fk(bs('純資産合計') - M.rv('BS (貸方)', '資本金', 'F'))), ('純資産合計', fk(bs('純資産合計'))), ('', ''), ('負債・純資産合計', fk(bs('負債・純資産合計')))]
rows = [[PS[3] + '（2026年3月31日）', None, None, None]]
for (a, b), (c, d) in zip(L, Rr):
    rows.append([a, b, c, d])
bold_r = [6, 11, 12]
t = table(s, X0, 1.78, [1.55, 0.95, 1.55, 0.95], rows, row_h=0.32, size=10, num_cols={1, 3}, label_cols=0,
          bold_rows=[6, 11, 12], fill_rows={6: GRAY, 11: GRAY, 12: LBLUE}, merge=[(0, 0, 0, 3)])
# 負債側の小計行を太字に
for r_ in (4, 6, 7, 10):
    for j in (2, 3):
        c_ = t.table.cell(r_, j)
        c_.text_frame.paragraphs[0].runs[0].font.bold = True
        c_.fill.solid()
        c_.fill.fore_color.rgb = GRAY
land_ev = M.rv('土地', '土地', 'G')
bld_ev = M.rv('土地', '建物', 'G')
na_fair = M.rv('BSハイライト', '時価純資産', 'F')
loans = M.wb['借入金']
comment_box(s, 5.6, 1.62, 5.7, 3.75, [
    ('現預金・保険積立金', [f'現預金{fk(bs("現預金（保険積立除く）"))}千円（うち定期預金{fk(M.rv("BS(借方)", "定期預金", "F"))}千円）。',
                       f'生命保険積立金{fk(bs("保険積立金"))}千円は決算書上「現預金（諸掛金）」に含めて表示されている。']),
    ('固定資産', [f'越谷工場の土地{fk(bs("土地"))}千円（固定資産税評価額{fk(land_ev)}千円、2,145㎡）、建物{fk(M.rv("BS(借方)", "建物", "F"))}千円（同{fk(bld_ev)}千円）。',
              f'工場建物（1991年建設）は会計上の減価償却を行っておらず、当期の償却不足額は{fk(M.rv("減価償却", "建物", "G"))}千円。',
              f'機械{fk(M.rv("BS(借方)", "機械", "F"))}千円は製氷設備・冷凍機等（127件）。']),
    ('借入金', [f'金融機関6行庫で{fk(bs("借入金（短期・長期）") - 14000000)}千円、代表者借入14,000千円。',
             '日本政策金融公庫161,252千円（うち資本性ローン105,420千円）、栃木銀行114,101千円、商工中金79,327千円等。',
             '代表者が連帯保証・個人所有土地の担保提供を行っている。']),
    ('純資産', [f'簿価純資産{fk(bs("純資産合計"))}千円。建物を固定資産税評価額、土地を同評価額÷0.7で評価替えした場合の時価純資産は{fk(na_fair)}千円と試算。']),
], size=9)
debt4 = [rk(M.rv('BSハイライト', '借入金（短期・長期）', c)) for c in C4]
na4 = [rk(M.rv('BSハイライト', '純資産合計', c)) for c in C4]
sub_head(s, 5.6, 5.45, 5, '借入金・純資産の推移（千円）', size=10)
bar_line_chart(s, 5.6, 5.72, 5.7, 1.45, PS, [('借入金', debt4)], [('純資産（右軸）', na4, '0889C9', '#,##0')],
               number_format='#,##0', bar_colors=(NAVY,), font_size=7.5, legend=True, gap=120)

# ============================================================ 9 譲渡希望条件
s = D.use(9)
for sh in s.shapes:
    if not sh.has_text_frame:
        continue
    if sh.name == 'テキスト ボックス 18':
        set_runs_text(sh.text_frame, '株式100％譲渡')
    elif sh.name == 'テキスト ボックス 55':
        set_runs_text(sh.text_frame, '応相談')
    elif sh.name == 'テキスト ボックス 1':
        set_runs_text(sh.text_frame, '後継者不在')
    elif sh.name == 'テキスト ボックス 255':
        tf = sh.text_frame
        items = ['従業員の継続雇用', '代表者は引継ぎ完了後に退任予定（引継ぎ期間・条件は譲受企業と協議）',
                 '代表者の連帯保証・個人所有土地の担保提供の解除', '代表者借入金（1,400万円）の取扱い・代表者所有駐車場の賃貸借は要協議']
        paras = tf.paragraphs
        for i, p in enumerate(paras):
            if i < len(items):
                p.runs[0].text = items[i]
                for r_ in p.runs[1:]:
                    r_._r.getparent().remove(r_._r)
            else:
                p._p.getparent().remove(p._p)

# ============================================================ 10 中扉 事業概要
D.use(10)

# ============================================================ 11 事業内容（商流）
s = D.new('事業内容（商流）', BZ,
          '包装資材・商品を仕入れ、越谷工場で製造した純氷を加工・袋詰めし、首都圏の食品卸・氷卸・業務用顧客へ配送。配送は自社便と運送会社への委託を併用する。')
for xx, w_, ttl in [(X0, 2.45, '主要仕入先'), (3.1, 5.2, '対象会社'), (8.55, 2.75, '主要販売先')]:
    rect(s, xx, 1.6, w_, 4.45, fill=GRAY if ttl != '対象会社' else LBLUE)
    line(s, xx, 1.6, xx + w_, 1.6, NAVY, 3)
    text_box(s, xx, 1.66, w_, 0.3, [[(ttl, {'bold': True, 'size': 13, 'color': NAVY})]], align=PP_ALIGN.CENTER)
sup = [('包装資材（段ボール・袋）', '鈴木段ボール工業(株)\n大三紙業(株)／(株)シモジマ\n(株)一陽パック'),
       ('商品（ドライアイス等）', '日本液炭(株)'),
       ('原水・電力', '越谷松伏水道企業団\n新電力（電力小売）')]
for i, (h, b) in enumerate(sup):
    yy = 2.05 + i * 1.3
    rect(s, X0 + 0.1, yy, 2.25, 0.28, fill=NAVY, text=h, size=8.5, color=WHITE, bold=True)
    rect(s, X0 + 0.1, yy + 0.28, 2.25, 0.85, fill=WHITE, text=b, size=8.5)
cust = [('食品卸', '(株)日本アクセス\n三菱食品(株)\n(株)日本橋冷凍手島商店\n金木屋商事(株)'),
        ('氷卸・業務用', '(株)クールテックサガワ\n(株)ニチレイ・アイス\n(株)小野田商店'),
        ('その他', 'スーパー・飲食店・イベント等')]
for i, (h, b) in enumerate(cust):
    yy = 2.05 + [0, 1.45, 2.65][i]
    hh = [1.05, 0.85, 0.6][i]
    rect(s, 8.65, yy, 2.55, 0.28, fill=NAVY, text=h, size=8.5, color=WHITE, bold=True)
    rect(s, 8.65, yy + 0.28, 2.55, hh, fill=WHITE, text=b, size=8.5)
# 対象会社内部
rect(s, 3.25, 2.05, 4.9, 0.3, fill=NAVY, text='越谷工場（埼玉県越谷市・平方工業団地）', size=10, color=WHITE, bold=True)
proc = ['純水製造', 'アイス缶\n凍結', '揚氷\n（角氷）', '加工・袋詰', '冷蔵保管']
for i, p_ in enumerate(proc):
    rect(s, 3.3 + i * 0.97, 2.45, 0.9, 0.62, fill=WHITE, line=BLUE, text=p_, size=8.5, bold=True, color=NAVY)
    if i < 4:
        arrow(s, 3.3 + i * 0.97 + 0.9, 2.68, 0.07, 0.14, fill=BLUE, direction='right')
text_box(s, 3.3, 3.12, 4.8, 0.4, '製氷設備4台（日産40トン）・袋詰砕氷設備（日産2,500ケース）・冷蔵設備（F級1,000トン）', size=8, color=DGRAY, align=PP_ALIGN.CENTER)
rect(s, 3.25, 3.6, 2.4, 0.3, fill=NAVY, text='松戸営業所（千葉県松戸市）', size=9, color=WHITE, bold=True)
rect(s, 3.25, 3.9, 2.4, 0.62, fill=WHITE, text='角氷・加工氷・凍氷雑品の販売拠点\n従業員5名', size=8.5)
rect(s, 5.75, 3.6, 2.4, 0.3, fill=NAVY, text='商品販売・冷蔵保管', size=9, color=WHITE, bold=True)
rect(s, 5.75, 3.9, 2.4, 0.62, fill=WHITE, text='ドライアイス等の商品販売\n冷蔵庫での寄託物品の保管', size=8.5)
# 売上構成
sales_tbl = [['拠点別売上（26/3期）', '千円', '構成比']]
tot = sum(x[3] for x in SITE_SALES_67)
for nm, _, _, sv, _, _, _ in SITE_SALES_67[1:]:
    sales_tbl.append([nm, f'{rk(sv):,}', fp(sv / tot)])
table(s, 3.25, 4.65, [2.4, 1.3, 1.2], sales_tbl, row_h=0.26, size=8.5, num_cols={1, 2}, hdr_size=8.5)
text_box(s, 3.25, 5.48, 4.9, 0.45, '出所：売上高等の事業所別内訳書（第67期）。登記上の本店（越谷市蒲生寿町）は管理のみ', size=7.5, color=DGRAY)
# 矢印とラベル
arrow(s, 2.84, 3.25, 0.24, 0.36, fill=BLUE)
text_box(s, 2.83, 2.25, 0.26, 0.9, '仕\n入', size=8.5, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
arrow(s, 8.31, 3.25, 0.24, 0.36, fill=BLUE)
text_box(s, 8.3, 2.25, 0.26, 0.9, '販\n売', size=8.5, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
# 下段：外注・委託
rect(s, 3.1, 6.2, 5.2, 0.9, fill=GRAY)
line(s, 3.1, 6.2, 8.3, 6.2, NAVY, 3)
text_box(s, 3.1, 6.23, 5.2, 0.26, [[('主要外注・委託先', {'bold': True, 'size': 11, 'color': NAVY})]], align=PP_ALIGN.CENTER)
text_box(s, 3.2, 6.52, 5.0, 0.55, ['配送委託：倉橋運送・鴻池運輸(株) 等（運賃 26/3期 27,570千円）', '人材派遣：(株)ヒューマンアイ・(株)スタッフサービス 等（繁忙期の製造要員）'], size=8.5)
arrow(s, 5.55, 6.07, 0.3, 0.13, fill=BLUE, direction='tri')
text_box(s, 8.55, 6.2, 2.75, 0.9, ['【取引条件】', '販売：月末締め翌月末回収（掛売99%）', '仕入・外注：月末締め翌月20日払い'], size=8, color=DGRAY)
text_box(s, X0, 6.2, 2.45, 0.9, ['※(株)日本橋冷凍手島商店は販売先かつ資材の仕入先。', '※(株)小野田商店との資本・人的関係は要確認。'], size=7.5, color=DGRAY)

# ============================================================ 12 純氷の製造工程
s = D.new('純氷の製造工程', BZ,
          'ろ過した純水をアイス缶に注ぎ、低温のブラインで約48時間かけてゆっくり凍らせる伝統製法を採用。空気で撹拌しながら不純物を中心に集めて除去するため、透明度が高く硬く溶けにくい純氷ができる。')
steps = [('①', '原水・純水製造', '原水を自社の純水装置でろ過。1964年にイオン交換方式、1990年に逆浸透膜（RO）ろ過装置を導入'),
         ('②', '注水', '純水をアイス缶（製氷缶）へ注水し、缶グリッドごとブライン槽へ'),
         ('③', '凍結', '約－12℃のブラインで約48時間かけて凍結。空気撹拌で不純物を中央に集めて吸い取る'),
         ('④', '揚氷・脱缶', '約135kgの角氷として揚氷。氷の基本となる製品'),
         ('⑤', '加工・袋詰', 'カット・砕氷・袋詰（キングアイス1.1kg〜）、手切りのダイヤアイス、板氷 等'),
         ('⑥', '保管・出荷', '冷蔵設備で保管し、自社便・委託便で配送')]
for i, (no, h, b) in enumerate(steps):
    xx = X0 + i * 1.83
    rect(s, xx, 1.65, 1.78, 0.48, fill=NAVY if i % 2 == 0 else BLUE, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON,
         text=f'{no} {h}', size=9.5, color=WHITE, bold=True)
    photo(s, f'proc_{i + 1}', xx + 0.05, 2.25, 1.65, 1.2, h + '（king-ice.com/junhyo.html）', IMG)
    rect(s, xx + 0.05, 3.55, 1.65, 1.35, fill=WHITE, line=LINE, text=b, size=8.5, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
sub_head(s, X0, 5.08, 6, '主要設備（法人事業概況説明書）')
eq = [['設備', '台数', '能力']] + [list(e) for e in EQUIPMENT_MAIN]
table(s, X0, 5.42, [1.7, 0.8, 1.6], eq, row_h=0.3, size=9.5)
rect(s, 4.85, 5.1, 6.45, 1.95, fill=LBLUE)
text_box(s, 5.0, 5.18, 6.2, 1.85, [
    [('純氷の特徴（同社公式サイト等の公開情報）', {'bold': True, 'size': 10.5, 'color': NAVY})],
    [('・結晶の並びが規則的で透明度が高く、硬く割れにくい', {'size': 9.5})],
    [('・熱が伝わりにくく溶けにくい、溶け方が一定で飲料の味を損ねにくい', {'size': 9.5})],
    [('・料亭・日本料理店などの飲食店や、ぶっかき氷・ロックアイス用途で使用', {'size': 9.5})],
    [('・自社ブランド「キングアイス」（1.1kg／2kg／3kg袋、板氷、かき氷用キューブ 等）', {'size': 9.5})]])
footnote(s, '※原水の取水方法（上水道か地下水か）、製氷缶の数量・1回あたり生産量、稼働率は要確認。', y=7.2)

# ============================================================ 13 月別売上（季節性）
mws = M.wb['月次売上']
rm = M.ref['月次売上']
months = [mws.cell(row=4, column=3 + j).value for j in range(12)]
sales_m = [mws.cell(row=rm['凍氷'], column=3 + j).value for j in range(12)]
store_m = [mws.cell(row=rm['保管料'], column=3 + j).value for j in range(12)]
tot_m = mws.cell(row=rm['凍氷'], column=15).value
summer = sum(sales_m[3:5])
s = D.new('月別売上高（季節性）', BZ,
          f'売上は気温に連動する季節性が高く、7〜8月の2か月で年間の{summer / tot_m * 100:.1f}%を占める。冬季（11〜2月）は売上が夏季の約1/4に落ち込む一方、冷蔵保管料が増加し閑散期の設備稼働を補っている。')
bar_line_chart(s, X0, 1.6, 7.0, 3.6, months, [('売上（凍氷）', sales_m)], [('冷蔵保管料（右軸）', store_m, '0889C9', '#,##0')],
               number_format='#,##0', bar_colors=(NAVY,))
unit(s, 5.5, 1.45, 1.8)
rows = [['', *months, '合計']]
for k, nm in [('凍氷', '売上（凍氷）'), ('保管料', '保管料'), ('原材料仕入', '原材料仕入'), ('人件費', '人件費'), ('従事員数', '従事員数（人）')]:
    vals = [mws.cell(row=rm[k], column=3 + j).value for j in range(12)]
    tot = mws.cell(row=rm[k], column=15).value
    rows.append([nm, *[f'{v:,}' for v in vals], f'{tot:,}' if tot else '-'])
table(s, X0, 5.35, [1.25] + [0.62] * 12 + [0.78], rows, row_h=0.27, size=8, num_cols=set(range(1, 14)), hdr_size=8, bold_rows=[1])
comment_box(s, 7.55, 1.6, 3.75, 3.6, [
    ('季節性', [f'7月{sales_m[3]:,}千円・8月{sales_m[4]:,}千円がピーク。6〜9月の4か月で年間の{sum(sales_m[2:6]) / tot_m * 100:.1f}%。',
             f'11〜2月は月14〜20百万円程度。']),
    ('費用・人員', ['人件費は7月・12月に賞与支給で増加。従事員数は23〜26名でほぼ一定であり、繁忙期は人材派遣で補完。']),
    ('冷蔵保管', [f'保管料は11〜12月に増加（12月{store_m[8]:,}千円）。年末商戦期の食品保管需要と推察。']),
], size=8.5)
footnote(s, '※出所：第67期（2026年3月期）法人事業概況説明書。単位：千円。', y=7.25)

# ============================================================ 14 主要販売先
tws = M.wb['取引先']
rt = M.ref['取引先']
s = D.new('主要販売先（期末売掛金残高）', BZ,
          '大手食品卸の(株)日本アクセス・三菱食品(株)をはじめ、食品卸・氷卸との継続的な取引関係を有する。上位7社で期末売掛金の約7割を占め、特定先への過度な依存はない。')
unit(s, 9.4, 1.48, 1.9)
rows = [['取引先', *[f'{p}\n残高' for p in PS], '26/3期\n構成比']]
order_ = ['(株)日本橋冷凍手島商店', '(株)日本アクセス', '三菱食品(株)', '金木屋商事(株)', '(株)小野田商店', '(株)クールテックサガワ', '(株)ニチレイ・アイス', '(株)双日食品']
for i, nm in enumerate(order_):
    r_ = rt[nm]
    rows.append([f'{i + 1}. {nm}', *[fk(tws.cell(row=r_, column=c).value) for c in (3, 5, 7, 9)], fp(tws.cell(row=r_, column=10).value)])
rows.append(['その他（複数社）', *[fk(tws.cell(row=rt['その他'], column=c).value) for c in (3, 5, 7, 9)], fp(tws.cell(row=rt['その他'], column=10).value)])
rows.append(['期末売掛金 合計', *[fk(tws.cell(row=rt['合計'], column=c).value) for c in (3, 5, 7, 9)], '100.0%'])
table(s, X0, 1.72, [3.3, 1.45, 1.45, 1.45, 1.45, 1.3], rows, row_h=0.36, size=10, num_cols={1, 2, 3, 4, 5},
      bold_rows=[len(rows) - 1], fill_rows={len(rows) - 1: LBLUE}, hdr_size=9)
top7 = sum(tws.cell(row=rt[n], column=9).value for n in order_[:7]) / tws.cell(row=rt['合計'], column=9).value
footnote(s, ['※取引先別の売上高データは未受領のため、期末売掛金残高（月末締め翌月末回収）を取引規模の代理指標として表示。取引先別売上高はデータルームにて開示予定。',
             f'※26/3期の上位7社の残高構成比は{top7 * 100:.1f}%。23/3〜25/3期は受領資料、26/3期は勘定科目内訳書による。',
             '※26/3期の内訳書は上位先の合計が残高を1,724千円上回るため、その他は残高との差額で表示（(株)小野田商店の計上区分は要確認）。'], y=6.25)

# ============================================================ 15 主要仕入先・経費支払先
uws = M.wb['仕入先']
ru = M.ref['仕入先']
s = D.new('主要仕入先・主要経費', BZ,
          '仕入は包装資材（段ボール・袋）とドライアイス等の商品が中心。原価の主要項目は人件費・電力料・原材料で、販管費では配送運賃の比重が大きい。')
sub_head(s, X0, 1.55, 5, '主要仕入先（26/3期末 買掛金残高）')
rows = [['仕入先', '内容', '残高(千円)', '構成比']]
r_ = 5
while uws.cell(row=r_, column=2).value and r_ < ru['ap_total']:
    rows.append([uws.cell(row=r_, column=2).value, uws.cell(row=r_, column=3).value, fk(uws.cell(row=r_, column=4).value), fp(uws.cell(row=r_, column=5).value)])
    r_ += 1
rows.append(['買掛金 合計', '', fk(uws.cell(row=ru['ap_total'], column=4).value), '100.0%'])
table(s, X0, 1.88, [2.3, 1.75, 0.95, 0.7], rows, row_h=0.32, size=9, num_cols={2, 3}, bold_rows=[len(rows) - 1], fill_rows={len(rows) - 1: LBLUE})
sub_head(s, 6.3, 1.55, 5, '主要経費支払先（26/3期末 未払金残高）')
rows = [['支払先', '内容', '残高(千円)']]
r_ = ru['payees']
for _ in range(8):
    rows.append([uws.cell(row=r_, column=2).value, uws.cell(row=r_, column=3).value, fk(uws.cell(row=r_, column=4).value)])
    r_ += 1
table(s, 6.3, 1.88, [2.6, 1.45, 0.95], rows, row_h=0.32, size=8.5, num_cols={2})
sub_head(s, X0, 5.0, 6, '主要費用の推移（千円）')
cost = [['費目', *PS]]
for nm, f in [('原材料仕入高（包装資材等）', lambda c: M.rv('CR', '原材料費/原材料仕入高', c)),
              ('労務費＋販管人件費', lambda c: M.rv('CR', 'g_労務費', c) + M.rv('SGA', 'g_給与関係費', c)),
              ('電力料（動力費・製造＋販管）', lambda c: M.rv('CR', '製造経費/動力費', c) + M.rv('SGA', '資産管理費/動力費', c)),
              ('配送運賃', lambda c: M.rv('SGA', '輸送関係費/運賃', c)),
              ('外注作業費', lambda c: M.rv('CR', '外注加工費/外注作業費', c))]:
    cost.append([nm, *[fk(f(c)) for c in C4]])
table(s, X0, 5.32, [3.0, 1.25, 1.25, 1.25, 1.25], cost, row_h=0.28, size=9, num_cols={1, 2, 3, 4})
comment_box(s, 8.55, 5.0, 2.75, 2.1, [('コスト構造', ['電力料は23/3期に42,894千円まで上昇後、24/3期以降は32〜34百万円台。', '配送運賃は売上拡大に伴い増加傾向。'])], size=8, title='Point', title_size=11)

# ============================================================ 16 特徴・強み
s = D.use(16)
texts = {'テキスト ボックス 13': ('65年超の業歴を持つ純氷メーカー・自社ブランド「キングアイス」',
                              '1960年創業。自社の純水設備（イオン交換・RO膜）でろ過した水を、アイス缶で約48時間かけて凍結する伝統製法により、透明度が高く溶けにくい純氷を製造。角氷から袋氷・手切りキューブまで幅広い商品を自社で加工する。'),
         'テキスト ボックス 19': ('大手食品卸を含む首都圏の安定した販路と一貫供給体制',
                              '(株)日本アクセス・三菱食品(株)等の食品卸や氷卸と継続取引（掛売99%）。製氷・加工・冷蔵保管・配送を自社で一貫して担い、夏季の需要ピークにも製氷能力日産40トン・冷蔵1,000トンの設備で対応する。'),
         'テキスト ボックス 28': ('猛暑による需要拡大と設備・物流網を活かした拡張余地（仮説）',
                              '夏季需要が売上の約3割を占め、猛暑の長期化は追い風。冬季は冷蔵保管で稼働を補完しており、物流・食品関連の買手の販路・配送網との連携により、閑散期の稼働向上や取扱商品拡大の余地があると想定。')}
for sh in s.shapes:
    if sh.name in texts and sh.has_text_frame:
        h, b = texts[sh.name]
        ps = sh.text_frame.paragraphs
        ps[0].runs[0].text = h
        for r_ in ps[0].runs[1:]:
            r_._r.getparent().remove(r_._r)
        ps[1].runs[0].text = b
        for r_ in ps[1].runs[1:]:
            r_._r.getparent().remove(r_._r)
        for p_ in ps[2:]:
            p_._p.getparent().remove(p_._p)
    if sh.name == 'タイトル 2':
        set_runs_text(sh.text_frame, '特徴・強み・成長性')

# ============================================================ 17 中扉
D.use(20)

# ============================================================ 18 組織図
s = D.new('組織図', ORG, '代表取締役の下、工場長が越谷工場・松戸営業所の現場を統括する体制。経理・総務は一部を外部（会計事務所）に委託している。')
rect(s, 3.6, 1.6, 2.6, 0.5, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, text='株主総会', size=12, color=WHITE, bold=True)
rect(s, 3.6, 2.35, 2.6, 0.62, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=[[('代表取締役', {'size': 12, 'bold': True, 'color': WHITE})], [('小野田 洋', {'size': 10, 'color': WHITE})]])
rect(s, 3.6, 3.22, 2.6, 0.55, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=[[('工場長', {'size': 11, 'bold': True, 'color': WHITE})], [('S氏（本社・工場・松戸営業所の責任者）', {'size': 7.5, 'color': WHITE})]])
for y1, y2 in [(2.1, 2.35), (2.97, 3.22), (3.77, 4.05)]:
    line(s, 4.9, y1, 4.9, y2, NAVY, 1.25)
line(s, 1.15, 4.05, 9.5, 4.05, NAVY, 1.25)
depts = [('財務・経理・総務', '2名（C氏・S氏）\n※一部外部委託', 0.35), ('営業部門', '1名（M氏）', 2.55),
         ('越谷工場', '製造・配送・品質検査', 4.75), ('松戸営業所', '責任者（M氏）\n従業員5名', 8.4)]
for nm, sub, xx in depts:
    w_ = 3.3 if nm == '越谷工場' else 1.95
    line(s, xx + w_ / 2, 4.05, xx + w_ / 2, 4.3, NAVY, 1.25)
    rect(s, xx, 4.3, w_, 0.4, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=nm, size=10.5, color=WHITE, bold=True)
    if nm != '越谷工場':
        rect(s, xx, 4.75, w_, 0.55, fill=WHITE, line=LINE, text=sub, size=8.5)
sub3 = [('製造部門', '揚氷（H氏）\n加工（M氏）'), ('配送部門', 'K氏'), ('品質検査', '代表者が担当')]
for i, (nm, sub) in enumerate(sub3):
    xx = 4.75 + i * 1.12
    line(s, xx + 0.5, 4.7, xx + 0.5, 4.85, NAVY, 1)
    rect(s, xx, 4.85, 1.05, 0.34, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=nm, size=8.5, color=WHITE, bold=True)
    rect(s, xx, 5.22, 1.05, 0.52, fill=WHITE, line=LINE, text=sub, size=8)
rect(s, 7.0, 1.6, 4.3, 1.75, fill=LBLUE)
text_box(s, 7.15, 1.68, 4.05, 1.62, [
    [('【人員構成】（2026年3月期末）', {'bold': True, 'size': 11, 'color': NAVY})],
    [('　役員　　　　　　：1名（代表取締役）', {'size': 10})],
    [('　従業員（越谷工場）：19名', {'size': 10})],
    [('　従業員（松戸営業所）：5名', {'size': 10})],
    [('　合計　　　　　　：25名', {'size': 10, 'bold': True})],
    [('　※月別従事員数は23〜26名で推移', {'size': 8.5, 'color': DGRAY})]])
footnote(s, ['※図中の氏名はイニシャル表記。役割分担は受領資料に基づく。品質検査部門は代表者が担当しているため、譲渡後の引継ぎ方法は要確認。',
             '※従業員一覧（年齢・勤続年数・資格・雇用形態）は未受領のため、データルームにて開示予定。'], y=6.05)

# ============================================================ 19 株主及び役員
s = D.new('株主及び役員', ORG, '代表取締役の小野田洋氏が発行済株式の100%を保有するオーナー企業。取締役会・監査役は設置しておらず、株式の譲渡には株主総会の承認を要する。')
sub_head(s, X0, 1.6, 5, '株主')
table(s, X0, 1.95, [0.5, 2.2, 1.8, 1.6, 1.0, 3.8], [['No', '氏名', '役職', '保有株式数', '比率', '備考'],
                                                  ['1', '小野田 洋', '代表取締役', '22,000株', '100.0%', '株式数は登記記載（要確認※）'],
                                                  ['', '合計', '', '22,000株', '100.0%', '']],
      row_h=0.34, size=10, num_cols={3, 4}, label_cols=0, bold_rows=[2], fill_rows={2: LBLUE})
sub_head(s, X0, 3.25, 5, '役員')
ows = M.wb['取締役']
table(s, X0, 3.6, [0.5, 2.2, 1.8, 1.6, 2.0, 2.8], [['No', '氏名', '役職', '役員報酬(年)', 'M&A後の進退', '備考'],
                                                  ['1', '小野田 洋', '代表取締役', f'{fk(ows["E5"].value)}千円', '引継ぎ後に退任予定', '2020年5月重任']],
      row_h=0.34, size=10, num_cols={3}, label_cols=0)
rect(s, X0, 4.55, 5.2, 1.1, fill=GRAY)
text_box(s, X0 + 0.12, 4.6, 5.0, 1.05, ['株券発行有無：株券発行会社（定款の定め）', '譲渡承認機関：株主総会', '取締役会　　：設置なし', '監査役　　　：設置なし'], size=10)
footnote(s, ['※発行済株式数は登記上22,000株。決算書の株主資本等変動計算書注記（200株）・法人税申告書別表二（220株）と記載が異なるため、株主名簿・定款で確認中（要確認）。',
             '※株券発行会社のため、株式譲渡には株券の交付を要する。株券の現存状況は要確認。'], y=6.4)

# ============================================================ 20 拠点概要
s = D.new('拠点概要', ORG, '越谷市平方工業団地内の自社工場（土地2,145㎡・建物約2,094㎡）を中核拠点とし、千葉県松戸市に販売拠点の営業所を賃借している。')
rows = [['拠点', '所在地', '用途', '所有／賃借', '従業員', '延床面積']]
rows += [['越谷工場', '埼玉県越谷市大字平方999-4', '製氷・加工・冷蔵保管・配送', '所有（土地・建物）', '19名', '2,065㎡'],
         ['松戸営業所', '千葉県松戸市稔台5-14-3', '角氷・加工氷等の販売', '賃借（年1,200千円）', '5名', '400㎡'],
         ['本店（登記）', '埼玉県越谷市蒲生寿町19-15', '管理（代表者住所）', '—', '1名', '10㎡'],
         ['駐車場', '埼玉県越谷市蒲生寿町19-15', '駐車場', '代表者から賃借（年545千円）', '—', '—']]
table(s, X0, 1.62, [1.4, 2.9, 2.3, 2.05, 0.8, 1.5], rows, row_h=0.34, size=9.5, num_cols={4, 5})
photo(s, 'factory', X0, 3.45, 5.35, 3.4, '越谷工場（外観）（king-ice.com/corporate.html）', IMG)
photo(s, 'map', 5.95, 3.45, 5.35, 3.4, '所在地地図（越谷工場・松戸営業所）', IMG)
footnote(s, '※延床面積は事業所別内訳書の記載。越谷工場建物は登記上2,093.87㎡で差異あり（要確認）。本店所在地は代表者住所と同一で、実務上の拠点は越谷工場。', y=7.05)

# ============================================================ 21 不動産概況
s = D.new('不動産概況', ORG, f'越谷工場の土地・建物を会社所有。固定資産税評価額は土地{fk(land_ev)}千円・建物{fk(bld_ev)}千円で、工場は日本政策金融公庫の担保に供されている。')
sub_head(s, X0, 1.6, 5, '所有不動産')
zws = M.wb['不動産(所有)']
rows = [['No', '種類', '区分', '所在地', '面積', '26/3期末簿価(千円)', '固定資産税評価額(千円)', '担保']]
rows.append(['1', '土地', '越谷工場', '埼玉県越谷市大字平方999-4', '2,145.18㎡', fk(zws['F5'].value), fk(zws['G5'].value), '日本公庫'])
rows.append(['2', '建物', '越谷工場', '同上', '2,093.87㎡', fk(zws['F6'].value), fk(zws['G6'].value), '日本公庫'])
rows.append(['', '合計', '', '', '', fk(zws['F7'].value), fk(zws['G7'].value), ''])
table(s, X0, 1.95, [0.45, 0.7, 1.0, 2.9, 1.15, 1.6, 1.85, 1.25], rows, row_h=0.36, size=9.5, num_cols={4, 5, 6}, label_cols=0,
      bold_rows=[3], fill_rows={3: LBLUE})
sub_head(s, X0, 3.65, 5, '賃借不動産')
rows = [['No', '区分', '所在地', '貸主', '年間賃料(千円)', '備考'],
        ['1', '店舗（松戸営業所）', '千葉県松戸市稔台5-14-3', '法人（第三者）', '1,200', ''],
        ['2', '駐車場', '埼玉県越谷市蒲生寿町19-15', '代表者', '545', '代表者からの賃借'],
        ['', '合計', '', '', '1,745', '']]
table(s, X0, 4.0, [0.45, 1.8, 2.9, 1.6, 1.4, 2.75], rows, row_h=0.34, size=9.5, num_cols={4}, label_cols=0, bold_rows=[3], fill_rows={3: LBLUE})
footnote(s, ['※固定資産税評価額は2025年3月期時点の資料による。最新年度の評価証明書は要確認。',
             '※建物の簿価は会計上の減価償却を行っていないため、固定資産税評価額を大きく上回る（内訳書記載の簿価280,673千円とBS残高259,242千円にも差異あり・要確認）。',
             '※代表者個人所有の土地を栃木銀行・商工組合中央金庫の借入担保に提供している（会社資産ではないため上表に含まない）。'], y=5.75)

# ============================================================ 22 保有設備
s = D.new('保有設備', ORG, '製氷設備4台・袋詰砕氷設備・冷蔵設備を保有。機械は127件・取得価額375百万円に上り、近年も製氷設備の更新・修繕投資を継続している。')
sub_head(s, X0, 1.6, 5, '主要設備')
table(s, X0, 1.95, [2.0, 0.9, 1.9, 5.0], [['設備', '台数', '能力', '備考']] + [[*e, '法人事業概況説明書（第67期）'] for e in EQUIPMENT_MAIN],
      row_h=0.34, size=10)
sub_head(s, X0, 3.5, 5, '償却資産の状況（第67期 減価償却計算書）')
dws = M.wb['減価償却']
rows = [['科目', '件数', '取得価額', '期末簿価', '当期償却額', '当期償却不足額']]
for r_ in range(5, 10):
    rows.append([dws.cell(row=r_, column=2).value, f'{dws.cell(row=r_, column=3).value}件'] + [fk(dws.cell(row=r_, column=c).value) for c in (4, 5, 6, 7)])
rr = M.row('減価償却', '合計')
rows.append(['合計', '185件'] + [fk(dws.cell(row=rr, column=c).value) for c in (4, 5, 6, 7)])
table(s, X0, 3.85, [1.8, 0.9, 1.7, 1.7, 1.7, 1.7], rows, row_h=0.32, size=10, num_cols={1, 2, 3, 4, 5}, bold_rows=[6], fill_rows={6: LBLUE})
unit(s, 7.8, 3.6, 1.7)
footnote(s, ['※建物（越谷工場、1991年3月取得）は当期の償却を行っておらず、当期償却不足額21,745千円。償却資産の明細（185件）はデータルームにて開示予定。',
             '※主要設備の更新時期・修繕計画は要確認。'], y=6.25)

# ============================================================ 23 許認可・取引金融機関
s = D.new('許認可・取引金融機関', ORG, '氷雪製造業・氷雪販売業の食品営業許可を取得。借入は日本政策金融公庫・地方銀行・信用金庫等の6行庫で、代表者が連帯保証・担保提供を行っている。')
sub_head(s, X0, 1.6, 5, '許認可')
table(s, X0, 1.95, [3.0, 1.9, 1.6, 4.4], [['名称', '番号', '許可日', '備考']] + [list(l) for l in LICENSES], row_h=0.34, size=9.5)
sub_head(s, X0, 3.15, 5, '借入金の状況（2026年3月期末）')
lws = M.wb['借入金']
rows = [['借入先', '期末残高(千円)', '支払利息(千円)', '担保・備考']]
for r_ in range(5, M.row('借入金', '合計')):
    nm = lws.cell(row=r_, column=2).value
    if '保証協会' in nm:
        continue
    rows.append([nm, fk(lws.cell(row=r_, column=3).value), fk(lws.cell(row=r_, column=4).value), lws.cell(row=r_, column=5).value])
rows.append(['保証料（埼玉県信用保証協会）', '-', fk(559650), '保証付融資あり'])
rr = M.row('借入金', '合計')
rows.append(['合計', fk(lws.cell(row=rr, column=3).value), fk(lws.cell(row=rr, column=4).value), ''])
table(s, X0, 3.5, [3.0, 1.5, 1.5, 4.9], rows, row_h=0.3, size=9.5, num_cols={1, 2}, bold_rows=[len(rows) - 1], fill_rows={len(rows) - 1: LBLUE})

# ============================================================ 24 中扉 財務
D.use(55)


# ============================================================ 25-28 財務詳細（4期）
def fin_rows(sheet, keys_labels, cols=C4):
    out = []
    for key, lab in keys_labels:
        out.append([lab, *[fk(M.rv(sheet, key, c)) for c in cols]])
    return out


# 貸借対照表
s = D.new('貸借対照表', FIN, '借入金の返済進捗と利益の積上げにより、純資産は4期で19,818千円から49,288千円へ増加。')

A = [('sub_cash', '（現金・預金）'), ('現金', '現金'), ('当座預金', '当座預金'), ('普通預金', '普通預金'), ('定期預金', '定期預金'), ('定期積金', '定期積金'),
     ('諸掛金（生命保険積立）', '諸掛金（保険積立）'), ('sub_quick', '（当座資産）'), ('売掛金', '売掛金'), ('sub_inv', '（棚卸資産）'), ('棚卸商品', '商品'),
     ('棚卸原材料', '原材料'), ('仕掛品', '仕掛品'), ('sub_oca', '（流動資産諸口）'), ('前渡金', '前渡金'), ('貸付金', '貸付金'), ('預け金', '預け金'),
     ('未収金', '未収金'), ('前払費用', '前払費用'), ('流動資産合計', '流動資産合計'), ('sub_ppe', '（有形固定資産）'), ('建物', '建物'), ('構築物', '構築物'),
     ('機械', '機械'), ('車両運搬具', '車両運搬具'), ('器具備品', '器具備品'), ('土地', '土地'), ('sub_intan', '（無形固定資産）'), ('sub_invest', '（投資等）'),
     ('固定資産合計', '固定資産合計'), ('資産合計', '資産合計')]
rows = [['資産の部', *PS]] + fin_rows('BS(借方)', A)
bolds = [i + 1 for i, (k, _) in enumerate(A) if k.startswith('sub_') or '合計' in k]
table(s, X0, 1.5, [1.65, 0.92, 0.92, 0.92, 0.92], rows, row_h=0.177, size=8, num_cols={1, 2, 3, 4}, bold_rows=bolds,
      fill_rows={i + 1: LBLUE for i, (k, _) in enumerate(A) if '合計' in k}, hdr_size=8)
Lb = [('sub_cl', '（流動負債）'), ('支払手形', '支払手形'), ('買掛金', '買掛金'), ('短期借入金', '短期借入金'), ('未払金', '未払金'), ('預り金', '預り金'),
      ('仮受金', '仮受金'), ('税金引当金（未払法人税等）', '未払法人税等'), ('sub_fl', '（固定負債）'), ('長期借入金', '長期借入金'), ('負債合計', '負債合計'),
      ('資本金', '資本金'), ('利益準備金', '利益準備金'), ('別途積立金', '別途積立金'), ('繰越利益剰余金', '繰越利益剰余金'), ('純資産合計', '純資産合計'),
      ('負債純資産合計', '負債・純資産合計')]
rows = [['負債・純資産の部', *PS]] + fin_rows('BS (貸方)', Lb)
bolds = [i + 1 for i, (k, _) in enumerate(Lb) if k.startswith('sub_') or '合計' in k]
table(s, 5.85, 1.5, [1.65, 0.92, 0.92, 0.92, 0.92], rows, row_h=0.177, size=8, num_cols={1, 2, 3, 4}, bold_rows=bolds,
      fill_rows={i + 1: LBLUE for i, (k, _) in enumerate(Lb) if '合計' in k}, hdr_size=8)
unit(s, 9.5, 1.3, 1.8)
text_box(s, 5.85, 4.85, 5.33, 1.6, [[('【留意事項】', {'bold': True, 'size': 9, 'color': NAVY})],
                                    [('・諸掛金は生命保険の積立金（67期：大同生命8,462千円・第一生命901千円）。', {'size': 8})],
                                    [('・建物は会計上の減価償却を行っておらず、4期とも同額。', {'size': 8})],
                                    [('・67期末の借入金461,024千円のうち代表者借入14,000千円、日本公庫の資本性ローン105,420千円。', {'size': 8})],
                                    [('・資本金は2026年5月15日に10,000千円へ減資（登記）。', {'size': 8})],
                                    [('・23/3〜25/3期は受領資料、26/3期は決算報告書による。', {'size': 8})]])

# 損益計算書
s = D.new('損益計算書', FIN, '売上総利益率は2023年3月期の16.2%から、2024年3月期以降は22〜26%へ回復して推移。')
plws = M.wb['PL']
rp = M.ref['PL']
keys = [('売上高', '売上高'), ('売上', '　売上'), ('保管料', '　保管料'), ('受入手数料', '　受入手数料'), ('売上原価', '売上原価'), ('期首商品', '　期首商品棚卸高'),
        ('商品仕入', '　商品仕入高'), ('期末商品', '　期末商品棚卸高'), ('製造原価', '　当期製品製造原価'), ('売上総利益', '売上総利益'), ('販管費', '販売費及び一般管理費'),
        ('販管_給与関係費', '　給与関係費'), ('販管_資産管理費', '　資産管理費'), ('販管_輸送関係費', '　輸送関係費'), ('販管_交際販売関係費', '　交際販売関係費'),
        ('販管_事務処理費', '　事務処理費'), ('販管_その他', '　税金公課・減価償却費'), ('支払手数料', '　支払手数料'), ('営業利益', '営業利益'), ('営業外収益', '営業外収益'),
        ('受取利息', '　受取利息・配当金'), ('雑収入', '　雑収入'), ('営業外費用', '営業外費用'), ('支払利息', '　支払利息'), ('雑損', '　雑損'), ('経常利益', '経常利益'),
        ('特別損失', '特別損益'), ('税引前', '税引前当期純利益'), ('法人税等', '法人税等'), ('当期純利益', '当期純利益')]
rows = [['科目'] + sum([[p, '売上比'] for p in PS], [])]
bolds, fills = [], {}
for key, lab in keys:
    vals = []
    for vc in (3, 5, 7, 9):
        v = plws.cell(row=rp[key], column=vc).value
        if key == '受取利息':
            v += plws.cell(row=rp['受取配当金'], column=vc).value
        if key == '特別損失':
            v = plws.cell(row=rp['特別利益'], column=vc).value - v
        ratio = v / plws.cell(row=rp['売上高'], column=vc).value
        vals += [fk(v), fp(ratio)]
    rows.append([lab, *vals])
    if not lab.startswith('　'):
        bolds.append(len(rows) - 1)
    if key in ('売上総利益', '営業利益', '経常利益', '当期純利益'):
        fills[len(rows) - 1] = LBLUE
t = table(s, X0, 1.5, [2.1] + [1.05, 0.62] * 4, rows, row_h=0.178, size=8, num_cols=set(range(1, 9)), bold_rows=bolds, fill_rows=fills, hdr_size=8)
for i in range(1, len(rows)):
    for j in (2, 4, 6, 8):
        c_ = t.table.cell(i, j).text_frame.paragraphs[0].runs[0]
        c_.font.size = Pt(7)
        c_.font.color.rgb = DGRAY
unit(s, 7.6, 1.3, 1.8)
footnote(s, '※65期の雑収入8,482千円・雑損5,192千円、66〜67期の支払手数料の内容は要確認。67期の雑収入には保険満期返戻金2,296千円を含む。', y=7.2)

# 製造原価報告書
s = D.new('製造原価報告書', FIN, '製造原価の主要項目は労務費・原材料（包装資材）・電力料（動力費）・減価償却費。')
crws = M.wb['CR']
rc = M.ref['CR']
keys = []
from data import CR as CRDEF
for g, items in CRDEF:
    keys.append(('g_' + g, g))
    for it, _ in items:
        if g == '製造経費' and all(v == 0 for v in dict(items)[it]):
            continue
        keys.append((g + '/' + it, '　' + it))
keys.append(('製造原価', '当期製品製造原価'))
rows = [['科目', *PS]]
bolds, fills = [], {}
for key, lab in keys:
    rows.append([lab, *[fk(crws[f'{c}{rc[key]}'].value) for c in C4]])
    if not lab.startswith('　'):
        bolds.append(len(rows) - 1)
fills[len(rows) - 1] = LBLUE
half = 22
table(s, X0, 1.5, [2.0, 0.85, 0.85, 0.85, 0.85], rows[:half], row_h=0.25, size=8.5, num_cols={1, 2, 3, 4}, bold_rows=[b for b in bolds if b < half], hdr_size=8.5)
table(s, 5.85, 1.5, [2.0, 0.85, 0.85, 0.85, 0.85], [rows[0]] + rows[half:], row_h=0.25, size=8.5, num_cols={1, 2, 3, 4},
      bold_rows=[b - half + 1 for b in bolds if b >= half], fill_rows={len(rows) - half: LBLUE}, hdr_size=8.5)
unit(s, 9.5, 1.3, 1.8)
footnote(s, ['※原材料は段ボール・ビニール袋等の包装資材。動力費は製氷・冷蔵設備の電力料。',
             '※減価償却費は64期（23/3期）にほぼ計上なし。67期の退職給与6,178千円は臨時的な支出と推察（要確認）。'], y=6.95)

# 販管費
s = D.new('販売費及び一般管理費', FIN, '販管費の中心は配送運賃を含む輸送関係費と給与関係費。運賃は売上拡大に伴い増加傾向。')
sws = M.wb['SGA']
rs = M.ref['SGA']
from data import SGA as SGADEF
keys = []
for g, items in SGADEF:
    keys.append(('g_' + g, g))
    for it, vals in items:
        if all(v == 0 for v in vals):
            continue
        keys.append((g + '/' + it, '　' + it))
keys.append(('合計', '販売費及び一般管理費 合計'))
rows = [['科目', *PS]]
bolds = []
for key, lab in keys:
    rows.append([lab, *[fk(sws[f'{c}{rs[key]}'].value) for c in C4]])
    if not lab.startswith('　'):
        bolds.append(len(rows) - 1)
half = [i for i, r_ in enumerate(rows) if r_[0] == '輸送関係費'][0]
table(s, X0, 1.5, [2.0, 0.85, 0.85, 0.85, 0.85], rows[:half], row_h=0.25, size=8.5, num_cols={1, 2, 3, 4}, bold_rows=[b for b in bolds if b < half], hdr_size=8.5)
table(s, 5.85, 1.5, [2.0, 0.85, 0.85, 0.85, 0.85], [rows[0]] + rows[half:], row_h=0.25, size=8.5, num_cols={1, 2, 3, 4},
      bold_rows=[b - half + 1 for b in bolds if b >= half], fill_rows={len(rows) - half: LBLUE}, hdr_size=8.5)
unit(s, 9.5, 1.3, 1.8)
footnote(s, ['※地代家賃は松戸営業所の店舗賃料1,200千円と代表者からの駐車場賃借545千円。支払手数料（一般業務費）は損益計算書に別掲。'], y=7.1)

# ============================================================ Appendix
D.use(59)
s = D.new('沿革', APX, '1960年に越谷で創業。早くから純水設備を導入し、品質にこだわった純氷の製造を65年以上にわたり継続している。')
hws = M.wb['沿革']
rows = [['年月', '主な出来事']]
r_ = 5
while hws.cell(row=r_, column=2).value:
    rows.append([hws.cell(row=r_, column=2).value, hws.cell(row=r_, column=3).value])
    r_ += 1
table(s, X0, 1.65, [2.6, 8.3], rows, row_h=0.42, size=10.5)
footnote(s, '※1960〜1990年の記載は同社公式サイト等の公開情報、その他は登記・減価償却計算書による。代表者の経歴は要確認。', y=6.9)
D.use(63)
D.use(64)

D.finalize(OUT)
print('saved', OUT, len(D.order), 'slides')
