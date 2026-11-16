"""鶴岡中央青果 株式価値評価書 261019 → 261116 の差替え（賞与の正常化を廃止：依頼者方針 2026/11）

評価書261019は旧マスターから作成済みのため、テンプレートから組み直さず、賞与の正常化に連動する
頁（P1・P7・P15・P20・P22・P24・P25）の表と文言を、案件マスター261115の値で差し替える。
土地は固定資産税評価額÷0.7で減額・税効果なし（依頼者確認済）、非流動性ディスカウント30%は維持。

元の261019版は git show 5165eb1:"株式会社鶴岡中央青果/株式価値評価書_鶴岡中央青果_261019.pptx" で復元できる。
使い方: python update_vr_261116.py <評価書261019.pptx> <案件マスター.xlsx> <出力.pptx>
"""
import sys, os
from copy import deepcopy
from openpyxl import load_workbook
from pptx import Presentation

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../.claude/skills/ma-deal-docs/scripts'))
from im_paste import set_text, scrub_doc_props  # noqa: E402

src, master, out = sys.argv[1:4]
DATE = '2026年11月16日'
wb = load_workbook(master, data_only=True)
v = lambda sh, a: wb[sh][a].value
K = lambda x: (f'{round(x / 1000):,}' if round(x / 1000) >= 0 else f'△{-round(x / 1000):,}')
Y, E = '年買法', 'EV EBITDAマルチプル'

prs = Presentation(src)
S = lambda n: prs.slides[n - 1]


def shp(slide, sid):
    return next(s for s in slide.shapes if s.shape_id == sid)


def cell(tbl, r, c, text):
    set_text(tbl.cell(r, c), text)


def replace_in(slide, old, new):
    hit = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new); hit += 1
    return hit


# P1 表紙・P7 算定日
assert replace_in(S(1), '2026年10月19日', DATE) == 1
t = shp(S(7), 15).table
cell(t, 1, 1, DATE)

# P7 試算結果
lo, hi = v(Y, 'K6'), v(Y, 'M6')
set_text(shp(S(7), 30), f'{K(lo)} ～ {K(hi)}（千円）')
# P7 レンジ図（積上げ横棒：系列 min＝下限、max＝上限−下限）。キャッシュと埋込ブックの両方を更新
import io
from pptx.oxml.ns import qn
from openpyxl import load_workbook as _lw
ch = next(s for s in S(7).shapes if s.has_chart).chart
want = {'min': lo, 'max': hi - lo}
for ser in ch._chartSpace.iter(qn('c:ser')):
    name = ser.find('.//' + qn('c:tx')).find('.//' + qn('c:v')).text
    if name in want:
        ser.find('.//' + qn('c:val')).find('.//' + qn('c:pt')).find(qn('c:v')).text = repr(float(want[name]))
wbp = ch.part.chart_workbook.xlsx_part
if wbp is not None:
    xb = _lw(io.BytesIO(wbp.blob))
    for ws in xb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value in want and c.column == 2:
                    ws.cell(c.row, 3).value = want[c.value]
    buf = io.BytesIO(); xb.save(buf); wbp._blob = buf.getvalue()
set_text(shp(S(7), 35), ['※マーケットアプローチはネットデット（有利子負債−現預金等）が事業価値を上回り、算定値がマイナスとなるため参考扱い。',
                         '※賞与は規定がなく業績連動で支給義務がないため、正常収益力の算定で正常化していない（P.20参照）。'])

# P15 賞与引当金
set_text(shp(S(15), 29), ['・2024/3期6,002千円、2025/3期5,692千円の賞与を支給。2026/3期は賞与の計上なし。',
                          '・賞与規定はなく、業績に応じて支給している（代表者ヒアリング）。基準日時点で支給義務がないため、賞与引当金は計上していない（修正額0円）。',
                          '・業績連動賞与であり支給義務がないため、正常収益力の算定でも過去実績による正常化（費用計上）は行っていない。'])

# P20 販管費の修正
t = shp(S(20), 15).table
lab = {t.cell(r, 0).text.strip(): r for r in range(len(t.rows))}
cell(t, lab['賞与'], 4, '業績連動のため修正なし')
cell(t, lab['法定福利費'], 4, '①に社保15%を連動')
cell(t, lab['寄付金'], 4, '②全額加算')
set_text(shp(S(20), 16), ['役員報酬・法定福利費：退任役員（代表者・配偶者・監査役・退任済みの常務）の報酬を控除（社保15%連動）し、代表者の顧問料（年2,160千円）を計上（別頁参照）',
                          '賞与：規定はなく業績連動で支給義務がないため修正しない／寄付金：譲渡後に発生しないため加算',
                          '支払保険料・接待交際費：オーナー個人に帰属する費用の有無が未確認のため修正していない（要確認）'])

# P22 修正PLの集計：②賞与の行を削除し、③④を②③に繰り上げ
gf = shp(S(22), 16); t = gf.table
rows = [t.cell(r, 0).text.strip() for r in range(len(t.rows))]
r_bonus = next(i for i, x in enumerate(rows) if x.startswith('②賞与'))
h = t.rows[r_bonus].height
t._tbl.remove(t._tbl.tr_lst[r_bonus]); gf.height -= h
rows = [t.cell(r, 0).text.strip() for r in range(len(t.rows))]
R = {x: i for i, x in enumerate(rows)}
cell(t, R['③寄付金の加算'], 0, '②寄付金の加算')
cell(t, R['④受取奨励金（純額）の加算'], 0, '③受取奨励金（純額）の加算')
for c, col in zip((1, 2, 3), 'FGH'):
    cell(t, R['修正額 計'], c, K(v(Y, f'{col}5')))
    cell(t, R['修正後営業利益'], c, K(v(Y, f'{col}6')))
    cell(t, R['税考慮額'], c, K(v(Y, f'{col}7')))
    cell(t, R['修正後税考慮後営業利益'], c, K(v(Y, f'{col}8')))
    cell(t, R['修正後EBITDA'], c, K(v(E, f'{col}8')))
cell(t, R['基準営業利益（税考慮後・加重平均）'], 3, K(v(Y, 'H10')))
cell(t, R['基準EBITDA（加重平均）'], 3, K(v(E, 'H10')))
replace_in(S(22), '②賞与の正常化', '（賞与は業績連動のため正常化しない）')

# P24 年買法
t = shp(S(24), 71).table
R = {t.cell(r, 0).text.strip(): r for r in range(len(t.rows))}
for c, col in zip((1, 2, 3), 'FGH'):
    cell(t, R['修正額'], c, K(v(Y, f'{col}5')))
    cell(t, R['修正後営業利益'], c, K(v(Y, f'{col}6')))
    cell(t, R['税考慮額'], c, K(v(Y, f'{col}7')))
    cell(t, R['修正後税考慮後営業利益'], c, K(v(Y, f'{col}8')))
cell(t, R['基準営業利益'], 3, K(v(Y, 'H10')))
t = shp(S(24), 72).table
R = {t.cell(r, 0).text.strip(): r for r in range(len(t.rows))}
for c, col in zip((1, 2, 3), 'KLM'):
    cell(t, R['時価純資産'], c, K(v(Y, 'K4')))
    cell(t, R['営業権'], c, K(v(Y, f'{col}5')))
    cell(t, R['想定株式価値'], c, K(v(Y, f'{col}6')))

# P25 EV/EBITDAマルチプル
t = shp(S(25), 79).table
R = {t.cell(r, 0).text.strip(): r for r in range(len(t.rows))}
for c, col in zip((1, 2, 3), 'FGH'):
    cell(t, R['修正額'], c, K(v(E, f'{col}5')))
    cell(t, R['修正後営業利益'], c, K(v(E, f'{col}6')))
    cell(t, R['修正後EBITDA'], c, K(v(E, f'{col}8')))
cell(t, R['基準EBITDA'], 3, K(v(E, 'H10')))
t = shp(S(25), 80).table
R = {t.cell(r, 0).text.strip(): r for r in range(len(t.rows))}
for c, col in zip((1, 2, 3), 'KLM'):
    cell(t, R['EV（EBITDA×倍率）'], c, K(v(E, f'{col}4')))
    cell(t, R['非流動性ディスカウント'], c, K(v(E, f'{col}5')))
    cell(t, R['ネットキャッシュ'], c, K(v(E, 'K6')))
    cell(t, R['想定株式価値（算定値）'], c, K(v(E, f'{col}7')))

prs.save(out)
scrub_doc_props(out, title='株式価値評価書')
print('saved', out)
