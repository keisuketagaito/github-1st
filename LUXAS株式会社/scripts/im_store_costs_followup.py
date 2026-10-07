# -*- coding: utf-8 -*-
"""店舗別の費用構造ページ（P.14）の追加に伴う、既存ページ・案件マスターの追記。

・案件マスター『その他』：論点7の事実を定量化し、緑店の契約形態の論点を追加
・概要書 P.13（店舗別損益）：費用科目の内訳はP.14である旨を注記
・概要書 P.39（要確認事項）：店舗別損益の解釈／契約の各行を更新

冪等。既に追記済みであれば何もしない。
"""
import openpyxl
from pptx import Presentation

IM     = '/home/user/github-1st/LUXAS株式会社/企業概要書_LUXAS株式会社_20260923.pptx'
MASTER = '/home/user/github-1st/LUXAS株式会社/案件マスター_LUXAS株式会社_20260923.xlsx'

M7_FACT = ('進行期7か月で売上25,109千円に対し販管費18,977千円（75.6%）。'
           '科目別に分解すると差の大半は従業員給与11,071千円（売上高比44.1%）であり、'
           '他のFC加盟店（伊勢20.5%・久居23.4%・緑18.1%）の2倍前後。'
           '松阪管理センターの人件費が同店に計上されていると推察される'
           '（棚卸表でも同センターの在庫19,669千円が明和店に含めて計上されている）')
M7_HOW  = '対象会社ヒアリング（部門設定の考え方・松阪管理センターの人員と費用の計上先）'

M19 = ('19', '契約', 'ドン・キホーテ緑店のみロイヤリティの負担がない理由',
       '進行期7か月の部門別試算表では、ロイヤリティを負担しているのはワンラブFC加盟店のうち'
       '明和・伊勢・久居（および撤退2店）で、売上高比12.0〜13.8%。'
       '一方ドン・キホーテ緑店はロイヤリティの計上がなく、代わりに地代家賃5,407千円'
       '（同14.6%）を負担しており、同店のみ費用構造が異なる。'
       'FC本部からの転借か直接賃借かによって、COC条項や譲渡後の賃料負担が変わる',
       'フランチャイズ契約書／緑店の賃貸借契約書（必要資料No.18・19）')

P13_NOTE = '※販管費の科目別の内訳はP.14「店舗別の費用構造」を参照。'

P39_ROW3 = ('①共通部門に残る本社仕入40,994千円・外注費5,610千円の店舗への配賦方法　'
            '②明和店の販管費が売上高比75.6%と高い要因。'
            '科目別では従業員給与が売上高比44.1%と他のFC加盟店（18.1〜23.4%）の2倍前後であり、'
            '松阪管理センターの人件費が含まれていると推察（P.14参照）')
P39_ROW5 = ('①フランチャイズ契約の内容とCOC条項への対応　'
            '②店舗賃貸借契約書および定期建物賃貸借（岐阜総本店2033年1月／大阪総本店2027年11月満了）'
            'の再契約見通し　③撤退2店舗の原状回復費用・違約金の有無　'
            '④ドン・キホーテ緑店のみロイヤリティの負担がなく地代家賃を負担している理由'
            '（FC本部からの転借か直接賃借かの別／P.14参照）')


def set_cell_text(cell, text):
    """セルの1段落目の最初のrunに書式を残したまま文字を差し替える。"""
    tf = cell.text_frame
    ps = list(tf.paragraphs)
    for p in ps[1:]:
        p._p.getparent().remove(p._p)
    runs = ps[0].runs
    if not runs:
        tf.text = text
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def master():
    wb = openpyxl.load_workbook(MASTER)
    ws = wb['その他']
    for r in range(5, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == '7':
            ws.cell(row=r, column=5).value = M7_FACT
            ws.cell(row=r, column=6).value = M7_HOW
        if str(ws.cell(row=r, column=2).value) == '19':
            print('その他シートは追記済み')
            wb.save(MASTER)
            return
    last = ws.max_row
    src = last                                        # 直前行の書式を引き継ぐ
    for i, v in enumerate(M19):
        c = ws.cell(row=last + 1, column=2 + i, value=v)
        s = ws.cell(row=src, column=2 + i)
        c.font, c.alignment, c.border, c.fill = (s.font.copy(), s.alignment.copy(),
                                                 s.border.copy(), s.fill.copy())
    ws.row_dimensions[last + 1].height = ws.row_dimensions[src].height
    wb.save(MASTER)
    print('その他シートに論点19を追加（行%d）' % (last + 1))


def deck():
    prs = Presentation(IM)
    s13 = prs.slides[12]
    note = [sh for sh in s13.shapes if sh.has_text_frame
            and sh.text_frame.text.startswith('（単位：千円）')][0]
    if P13_NOTE not in note.text_frame.text:
        p = note.text_frame.paragraphs[-1]
        new = p._p.makeelement(p._p.tag, {})
        p._p.addnext(new)
        from copy import deepcopy
        new.getparent().replace(new, deepcopy(p._p))
        tgt = note.text_frame.paragraphs[-1]
        tgt.runs[0].text = P13_NOTE
        for r in tgt.runs[1:]:
            r._r.getparent().remove(r._r)

    s39 = prs.slides[38]
    t = [sh for sh in s39.shapes if sh.has_table][0].table
    set_cell_text(t.cell(3, 1), P39_ROW3)
    set_cell_text(t.cell(5, 1), P39_ROW5)
    prs.save(IM)
    print('概要書 P.13 注記・P.39 要確認事項を更新')


if __name__ == '__main__':
    master()
    deck()
