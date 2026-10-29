# -*- coding: utf-8 -*-
"""ノンネームシート（匿名ティーザー・1枚）。社名・所在地詳細・取引先名・ブランド名は記載しない。"""
import sys
sys.path.insert(0, '.')
from deck import *

TEMPLATE, MASTER, OUT = sys.argv[1:4]
M = Master(MASTER)
D = Deck(TEMPLATE)
PER = [M.v('BS(借方)', f'{c}4') for c in 'CDEF']
PS = [p.replace('年', '/').replace('月期', '期')[2:] for p in PER]
C4 = 'CDEF'
ph = M.wb['PLハイライト']
RH = M.ref['PLハイライト']


def mil(v):
    """千円単位を百万円（四捨五入）"""
    return int(Decimal(abs(v) / 1_000_000).quantize(Decimal('1'), rounding=ROUND_HALF_UP)) * (1 if v >= 0 else -1)


def approx(v, step):
    """概数表示（例：約4億円）。ノンネームでは実額を出さない"""
    return round(v / step) * step


sales = ph[f'F{RH["売上高"]}'].value
avg3 = ph[f'F{RH["avg3"]}'].value

s = D.new('ノンネームシート（案件概要）', '案件No.', '関東地方で業歴60年超の純氷メーカー。自社工場で純水から製氷し、大手食品卸を含む安定した販路を有する。後継者不在のため株式譲渡による事業承継を希望。')
section_bar(s, 0.37, 1.62, 5.35, '対象会社の概要')
rows = [['項目', '内容'],
        ['業種', '製氷業（純氷・加工氷の製造販売）、冷蔵保管'],
        ['所在地', '関東地方'],
        ['業歴', '1960年代創業（業歴60年超）'],
        ['売上高', f'約{approx(sales / 100_000_000, 0.5):.1f}億円（直近期）'],
        ['調整後EBITDA', f'約{approx(avg3 / 1_000_000, 5):.0f}百万円（直近3期平均）'],
        ['従業員数', '約25名'],
        ['主要資産', '自社工場（土地・建物）、製氷・冷蔵設備'],
        ['主要販売先', '大手食品卸、氷卸、業務用顧客'],
        ['譲渡スキーム', '株式譲渡（100%）'],
        ['譲渡理由', '後継者不在'],
        ['希望価格', '応相談']]
table(s, 0.37, 2.0, [1.35, 4.0], rows, row_h=0.36, size=10)
section_bar(s, 6.0, 1.62, 5.3, '業績推移（百万円）')
bar_line_chart(s, 6.0, 1.98, 5.3, 2.55, PS, [('売上高', [mil(ph[f'{c}{RH["売上高"]}'].value) for c in C4])],
               [('調整後EBITDA（右軸）', [mil(ph[f'{c}{RH["調整後EBITDA"]}'].value) for c in C4], '0889C9', '#,##0;△#,##0')],
               number_format='#,##0', bar_colors=(NAVY,), font_size=8)
section_bar(s, 6.0, 4.62, 5.3, '特徴・強み')
pts = [('伝統製法の純氷', '自社の純水設備でろ過した水を約48時間かけて凍結。透明度が高く溶けにくい高品質な氷を製造'),
       ('安定した販路', '大手食品卸・氷卸と継続取引。掛売中心で回収も安定'),
       ('一貫体制', '製氷・加工・冷蔵保管・配送を自社で一貫。製氷能力は日産約40トン'),
       ('シナジー余地（仮説）', '食品・物流・飲料関連企業の販路や配送網との連携で、閑散期の稼働向上と商品拡大が見込める')]
for i, (h, b) in enumerate(pts):
    yy = 5.0 + i * 0.5
    rect(s, 6.0, yy, 0.32, 0.32, fill=BLUE, shape=MSO_SHAPE.OVAL, text='✓', size=11, color=WHITE, bold=True)
    text_box(s, 6.4, yy - 0.04, 4.9, 0.5, [[(h, {'bold': True, 'size': 9.5, 'color': NAVY})], [(b, {'size': 8.5})]])
footnote(s, ['※本資料は対象会社を特定できないよう情報を加工しています。', '※詳細資料（企業概要書）は秘密保持契約の締結後に開示いたします。',
             '※数値は概数。調整後EBITDA＝営業利益＋減価償却費＋臨時的費用の調整（詳細は企業概要書参照）。'], y=6.55, w=5.4, size=7.5)
D.finalize(OUT)
print('saved', OUT)
