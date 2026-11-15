# -*- coding: utf-8 -*-
"""ノンネームシート：当社テンプレート（assets/nonname_template.pptx、A4縦1枚）にラベルで値を流し込む。
社名が特定される語（社名・ブランド・所在地・取引先・代表者名・創業年）は書かない。売上高・従業員はレンジ。

    python3 build_nn.py <案件マスター.xlsx> <出力.pptx>
"""
import os
import sys

SK = os.environ.get('MA_SKILL', '/root/.claude/skills/synced/88fd3bc4-f86f-4f70-8539-650c944222ea_4f59f5da-cf81-4443-8419-ba745686a7ff/ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
from nn_builder import *
from openpyxl import load_workbook

MASTER, OUT = sys.argv[1], sys.argv[2]
wb = load_workbook(MASTER, data_only=True)
sales = wb['PLハイライト']['F4'].value
avg3 = wb['PLハイライト']['F26'].value
emp = 25
assert 3e8 <= sales < 5e8 and 20 <= emp < 30 and 3e7 <= avg3 < 4e7     # レンジ表記の前提（マスターが変わったら見直す）

prs = open_nn_template()
miss = fill_nonname(prs, {
    '業種': '製氷業（純氷の製造・卸売）',
    'エリア': '関東',
    '売上高': '3億円以上　／　5億円未満',
    '従業員': '20名以上　／　30名未満',
    'ビジネス面': ['純水をろ過し時間をかけて凍らせる伝統製法の純氷メーカー。業歴60年超で、透明で溶けにくい高品質な氷を自社ブランドで展開',
                   '大手食品卸・氷卸・業務用顧客と継続取引（掛売中心）。製氷・加工・冷蔵保管・配送を自社で一貫して担う',
                   '夏季に需要が集中する季節性の高い事業。冬季は冷蔵保管で設備稼働を補完（食品・物流関連とのシナジーが見込める）'],
    '財務面': ['売上高は直近4期連続で増加。臨時的費用を調整した調整後EBITDAは直近3期とも3千万円台で安定',
               '工場の土地・建物・製氷設備を自社保有。有利子負債が大きく、借入金の承継・代表者保証の解除と一体で条件を検討'],
    '組織面': ['代表者の下、工場長が現場を統括。製造・配送・販売の担当が揃っており、引継ぎ期間を設けて円滑な承継が可能（弊社の評価）'],
    '譲渡理由': '後継者不在',
    'スキーム': '株式100％譲渡',
    '希望譲渡価額': '応相談',
})
assert not miss, miss
place_table(find_table(prs, 'ビジネス面'), y=4.13, heights=[1.55, 0.95, 0.62])
move_text(prs, '補足事項', 7.40)
place_table(find_table(prs, '譲渡理由'), y=7.77)
BANNED = ['越ヶ谷', '越谷', 'キングアイス', 'king-ice', '小野田', '埼玉', '松戸', '千葉', '平方', '蒲生',
          '日本アクセス', '三菱食品', '手島', 'ニチレイ', 'クールテック', '金木屋', '1960', '昭和35', '65年', '日産40', '48時間']
ng = check_anonymity(prs, BANNED)
assert not ng, ng
cp = prs.core_properties          # テンプレート作成者の個人名を残さない
cp.author = '株式会社M&A承継機構'; cp.last_modified_by = ''; cp.title = ''
prs.save(OUT)
print('saved', OUT)
