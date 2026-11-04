# -*- coding: utf-8 -*-
"""ノンネームシートのビルドスクリプト（当社テンプレート nonname_template.pptx を複製）。

数値は案件マスター（案件マスター_ウェルカム・バスケット.xlsx）から読み、レンジへの丸めはここで行う。
社名が特定される語は BANNED に列挙し、check_anonymity() で残存ゼロを確認してから保存する。
ファイル名には社名を入れない。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.path.join(HERE, '..', '.claude', 'skills', 'ma-deal-docs')
sys.path.insert(0, os.path.join(SK, 'scripts'))
from nn_builder import *  # noqa: E402,F401
from openpyxl import load_workbook  # noqa: E402

OUT = os.path.join(HERE, 'ノンネームシート_引越運送業_東海_261104.pptx')

# ---------------------------------------------------------------- マスターから読む
wb = load_workbook(os.path.join(HERE, '案件マスター_ウェルカム・バスケット.xlsx'), data_only=True)
pl, sga, cr = wb['PL'], wb['SGA'], wb['CR']
bsa, bsl = wb['BS(借方)'], wb['BS (貸方)']
LAST = 'I'           # 2025年12月期（PL/SGA/CR の帳簿価額列）
LASTB = 'F'          # 2025年12月期（BS）
sales = [pl[c + '11'].value for c in 'CEGI']
gp = pl[LAST + '17'].value
op = pl[LAST + '19'].value
dep = cr[LAST + '34'].value + sga[LAST + '29'].value      # 原価・販管費の減価償却費
ebitda = op + dep
net_assets = bsl[LASTB + '50'].value
equity_ratio = net_assets / bsa[LASTB + '61'].value
assert op > 0 and sga['B29'].value == '減価償却費' and cr['B34'].value == '減価償却費'


def oku_range(v):
    """売上高のレンジ表記。"""
    o = v / 1e8
    return '10億円以上' if o >= 10 else ('5億円以上10億円未満' if o >= 5 else '5億円未満')


def man(v):           # 千万円単位の概数（「約3千万円」）
    return '約%d千万円' % round(v / 1e7)


peak = max(sales)
gp_margin = gp / sales[-1]
print('売上', sales, 'GP率', gp_margin, 'OP', op, 'EBITDA', ebitda, '純資産', net_assets, equity_ratio)

DATA = {
    '業種': '引越運送業（個人・法人向け引越および付帯サービス）',
    'エリア': '東海地方',
    '売上高': '%s（直近期）' % oku_range(sales[-1]),
    '従業員': '50名以上（パート・アルバイトを含む）',
    'ビジネス面': [
        '個人向け引越を主力に、オフィス・医療機関等の法人移転も手掛ける。'
        '業歴20年超、地域内に複数拠点',
        '電気工事（エアコン等の取付）・ピアノ輸送・不用品回収・家電販売等の'
        '付帯サービスをワンストップで提供',
        '不動産仲介・住宅関連企業等からの紹介による受注チャネルを確立。'
        '地域での広告展開により一定の認知度を有する',
        '繁忙期は臨時人員・協力会社を活用し、固定費を抑えた機動的な運営',
    ],
    '財務面': [
        'ピーク時%d億円台から直近期%d億円台へ減収も、外注・臨時人員と販管費の圧縮で'
        '営業黒字・EBITDA%s（調整前）に回復'
        % (int(peak // 1e8), int(sales[-1] // 1e8), man(ebitda)),
        '純資産%d億円超（自己資本比率%d%%台）。金融機関借入あり。'
        '事業拠点は主に賃借で、重い不動産は保有しない'
        % (int(net_assets // 1e8), int(equity_ratio * 100) // 10 * 10),
    ],
    '組織面': [
        '常勤役員（管理・営業）と拠点ごとの現場責任者による運営体制。'
        'オーナーの日常業務への関与は限定的と推察',
        '社内研修・資格制度で作業品質を標準化し、リピート・紹介受注に寄与',
        '外注中の電気工事等の内製化、法人移転の拡大に伸び代（弊社の評価）',
    ],
    '譲渡理由': '●●●（売主ヒアリング後に記載・要確認）',
    'スキーム': '株式譲渡（100%）を想定',
    '希望譲渡価額': '応相談',
}

# 社名・所在地・人名・ブランド・取引先・固有数値など、特定につながる語
BANNED = [
    'ウェルカム', 'ウエルカム', 'バスケット', 'WELCOME', 'Welcome', 'BASKET', 'welcome-basket',
    '天野', '山本', '澤井', '山口', '重富', 'つばめ', 'ツバメ', 'あんしんネット', 'タクシー',
    '名古屋', '愛知', '岐阜', '日進', '南区', '中川区', '名東', '名西', '名南', '岐南', '阿原', '赤池', '服部',
    'ミニミニ', 'ニッショー', 'トヨタ', '大和ハウス', '野村不動産', '三井不動産', 'ミサワ', '積水', 'ユニホー',
    'ティーファス', 'エディオン', 'グリーンソイル', 'スマっこし', 'オリコン', '三菱ＵＦＪ', '三菱UFJ',
    'みずほ', '商工中金', '2001年', '平成13', '派23', '派２３', '119名', '1,074', '10.7億', '17.5億',
]

prs = open_nn_template()
miss = fill_nonname(prs, DATA)
assert not miss, miss

# 行高の取り直し（文字量に合わせる）。描画で確認しながら決めた値
place_table(find_table(prs, '業種'), y=1.36, heights=[0.46, 0.40, 0.40, 0.40])
move_text(prs, '特徴・強み', 3.20)
place_table(find_table(prs, 'ビジネス面'), y=3.57, heights=[1.78, 1.00, 1.15])
move_text(prs, '補足事項', 7.66)
place_table(find_table(prs, '譲渡理由'), y=8.03, heights=[0.36, 0.36, 0.36])

ng = check_anonymity(prs, BANNED)
assert not ng, ng
prs.save(OUT)
print('saved', OUT)
