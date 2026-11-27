# -*- coding: utf-8 -*-
"""対象会社HP（https://www.welcome-basket.co.jp/）の公開画像から、概要書用の写真を切り出す。

ダウンロードにはネットワーク許可（www.welcome-basket.co.jp）が要る。写真/ に結果があれば再取得しない。
グループ名（つばめグループ）のロゴ文字が入る範囲は切り落とす。
"""
import io
import os
import urllib.request

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '写真')
BASE = 'https://www.welcome-basket.co.jp/build/assets/'
os.makedirs(OUT, exist_ok=True)

# 出力名: (元ファイル, 切り出し範囲 (左, 上, 右, 下)・元画像のピクセル。None は全体, 説明)
SPEC = {
    'fleet.jpg': ('img1-4925b98a.png', None, '車庫に並ぶ自社トラック（HP 会社案内）'),
    'branch_yellow.jpg': ('img2-5b3dd393.png', None, '支店社屋（HP 会社案内）'),
    'branch_blue.jpg': ('img3-bfd7b19c.png', None, '支店社屋（HP 会社案内）'),
    'staff.jpg': ('top_3-54f03445.png', (975, 95, 1700, 836), '女性スタッフ（HP トップ「シーズムービング」）'),
    'phone.jpg': ('top_1-9aae55ab.png', (1235, 120, 1820, 1000), '確定見積り・予約サイト「スマっこし」（HP トップ）'),
    'box.png': ('cardboard-after-60d0ae5a.png', None, '自社デザインの取っ手付きダンボール（HP）'),
    'char.png': ('char-after-bc99a0b4.png', None, '自社キャラクター（HP）'),
}


def fetch(name):
    with urllib.request.urlopen(BASE + name, timeout=30) as r:
        return Image.open(io.BytesIO(r.read()))


if __name__ == '__main__':
    lines = ['# 写真（出典：対象会社HP。概要書の「東海No.1の引越会社」「選ばれる理由」頁で使用）', '',
             '| ファイル | 内容 | 元画像 |', '|---|---|---|']
    for out, (src, box, desc) in SPEC.items():
        p = os.path.join(OUT, out)
        if not os.path.exists(p):
            im = fetch(src)
            if box:
                im = im.crop(box)
            if out.endswith('.jpg'):
                im.convert('RGB').save(p, quality=90)
            else:
                im.save(p)
        lines.append(f'| {out} | {desc} | {BASE}{src} |')
    lines += ['', '写真の差し替え：同じファイル名で置き換えて `bash make_all.sh` を再実行する。',
              '売主から高解像度の写真（トラック・作業風景・拠点）を受領したら差し替える。']
    open(os.path.join(OUT, 'README.md'), 'w').write('\n'.join(lines) + '\n')
    print('ok', sorted(os.listdir(OUT)))
