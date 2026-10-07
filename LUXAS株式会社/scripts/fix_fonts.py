# -*- coding: utf-8 -*-
"""recalc.py（LibreOffice経由）が書き戻した際のフォント置換を元に戻す。

LibreOffice は未インストールの和文フォントを環境のCJKフォントへ置換して保存するため、
recalc を通すと一部セルの游ゴシックが 'WenQuanYi Zen Hei' 等になる。
そのまま納品すると客先のExcelで既定フォントにフォールバックし、游ゴシックのセルと
混在して見た目が崩れる。recalc の直後に必ず通すこと。
"""
import re
import shutil
import sys
import zipfile

SUBST = {
    'WenQuanYi Zen Hei': '游ゴシック',
    'Noto Sans CJK JP': '游ゴシック',
    'Noto Sans CJK SC': '游ゴシック',
    'Noto Sans CJK TC': '游ゴシック',
    'Noto Sans CJK KR': '游ゴシック',
    'Noto Serif CJK JP': '游ゴシック',
    'IPAGothic': '游ゴシック',
    'IPAPGothic': '游ゴシック',
    'IPAMincho': '游ゴシック',
}


def fix(path):
    tmp = path + '.tmp'
    hits = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith('.xml'):
                text = data.decode('utf-8')
                for bad, good in SUBST.items():
                    n = text.count(bad)
                    if n:
                        hits += n
                        text = text.replace(bad, good)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return hits


if __name__ == '__main__':
    targets = sys.argv[1:] or ['/home/user/github-1st/LUXAS株式会社/案件マスター_LUXAS株式会社_20260923.xlsx']
    for t in targets:
        print(f'{t}: {fix(t)} 件のフォント置換を是正')
