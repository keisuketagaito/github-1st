#!/usr/bin/env python3
"""和文フォントの取りこぼしを塞ぐ。

pptxgenjs は <a:ea> に <a:latin> と同じ書体を書く。欧文ラベル用に Arial を
指定した run に和文が混ざると（例：「件+」「無料」）、和文が Arial 指定となり
PowerPoint 側の既定フォントへ落ちる。ea だけ 游ゴシック に寄せれば、
欧文は Arial のまま和文だけが正しい書体になる。
"""
import re
import sys
import zipfile

SRC, DST = sys.argv[1], sys.argv[2]
JP = "游ゴシック"

ea_re = re.compile(r'<a:ea typeface="(?!游)([^"]*)"([^/>]*)/>')

zin = zipfile.ZipFile(SRC)
names = zin.namelist()
data = {n: zin.read(n) for n in names}
zin.close()

touched = subs = 0
for n in names:
    if not (n.startswith("ppt/slides/slide")
            or n.startswith("ppt/slideMasters/")
            or n.startswith("ppt/slideLayouts/")
            or n.startswith("ppt/theme/")):
        continue
    s = data[n].decode("utf-8")
    new, k = ea_re.subn(lambda m: f'<a:ea typeface="{JP}"{m.group(2)}/>', s)
    if k:
        data[n] = new.encode("utf-8")
        touched += 1
        subs += k

with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.writestr(n, data[n])

print(f"ea -> {JP}: {subs} runs in {touched} parts -> {DST}")
