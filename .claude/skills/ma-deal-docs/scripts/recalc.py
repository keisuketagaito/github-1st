# -*- coding: utf-8 -*-
"""Excelの数式を LibreOffice Calc で再計算し、キャッシュ値を書き戻す。

案件マスターは数式で持ち、パワポ側は data_only=True で読む運用のため、
数式を書き換えたら必ずこれを通す。
"""
import json, os, shutil, subprocess, sys, tempfile

PROFILE = '/tmp/lo-recalc-profile'
XCU = '''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load">
  <prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop>
</item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load">
  <prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop>
</item>
</oor:items>
'''

def recalc(path):
    d = os.path.join(PROFILE, 'user')
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, 'registrymodifications.xcu'), 'w') as f:
        f.write(XCU)
    out = tempfile.mkdtemp()
    r = subprocess.run(['soffice', '--headless', '--norestore',
                        '-env:UserInstallation=file://' + PROFILE,
                        '--convert-to', 'xlsx', '--outdir', out, path],
                       capture_output=True, text=True, timeout=600)
    name = os.path.splitext(os.path.basename(path))[0] + '.xlsx'
    dst = os.path.join(out, name)
    if not os.path.exists(dst):
        return {'status': 'failed', 'stdout': r.stdout[-800:], 'stderr': r.stderr[-800:]}
    shutil.copyfile(dst, path)
    shutil.rmtree(out, ignore_errors=True)

    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    errs = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#') and c.value.endswith(('!', '?', '0', 'E', 'A')):
                    errs.append(f'{ws.title}!{c.coordinate}={c.value}')
    return {'status': 'success', 'errors': len(errs), 'sample': errs[:10]}

if __name__ == '__main__':
    print(json.dumps(recalc(sys.argv[1]), ensure_ascii=False, indent=1))
