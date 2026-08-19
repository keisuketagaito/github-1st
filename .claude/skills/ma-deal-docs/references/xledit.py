# -*- coding: utf-8 -*-
"""グラフ・画像を壊さずに .xlsx のセルを書き換えるための最小ヘルパー。

openpyxl でブックを開き直すとグラフ・図形・一部の書式が失われるため、
zip 内のシートXMLを直接編集する。数式は <v>（キャッシュ値）を書かずに
fullCalcOnLoad を立てておき、Excel／LibreOffice 側で再計算させる。
"""
import re
import zipfile
from lxml import etree

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NSMAP = {'m': M}


def q(tag):
    return '{%s}%s' % (M, tag)


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return m.group(1), int(m.group(2))


class XlsxEditor:
    def __init__(self, path):
        self.zin = zipfile.ZipFile(path)
        self.raw = {n: self.zin.read(n) for n in self.zin.namelist()}
        self.infos = self.zin.infolist()
        self.zin.close()
        wb = etree.fromstring(self.raw['xl/workbook.xml'])
        rels = etree.fromstring(self.raw['xl/_rels/workbook.xml.rels'])
        rid = {r.get('Id'): r.get('Target') for r in rels}
        self.sheet_path = {}
        for s in wb.findall('m:sheets/m:sheet', NSMAP):
            t = rid[s.get('{%s}id' % R)]
            self.sheet_path[s.get('name')] = 'xl/' + t.lstrip('/').replace('xl/', '', 1)
        self.trees = {}

    def _tree(self, sheet):
        if sheet not in self.trees:
            self.trees[sheet] = etree.fromstring(self.raw[self.sheet_path[sheet]])
        return self.trees[sheet]

    def _row(self, root, rnum):
        data = root.find(q('sheetData'))
        prev = None
        for row in data.findall(q('row')):
            n = int(row.get('r'))
            if n == rnum:
                return row
            if n < rnum:
                prev = row
        row = etree.SubElement(data, q('row'))
        row.set('r', str(rnum))
        data.remove(row)
        data.insert((list(data).index(prev) + 1) if prev is not None else 0, row)
        return row

    def _cell(self, sheet, ref):
        col, rnum = split_ref(ref)
        cn = col_to_num(col)
        row = self._row(self._tree(sheet), rnum)
        prev = None
        for c in row.findall(q('c')):
            ccol, _ = split_ref(c.get('r'))
            k = col_to_num(ccol)
            if k == cn:
                return c
            if k < cn:
                prev = c
        c = etree.Element(q('c'))
        c.set('r', ref)
        # 同じ列の上下のセルから書式（s 属性）を引き継ぐ
        style = self._nearby_style(sheet, col, rnum)
        if style is not None:
            c.set('s', style)
        row.insert((list(row).index(prev) + 1) if prev is not None else 0, c)
        return c

    def _nearby_style(self, sheet, col, rnum):
        root = self._tree(sheet)
        for delta in (1, -1, 2, -2, 3, -3):
            for r in root.iter(q('row')):
                if r.get('r') != str(rnum + delta):
                    continue
                for c in r.findall(q('c')):
                    if split_ref(c.get('r'))[0] == col and c.get('s'):
                        return c.get('s')
        return None

    def set(self, sheet, ref, formula=None, number=None, text=None):
        """数式・数値・文字列のいずれかを書き込む（既存の内容は置き換える）。"""
        c = self._cell(sheet, ref)
        c.attrib.pop('t', None)
        for ch in list(c):
            c.remove(ch)
        if formula is not None:
            etree.SubElement(c, q('f')).text = formula.lstrip('=')
        elif number is not None:
            etree.SubElement(c, q('v')).text = repr(number) if isinstance(number, float) else str(number)
        elif text is not None:
            c.set('t', 'inlineStr')
            is_ = etree.SubElement(c, q('is'))
            etree.SubElement(is_, q('t')).text = text

    def clear(self, sheet, ref):
        c = self._cell(sheet, ref)
        c.attrib.pop('t', None)
        for ch in list(c):
            c.remove(ch)

    def save(self, out):
        for sheet, root in self.trees.items():
            self.raw[self.sheet_path[sheet]] = etree.tostring(
                root, xml_declaration=True, encoding='UTF-8', standalone=True)

        wbx = self.raw['xl/workbook.xml'].decode('utf-8')
        if '<calcPr' in wbx:
            wbx = re.sub(r'<calcPr([^>]*?)/>',
                         lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                             r'\s*fullCalcOnLoad="[^"]*"', '', m.group(1)), wbx, count=1)
        else:
            wbx = wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.raw['xl/workbook.xml'] = wbx.encode('utf-8')

        # 数式を追加したので計算チェーンは破棄する（Excelが再構築する）
        drop = {'xl/calcChain.xml'}
        if 'xl/calcChain.xml' in self.raw:
            ct = self.raw['[Content_Types].xml'].decode('utf-8')
            ct = re.sub(r'<Override[^>]*calcChain\.xml"[^>]*/>', '', ct)
            self.raw['[Content_Types].xml'] = ct.encode('utf-8')
            rels = self.raw['xl/_rels/workbook.xml.rels'].decode('utf-8')
            rels = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '', rels)
            self.raw['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for i in self.infos:
                if i.filename in drop:
                    continue
                z.writestr(i, self.raw[i.filename])
