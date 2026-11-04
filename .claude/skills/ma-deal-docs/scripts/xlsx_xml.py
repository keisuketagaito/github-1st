# -*- coding: utf-8 -*-
"""案件マスター（グラフ入りExcel）を壊さずに編集するための最小ヘルパー。

openpyxl で保存するとグラフのスタイル・配色パーツ（chartStyle／colors）が失われ、
グラフの見た目が崩れる。本モジュールは zip 内の XML を直接書き換え、触らないパーツは
バイト単位でそのまま残す。

    from xlsx_xml import Book
    b = Book('案件マスター.xlsx')
    b.set_cell('土地', 'F13', 1029381.43)              # 数値（書式は既存セルの s= を維持）
    b.set_cell('土地', 'E13', formula='F13-D13')       # 数式
    b.set_cell('土地', 'G13', '固定資産税評価額÷0.7')  # 文字（inlineStr）
    b.append_rows('不動産(所有) ②', rows)              # rows = [[(値, style_id), ...], ...]
    b.add_sheet('個人不動産_買取試算', rows, widths={'A':7,'B':30})
    b.save('案件マスター_更新.xlsx')

落とし穴（すべて対処済み）:
- 新シートの rId は workbook.xml.rels 全体から採番する。workbook.xml の sheet の r:id だけを
  見ると theme / styles / sharedStrings の rId と衝突し、ファイルが開けなくなる。
- 数式を変えたら calcChain.xml を削除し、[Content_Types].xml と rels の参照も消す。
  workbook.xml の calcPr に fullCalcOnLoad="1" を付け、開いたときに再計算させる。
- 一時ファイルに書いてから置き換える（書き込み途中で失敗すると原本が壊れる）。
- 文字列は sharedStrings を触らず inlineStr で書く（他シートの同じ文字列に波及しない）。
- 表示が「#,##0,」（千円表示）の書式IDに円単位で入れる。マスターは円入力・千円表示が原則。
"""
import zipfile, re, os
from xml.sax.saxutils import escape


def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def cell_xml(ref, value=None, style=None, formula=None):
    s = ' s="%s"' % style if style is not None else ''
    if formula is not None:
        return '<c r="%s"%s><f>%s</f></c>' % (ref, s, escape(formula))
    if value is None or value == '':
        return '<c r="%s"%s/>' % (ref, s)
    if isinstance(value, bool):
        value = str(value)
    if isinstance(value, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(value) if isinstance(value, float) else value)
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(str(value)))


class Book:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.parts = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
        z.close()
        self.new = {}
        self.formula_changed = False

    def _get(self, name): return self.parts[name][1].decode('utf8')
    def _put(self, name, text): self.parts[name] = (self.parts[name][0], text.encode('utf8'))

    def sheet_path(self, sheet_name):
        wb, rels = self._get('xl/workbook.xml'), self._get('xl/_rels/workbook.xml.rels')
        m = re.search(r'<sheet name="%s" sheetId="\d+"[^>]*r:id="([^"]+)"' % re.escape(escape(sheet_name, {'"': '&quot;'})), wb)
        if not m: raise KeyError(sheet_name)
        t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels).group(1)
        return 'xl/' + t.lstrip('/').replace('xl/', '')

    def set_cell(self, sheet, ref, value=None, formula=None, style=None):
        p = self.sheet_path(sheet); x = self._get(p)
        m = re.search(r'<c r="%s"( s="\d+")?(?: t="\w+")?(?:/>|>.*?</c>)' % ref, x, re.S)
        if not m: raise KeyError('%s!%s が存在しない（行ごと append_rows で追加する）' % (sheet, ref))
        st = style if style is not None else (m.group(1) or '').replace(' s="', '').replace('"', '') or None
        x = x[:m.start()] + cell_xml(ref, value, st, formula) + x[m.end():]
        self._put(p, x)
        if formula is not None: self.formula_changed = True

    def append_rows(self, sheet, rows, start_row=None, height=None):
        p = self.sheet_path(sheet); x = self._get(p)
        last = max([int(r) for r in re.findall(r'<row r="(\d+)"', x)] or [0])
        r = start_row or last + 1
        out = []
        for row in rows:
            cells = ''.join(cell_xml('%s%d' % (col_letter(j + 1), r), *(c if isinstance(c, tuple) else (c,)))
                            for j, c in enumerate(row) if c is not None)
            h = ' ht="%s" customHeight="1"' % height if height else ''
            out.append('<row r="%d"%s>%s</row>' % (r, h, cells)); r += 1
        x = x.replace('</sheetData>', ''.join(out) + '</sheetData>')
        x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"/>', lambda m: '<dimension ref="%s:%s%d"/>' % (m.group(1), m.group(2), r - 1), x)
        self._put(p, x)

    def add_sheet(self, name, rows, widths=None, row_heights=None, merges_single_cell_rows=True, ncols=None):
        """rows = [[(値, style) or (値, style, 数式) ...], ...]。1セルだけの行は全幅で結合。"""
        ncols = ncols or max((len(r) for r in rows), default=1)
        xs = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
              '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
              '<dimension ref="A1:%s%d"/>' % (col_letter(ncols), max(len(rows), 1)),
              '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>']
        if widths:
            xs.append('<cols>%s</cols>' % ''.join(
                '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (ord(k) - 64, ord(k) - 64, v) for k, v in sorted(widths.items())))
        xs.append('<sheetData>')
        merges = []
        for i, row in enumerate(rows, 1):
            if not row: xs.append('<row r="%d"/>' % i); continue
            h = ' ht="%s" customHeight="1"' % row_heights[i] if row_heights and i in row_heights else ''
            cells = ''
            for j, c in enumerate(row):
                v, s, f = (tuple(c) + (None, None))[:3] if isinstance(c, tuple) else (c, None, None)
                cells += cell_xml('%s%d' % (col_letter(j + 1), i), v, s, f)
                if f: self.formula_changed = True
            xs.append('<row r="%d"%s>%s</row>' % (i, h, cells))
            if merges_single_cell_rows and len(row) == 1 and ncols > 1:
                merges.append('A%d:%s%d' % (i, col_letter(ncols), i))
        xs.append('</sheetData>')
        if merges:
            xs.append('<mergeCells count="%d">%s</mergeCells>' % (len(merges), ''.join('<mergeCell ref="%s"/>' % m for m in merges)))
        xs.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>')
        rels = self._get('xl/_rels/workbook.xml.rels')
        rid = max(int(i) for i in re.findall(r'Id="rId(\d+)"', rels)) + 1          # rels 全体から採番
        used = [int(i) for i in re.findall(r'worksheets/sheet(\d+)\.xml', rels)] + \
               [int(re.search(r'sheet(\d+)', k).group(1)) for k in self.new]
        num = max(used) + 1
        path = 'xl/worksheets/sheet%d.xml' % num
        self._put('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>',
            '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet%d.xml"/></Relationships>' % (rid, num)))
        wb = self._get('xl/workbook.xml')
        sid = max(int(i) for i in re.findall(r'sheetId="(\d+)"', wb)) + 1
        self._put('xl/workbook.xml', wb.replace('</sheets>', '<sheet name="%s" sheetId="%d" r:id="rId%d"/></sheets>' % (escape(name), sid, rid)))
        ct = self._get('[Content_Types].xml')
        self._put('[Content_Types].xml', ct.replace('</Types>',
            '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % path))
        self.new[path] = ''.join(xs).encode('utf8')
        return path

    def save(self, dst):
        if self.formula_changed or self.new:
            for name, pat in (('[Content_Types].xml', r'<Override PartName="/xl/calcChain\.xml"[^>]*/>'),
                              ('xl/_rels/workbook.xml.rels', r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>')):
                self._put(name, re.sub(pat, '', self._get(name)))
            wb = self._get('xl/workbook.xml')
            if '<calcPr' in wb:
                wb = re.sub(r'<calcPr([^>]*?)\s*/>', lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % m.group(1).replace(' fullCalcOnLoad="1"', ''), wb)
            else:
                wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
            self._put('xl/workbook.xml', wb)
        tmp = dst + '.tmp'
        z = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
        for name, (info, blob) in self.parts.items():
            if name == 'xl/calcChain.xml' and (self.formula_changed or self.new):
                continue
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type, zi.external_attr = info.compress_type, info.external_attr
            zi.internal_attr, zi.create_system = info.internal_attr, info.create_system
            z.writestr(zi, blob)
        for name, blob in self.new.items():
            z.writestr(name, blob)
        z.close()
        os.replace(tmp, dst)
