import sys
from openpyxl import load_workbook
wb=load_workbook(sys.argv[1])
for name in sys.argv[2:]:
    ws=wb[name]
    print('=====',name,ws.dimensions,'merged:',len(ws.merged_cells.ranges))
    for row in ws.iter_rows():
        cells=[f'{c.coordinate}={str(c.value)[:40]!s}' for c in row if c.value is not None]
        if cells: print(' | '.join(cells))
