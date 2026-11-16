import os
from openpyxl import load_workbook
wb=load_workbook(os.environ.get('MA_WORK', '.') + '/master/master.xlsx',data_only=True)
err=[]
for ws in wb:
    for r in ws.iter_rows():
        for c in r:
            if isinstance(c.value,str) and c.value.startswith('#') and c.value[1:4] in ('REF','DIV','VAL','NAM','N/A','NUL','NUM'):
                err.append(f'{ws.title}!{c.coordinate}={c.value}')
            if c.value is False:
                err.append(f'FALSE {ws.title}!{c.coordinate}')
print('errors/false:',[e for e in err if not e.startswith(('SGA!D','PLハイライト!C'))], len(err))
g=lambda s,a: wb[s][a].value
r0=lambda x: round(x) if isinstance(x,(int,float)) else x
print('PLH EBITDA', [r0(g('PLハイライト',f'{c}25')) for c in 'DEFG'])
print('PLH adj', [[r0(g('PLハイライト',f'{c}{r}') or 0) for c in 'DEFG'] for r in range(17,24)])
print('PLH G col', [r0(g('PLハイライト',f'G{r}')) for r in (4,6,8,10,12,15)], g('PLハイライト','G5'))
print('avg', r0(g('PLハイライト','C29')), r0(g('PLハイライト','C30')), g('PLハイライト','C31'))
print('時価純資産', r0(g('年買法','C7')), [r0(g('年買法',f'{c}6')) for c in 'KLM'], r0(g('年買法','H10')))
print('EV', [r0(g('EV EBITDAマルチプル',f'{c}7')) for c in 'KLM'], r0(g('EV EBITDAマルチプル','H10')), g('EV EBITDAマルチプル','C6'))
print('NC', g('ネットキャッシュ','C15'), g('ネットキャッシュ','F15'))
print('会社概要', g('会社概要','C7'), g('会社概要','C13'), g('会社概要','C14'))
print('SGA chk', [g('SGA',f'{c}43') for c in 'EGIK'], 'PL chk', [g('PL',f'{c}59') for c in 'EGIK'], [r0(g('PL',f'{c}19')) for c in 'EGIK'])
print('emp', [r0(g('従業員',f'E{r}')) for r in range(37,43)])
print('株式価値', [[r0(g('株式価値 ',f'{c}{r}')) for c in 'CE'] for r in (5,6)])
print('不動産', [r0(g('不動産(所有)',f'{c}15')) for c in 'EHIJ'], g('不動産(所有)','H17'), [r0(g('不動産概況',f'{c}9')) for c in 'DGHI'], g('不動産概況','G11'))
print('設備chk', [ (c.coordinate,c.value) for row in wb['設備'].iter_rows() for c in row if c.value is True or c.value is False])
print('取締役chk', g('取締役','F17'), '株主', g('株主','F17'), r0(g('株主','G18')))
print('売上構成', [r0(g('売上構成',f'C{r}')) for r in range(6,12)], g('売上構成','C13'))
