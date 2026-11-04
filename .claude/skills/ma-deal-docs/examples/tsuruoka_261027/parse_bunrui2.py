import re,sys,json
out={}
for f in sys.argv[1:]:
    t=open(f).read()
    parts=t.split('＊＊ 大 分 類 計 ＊＊')
    first=None
    names=['301 根菜類']
    d={}
    for p in parts[1:]:
        nums=re.findall(r'-?[\d,]+\.?\d*',p.strip())[:12]
        v=[float(x.replace(',','')) for x in nums]
        m=re.search(r'^\s*(?:-?[\d,.]+\s+){12}.*?(?<![\d/,.])(\d{3}) ([^\s\d/][^\s]*)',p,re.S)
        d[names[-1]]=(v[11],v[7],v[9])
        names.append(m.group(1)+' '+m.group(2) if m else '?')
    out[f]=d
    print(f, 'first',names[0], 'n',len(d),'total',f"{sum(x[0] for x in d.values()):,.0f}", '県内',f"{sum(x[1] for x in d.values()):,.0f}",'県外',f"{sum(x[2] for x in d.values()):,.0f}")
keys=sorted(set(k[:3] for d in out.values() for k in d))
fs=list(out)
for k in keys:
    row=[]
    nm=''
    for f in fs:
        hit=[(n,v) for n,v in out[f].items() if n[:3]==k]
        if hit: nm=hit[0][0]; row.append(hit[0][1][0])
        else: row.append(0)
    print(f'{nm[:22]:22s}',' '.join(f'{x:>14,.0f}' for x in row))
json.dump({f:{k:v for k,v in d.items()} for f,d in out.items()},open(sys.argv[1].rsplit('/',1)[0]+'/bunrui.json','w'),ensure_ascii=False)
