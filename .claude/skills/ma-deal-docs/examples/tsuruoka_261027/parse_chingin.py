import re,sys
t=open(sys.argv[1]).read()
pages=re.split(r'令和 8年 4月～令和 8年12月 賃金台帳\s+\d+ 頁',t)
for p in pages[1:]:
    m=re.search(r'部門：(\d+) (\S+) 従業員名：(\d+) (\S+ \S+) 【(\S+)】',p)
    if not m: 
        continue
    months=re.search(r'^\s*((?:\S+月 )+)給与合計',p.split('部門')[1].split('】',1)[1].strip()+' ')
    hdr=re.findall(r'(\d+)月',p.split('給与合計')[0])
    tot=re.search(r'実総支給額 ([\d, ]+)',p)
    vals=tot.group(1).split() if tot else []
    bonus=re.search(r'賞与合計.*?賞与 支給 (.*?)総支給額 ([\d,]+) ([\d,]+)',p,re.S)
    bon=bonus.group(2) if bonus else ''
    hours=re.search(r'時間外\S* ([\d. ]+)',p)
    print(m.group(1),m.group(2),m.group(3),m.group(4),m.group(5),'|', dict(zip(hdr,vals)),'| 賞与',bon)
