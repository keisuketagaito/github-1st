# -*- coding: utf-8 -*-
"""登記情報提供サービスの全部事項証明PDF（土地・建物）を解析する。

使い方:
    python registry_parse.py <PDFフォルダ> [-o registry.json]

出力（1通＝1レコード）:
    kind        土地／建物
    shozai      所在
    kaoku       家屋番号（建物）
    fudosan     13桁不動産番号
    chimoku     登記地目（土地・最新）／種類（建物）
    area        地積（土地・最新）／床面積合計（建物＝主たる建物＋附属建物）
    owner       甲区の現在の所有者（最後の所有権登記）
    cause       取得原因と日付（例「令和1年9月22日相続」）
    uketsuke    受付年月日（登記日）
    liens       乙区で抹消されていない（根）抵当権のリスト
    kyodo       共同担保目録の記載物件

注意:
- 抹消事項は原本では下線で示されるが、テキスト抽出では判別できない。
  「N番（根）抵当権抹消」の登記を拾って、順位番号Nの担保を抹消済みと判定する。
- 共同担保目録の見出しは「共 同 担 保 目 録」の1行（空白除去後に完全一致）で判定する。
  乙区本文中の「共同担保 目録第N号」を見出しと誤認すると、以降の乙区が切り捨てられる。
- 地積・地目は表題部の最後の値が現在値（それ以前は変更履歴）。
- 画像PDF（テキストなし）は解析できない。その場合は画像化して目視で読む。
- テキスト層の都合で氏名の1文字が欠けることがある（例「鳥塚康弘」→「鳥康弘」）。
  所有者名は同一住所の他の謄本と突き合わせ、疑わしければPDFを目視する。
- 建物の床面積は主たる建物と附属建物の各階を合算した値。1棟ごとの内訳は原本で確認する。
"""
import sys, os, re, json, glob, argparse
import pypdf

Z = '０１２３４５６７８９'
def z2h(s):
    for i, c in enumerate(Z): s = s.replace(c, str(i))
    return s.replace('，', ',').replace('．', '.').replace('－', '-').replace('％', '%')

BOX = r'[┏┓┗┛┠┨┯┷┼├┤─│━┃┌┐└┘┬┴╂　 ]+'
NOISE = re.compile(r'現在の情報です|権利関係を公示するもの|下線のあるものは抹消事項|相続人申告」と記載')

def lines_of(path):
    txt = '\n'.join((p.extract_text() or '') for p in pypdf.PdfReader(path).pages)
    out = []
    for ln in txt.split('\n'):
        ln = re.sub(BOX, '|', ln).strip('|')
        ln = re.sub(r'\|+', '|', ln)
        if not ln or re.fullmatch(r'[|：]*', ln) or ln.startswith('＊') or NOISE.search(ln):
            continue
        out.append(ln)
    return out

def flat(s): return s.replace('|', '')

def section(L, head, start=0, exact=False):
    for i in range(start, len(L)):
        t = flat(L[i])
        if (t == head) if exact else t.startswith(head):
            return i
    return len(L)

def parse(path):
    L = lines_of(path)
    if not L:
        return {'file': os.path.basename(path), 'error': 'テキストなし（画像PDF）'}
    i_kou = section(L, '権利部（甲区）')
    i_otsu = section(L, '権利部（乙区）', i_kou)
    i_kyo = section(L, '共同担保目録', max(i_kou, i_otsu if i_otsu < len(L) else i_kou), exact=True)
    head, kou = L[:i_kou], L[i_kou + 2:min(i_otsu, i_kyo)]
    otsu, kyo = L[i_otsu + 2:i_kyo], L[i_kyo + 1:]
    kind = '建物' if any('建物の表示' in flat(x) for x in head) else '土地'

    shozai = [z2h(flat(x)[2:]) for x in head if flat(x).startswith('所在')]
    kaoku = next((z2h(flat(x)[4:]) for x in head if flat(x).startswith('家屋番号')), '')
    m = re.search(r'不動産番号([０-９]{13})', ''.join(flat(x) for x in head))
    fudosan = z2h(m.group(1)) if m else ''

    nums = []
    for x in head:
        for mm in re.finditer(r'([０-９]+)：([０-９]*)', x):
            nums.append(float(z2h(mm.group(1)) + ('.' + z2h(mm.group(2)) if mm.group(2) else '')))
    if kind == '土地':
        chimoku = ([c for x in head for c in re.findall(r'(田|畑|宅地|雑種地|池沼|山林|原野|公衆用道路|保安林)', x)] or [''])[-1]
        area = nums[-1] if nums else None
    else:
        chimoku = '・'.join(dict.fromkeys(c for x in head for c in re.findall(r'(居宅|倉庫|作業所|寄宿舎|事務所|工場|車庫|物置|店舗|便所|ボイラー室?)', x)))
        # 主たる建物と附属建物の各階床面積をすべて合算（変更前の値が併記される場合は要目視）
        area = round(sum(nums), 2) if nums else None

    # 甲区：最後の所有権登記
    ent, cur = [], None
    for x in kou:
        s = flat(x)
        m = re.match(r'^([０-９]+)(?![０-９])(.*)$', s)
        if m and not s.startswith('第'):
            cur = {'no': z2h(m.group(1)), 'body': [m.group(2)]}; ent.append(cur)
        elif cur: cur['body'].append(s)
    owner = cause = uketsuke = ''
    for e in ent:
        if '名義人' in ''.join(e['body'])[:20]:      # 住所・氏名変更の付記登記は所有者を変えない
            continue
        b = z2h(''.join(e['body'])).split('付記')[0]     # 付記登記（住所変更等）は原因に含めない
        mo = re.search(r'(?:所有者|共有者)(.*)', b)
        if not mo: continue
        owner = re.split(r'順位|(?:昭和|平成)\d+年法務省令|昭和\d+年法律|管轄転属|\d+所有権|付記\d+号|'
                         r'(?:令和|平成|昭和)\d+年\d+月\d+日登記', mo.group(1))[0]
        mu = re.match(r'^[^0-9]*?((?:令和|平成|昭和|大正|明治)\d+年\d+月\d+日)', b)
        uketsuke = mu.group(1) if mu else ''
        mc = re.search(r'原因((?:令和|平成|昭和|大正|明治)(?:元|\d+)年\d+月\d+日[^\d第所共]*)', b)
        cause = mc.group(1) if mc else ''

    # 乙区：設定と抹消を突き合わせて現存担保を抽出
    oe, cur = [], None
    for x in otsu:
        s = flat(x)
        m0 = re.match(r'^([０-９]+)((?:根)?抵当権設定)', s)
        mx = re.match(r'^([０-９]+)([０-９]+)番(?:根)?抵当権抹消', s)
        if m0: cur = {'no': z2h(m0.group(1)), 'type': m0.group(2), 'lines': [s]}; oe.append(cur)
        elif mx: cur = {'no': z2h(mx.group(1)), 'type': '抹消', 'target': z2h(mx.group(2)), 'lines': [s]}; oe.append(cur)
        elif cur: cur['lines'].append(s)
    dead = {e['target'] for e in oe if e['type'] == '抹消'}
    liens = []
    for e in oe:
        if e['type'] == '抹消' or e['no'] in dead: continue
        b = z2h(''.join(e['lines']))
        gk = re.search(r'(極度額|債権額)金([0-9,億万]+円)', b)
        who = re.search(r'(?:根抵当権者|抵当権者)(.*?)(?:共同担保|順位|付記|$)', b)
        deb = re.search(r'債務者(.*?)(?:根?抵当権者|共同担保|$)', b)
        liens.append({'順位': e['no'], '種類': e['type'].replace('設定', ''),
                      '金額': (gk.group(1) + ' ' + gk.group(2)) if gk else '',
                      '権利者': who.group(1) if who else '', '債務者': deb.group(1) if deb else '',
                      '共同担保目録': re.findall(r'目録第([0-9／]+)号', b)})
    kyodo = [z2h(flat(x)) for x in kyo if re.match(r'^[０-９]+', flat(x))]
    return dict(file=os.path.basename(path), kind=kind, shozai=' / '.join(shozai), kaoku=kaoku,
                fudosan=fudosan, chimoku=chimoku, area=area, owner=owner, cause=cause,
                uketsuke=uketsuke, liens=liens, kyodo=kyodo)

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('folder'); ap.add_argument('-o', default='registry.json')
    a = ap.parse_args()
    files = sorted(glob.glob(os.path.join(a.folder, '**', '*.[Pp][Dd][Ff]'), recursive=True))
    recs = [parse(f) for f in files]
    json.dump(recs, open(a.o, 'w'), ensure_ascii=False, indent=1)
    for r in recs:
        if 'error' in r: print('!!', r['file'], r['error']); continue
        print('%-4s %-13s %-6s %9s  %-24s %s' % (r['kind'], r['fudosan'], r['chimoku'][:6], r['area'],
              r['owner'][-16:], '; '.join('%s%s %s' % (l['順位'], l['種類'], l['金額']) for l in r['liens']) or '担保なし'))
    print('解析 %d通' % len(recs))
