"""公式マスターの VR→ シートを評価書に貼れる形に整える（再計算後に実行）。

鳥塚案件の評価書は、マスターの修正BS・修正PL・修正SGA等を「使っていない行・列を非表示」にしたうえで
そのまま貼っている（paste_range は非表示の行・列を貼らない）。公式マスターは科目行が多めに用意され、
空行（科目名 0）や1期目（●●年●月期）列が残るので、値を見て非表示を決める必要がある。

使い方（recalc.sh の LibreOffice 再計算と inject_cache の間で実行する）:
    python vr_prep.py <数式版.xlsx> <再計算版.xlsx> <出力（数式版・非表示設定済み）.xlsx>
    python inject_cache.py <出力.xlsx> <再計算版.xlsx> <マスター.xlsx>

非表示は値を変えないので、再計算版のキャッシュはそのまま使える。
"""
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# シート名, 判定する行範囲, 値の列, 常に残す行, 非表示にする列, 貼付範囲, 評価書の頁
SPEC = [
    ('修正BS(借方)', 5, 61, 'CDE', (3, 4), '', 'B3:E61', '修正BS［資産］'),
    ('修正BS(貸方) ', 5, 53, 'CDE', (3, 4, 34), '', 'B3:E53', '修正BS［負債・純資産］'),
    ('修正PL', 5, 56, 'FGHIJKLMN', (3, 4), 'CDEOPQ', 'B3:N56', '修正PLの算定［正常収益力の集計］'),
    ('修正SGA', 5, 39, 'FGHIJKLMN', (3, 4), 'CDEOPQ', 'B3:N39', '販管費及び一般管理費の修正'),
]


def _ungroup_cols(ws, upto=40):
    """列幅定義が C:Q のように範囲でまとまっていると、1列の非表示で範囲全体が消える。列ごとに分解する。"""
    from openpyxl.worksheet.dimensions import ColumnDimension
    for key, dim in list(ws.column_dimensions.items()):
        if dim.min and dim.max and dim.max > dim.min and dim.min <= upto:
            w, hid = dim.width, dim.hidden
            del ws.column_dimensions[key]
            for i in range(dim.min, min(dim.max, upto) + 1):
                L = get_column_letter(i)
                ws.column_dimensions[L] = ColumnDimension(ws, index=L, width=w, hidden=hid)


def _empty(v):
    return v is None or v == 0 or (isinstance(v, str) and v.strip() in ('', '0'))


def prep(fx_path, val_path, out_path, spec=SPEC):
    wf = load_workbook(fx_path)
    wv = load_workbook(val_path, data_only=True)
    for name, r1, r2, cols, keep, hide_cols, rng, page in spec:
        if name not in wf.sheetnames:
            continue
        f, v = wf[name], wv[name]
        n = 0
        for r in range(r1, r2 + 1):
            if r in keep:
                continue
            lab = v.cell(r, 2).value
            vals = [v[f'{c}{r}'].value for c in cols]
            hidden = _empty(lab) or all(_empty(x) for x in vals)
            f.row_dimensions[r].hidden = hidden
            n += hidden
        _ungroup_cols(f)
        if hide_cols:   # 範囲定義に引きずられて非表示になった列を戻し、指定列だけ隠す
            from openpyxl.utils import range_boundaries
            c1, _, c2, _ = range_boundaries(rng)
            for i in range(c1, c2 + 1):
                f.column_dimensions[get_column_letter(i)].hidden = False
        for c in hide_cols:
            f.column_dimensions[c].hidden = True
        f['A1'] = f'■ 貼付範囲 {rng} → 評価書「{page}」（非表示の行・列は貼られない。数字はここを直す）'
        print(f'{name.strip()}: {n}行を非表示, 列{hide_cols or "なし"}')
    wf.save(out_path)


if __name__ == '__main__':
    prep(*sys.argv[1:4])
