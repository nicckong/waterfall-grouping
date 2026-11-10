"""Independent check of an exported file: recalculates every formula from scratch (not using the
tool's own numbers) and compares with the saved values and with figures derived from the source."""
import re, sys, openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl

src_path, out_path = sys.argv[1], sys.argv[2]
groups, combine, sort = {}, [], None
for g in sys.argv[3:]:
    if g.startswith('--sort='):
        sort = g.split('=', 1)[1].split('|'); continue
    if g.startswith('--combine='):
        combine = g.split('=', 1)[1].split(','); continue
    parent, labels = g.split('=')
    groups[parent] = labels.split('+')

sf, sv = openpyxl.load_workbook(src_path), openpyxl.load_workbook(src_path, data_only=True)
of, ov = openpyxl.load_workbook(out_path), openpyxl.load_workbook(out_path, data_only=True)
assert sf.sheetnames == of.sheetnames
fails = 0
def bad(*a):
    global fails
    fails += 1
    if fails < 25: print('FAIL', *a)

def recalc(ws):
    memo = {}
    def val(col, row):
        k = (col, row)
        if k in memo: return memo[k]
        v = ws.cell(row, col).value
        if isinstance(v, str) and v.startswith('='):
            memo[k] = ev(v[1:])
        else:
            memo[k] = v if isinstance(v, (int, float)) else 0
        return memo[k]
    def ev(f):
        f = re.sub(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)',
                   lambda m: '(' + '+'.join(f'{m[1]}{r}' for r in range(int(m[2]), int(m[4]) + 1)) + ')', f)
        f = re.sub(r'\b([A-Z]{1,3})(\d+)\b', lambda m: f'val({ci(m[1])},{m[2]})', f)
        f = re.sub(r'\(val\((\d+),(\d+)\)>val\((\d+),(\d+)\)\)',
                   lambda m: f'(round(val({m[1]},{m[2]}),9)>round(val({m[3]},{m[4]}),9))', f)
        return eval(f, {'val': val, 'round': round, 'ROUND': round})
    return val

for name in sf.sheetnames:
    sF, sV, oF, oV = sf[name], sv[name], of[name], ov[name]
    src_rows = {sF.cell(r, 2).value: r for r in range(1, sF.max_row + 1) if sF.cell(r, 2).value}
    out_rows = {oF.cell(r, 2).value: r for r in range(1, oF.max_row + 1) if oF.cell(r, 2).value}

    # combined groups: in source (table) order, joined with ' & '
    parts = sorted(combine, key=lambda p: src_rows[p])
    combo = ' & '.join(parts) if len(parts) >= 2 else None

    # 1. expected row count and labels
    gone = {l for ls in groups.values() for l in ls} | (set(parts) if combo else set())
    want = sum(1 for l in src_rows if l not in gone) + len(groups) + (1 if combo else 0)
    if len(out_rows) != want: bad(name, 'label count', len(out_rows), want)
    for l in gone:
        if l in out_rows: bad(name, 'still present', l)

    # 2. saved value == independent recalculation of formula, every cell
    val = recalc(oF)
    n = 0
    for row in oF.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                n += 1
                saved = oV[c.coordinate].value
                if saved is None or abs(val(c.column, c.row) - saved) > 1e-6:
                    bad(name, c.coordinate, c.value, 'saved', saved, 'recalc', val(c.column, c.row))
    print(name, 'formula cells recalculated:', n)

    # 3. untouched rows match the source exactly (values), parents and totals included
    for label, r in src_rows.items():
        if label in gone or label not in out_rows: continue
        for col in range(3, sF.max_column + 1):
            a, b = sV.cell(r, col).value, oV.cell(out_rows[label], col).value
            if a != b and not (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-9):
                bad(name, label, cl(col), a, b)
            if sF.cell(r, col).font.color and sF.cell(r, col).font.color.rgb != (oF.cell(out_rows[label], col).font.color.rgb if oF.cell(out_rows[label], col).font.color else None):
                bad(name, label, cl(col), 'font colour')
            if sF.cell(r, col).number_format != oF.cell(out_rows[label], col).number_format:
                bad(name, label, cl(col), 'number format')

    # 4. Others row from the SOURCE numbers: Pro/Det sum, net diff, rank across the 6 brands, chg
    starts = [c for c in range(3, sF.max_column + 1) if sF.cell(1, c).value]
    print(name, 'brands:', [sF.cell(1, c).value for c in starts])
    made = [(f'{p} - Others', ls) for p, ls in groups.items()] + ([(combo, parts)] if combo else [])
    for parent, labels in made:
        row = out_rows[parent]
        for pi in (0, 4):
            nets = []
            for s in starts:
                p = sum(sV.cell(src_rows[l], s + pi + 1).value for l in labels)
                d = sum(sV.cell(src_rows[l], s + pi + 2).value for l in labels)
                nets.append(p - d)
                if abs(oV.cell(row, s + pi + 1).value - p) > 1e-9: bad(name, parent, 'pro', s + pi)
                if abs(oV.cell(row, s + pi + 2).value - d) > 1e-9: bad(name, parent, 'det', s + pi)
                if abs(oV.cell(row, s + pi + 3).value - (p - d)) > 1e-9: bad(name, parent, 'net', s + pi)
            for s, me in zip(starts, nets):
                rank = 1 + sum(round(x, 9) > round(me, 9) for x in nets)
                if oV.cell(row, s + pi).value != rank: bad(name, parent, 'rank', cl(s + pi), oV.cell(row, s + pi).value, rank)
        for s in starts:
            if oV.cell(row, s + 8).value != oV.cell(row, s).value - oV.cell(row, s + 4).value: bad(name, parent, 'rank chg', s)
            for m in (1, 2, 3):
                if abs(oV.cell(row, s + 8 + m).value - (oV.cell(row, s + 4 + m).value - oV.cell(row, s + m).value)) > 1e-9: bad(name, parent, 'chg', s, m)

    # 5. workbook features survive
    delta = len(out_rows) - len(src_rows)
    last = max(r for r in out_rows.values() if isinstance(oF.cell(r, 4).value, (int, float, str)))
    if oF.freeze_panes != sF.freeze_panes: bad(name, 'freeze', oF.freeze_panes)
    chg_cols = [s + 8 + m for s in starts for m in range(4)]
    want_refs = {f'{cl(c)}5:{cl(c)}{last}' for c in chg_cols}
    got = {str(k.sqref): k.rules for k in oF.conditional_formatting}
    if set(got) != want_refs: bad(name, 'CF ranges', sorted(set(got) ^ want_refs)[:4])
    for ref, rules in got.items():
        c = ref.split(':')[0]
        kinds = sorted((r.type, r.formula[0], r.dxf.fill.bgColor.rgb) for r in rules)
        is_det = str(oF[c.rstrip('0123456789') + '3'].value).lower().startswith('det')
        below, above = ('FFA9DDB0', 'FFF8B4B4') if is_det else ('FFF8B4B4', 'FFA9DDB0')
        if kinds != [('expression', f'ROUND({c},6)<0', below), ('expression', f'ROUND({c},6)>0', above)]: bad(name, 'CF rule', ref, kinds)

    # 6. Rank on every data row is standard competition ranking: ties share, next rank skips
    tied = 0
    for r in range(5, last + 1):
        if oF.cell(r, 2).value is None: continue  # blank spacer row
        for pi in (0, 4):
            nets = [round(oV.cell(r, s + pi + 3).value, 9) for s in starts]
            tied += len(set(nets)) < len(nets)
            for s, me in zip(starts, nets):
                want_rank = 1 + sum(x > me for x in nets)
                if oV.cell(r, s + pi).value != want_rank: bad(name, 'competition rank', r, cl(s + pi), oV.cell(r, s + pi).value, want_rank)
    print(' ', name, 'rows with a tie, all ranked correctly:', tied)
    # 7. Column A: each group's merge still covers its parent and everything now under it
    want_merges = set()
    parents_in_a = [(sF.cell(m.min_row, 1).value, m) for m in sF.merged_cells.ranges if m.min_col == 1]
    if combo:
        kids = [out_rows[l] for p, m in parents_in_a if p in parts for l, r in src_rows.items() if m.min_row < r <= m.max_row and l not in gone]
        if combo in groups: kids.append(out_rows[f'{combo} - Others'])
        top = out_rows[combo]
        want_merges.add(f'A{top}:A{max(kids)}')
        if oV.cell(top, 1).value != combo: bad(name, 'A label', combo, oV.cell(top, 1).value)
        # 8. the combined group's drivers sit directly under it, nothing else in between
        if sorted(kids) != list(range(top + 1, top + 1 + len(kids))): bad(name, 'combined drivers not contiguous', sorted(kids), top)
        print(' ', name, f'{combo}: row {top}, drivers on rows {min(kids)}-{max(kids)}')
    # 9. Overall equals the sum of the parent rows now in the table (combined row counted once)
    has_a = any(sF.cell(r, 1).value for r in range(5, 31))
    is_parent = (lambda l: sF.cell(src_rows[l], 1).value == l) if has_a else (lambda l: bool(sF.cell(src_rows[l], 2).font.b))
    out_parents = [l for l in src_rows if l in out_rows and l not in gone and is_parent(l) and src_rows[l] != 5] + ([combo] if combo else [])
    for s0 in starts:
        for m in (1, 2, 5, 6):
            tot = sum(oV.cell(out_rows[p], s0 + m).value for p in out_parents)
            if abs(tot - oV.cell(out_rows[sF.cell(5, 2).value], s0 + m).value) > 1e-6: bad(name, 'Overall != sum of parents', cl(s0 + m), tot)
    # 10. Sorting: inside every group, drivers follow the sort column; Others is last; groups keep their order
    if sort:
        brand, period, metric, dirn = sort
        s0 = [c for c in starts if sF.cell(1, c).value == brand][0]
        col = [c for c in range(s0, s0 + 12) if oF.cell(2, c).value == period][0] + {'rank': 0, 'pro': 1, 'det': 2, 'net': 3}[metric]
        prow = sorted(out_rows[p] for p in out_parents)
        if [oF.cell(r, 2).value for r in prow] != sorted(out_parents, key=lambda p: out_rows[p]): bad(name, 'group order')
        src_parent_order = [l for l in sorted(src_rows, key=src_rows.get) if l in out_parents or l in parts]
        for i, r0 in enumerate(prow):
            end = prow[i + 1] if i + 1 < len(prow) else last + 1
            kids = [r for r in range(r0 + 1, end) if oF.cell(r, 2).value is not None]
            labels = [oF.cell(r, 2).value for r in kids]
            others = [l for l in labels if str(l).endswith(' - Others')]
            if others and labels[-1] != others[0]: bad(name, 'Others not last', oF.cell(r0, 2).value, labels)
            vals = [round(oV.cell(r, col).value, 9) for r in kids if not str(oF.cell(r, 2).value).endswith(' - Others')]
            want_vals = sorted(vals, reverse=(dirn == 'desc'))
            if vals != want_vals: bad(name, 'sort order', oF.cell(r0, 2).value, vals)
            if name == sf.sheetnames[0] and i < 2: print(f'  sorted {oF.cell(r0, 2).value}:', list(zip(labels, [oV.cell(r, col).value for r in kids])))
    if name == sf.sheetnames[0]: print('  parents now:', out_parents, '| Overall formula:', oF.cell(out_rows[sF.cell(5, 2).value], 4).value)
    for label, m in parents_in_a:
        if label in parts and combo: continue
        top = out_rows[label]
        below = [out_rows[l] for l, r in src_rows.items() if m.min_row < r <= m.max_row and l in out_rows and l not in gone]
        if label in groups: below.append(out_rows[f'{label} - Others'])
        want_merges.add(f'A{top}:A{max(below or [top])}')
        if oV.cell(top, 1).value != label: bad(name, 'A label', label, oV.cell(top, 1).value)
    got_merges = {str(m) for m in oF.merged_cells.ranges if m.min_col == 1}
    if got_merges != want_merges: bad(name, 'merges', sorted(got_merges), sorted(want_merges))
    if name == sf.sheetnames[0]: print('  column A merges:', sorted(got_merges, key=lambda x: int(x[1:].split(':')[0])))
    if oF.column_dimensions['A'].width != sF.column_dimensions['A'].width: bad(name, 'A width')
    if oF.column_dimensions['B'].width != sF.column_dimensions['B'].width: bad(name, 'width')
    if oF.max_row != sF.max_row + delta: print('  note: max_row', oF.max_row, 'expected', sF.max_row + delta)
    if name == sf.sheetnames[0]:
        for parent, _ in made:
            r = out_rows[parent]
            print(' ', parent, 'row', r, [oF.cell(r, c).value for c in range(2, 15)])
            print('   values', [oV.cell(r, c).value for c in range(3, 15)])
        print('  parent', [oF.cell(out_rows[p], 4).value for p in groups if p in out_rows])
print('RESULT:', 'ALL PASS' if not fails else f'{fails} FAILURES')
sys.exit(1 if fails else 0)
