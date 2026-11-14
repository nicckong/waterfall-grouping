"""Independent check of an exported file: recalculates every formula from scratch (not using the tool's own
numbers) and re-derives every figure, row position, merge and colour rule from the source file.

Usage: python3 tests/verify.py source.xlsx export.xlsx [options] ["Row name=Driver+Driver" ...]
  "Row name=Driver+Driver"   a named grouping and the drivers in it
  --pinned=Name,Name         groupings kept at the bottom of their group (default: names ending in "Others")
  --combine=Journey,Product  groups that were combined;  --combo-name=Name  its name (default "Journey & Product")
  --sort=Brand|Period|metric|asc/desc   the sort that was applied
"""
import re, sys, openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl

src_path, out_path = sys.argv[1], sys.argv[2]
groups, combine, sort, combo_name, pinned_arg = {}, [], None, None, None
for a in sys.argv[3:]:
    if a.startswith('--sort='): sort = a.split('=', 1)[1].split('|')
    elif a.startswith('--combine='): combine = a.split('=', 1)[1].split(',')
    elif a.startswith('--combo-name='): combo_name = a.split('=', 1)[1]
    elif a.startswith('--pinned='): pinned_arg = [p for p in a.split('=', 1)[1].split(',') if p]
    else:
        name, members = a.split('=', 1)
        groups[name] = members.split('+')
pinned = set(pinned_arg) if pinned_arg is not None else {g for g in groups if g.endswith('Others')}
METRIC = {'rank': 0, 'pro': 1, 'det': 2, 'net': 3}

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
        memo[k] = ev(v[1:]) if isinstance(v, str) and v.startswith('=') else (v if isinstance(v, (int, float)) else 0)
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
    root = sF.cell(5, 2).value
    starts = [c for c in range(3, sF.max_column + 1) if sF.cell(1, c).value]
    data_rows = [r for r in range(5, sF.max_row + 1) if sF.cell(r, 2).value and isinstance(sV.cell(r, 4).value, (int, float))]

    # Source hierarchy, worked out here independently: groups merged in column A, else bold labels.
    has_a = any(sF.cell(r, 1).value for r in data_rows)
    is_parent = (lambda r: sF.cell(r, 1).value == sF.cell(r, 2).value) if has_a else (lambda r: bool(sF.cell(r, 2).font.b))
    src_parents = [sF.cell(r, 2).value for r in data_rows if r != 5 and is_parent(r)]
    drivers_of = {p: [] for p in src_parents}
    current = None
    for r in data_rows:
        l = sF.cell(r, 2).value
        if r == 5: continue
        if l in drivers_of: current = l
        elif current: drivers_of[current].append(l)
    parent_of = {d: p for p, ds in drivers_of.items() for d in ds}

    parts = sorted(combine, key=lambda p: src_rows[p])
    combo = (combo_name or ' & '.join(parts)) if len(parts) >= 2 else None
    effective = lambda p: combo if combo and p in parts else p
    group_parent = {g: effective(parent_of[ms[0]]) for g, ms in groups.items()}

    # 1. the right rows are there
    gone = {l for ms in groups.values() for l in ms} | (set(parts) if combo else set())
    want = sum(1 for l in src_rows if l not in gone) + len(groups) + (1 if combo else 0)
    if len(out_rows) != want: bad(name, 'label count', len(out_rows), want)
    for l in gone:
        if l in out_rows: bad(name, 'still present', l)

    # 2. every saved formula result equals an independent recalculation
    val = recalc(oF)
    n = 0
    for row in oF.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                n += 1
                saved = oV[c.coordinate].value
                if saved is None or abs(val(c.column, c.row) - saved) > 1e-6:
                    bad(name, c.coordinate, c.value, 'saved', saved, 'recalc', val(c.column, c.row))
    print(name, 'formula cells recalculated:', n, '| brands:', [sF.cell(1, c).value for c in starts])

    # 3. untouched rows match the source exactly (values, font colours, number formats)
    for label, r in src_rows.items():
        if label in gone or label not in out_rows: continue
        for col in range(3, sF.max_column + 1):
            a, b = sV.cell(r, col).value, oV.cell(out_rows[label], col).value
            if a != b and not (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-9):
                bad(name, label, cl(col), a, b)
            fa, fb = sF.cell(r, col).font.color, oF.cell(out_rows[label], col).font.color
            if fa and fa.rgb != (fb.rgb if fb else None): bad(name, label, cl(col), 'font colour')
            if sF.cell(r, col).number_format != oF.cell(out_rows[label], col).number_format: bad(name, label, cl(col), 'number format')

    # 4. each grouping (and the combined row) from the SOURCE numbers: sums, Pro-Det, rank across brands, chg
    made = list(groups.items()) + ([(combo, parts)] if combo else [])
    for label, members in made:
        row = out_rows[label]
        for pi in (0, 4):
            nets = []
            for s in starts:
                p = sum(sV.cell(src_rows[l], s + pi + 1).value for l in members)
                d = sum(sV.cell(src_rows[l], s + pi + 2).value for l in members)
                nets.append(p - d)
                if abs(oV.cell(row, s + pi + 1).value - p) > 1e-9: bad(name, label, 'pro', s + pi)
                if abs(oV.cell(row, s + pi + 2).value - d) > 1e-9: bad(name, label, 'det', s + pi)
                if abs(oV.cell(row, s + pi + 3).value - (p - d)) > 1e-9: bad(name, label, 'net', s + pi)
            for s, me in zip(starts, nets):
                rank = 1 + sum(round(x, 9) > round(me, 9) for x in nets)
                if oV.cell(row, s + pi).value != rank: bad(name, label, 'rank', cl(s + pi), oV.cell(row, s + pi).value, rank)
        for s in starts:
            if oV.cell(row, s + 8).value != oV.cell(row, s).value - oV.cell(row, s + 4).value: bad(name, label, 'rank chg', s)
            for m in (1, 2, 3):
                if abs(oV.cell(row, s + 8 + m).value - (oV.cell(row, s + 4 + m).value - oV.cell(row, s + m).value)) > 1e-9: bad(name, label, 'chg', s, m)

    # 5. frozen panes and the flat red/green chg colours over the whole table
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

    # 6. rank on every row is competition ranking: ties share a rank, the next one skips
    tied = 0
    for r in range(5, last + 1):
        if oF.cell(r, 2).value is None: continue
        for pi in (0, 4):
            nets = [round(oV.cell(r, s + pi + 3).value, 9) for s in starts]
            tied += len(set(nets)) < len(nets)
            for s, me in zip(starts, nets):
                if oV.cell(r, s + pi).value != 1 + sum(x > me for x in nets): bad(name, 'competition rank', r, cl(s + pi))

    # 7. each parent's block: the right rows, in the right order
    out_parents = [p for p in src_parents if p in out_rows] + ([combo] if combo else [])
    out_parents.sort(key=lambda p: out_rows[p])
    sort_col = None
    if sort:
        brand, period, metric, dirn = sort
        s0 = [c for c in starts if sF.cell(1, c).value == brand]
        if s0: sort_col = [c for c in range(s0[0], s0[0] + 12) if oF.cell(2, c).value == period][0] + METRIC[metric]
    block_rows = {}
    for i, p in enumerate(out_parents):
        end = out_rows[out_parents[i + 1]] if i + 1 < len(out_parents) else last + 1
        rows = [r for r in range(out_rows[p] + 1, end) if oF.cell(r, 2).value is not None]
        block_rows[p] = rows
        got_labels = [oF.cell(r, 2).value for r in rows]
        src_drivers = [d for q in (parts if p == combo else [p]) for d in drivers_of[q]]
        mine = [g for g in groups if group_parent[g] == p]
        expected, seen = [], set()
        for d in src_drivers:
            g = next((g for g in mine if d in groups[g]), None)
            if g is None: expected.append(d)
            elif g not in pinned and g not in seen: seen.add(g); expected.append(g)
        expected += [g for g in mine if g in pinned]
        if sorted(got_labels) != sorted(expected): bad(name, 'block rows', p, got_labels, expected); continue
        if sort_col is None:
            if got_labels != expected: bad(name, 'block order', p, got_labels, expected)
        else:
            n_pin = sum(1 for g in mine if g in pinned)
            free, tail = got_labels[:len(got_labels) - n_pin], got_labels[len(got_labels) - n_pin:]
            if tail != [g for g in mine if g in pinned]: bad(name, 'pinned not last', p, got_labels)
            vals = [round(oV.cell(out_rows[l], sort_col).value, 9) for l in free]
            if vals != sorted(vals, reverse=(dirn == 'desc')): bad(name, 'sort order', p, list(zip(free, vals)))
        if name == sf.sheetnames[0] and (mine or p == combo): print(f'  {p}: {got_labels}')

    # 8. Overall equals the sum of the parent rows now in the table (combined row counted once)
    for s0 in starts:
        for m in (1, 2, 5, 6):
            tot = sum(oV.cell(out_rows[p], s0 + m).value for p in out_parents)
            if abs(tot - oV.cell(out_rows[root], s0 + m).value) > 1e-6: bad(name, 'Overall != sum of parents', cl(s0 + m), tot)

    # 9. column A: one merge per parent covering it and everything now under it, carrying its name
    if has_a:
        want_merges = set()
        for p in out_parents:
            rows = block_rows[p]
            if rows: want_merges.add(f'A{out_rows[p]}:A{max(rows)}')
            if oV.cell(out_rows[p], 1).value != p: bad(name, 'A label', p, oV.cell(out_rows[p], 1).value)
        got_merges = {str(m) for m in oF.merged_cells.ranges if m.min_col == 1}
        if got_merges != want_merges: bad(name, 'merges', sorted(got_merges), sorted(want_merges))
    for col in 'AB':
        if oF.column_dimensions[col].width != sF.column_dimensions[col].width: bad(name, 'width', col)
    if name == sf.sheetnames[0]:
        print('  tied rows ranked correctly:', tied, '| parents:', out_parents, '| Overall Pro formula:', oF.cell(out_rows[root], 4).value)
print('RESULT:', 'ALL PASS' if not fails else f'{fails} FAILURES')
sys.exit(1 if fails else 0)
