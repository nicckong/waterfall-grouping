/* Grouping logic. No DOM here: it takes ExcelJS worksheets, so it can be tested on its own. */
const GroupingCore = (() => {
  const clean = (x, d = 10) => { const p = 10 ** d; const v = Math.round(x * p) / p; return v === 0 ? 0 : v; };
  const text = (cell) => String(cell.text ?? '').trim();
  const num = (cell) => {
    const v = cell.value;
    if (typeof v === 'number') return v;
    // cell.value drops a formula result of exactly 0, so read cell.result.
    if (v && typeof v === 'object' && typeof cell.result === 'number') return cell.result;
    return null;
  };
  const isFormulaCell = (cell) => { const v = cell.value; return !!(v && typeof v === 'object' && (v.formula || v.sharedFormula)); };
  const formulaOf = (cell) => (isFormulaCell(cell) ? String(cell.formula || '') : '');

  const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  const colNum = (s) => [...s.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);

  function metricOf(t) {
    const s = t.toLowerCase().replace(/\s+/g, '');
    if (/rank/.test(s)) return 'rank';
    if (/^pro.*det/.test(s)) return 'net';
    if (/^pro/.test(s)) return 'pro';
    if (/^det/.test(s)) return 'det';
    return null;
  }

  /* Finds the Brand / Period / Metric header rows, the label column and the row hierarchy. */
  function detectLayout(ws) {
    const maxCol = ws.columnCount, maxRow = ws.rowCount;
    let metricRow = 0, metrics = {};
    for (let r = 1; r <= Math.min(20, maxRow); r++) {
      const m = {};
      for (let c = 1; c <= maxCol; c++) { const k = metricOf(text(ws.getCell(r, c))); if (k) m[c] = k; }
      const found = new Set(Object.values(m));
      if (['rank', 'pro', 'det', 'net'].every((k) => found.has(k))) { metricRow = r; metrics = m; break; }
    }
    if (metricRow < 3) return null;

    const brands = [];
    let brandName = '', periodName = '';
    const cols = Object.keys(metrics).map(Number).sort((a, b) => a - b);
    for (const c of cols) {
      brandName = text(ws.getCell(metricRow - 2, c)) || brandName;
      periodName = text(ws.getCell(metricRow - 1, c)) || periodName;
      if (!brandName || !periodName) return null;
      let b = brands.find((x) => x.name === brandName);
      if (!b) { b = { name: brandName, periods: [] }; brands.push(b); }
      let p = b.periods.find((x) => x.name === periodName);
      if (!p) { p = { name: periodName, isChg: /^(chg|change|diff|var)/i.test(periodName), cols: {} }; b.periods.push(p); }
      p.cols[metrics[c]] = c;
    }
    const complete = brands.every((b) => b.periods.every((p) => ['rank', 'pro', 'det', 'net'].every((k) => p.cols[k])));
    if (!complete || !cols[0] || cols[0] < 2) return null;
    const labelCol = cols[0] - 1;

    const rows = [];
    let started = false;
    for (let r = metricRow + 1; r <= maxRow; r++) {
      const label = text(ws.getCell(r, labelCol));
      const hasData = cols.some((c) => num(ws.getCell(r, c)) !== null || isFormulaCell(ws.getCell(r, c)));
      if (label && hasData) { rows.push({ row: r, label, level: 'other', parent: null }); started = true; }
      else if (!label && !hasData) continue; // blank spacer row: the table may carry on below it
      else if (started) break; // e.g. a footnote: the table has ended
    }
    if (!rows.length) return null;

    assignHierarchy(ws, rows, brands, labelCol);
    const uncalculated = rows.some((x) => cols.some((c) => isFormulaCell(ws.getCell(x.row, c)) && num(ws.getCell(x.row, c)) === null));
    return { metricRow, labelCol, brands, rows, cols, uncalculated };
  }

  /* The column left of the labels can state the hierarchy: A14:A20 = "Product" means row 14 is the parent
     and rows 15-20 its drivers. An unmerged A5 = "Overall" beside B5 = "Overall" is a total. Used only if
     every merge fits that pattern; otherwise the other methods decide. */
  function readFromMerges(ws, rows, byRow, labelCol) {
    if (labelCol < 2) return false;
    const col = labelCol - 1, plan = [];
    for (const x of rows) {
      const cell = ws.getCell(x.row, col);
      if (!cell.isMerged) {
        if (text(cell) && text(cell) === x.label) plan.push({ x, level: 'root' });
        continue;
      }
      const master = cell.master, parent = byRow.get(master.row);
      if (!parent || text(master) !== parent.label) return false;
      if (master.row === x.row) plan.push({ x, level: 'parent' });
      else plan.push({ x, level: 'child', parent });
    }
    if (!plan.some((p) => p.level === 'child')) return false;
    for (const { x, level, parent } of plan) {
      x.level = level;
      if (parent) { x.parent = parent.label; x.parentRow = parent.row; }
    }
    return true;
  }

  /* A parent's Pro and Det equal the sum of the rows under it, for every brand and period. Splits the
     whole table into parent + drivers blocks that all add up; if no such split exists, returns false.
     Exact sums are tried first, then a rounding allowance of 0.05 per row added. */
  function inferFromNumbers(ws, rows, brands) {
    const cols = brands.flatMap((b) => b.periods.filter((p) => !p.isChg).flatMap((p) => [p.cols.pro, p.cols.det]));
    const vec = rows.map((x) => cols.map((c) => num(ws.getCell(x.row, c)) ?? 0));
    const adds = (target, parts, perRow) => target.every((v, i) => Math.abs(parts.reduce((a, p) => a + p[i], 0) - v) <= perRow * (parts.length + 1) + 1e-6);

    const split = (from, perRow) => {
      const failed = new Set();
      const walk = (i) => {
        if (i === rows.length) return [];
        if (failed.has(i)) return null;
        for (let j = i + 1; j < rows.length; j++) {
          if (adds(vec[i], vec.slice(i + 1, j + 1), perRow)) {
            const rest = walk(j + 1);
            if (rest) return [{ parent: i, last: j }, ...rest];
          }
        }
        failed.add(i);
        return null;
      };
      return walk(from);
    };
    const apply = (blocks) => {
      for (const { parent, last } of blocks) {
        const p = rows[parent];
        p.level = 'parent';
        for (let k = parent + 1; k <= last; k++) Object.assign(rows[k], { level: 'child', parent: p.label, parentRow: p.row });
      }
      return true;
    };
    for (const perRow of [0, 0.05]) {
      // With a total on the first row (e.g. Overall), which equals the sum of the parents.
      const withTotal = rows.length > 2 ? split(1, perRow) : null;
      if (withTotal && adds(vec[0], withTotal.map((b) => vec[b.parent]), perRow)) { rows[0].level = 'root'; return apply(withTotal); }
      const noTotal = split(0, perRow);
      if (noTotal) return apply(noTotal);
    }
    return false;
  }

  function assignHierarchy(ws, rows, brands, labelCol) {
    const firstPeriod = brands[0].periods.find((p) => !p.isChg) || brands[0].periods[0];
    const proCol = firstPeriod.cols.pro;
    const byRow = new Map(rows.map((x) => [x.row, x]));
    if (readFromMerges(ws, rows, byRow, labelCol)) return;
    const adopt = (parent, rowNums) => {
      parent.level = 'parent';
      for (const n of rowNums) { const kid = byRow.get(n); if (kid && kid !== parent) { kid.level = 'child'; kid.parent = parent.label; kid.parentRow = parent.row; } }
    };
    // Parents are normally =SUM(D15:D20).
    for (const x of rows) {
      const m = formulaOf(ws.getCell(x.row, proCol)).match(/^\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?\1\$?(\d+)\s*\)\s*$/i);
      if (m) { const nums = []; for (let n = +m[2]; n <= +m[3]; n++) nums.push(n); adopt(x, nums); }
    }
    // A =D6+D14+... row adds up other parents (a total), or adds up plain drivers (a parent).
    for (const x of rows) {
      if (x.level !== 'other') continue;
      const f = formulaOf(ws.getCell(x.row, proCol));
      if (!f) continue;
      const refs = [...f.matchAll(/\$?([A-Z]{1,3})\$?(\d+)/gi)].map((m) => +m[2]).filter((n) => byRow.has(n));
      if (!refs.length) continue;
      if (refs.every((n) => byRow.get(n).level === 'parent')) x.level = 'root';
      else adopt(x, refs);
    }
    // No formulas (e.g. a file written by a script): work it out from the numbers.
    if (!rows.some((x) => x.level === 'parent') && inferFromNumbers(ws, rows, brands)) return;
    // Numbers don't settle it: fall back on bold labels. Bold followed by bold is a total.
    if (!rows.some((x) => x.level === 'parent')) {
      const bold = (x) => !!ws.getCell(x.row, labelCol).font?.bold;
      let current = null;
      rows.forEach((x, i) => {
        if (bold(x)) {
          const next = rows[i + 1];
          if (next && !bold(next)) { x.level = 'parent'; current = x; } else { x.level = 'root'; current = null; }
        } else if (current) { x.level = 'child'; x.parent = current.label; x.parentRow = current.row; }
      });
    }
  }

  const keyOf = (parent, label) => parent + '\u0000' + label;

  /* sort = { brand, period, metric, dir: 'asc' | 'desc' }, matched by name so it works on every sheet. */
  function sortColumn(layout, sort) {
    if (!sort) return null;
    const b = layout.brands.find((x) => x.name === sort.brand);
    const p = b && b.periods.find((x) => x.name === sort.period);
    return (p && p.cols[sort.metric]) || null;
  }

  /* Works out the new row order.
     groupings: [{ name, keys: [keyOf(parent, driver)], pinLast }]. Each named grouping becomes one row, placed
       where its first driver was (or at the bottom of its parent if pinLast), and must sit inside one parent.
       A grouping whose drivers span parents that are not combined is skipped (listed in plan.skipped).
     combine: parent labels to merge into one group, named comboName (default "Journey & Product").
     sort: drivers and groupings are sorted within their parent; pinned groupings stay last. */
  function buildPlan(ws, layout, groupings = [], combine = new Set(), sort = null, comboName = '') {
    const comboParents = layout.rows.filter((x) => x.level === 'parent' && combine.has(x.label));
    const combo = comboParents.length >= 2
      ? { combo: true, label: String(comboName || '').trim() || comboParents.map((p) => p.label).join(' & '),
          parts: comboParents.map((p) => p.label), ticked: comboParents.map((p) => p.row) }
      : null;
    const inCombo = new Set(combo ? combo.ticked : []);
    const info = new Map(layout.rows.map((x) => [x.row, x]));
    const kidsOf = new Map();
    for (const x of layout.rows) {
      if (x.level !== 'child') continue;
      if (!kidsOf.has(x.parentRow)) kidsOf.set(x.parentRow, []);
      kidsOf.get(x.parentRow).push(x);
    }
    const blockOf = (x) => (inCombo.has(x.parentRow) ? 'combo' : x.parentRow);
    const parentLabelOf = (x) => (inCombo.has(x.parentRow) ? combo.label : x.parent);

    // Named groupings that apply on this sheet.
    const byKey = new Map(layout.rows.filter((x) => x.level === 'child').map((x) => [keyOf(x.parent, x.label), x]));
    const active = [], skipped = [], groupOf = new Map();
    for (const def of groupings) {
      const members = [...new Set(def.keys)].map((k) => byKey.get(k)).filter((x) => x && !groupOf.has(x.row)).sort((p, q) => p.row - q.row);
      if (!members.length) continue;
      if (new Set(members.map(blockOf)).size > 1) { skipped.push(def.name); continue; }
      const g = { label: String(def.name).trim(), parent: parentLabelOf(members[0]), ticked: members.map((m) => m.row), pinLast: !!def.pinLast };
      active.push(g);
      for (const m of members) groupOf.set(m.row, g);
    }

    // Sorting: drivers and (unpinned) groupings are reordered within their parent only.
    const sortCol = sortColumn(layout, sort);
    const valueOf = (u) => (u.group ? computeOthers(ws, layout, { ...u.group, newRow: 0 }).values[sortCol] : num(ws.getCell(u.old, sortCol)));
    const unitsFor = (kids) => {
      let units = [];
      const seen = new Set();
      for (const k of kids) {
        const g = groupOf.get(k.row);
        if (!g) units.push({ old: k.row });
        else if (!g.pinLast && !seen.has(g)) { seen.add(g); units.push({ group: g }); }
      }
      if (sortCol) {
        const key = (u) => { const v = valueOf(u); return v === null || v === undefined ? null : Math.round(v * 1e9); };
        units = units.map((u, i) => ({ u, i, v: key(u) })).sort((a, b) =>
          (a.v === null) - (b.v === null) || (a.v === null || a.v === b.v ? 0 : sort.dir === 'asc' ? a.v - b.v : b.v - a.v) || a.i - b.i,
        ).map((o) => o.u);
      }
      const pinned = active.filter((g) => g.pinLast && kids.some((k) => groupOf.get(k.row) === g));
      return [...units, ...pinned.map((g) => ({ group: g }))];
    };

    // Each parent (or the combined parent) is followed by its block of drivers and groupings.
    const underParent = new Set([...kidsOf.values()].flat().map((k) => k.row));
    const seq = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      if (combo && r === combo.ticked[0]) {
        seq.push({ combo }, ...unitsFor(comboParents.flatMap((p) => kidsOf.get(p.row) || [])));
      } else if (!inCombo.has(r) && !underParent.has(r)) {
        seq.push({ old: r });
        if (kidsOf.has(r) && info.get(r)?.level === 'parent') seq.push(...unitsFor(kidsOf.get(r)));
      }
    }

    const order = [], map = {};
    for (const e of seq) {
      order.push(e);
      if (e.combo) combo.newRow = order.length;
      else if (e.group) e.group.newRow = order.length;
      else map[e.old] = order.length;
    }
    for (const g of active) for (const r of g.ticked) map[r] = g.newRow;
    if (combo) {
      const blockRows = comboParents.flatMap((p) => (kidsOf.get(p.row) || []).map((k) => map[k.row]));
      combo.endRow = Math.max(combo.newRow, ...blockRows);
      for (const r of combo.ticked) map[r] = combo.newRow;
    }
    const removed = new Set([...groupOf.keys(), ...inCombo]);
    const maxRow = ws.rowCount, delta = order.length - maxRow;
    return { order, map, active, skipped, combo, removed, maxRow, parentLabelOf, mapRow: (r) => map[r] ?? r + delta };
  }

  /* Pro and Det are summed; Pro-Det and chg follow from them; Rank is redone across brands. */
  function computeOthers(ws, layout, g) {
    const values = {}, formulas = {};
    const row = g.newRow;
    const addr = (c) => colName(c) + row;
    for (const b of layout.brands) {
      for (const p of b.periods) {
        if (p.isChg) continue;
        for (const m of ['pro', 'det']) {
          const c = p.cols[m];
          const parts = g.ticked.map((r) => num(ws.getCell(r, c))).filter((v) => v !== null);
          values[c] = clean(parts.reduce((a, v) => a + v, 0));
          if (parts.length > 1) formulas[c] = parts.map((v, i) => (i === 0 ? String(v) : (v < 0 ? '-' + -v : '+' + v))).join('');
        }
        values[p.cols.net] = clean(values[p.cols.pro] - values[p.cols.det]);
        // ROUND removes decimal noise (0.1+0.2 is not exactly 0.3), so equal scores tie in Rank.
        formulas[p.cols.net] = `ROUND(${addr(p.cols.pro)}-${addr(p.cols.det)},10)`;
      }
    }
    for (const b of layout.brands) {
      for (const p of b.periods) {
        if (p.isChg) continue;
        const me = p.cols.net;
        const rivals = layout.brands.filter((o) => o !== b).map((o) => o.periods.find((q) => q.name === p.name)).filter(Boolean).map((q) => q.cols.net);
        values[p.cols.rank] = 1 + rivals.filter((c) => values[c] > values[me]).length;
        formulas[p.cols.rank] = ['1', ...rivals.map((c) => `(${addr(c)}>${addr(me)})`)].join('+');
      }
      const base = b.periods.filter((p) => !p.isChg);
      const chg = b.periods.find((p) => p.isChg);
      if (chg && base.length >= 2) {
        const first = base[0], last = base[base.length - 1];
        for (const m of ['rank', 'pro', 'det', 'net']) {
          // Rank chg is first minus last, so a positive number means the brand moved up.
          const [a, z] = m === 'rank' ? [first, last] : [last, first];
          values[chg.cols[m]] = clean(values[a.cols[m]] - values[z.cols[m]]);
          formulas[chg.cols[m]] = addr(a.cols[m]) + '-' + addr(z.cols[m]);
        }
      }
    }
    return { values, formulas };
  }

  /* The result table for the screen: read-only, the worksheet is not changed. */
  function preview(ws, layout, groupings, combine, sort, comboName) {
    const plan = buildPlan(ws, layout, groupings, combine, sort, comboName);
    const info = new Map(layout.rows.map((x) => [x.row, x]));
    const rows = [];
    for (const e of plan.order) {
      if (e.group) {
        rows.push({ kind: 'others', label: e.group.label, values: computeOthers(ws, layout, e.group).values, parent: e.group.parent, pinLast: e.group.pinLast });
      } else if (e.combo) {
        rows.push({ kind: 'parent', label: e.combo.label, values: computeOthers(ws, layout, e.combo).values, combined: true });
      } else if (info.has(e.old)) {
        const x = info.get(e.old), values = {};
        for (const c of layout.cols) values[c] = num(ws.getCell(e.old, c));
        rows.push({ kind: x.level, label: x.label, values, parent: x.level === 'child' ? plan.parentLabelOf(x) : x.parent, row: e.old });
      }
    }
    return { rows, plan };
  }

  /* After grouping, the drivers should still add up to their parent row. */
  function reconcile(layout, rows) {
    const issues = [];
    const parents = rows.filter((x) => x.kind === 'parent');
    for (const par of parents) {
      const kids = rows.filter((x) => (x.kind === 'child' || x.kind === 'others') && x.parent === par.label);
      for (const b of layout.brands) for (const p of b.periods) {
        if (p.isChg) continue;
        for (const m of ['pro', 'det']) {
          const c = p.cols[m];
          if (par.values[c] === null) continue;
          const total = kids.reduce((a, k) => a + (k.values[c] || 0), 0);
          if (Math.abs(total - par.values[c]) > 1e-6) issues.push({ parent: par.label, brand: b.name, period: p.name, metric: m, total: clean(total), expected: par.values[c] });
        }
      }
    }
    return issues;
  }

  /* Rewrites row numbers in a formula. A range keeps covering the same rows, including the new one. */
  function remapFormula(formula, sheetName, plans, onSingleRef) {
    const re = /((?:'[^']+'|[A-Za-z0-9_.]+)!)?(?<![A-Za-z0-9_.$])(\$?[A-Z]{1,3})(\$?)(\d+)(?::(\$?[A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])/g;
    return formula.split(/("[^"]*")/).map((part, i) => (i % 2 ? part : part.replace(re, (all, sheet, c1, abs1, r1, c2, abs2, r2) => {
      const target = sheet ? sheet.slice(0, -1).replace(/^'|'$/g, '') : sheetName;
      const plan = plans.get(target);
      if (!plan) return all;
      if (r2 === undefined) {
        if (plan.removed.has(+r1)) onSingleRef?.(all);
        return `${sheet || ''}${c1}${abs1}${plan.mapRow(+r1)}`;
      }
      const [a, z] = mapSpan(plan, +r1, +r2);
      return `${sheet || ''}${c1}${abs1}${a}:${c2}${abs2}${z}`;
    }))).join('');
  }

  function mapSpan(plan, r1, r2) {
    let a = Infinity, z = -Infinity;
    for (let r = r1; r <= r2; r++) { const n = plan.mapRow(r); if (n < a) a = n; if (n > z) z = n; }
    return [a, z];
  }

  function mapRef(ref, plan) {
    return ref.split(/\s+/).filter(Boolean).map((one) => {
      const m = one.match(/^(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?$/);
      if (!m) return one;
      if (m[4] === undefined) return m[1] + plan.mapRow(+m[2]);
      const [a, z] = mapSpan(plan, +m[2], +m[4]);
      return `${m[1]}${a}:${m[3]}${z}`;
    }).join(' ');
  }

  const RED = 'FFF8B4B4', GREEN = 'FFA9DDB0';
  const flatFill = (argb) => ({ fill: { type: 'pattern', pattern: 'solid', bgColor: { argb } } });
  const columnsOf = (ref) => ref.split(/\s+/).filter(Boolean).flatMap((one) => {
    const m = one.match(/^\$?([A-Z]{1,3})\$?\d+(?::\$?([A-Z]{1,3})\$?\d+)?$/);
    if (!m) return [0];
    const out = [];
    for (let c = colNum(m[1]); c <= colNum(m[2] || m[1]); c++) out.push(c);
    return out;
  });

  /* chg columns: flat red for a worse move, flat green for a better one, no fill at zero. Replaces any gradient there.
     Det NPS is reversed, because fewer detractors is an improvement. */
  function applyChgColours(ws, layout, plan) {
    const chgCols = layout.brands.flatMap((b) => b.periods.filter((p) => p.isChg).flatMap((p) => Object.values(p.cols)));
    if (!chgCols.length) return;
    const isChg = new Set(chgCols);
    const kept = (ws.conditionalFormattings || []).map((cf) => {
      const onChgOnly = columnsOf(cf.ref).every((c) => isChg.has(c));
      // Drop gradients, and this tool's own red/green rules from an earlier export, so they don't pile up.
      const ours = (r) => r.type === 'expression' && (r.formulae || []).every((f) => /^ROUND\(\$?[A-Z]{1,3}\$?\d+,6\)[<>]0$/.test(String(f)));
      return onChgOnly ? { ...cf, rules: cf.rules.filter((r) => r.type !== 'colorScale' && !ours(r)) } : cf;
    }).filter((cf) => cf.rules.length);
    let priority = kept.flatMap((cf) => cf.rules).reduce((a, r) => Math.max(a, r.priority || 0), 0);
    const newRows = layout.rows.map((x) => plan.mapRow(x.row));
    const first = Math.min(...newRows), last = Math.max(...newRows);
    const detCols = new Set(layout.brands.flatMap((b) => b.periods.filter((p) => p.isChg).map((p) => p.cols.det)));
    for (const c of [...chgCols].sort((a, b) => a - b)) {
      const cell = colName(c) + first;
      const [below, above] = detCols.has(c) ? [GREEN, RED] : [RED, GREEN];
      // ROUND stops decimal noise from colouring a zero.
      kept.push({ ref: `${colName(c)}${first}:${colName(c)}${last}`, rules: [
        { type: 'expression', priority: ++priority, formulae: [`ROUND(${cell},6)<0`], style: flatFill(below) },
        { type: 'expression', priority: ++priority, formulae: [`ROUND(${cell},6)>0`], style: flatFill(above) },
      ] });
    }
    ws.conditionalFormattings = kept;
  }

  /* Applies the grouping to a loaded workbook, in place. Returns any warnings. */
  function applyToWorkbook(wb, groupings = [], combine = new Set(), sort = null, comboName = '') {
    const warnings = [];
    const layouts = new Map(), plans = new Map();
    wb.eachSheet((ws) => {
      const layout = detectLayout(ws);
      if (layout) { layouts.set(ws.name, layout); plans.set(ws.name, buildPlan(ws, layout, groupings, combine, sort, comboName)); }
    });
    wb.eachSheet((ws) => {
      const plan = plans.get(ws.name);
      if (!plan) return;
      const layout = layouts.get(ws.name);
      for (const name of plan.skipped) warnings.push(`${ws.name}: grouping "${name}" mixes drivers from different groups, so it was not applied. Combine those groups first.`);
      const others = new Map([...plan.active, ...(plan.combo ? [plan.combo] : [])].map((g) => [g, computeOthers(ws, layout, g)]));
      const onSingleRef = (ref) => warnings.push(`${ws.name}: a formula refers directly to grouped cell ${ref}. Check this formula in the export.`);
      const isKept = (r) => !plan.removed.has(r);
      const isRefSum = (f) => f.split('+').length > 1 && f.split('+').every((t) => /^\s*\$?[A-Z]{1,3}\$?\d+\s*$/.test(t));
      const remap = (f) => {
        const hits = [];
        const out = remapFormula(f, ws.name, plans, (ref) => hits.push(ref));
        // =D6+D14+D21+D27 where D6 and D14 are now one combined row: count that row once.
        if (isRefSum(f) && isRefSum(out)) {
          const refs = out.split('+').map((t) => t.trim());
          if (new Set(refs).size < refs.length && new Set(f.split('+').map((t) => t.trim())).size === refs.length) return [...new Set(refs)].join('+');
        }
        hits.forEach(onSingleRef);
        return out;
      };

      const maxCol = ws.columnCount;
      const merges = (ws.model.merges || []).slice();
      merges.forEach((m) => ws.unMergeCells(m));

      const snap = new Map();
      for (let r = 1; r <= plan.maxRow; r++) {
        const row = ws.getRow(r), cells = [];
        for (let c = 1; c <= maxCol; c++) {
          const cell = ws.getCell(r, c);
          cells[c] = { value: cell.value, formula: formulaOf(cell), result: cell.result, style: cell.style };
        }
        snap.set(r, { height: row.height, hidden: row.hidden, outlineLevel: row.outlineLevel, cells });
      }

      const total = Math.max(plan.order.length, plan.maxRow);
      for (let i = 1; i <= total; i++) {
        const e = plan.order[i - 1];
        const row = ws.getRow(i);
        const made = e && (e.group || e.combo);
        const src = e ? snap.get(made ? made.ticked[0] : e.old) : null;
        row.height = src ? src.height : undefined;
        row.hidden = src ? src.hidden : false;
        row.outlineLevel = src ? src.outlineLevel : 0;
        for (let c = 1; c <= maxCol; c++) {
          const cell = ws.getCell(i, c);
          if (!src) { cell.value = null; cell.style = {}; continue; }
          const s = src.cells[c];
          cell.style = s.style;
          if (made) {
            const o = others.get(made);
            if (c === layout.labelCol) cell.value = made.label;
            // Column A beside a combined row: it held the first group's name, so it gets the combined name.
            else if (o.values[c] === undefined) cell.value = e.combo && typeof s.value === 'string' && s.value.trim() === made.parts[0] ? made.label : null;
            else if (o.formulas[c]) cell.value = { formula: o.formulas[c], result: o.values[c] };
            else cell.value = o.values[c];
          } else if (s.formula && isKept(e.old)) {
            cell.value = { formula: remap(s.formula), result: s.result };
          } else {
            cell.value = s.value;
          }
        }
      }

      const comboCols = new Set();
      merges.forEach((m) => {
        const [, col1, top] = m.match(/^([A-Z]+)(\d+)/) || [];
        if (plan.combo && plan.combo.ticked.includes(+top)) { comboCols.add(col1); return; } // replaced by one merge below
        const mapped = mapRef(m, plan);
        const [, r1, r2] = m.match(/[A-Z]+(\d+):[A-Z]+(\d+)/) || [];
        if (r1 && plan.removed.has(+r1) && plan.removed.has(+r2)) return;
        try { ws.mergeCells(mapped); } catch (err) { warnings.push(`${ws.name}: could not restore merged cells ${m}.`); }
      });
      for (const col of comboCols) {
        const ref = `${col}${plan.combo.newRow}:${col}${plan.combo.endRow}`;
        try { if (plan.combo.endRow > plan.combo.newRow) ws.mergeCells(ref); } catch (err) { warnings.push(`${ws.name}: could not merge ${ref}.`); }
      }

      for (const cf of ws.conditionalFormattings || []) {
        cf.ref = mapRef(cf.ref, plan);
        for (const rule of cf.rules || []) {
          if (Array.isArray(rule.formulae)) rule.formulae = rule.formulae.map((f) => (typeof f === 'string' ? remap(f) : f));
        }
      }
      applyChgColours(ws, layout, plan);
    });
    wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
    return warnings;
  }

  return { detectLayout, preview, reconcile, applyToWorkbook, keyOf, colName, colNum };
})();
if (typeof module !== 'undefined') module.exports = GroupingCore;
