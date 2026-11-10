/* Generates made-up test workbooks with the same layout as the real files: brands across the top,
   periods (two periods + chg) under each, Rank / Pro NPS / Det NPS / Pro-Det NPS under each period,
   groups merged in column A, labels in column B, Overall on row 5. Numbers are random but consistent.
   Usage: node tests/make_fixtures.js <output dir> */
const ExcelJS = require('exceljs');
const path = require('path');

const L = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
const BANKS = ['Bank A', 'Bank B', 'Bank C', 'Bank D', 'Bank E', 'Bank F'];
const PERIODS = ['2026 1H', '2026 3Q YTD Interim', 'chg'];
const TREE = {
  Journey: ['Ease & Convenience', 'Informative', 'Digital Capability & Design', 'Efficiency', '(Digital) Service Quality', '(Fees & Rewards) Transparency', 'Easy Logon'],
  Product: ['Benefits & Rewards', 'Fees & Charges', 'Level of Return', 'Product Features & Design', 'Product Fit', 'Product Range'],
  Distribution: ['Accessibility', 'Personalized & Helpful Service', 'Proactive Staff', 'Professional Staff', 'Queue Time'],
  Brand: ['Brand Reputation', 'Trust & Security', 'Brand Image'],
};

async function make(file, { banks = BANKS, tree = TREE, valuesOnly = false, spacerBefore = null, seed = 7 } = {}) {
  const wb = new ExcelJS.Workbook();
  for (const [si, name] of ['Overall', 'Affluent', 'Mass'].entries()) {
    let st = seed * 97 + si * 13 + 1;
    const rnd = () => (st = (st * 16807) % 2147483647) / 2147483647;
    const r1 = (x) => Math.round(x * 10) / 10;
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 2, ySplit: 4 }] });
    ws.getColumn(1).width = 15; ws.getColumn(2).width = 34;
    const start = (i) => 3 + 12 * i;
    const head = { font: { name: 'Arial', size: 10, bold: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }, alignment: { horizontal: 'center' } };
    banks.forEach((b, i) => {
      ws.getCell(1, start(i)).value = b;
      ws.getCell(1, start(i)).font = { name: 'Arial', size: 11, bold: true };
      PERIODS.forEach((p, k) => {
        ws.getCell(2, start(i) + 4 * k).value = p;
        ['Rank', 'Pro NPS', 'Det NPS', 'Pro-Det NPS'].forEach((m, j) => { const c = ws.getCell(3, start(i) + 4 * k + j); c.value = m; c.style = head; });
      });
    });
    ws.getCell(4, 2).value = 'L3';

    // Rows: Overall, then each group's parent row followed by its drivers (optionally a blank row before a group).
    const rows = [{ label: 'Overall', kind: 'root', row: 5 }];
    let r = 6;
    for (const [p, kids] of Object.entries(tree)) {
      if (p === spacerBefore) r++;
      rows.push({ label: p, kind: 'parent', row: r, a: r + 1, b: r + kids.length });
      kids.forEach((k, i) => rows.push({ label: k, kind: 'kid', parent: p, row: r + 1 + i }));
      r += kids.length + 1;
    }
    const lastRow = r - 1;
    const parents = rows.filter((x) => x.kind === 'parent');
    const v = new Map(rows.map((x) => [x, {}]));
    const get = (x, i, p, m) => v.get(x)[`${i}|${p}|${m}`];
    const set = (x, i, p, m, val) => { v.get(x)[`${i}|${p}|${m}`] = val; };
    for (const x of rows.filter((q) => q.kind === 'kid')) banks.forEach((_, i) => [0, 4].forEach((p) => {
      set(x, i, p, 1, r1(rnd() * 3));
      set(x, i, p, 2, rnd() < 0.55 ? 0 : r1(rnd() * 2)); // plenty of exact zeros
    }));
    for (const x of parents) {
      const kids = rows.filter((q) => q.parent === x.label);
      banks.forEach((_, i) => [0, 4].forEach((p) => [1, 2].forEach((m) => set(x, i, p, m, kids.reduce((a, k) => a + get(k, i, p, m), 0)))));
    }
    banks.forEach((_, i) => [0, 4].forEach((p) => [1, 2].forEach((m) => set(rows[0], i, p, m, parents.reduce((a, k) => a + get(k, i, p, m), 0)))));
    const key = (x) => Math.round(x * 1e9);
    for (const x of rows) {
      banks.forEach((_, i) => [0, 4].forEach((p) => set(x, i, p, 3, get(x, i, p, 1) - get(x, i, p, 2))));
      banks.forEach((_, i) => [0, 4].forEach((p) => set(x, i, p, 0, 1 + banks.filter((__, j) => key(get(x, j, p, 3)) > key(get(x, i, p, 3))).length)));
    }

    for (const x of rows) {
      const bold = x.kind !== 'kid';
      const lab = ws.getCell(x.row, 2);
      lab.value = x.label; lab.font = { name: 'Arial', size: 10, bold }; lab.alignment = { horizontal: 'left' };
      if (x.kind === 'root') { ws.getCell(x.row, 1).value = x.label; ws.getCell(x.row, 1).font = { bold: true }; }
      banks.forEach((_, i) => {
        const s = start(i);
        const put = (col, val, f, fmt, blue) => {
          const c = ws.getCell(x.row, col);
          c.value = f && !valuesOnly ? { formula: f, result: val } : val;
          c.numFmt = fmt;
          c.font = { name: 'Arial', size: 10, bold, ...(blue ? { color: { argb: 'FF0000FF' } } : {}) };
          c.alignment = { horizontal: 'center' };
        };
        [0, 4].forEach((p) => {
          const rival = banks.map((__, j) => j).filter((j) => j !== i).map((j) => `(${L(start(j) + p + 3)}${x.row}>${L(s + p + 3)}${x.row})`);
          put(s + p, get(x, i, p, 0), '1+' + rival.join('+'), '0');
          [1, 2].forEach((m) => {
            let f = null;
            if (x.kind === 'parent') f = `SUM(${L(s + p + m)}${x.a}:${L(s + p + m)}${x.b})`;
            if (x.kind === 'root') f = parents.map((q) => L(s + p + m) + q.row).join('+');
            put(s + p + m, get(x, i, p, m), f, '0.0', x.kind === 'kid');
          });
          put(s + p + 3, get(x, i, p, 3), `${L(s + p + 1)}${x.row}-${L(s + p + 2)}${x.row}`, '0.0');
        });
        put(s + 8, get(x, i, 0, 0) - get(x, i, 4, 0), `${L(s)}${x.row}-${L(s + 4)}${x.row}`, '0;\\-0;0');
        [1, 2, 3].forEach((m) => put(s + 8 + m, get(x, i, 4, m) - get(x, i, 0, m), `${L(s + 4 + m)}${x.row}-${L(s + m)}${x.row}`, '0.0'));
      });
    }
    for (const p of parents) {
      ws.mergeCells(`A${p.row}:A${p.b}`);
      const a = ws.getCell(p.row, 1);
      a.value = p.label; a.font = { bold: true }; a.alignment = { vertical: 'middle', horizontal: 'center' };
    }
    // Gradient colours on the chg columns, as in the real files (the tool replaces them with flat red/green).
    banks.forEach((_, i) => [0, 1, 2, 3].forEach((m) => ws.addConditionalFormatting({
      ref: `${L(start(i) + 8 + m)}5:${L(start(i) + 8 + m)}${lastRow}`,
      rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'num', value: 0 }, { type: 'max' }], color: [{ argb: 'FFF8B4B4' }, { argb: 'FFFFFFFF' }, { argb: 'FFA9DDB0' }] }],
    })));
    ws.getCell(lastRow + 2, 2).value = 'Made-up test data.';
    ws.getCell(lastRow + 3, 2).value = 'Parent rows = sum of their drivers; Overall = sum of the parents.';
  }
  await wb.xlsx.writeFile(file);
}

(async () => {
  const out = process.argv[2] || path.join(__dirname, 'out');
  await make(path.join(out, 'formulas.xlsx'));
  await make(path.join(out, 'values_only.xlsx'), { valuesOnly: true });
  await make(path.join(out, 'five_banks.xlsx'), {
    banks: BANKS.slice(0, 5), seed: 3,
    tree: { Journey: ['Ease', 'Speed', 'Logon'], Product: ['Product Fit', 'Product Range', 'Fees'], Brand: ['Brand Trust', 'Other Brand Images'] },
  });
  await make(path.join(out, 'blank_row.xlsx'), { valuesOnly: true, spacerBefore: 'Distribution', seed: 11 });
  console.log('fixtures written to', out);
})().catch((e) => { console.error(e); process.exit(1); });
