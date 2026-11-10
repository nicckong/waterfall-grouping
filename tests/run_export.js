/* Runs the tool's grouping logic on a file, the same way the page does, and saves the result.
   Usage: node tests/run_export.js in.xlsx out.xlsx "Parent|Driver" ...
   Optional environment: COMBINE=Journey,Product   SORT="Bank A|2026 1H|net|desc" */
const ExcelJS = require('exceljs');
const core = require('../src/core.js');

// Same number-format fix as the build applies to the browser copy of ExcelJS.
const NumFmtXform = require('exceljs/lib/xlsx/xform/style/numfmt-xform');
const parseOpen = NumFmtXform.prototype.parseOpen;
NumFmtXform.prototype.parseOpen = function (node) {
  const handled = parseOpen.call(this, node);
  if (handled && node.name === 'numFmt') this.model.formatCode = node.attributes.formatCode;
  return handled;
};

(async () => {
  const [inp, out, ...picks] = process.argv.slice(2);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inp);
  const selected = new Set(picks.map((p) => { const [a, b] = p.split('|'); return core.keyOf(a, b); }));
  const combine = new Set((process.env.COMBINE || '').split(',').filter(Boolean));
  const [brand, period, metric, dir] = (process.env.SORT || '').split('|');
  const sort = process.env.SORT ? { brand, period, metric, dir } : null;
  wb.eachSheet((ws) => {
    const l = core.detectLayout(ws);
    if (!l) { console.log(ws.name, 'NO LAYOUT'); return; }
    const levels = l.rows.reduce((a, r) => ((a[r.level] = (a[r.level] || 0) + 1), a), {});
    const issues = core.reconcile(l, core.preview(ws, l, selected, combine, sort).rows).length;
    console.log(ws.name, 'brands', l.brands.map((b) => b.name).join(','), 'levels', JSON.stringify(levels), 'reconcile issues', issues);
  });
  console.log('warnings', core.applyToWorkbook(wb, selected, combine, sort));
  await wb.xlsx.writeFile(out);
})().catch((e) => { console.error(e); process.exit(1); });
