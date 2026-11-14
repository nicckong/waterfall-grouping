/* Runs the tool's grouping logic on a file, the same way the page does, and saves the result.
   Usage: node tests/run_export.js in.xlsx out.xlsx "Parent|Driver" ...
     Each "Parent|Driver" goes into that parent's "<Parent> - Others" grouping (kept at the bottom); under
     combined parents it goes into "<combined name> - Others". COMBINE lists parents in table order.
   Optional environment:
     GROUPS='[{"name":"Digital","members":["Journey|Efficiency","Journey|Easy Logon"],"pinLast":false}]'
     COMBINE=Journey,Product   COMBO_NAME="Customer Experience"   SORT="Bank A|2026 1H|net|desc" */
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
  const combineList = (process.env.COMBINE || '').split(',').filter(Boolean);
  const comboLabel = process.env.COMBO_NAME || combineList.join(' & ');
  const others = new Map();
  for (const p of picks) {
    const [parent, driver] = p.split('|');
    const owner = combineList.length >= 2 && combineList.includes(parent) ? comboLabel : parent;
    if (!others.has(owner)) others.set(owner, { name: `${owner} - Others`, keys: [], pinLast: true });
    others.get(owner).keys.push(core.keyOf(parent, driver));
  }
  const custom = JSON.parse(process.env.GROUPS || '[]').map((g) => ({
    name: g.name, pinLast: !!g.pinLast, keys: g.members.map((m) => { const [a, b] = m.split('|'); return core.keyOf(a, b); }),
  }));
  const groupings = [...others.values(), ...custom];
  const comboName = process.env.COMBO_NAME || '';
  const combine = new Set(combineList);
  const [brand, period, metric, dir] = (process.env.SORT || '').split('|');
  const sort = process.env.SORT ? { brand, period, metric, dir } : null;
  wb.eachSheet((ws) => {
    const l = core.detectLayout(ws);
    if (!l) { console.log(ws.name, 'NO LAYOUT'); return; }
    const levels = l.rows.reduce((a, r) => ((a[r.level] = (a[r.level] || 0) + 1), a), {});
    const issues = core.reconcile(l, core.preview(ws, l, groupings, combine, sort, comboName).rows).length;
    console.log(ws.name, 'brands', l.brands.map((b) => b.name).join(','), 'levels', JSON.stringify(levels), 'reconcile issues', issues);
  });
  console.log('warnings', core.applyToWorkbook(wb, groupings, combine, sort, comboName));
  await wb.xlsx.writeFile(out);
})().catch((e) => { console.error(e); process.exit(1); });
