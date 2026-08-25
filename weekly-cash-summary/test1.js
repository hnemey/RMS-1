const X = require('./harness.js');
const fs = require('fs'), vm = require('vm');

// ---- the ORIGINAL grid-based computeRollingDash, frozen here as the baseline ----
// The migration must reproduce this function's numbers exactly. Kept verbatim so
// the guarantee stays testable after the old code is gone.
const OLD_DASH_SRC = "function computeRollingDash(raOverride) {\n  const ra = raOverride || (state.rolling && state.rolling.ra);\n  if (!ra) return null;\n  const hdr = ra.grid[1] || {};\n  const weekCols = [];\n  for (let c = 1; c <= ra.maxCol; c++) {\n    if (hdr[c] && dateKey(hdr[c].v)) weekCols.push(c);\n    else if (weekCols.length) break;\n  }\n  if (!weekCols.length) return null;\n  let totalRow = 0;\n  for (let r = 2; r <= ra.maxRow; r++) {\n    const l = ra.grid[r] && ra.grid[r][0];\n    if (l && alnumKey(l.v) === 'TOTAL') { totalRow = r; break; }\n  }\n  if (!totalRow) totalRow = ra.maxRow + 1;\n  const cats = {};\n  for (let r = 2; r < totalRow; r++) {\n    const l = ra.grid[r] && ra.grid[r][0];\n    if (!l || l.v == null) continue;\n    const cat = rollingCategoryOf(alnumKey(String(l.v).split(',')[0]));\n    const c0 = cats[cat] || (cats[cat] = { series: new Array(weekCols.length).fill(0), avg: 0 });\n    let sum = 0, n = 0;\n    weekCols.forEach((c, i) => {\n      const cell = ra.grid[r][c];\n      const v = cell && typeof cell.v === 'number' ? cell.v : null;\n      if (v != null) { c0.series[i] += v; sum += v; n++; }\n    });\n    // the workbook's rolling average per row ignores blank weeks, and the\n    // grouping column sums the per-row averages — mirror that exactly\n    if (n) c0.avg += sum / n;\n  }\n  const labels = weekCols.map(c => String(hdr[c].v));\n  const totals = new Array(weekCols.length).fill(0);\n  for (const k in cats) cats[k].series.forEach((v, i) => { totals[i] += v; });\n  const avgTotal = Object.values(cats).reduce((a, c) => a + c.avg, 0);\n  return { labels, cats, totals, avgTotal };\n}";
const oc = { alnumKey: X.alnumKey, dateKey: X.dateKey, rollingCategoryOf: X.rollingCategoryOf, state: {}, Number, Object, Array, String, Math };
vm.createContext(oc);
vm.runInContext(OLD_DASH_SRC + '\n;globalThis.oldDash = computeRollingDash;', oc);
const oldDash = oc.oldDash;

// ---- build a synthetic legacy Rolling Ave workbook ----
const p2 = n => String(n).padStart(2,'0');
const fmt = d => p2(d.getUTCMonth()+1)+'/'+p2(d.getUTCDate())+'/'+String(d.getUTCFullYear()).slice(-2);
const DAY = 86400000;
const N_WK = 52;
const weeks = [];
let end = Date.UTC(2026, 6, 9);                       // a Thursday
for (let i = N_WK - 1; i >= 0; i--) {
  const e = new Date(end - i*7*DAY), s = new Date(end - i*7*DAY - 6*DAY);
  weeks.push(fmt(s)+'-'+fmt(e));
}
const LABELS = ['NORIDIAN AK-ID','NORIDIAN ID','STATE OF IDAHO U','STATE OF IDAHO S','BLUE CROSS IDAHO',
  'REGENCE BLUE SHIELD','SH SELECT HEALTH','UNITED HEALTH CARE','PACIFIC SOURCE','AETNA AS01','CIGNA',
  'ZP FIRST 1137','OTHER PHARMACY','SOME WEIRD PAYER LLC','ALL OTHER PAYORS*'];
let seed = 12345;
const rnd = () => (seed = (seed*1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;

const grid = {}; 
grid[1] = { 0: {v:'Payor', f:null} };
weeks.forEach((w,i) => { grid[1][i+1] = { v: w, f: null }; });
LABELS.forEach((lbl, li) => {
  const r = li + 2;
  grid[r] = { 0: { v: lbl, f: null } };
  weeks.forEach((w,i) => {
    if (rnd() < 0.12) return;                          // blank week — exercises the averaging rule
    grid[r][i+1] = { v: Math.round(rnd()*4_000_000)/100, f: null };
  });
});
const totalRow = LABELS.length + 2;
grid[totalRow] = { 0: { v: 'TOTAL', f: null } };
weeks.forEach((w,i) => { grid[totalRow][i+1] = { v: null, f: 'SUM('+X.colRef2(i+1)+'2:'+X.colRef2(i+1)+(totalRow-1)+')' }; });
const checkRow = totalRow + 1;
grid[checkRow] = { 0: { v: 'Check', f: null } };
weeks.forEach((w,i) => { grid[checkRow][i+1] = { v: 1, f: null }; });
// special items pinned to a couple of weeks
grid[checkRow+1] = { 0: { v: 'UPL PAYMENT', f: null } };
grid[checkRow+1][50] = { v: 1234567.89, f: null };
grid[checkRow+2] = { 0: { v: 'SETTLEMENT', f: null } };
grid[checkRow+2][52] = { v: null, f: '73381275' };     // the constant-formula annotation style
const ra = { name: '12 Mo Rolling Ave', grid, maxRow: checkRow+2, maxCol: N_WK };
const s1 = { name: 'Sheet1', grid: { 1: {} }, maxRow: 1, maxCol: 0 };

const bytes = X.writeSheetsXlsx([s1, ra]);
console.log('legacy workbook built:', bytes.length, 'bytes,', N_WK, 'weeks,', LABELS.length, 'payer rows');

(async () => {
  const baseline = oldDash(ra);
  const led = await X.parseRolling(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset+bytes.byteLength));
  X.state.rolling = led;
  const got = X.computeRollingDash(led);

  const near = (a,b) => Math.abs(a-b) < 0.005;
  let fails = 0;
  const chk = (name, cond, extra) => { if (!cond) { fails++; console.log('  FAIL', name, extra||''); } };

  chk('labels', JSON.stringify(baseline.labels) === JSON.stringify(got.labels));
  chk('windowWeeks', led.windowWeeks === N_WK, led.windowWeeks);
  const cA = Object.keys(baseline.cats).sort(), cB = Object.keys(got.cats).sort();
  chk('category set', JSON.stringify(cA) === JSON.stringify(cB), JSON.stringify(cA)+' vs '+JSON.stringify(cB));
  for (const c of cA) {
    if (!got.cats[c]) continue;
    chk('avg '+c, near(baseline.cats[c].avg, got.cats[c].avg), baseline.cats[c].avg+' vs '+got.cats[c].avg);
    for (let i=0;i<baseline.labels.length;i++)
      chk('series '+c+'['+i+']', near(baseline.cats[c].series[i], got.cats[c].series[i]));
  }
  for (let i=0;i<baseline.totals.length;i++) chk('total['+i+']', near(baseline.totals[i], got.totals[i]));
  chk('avgTotal', near(baseline.avgTotal, got.avgTotal), baseline.avgTotal+' vs '+got.avgTotal);

  console.log('\nold avgTotal :', baseline.avgTotal.toFixed(2));
  console.log('new avgTotal :', got.avgTotal.toFixed(2));
  console.log('unmapped labels carried over:', JSON.stringify(led.unmapped));
  console.log('special items found:', JSON.stringify(X.rollingSpecialItems()));
  console.log(fails ? '\nMIGRATION: ' + fails + ' MISMATCH(ES)' : '\nMIGRATION: numbers identical ✓');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
