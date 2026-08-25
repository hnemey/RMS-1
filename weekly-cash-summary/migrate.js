const X = require('./harness.js'), fs = require('fs'), vm = require('vm');
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset+u8.byteLength);

// the frozen pre-change grid dashboard, to prove the numbers don't move
const OLD = fs.readFileSync('test1.js','utf8').match(/const OLD_DASH_SRC = ([\s\S]*?);\nconst oc =/)[1];
const oc = { alnumKey:X.alnumKey, dateKey:X.dateKey, rollingCategoryOf:X.rollingCategoryOf, state:{}, Number, Object, Array, String, Math };
vm.createContext(oc);
vm.runInContext('const S=' + OLD + ';(0,eval)(S + ";globalThis.oldDash=computeRollingDash;")', oc);
const oldDash = oc.oldDash;

(async () => {
  const src = fs.readFileSync('source.xlsx');
  const sheets = await X.readSheetGrids(buf(new Uint8Array(src)));
  const ra = sheets.find(s => /rolling/i.test(s.name));
  const baseline = oldDash(ra);

  const led = await X.parseRolling(buf(new Uint8Array(src)));
  X.state.rolling = led;
  const got = X.computeRollingDash(led);

  const near = (a,b) => Math.abs(a-b) < 0.005;
  let bad = 0;
  const chk = (n,c,e) => { if (!c) { bad++; console.log('  MISMATCH', n, e||''); } };
  chk('labels', JSON.stringify(baseline.labels) === JSON.stringify(got.labels));
  const cats = Object.keys(baseline.cats).sort();
  chk('category set', JSON.stringify(cats) === JSON.stringify(Object.keys(got.cats).sort()),
      JSON.stringify(cats) + ' vs ' + JSON.stringify(Object.keys(got.cats).sort()));
  for (const c of cats) {
    if (!got.cats[c]) continue;
    chk('avg ' + c, near(baseline.cats[c].avg, got.cats[c].avg), baseline.cats[c].avg + ' vs ' + got.cats[c].avg);
    baseline.cats[c].series.forEach((v,i) => chk('series ' + c + '[' + i + ']', near(v, got.cats[c].series[i])));
  }
  baseline.totals.forEach((v,i) => chk('total[' + i + ']', near(v, got.totals[i])));
  chk('avgTotal', near(baseline.avgTotal, got.avgTotal), baseline.avgTotal + ' vs ' + got.avgTotal);

  console.log('=== MIGRATION OF THE REAL WORKBOOK ===');
  console.log('window weeks kept      :', led.windowWeeks);
  console.log('weeks stored in ledger :', led.weeks.size);
  const sorted = X.ledgerSorted(led);
  console.log('span                   :', sorted[0].range, '->', sorted[sorted.length-1].range);
  const bset = new Set(); let items = 0;
  for (const w of sorted) { for (const b of w.payors.keys()) bset.add(b); items += w.items.length; }
  console.log('payor buckets          :', bset.size);
  console.log('special item rows      :', items);
  console.log('cross-foot rows dropped:', led.droppedRows);
  console.log('unmapped labels        :', led.unmapped.length, JSON.stringify(led.unmapped));
  console.log('de-collided labels     :', led.split.length, JSON.stringify(led.split));
  console.log();
  console.log('12-mo avg total  OLD grid :', baseline.avgTotal.toFixed(2));
  console.log('12-mo avg total  NEW ledger:', got.avgTotal.toFixed(2));
  console.log(bad ? '\n*** ' + bad + ' MISMATCH(ES) ***' : '\nNumbers identical across all ' + cats.length + ' categories and ' + baseline.labels.length + ' weeks.');

  fs.writeFileSync('Rolling_Ave_Ledger.xlsx', Buffer.from(await X.writeRollingXlsx(led)));
  console.log('\nwrote Rolling_Ave_Ledger.xlsx');
  process.exit(bad ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
