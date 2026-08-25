const X = require('./harness.js'), fs = require('fs');
let fails = 0;
const chk = (n,c,e) => { if (c) console.log('  ok   '+n); else { fails++; console.log('  FAIL '+n+(e?' — '+e:'')); } };
const near = (a,b) => Math.abs(a-b) < 0.005;
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset+u8.byteLength);
const D = (y,m,d) => new Date(Date.UTC(y,m-1,d));

// ---- a mock linked folder ----
const disk = {};
const fileHandle = name => ({
  name,
  async getFile() { return { name, lastModified: 1, async arrayBuffer(){ return buf(disk[name]); } }; },
  async createWritable() {
    return { async write(blob) { this._b = blob; disk[name] = new Uint8Array(await blob.arrayBuffer()); }, async close(){} };
  },
});
X.state.folderHandle = { name: 'Cash', async getFileHandle(n) { return fileHandle(n); } };
X.state.rollingHandle = fileHandle('Rolling Ave.xlsx');
X.state.rollingName = 'Rolling Ave.xlsx';

// ---- seed the ledger with 52 weeks of history ----
const p2 = n => String(n).padStart(2,'0');
const fmt = d => p2(d.getUTCMonth()+1)+'/'+p2(d.getUTCDate())+'/'+String(d.getUTCFullYear()).slice(-2);
const DAY = 86400000, end = Date.UTC(2026, 6, 2);          // history runs through Thu 07/02/26
const led = X.ledgerNew();
for (let i = 51; i >= 0; i--) {
  const e = end - i*7*DAY;
  X.ledgerUpsertWeek(led, fmt(new Date(e-6*DAY))+'-'+fmt(new Date(e)),
    new Map([['NORIDIANAKID', 100000], ['CIGNA', 20000], ['ALLOTHERPAYORS', 5000]]), []);
}
led.legacy = true;                                          // exercise the backup + conversion path
X.state.rolling = led;
const oldestKey = X.ledgerSorted(led)[0].key;

// ---- a week of report rows: Fri 07/03/26 .. Thu 07/09/26 ----
// Payer = the bank's payer text, PayorName = the payor class the app assigns
const row = (id, day, payer, cls, amt) => ({
  _id: 'r'+id, BankDetailSID: id, LockBoxDate: D(2026,7,day),
  Payer: payer, PayorName: cls, Comments: payer,
  DetailLineItemDepositAmount: amt, BatchAmount: amt, BankBatchNo: 'B'+id, BankSID: id,
});
X.state.rows = [
  row(1, 3, 'NORIDIAN AK-ID', 'MEDICARE',  250000),
  row(2, 6, 'NORIDIAN ID',    'MEDICARE',   40000),
  row(3, 7, 'CIGNA',          'CIGNA',      15000),
  row(4, 9, 'WHOEVER LLC',    'COMMERCIAL',  3000),
];
X.state.adjustments = [{ id:'a1', desc:'UPL PAYMENT', cat:'MEDICARE', amt: 50000, week: '07/03/26-07/09/26' }];

(async () => {
  disk['Rolling Ave.xlsx'] = await X.writeRollingXlsx(led);   // stand-in for the on-disk original
  const before = X.computeRollingDash();
  console.log('\n-- pre-save --');
  chk('52 weeks in window', before.labels.length === 52);
  chk('window ends 07/02/26', before.labels[51] === '06/26/26-07/02/26', before.labels[51]);

  console.log('\n-- preview (wcsRolling) folds the open report in --');
  const prev = X.wcsRolling ? X.wcsRolling() : null;

  console.log('\n-- saveToRolling --');
  await X.saveToRolling();

  const after = X.computeRollingDash();
  chk('new week is now the last column', after.labels[51] === '07/03/26-07/09/26', after.labels[51]);
  chk('window still 52', after.labels.length === 52);
  chk('OLDEST WEEK NOT DESTROYED', X.state.rolling.weeks.has(oldestKey));
  chk('stored weeks grew to 53', X.state.rolling.weeks.size === 53, X.state.rolling.weeks.size);
  chk('legacy flag cleared', X.state.rolling.legacy === false);
  chk('backup file written', Object.keys(disk).some(k => /pre-ledger backup/.test(k)), Object.keys(disk).join(', '));

  const wk = X.state.rolling.weeks.get(X.dateKey('07/03/26-07/09/26'));
  chk('MEDICARE netted of the adjustment (250000-50000)', near(wk.payors.get('NORIDIANAKID'), 200000), wk.payors.get('NORIDIANAKID'));
  chk('NORIDIAN ID bucketed separately', near(wk.payors.get('NORIDIANID'), 40000), wk.payors.get('NORIDIANID'));
  chk('CIGNA bucketed', near(wk.payors.get('CIGNA'), 15000));
  chk('unknown payer -> ALL OTHER PAYORS', near(wk.payors.get('ALLOTHERPAYORS'), 3000));
  chk('adjustment recorded as a special item', JSON.stringify(wk.items) === JSON.stringify([{desc:'UPL PAYMENT', amt:50000}]));

  console.log('\n-- what landed on disk --');
  const reread = await X.parseRolling(buf(disk['Rolling Ave.xlsx']));
  const d2 = X.computeRollingDash(reread);
  chk('file re-reads to 53 weeks', reread.weeks.size === 53, reread.weeks.size);
  chk('dashboard identical after reload', near(d2.avgTotal, after.avgTotal), d2.avgTotal + ' vs ' + after.avgTotal);
  chk('preview matched the saved result', prev && near(prev.avgTotal, after.avgTotal), prev ? prev.avgTotal + ' vs ' + after.avgTotal : 'no preview');

  console.log('\n-- re-saving the SAME week after deleting the adjustment --');
  X.state.adjustments = [];
  await X.saveToRolling();
  const wk2 = X.state.rolling.weeks.get(X.dateKey('07/03/26-07/09/26'));
  chk('stale adjustment row is gone', wk2.items.length === 0, JSON.stringify(wk2.items));
  chk('MEDICARE back to the gross 250000', near(wk2.payors.get('NORIDIANAKID'), 250000), wk2.payors.get('NORIDIANAKID'));
  chk('still 53 weeks (no duplicate)', X.state.rolling.weeks.size === 53, X.state.rolling.weeks.size);

  console.log('\n-- a stray old lockbox date no longer evicts history --');
  X.state.rows.push(row(9, 3, 'CIGNA', 'CIGNA', 111));
  X.state.rows[4].LockBoxDate = D(2026,1,15);
  await X.saveToRolling();
  chk('stray week added, nothing dropped', X.state.rolling.weeks.has(oldestKey));
  chk('window still 52', X.computeRollingDash().labels.length === 52);

  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
