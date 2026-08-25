const X = require('./harness.js');
let fails = 0;
const chk = (n,c,e) => { if (c) console.log('  ok   '+n); else { fails++; console.log('  FAIL '+n+(e?' — '+e:'')); } };
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset+u8.byteLength);
const p2 = n => String(n).padStart(2,'0');
const fmt = d => p2(d.getUTCMonth()+1)+'/'+p2(d.getUTCDate())+'/'+String(d.getUTCFullYear()).slice(-2);
const DAY = 86400000;

(async () => {
  console.log('\n-- saveRollingEdits writes items into the ledger --');
  const disk = {};
  X.state.rollingHandle = { async getFile(){ return { async arrayBuffer(){ return buf(disk.f); } }; },
    async createWritable(){ return { async write(b){ disk.f = new Uint8Array(await b.arrayBuffer()); }, async close(){} }; } };
  X.state.rollingName = 'Rolling.xlsx';
  const led = X.ledgerNew();
  const end = Date.UTC(2026,6,9);
  const wkA = fmt(new Date(end-6*DAY))+'-'+fmt(new Date(end));
  const wkB = fmt(new Date(end-13*DAY))+'-'+fmt(new Date(end-7*DAY));
  X.ledgerUpsertWeek(led, wkB, new Map([['CIGNA', 10]]), []);
  X.ledgerUpsertWeek(led, wkA, new Map([['CIGNA', 20]]), [{ desc:'SETTLEMENT', amt: 100 }]);
  X.state.rolling = led;
  X.state.rollItemEdits = { [X.dateKey(wkA)+'|SETTLEMENT']: { wk: X.dateKey(wkA), desc:'SETTLEMENT', amt: 555 } };
  X.state.rollItems = [
    { id:'i1', week: wkB, desc:'UPL PAYMENT', amt: 777 },
    { id:'i2', week: '01/01/99-01/07/99', desc:'GHOST', amt: 1 },   // no such week -> skipped
  ];
  await X.saveRollingEdits();
  const a = X.state.rolling.weeks.get(X.dateKey(wkA)), b = X.state.rolling.weeks.get(X.dateKey(wkB));
  chk('existing item edited in place, not duplicated', JSON.stringify(a.items) === JSON.stringify([{desc:'SETTLEMENT',amt:555}]), JSON.stringify(a.items));
  chk('new item added to its week', JSON.stringify(b.items) === JSON.stringify([{desc:'UPL PAYMENT',amt:777}]), JSON.stringify(b.items));
  chk('pending edits cleared', Object.keys(X.state.rollItemEdits).length === 0 && X.state.rollItems.length === 0);
  chk('payor figures untouched', a.payors.get('CIGNA') === 20 && b.payors.get('CIGNA') === 10);
  const rr = await X.parseRolling(buf(disk.f));
  chk('items survive the file round-trip', JSON.stringify([...rr.weeks.values()].map(w=>w.items)) === JSON.stringify([b.items, a.items]));

  console.log('\n-- cross-foot row detected by behaviour, not position --');
  // a row named "Total Deposits" that restates the column total, parked 4 rows
  // below TOTAL (the old positional heuristic would have missed it)
  const weeks = [];
  for (let i = 5; i >= 0; i--) { const e = end - i*7*DAY; weeks.push(fmt(new Date(e-6*DAY))+'-'+fmt(new Date(e))); }
  const grid = { 1: { 0: {v:'Payor',f:null} } };
  weeks.forEach((w,i) => { grid[1][i+1] = {v:w,f:null}; });
  grid[2] = { 0:{v:'CIGNA',f:null} }; grid[3] = { 0:{v:'AETNA AS01',f:null} };
  weeks.forEach((w,i) => { grid[2][i+1] = {v: 100+i, f:null}; grid[3][i+1] = {v: 200+i, f:null}; });
  grid[4] = { 0:{v:'TOTAL',f:null} };
  grid[5] = { 0:{v:'',f:null} };
  grid[6] = { 0:{v:'REAL SETTLEMENT',f:null} }; grid[6][3] = {v: 9999, f:null};
  grid[7] = { 0:{v:'Total Deposits',f:null} };
  weeks.forEach((w,i) => { grid[7][i+1] = {v: 300+2*i, f:null}; });      // == CIGNA+AETNA
  const ra = { name:'12 Mo Rolling Ave', grid, maxRow:7, maxCol:weeks.length };
  const bytes = X.writeSheetsXlsx([{name:'Sheet1',grid:{1:{}},maxRow:1,maxCol:0}, ra]);
  const mig = await X.parseRolling(buf(bytes));
  const items = [].concat(...[...mig.weeks.values()].map(w => w.items.map(i=>i.desc)));
  chk('cross-foot row dropped', !items.includes('Total Deposits'), JSON.stringify(items));
  chk('real settlement kept', items.includes('REAL SETTLEMENT'));
  chk('droppedRows reported', mig.droppedRows === 1, mig.droppedRows);

  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
