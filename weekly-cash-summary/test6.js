const X = require('./harness.js');
let fails = 0;
const chk = (n,c,e) => { if (c) console.log('  ok   '+n); else { fails++; console.log('  FAIL '+n+(e?' — '+e:'')); } };
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset+u8.byteLength);
const D = (y,m,d) => new Date(Date.UTC(y,m-1,d));
const p2 = n => String(n).padStart(2,'0');
const fmt = d => p2(d.getUTCMonth()+1)+'/'+p2(d.getUTCDate())+'/'+String(d.getUTCFullYear()).slice(-2);
const DAY = 86400000;

const disk = {};
X.state.rollingHandle = {
  async getFile(){ return { async arrayBuffer(){ return buf(disk.f); } }; },
  async createWritable(){ return { async write(b){ disk.f = new Uint8Array(await b.arrayBuffer()); }, async close(){} }; } };
X.state.rollingName = 'Rolling Ave.xlsx';
X.state.folderHandle = null;

const hEnd = Date.UTC(2026,6,2);
const wkOf = e => fmt(new Date(e-6*DAY))+'-'+fmt(new Date(e));
const led = X.ledgerNew();
for (let i = 51; i >= 0; i--)
  X.ledgerUpsertWeek(led, wkOf(hEnd - i*7*DAY), new Map([['NORIDIANAKID', 100000], ['CIGNA', 8000]]), []);
X.state.rolling = led;
X.state.rows = [{ _id:'r1', BankDetailSID:1, LockBoxDate: D(2026,7,6), Payer:'NORIDIAN AK-ID',
  PayorName:'MEDICARE', Comments:'NORIDIAN AK-ID', DetailLineItemDepositAmount: 250000,
  BatchAmount: 250000, BankBatchNo:'B1', BankSID:1 }];

const PAST = wkOf(hEnd - 10*7*DAY);          // 04/17/26-04/23/26, in the ledger, not in the report
const REPORTWK = '07/03/26-07/09/26';
const pastPayor = () => X.state.rolling.weeks.get(X.dateKey(PAST)).payors.get('NORIDIANAKID');
const pastItems = () => JSON.stringify(X.state.rolling.weeks.get(X.dateKey(PAST)).items);

(async () => {
  console.log('\n-- an earlier week now actually applies --');
  X.state.adjustments = [{ id:'a1', desc:'Prior-period UPL', cat:'MEDICARE', amt:'12345', week: PAST }];
  await X.saveToRolling();
  chk('payor figure moved 100000 -> 87655', pastPayor() === 87655, pastPayor());
  chk('item row written', pastItems() === JSON.stringify([{desc:'Prior-period UPL', amt:12345}]), pastItems());
  const onDisk = await X.parseRolling(buf(disk.f));
  chk('it reached the .xlsx', onDisk.weeks.get(X.dateKey(PAST)).payors.get('NORIDIANAKID') === 87655);

  console.log('\n-- re-saving does not double-subtract --');
  await X.saveToRolling();
  await X.saveToRolling();
  chk('still 87655 after two more saves', pastPayor() === 87655, pastPayor());
  chk('still one item row', X.state.rolling.weeks.get(X.dateKey(PAST)).items.length === 1);

  console.log('\n-- correcting the amount applies only the delta --');
  X.state.adjustments[0].amt = '20000';
  await X.saveToRolling();
  chk('payor figure now 80000', pastPayor() === 80000, pastPayor());
  chk('item row updated, not duplicated', pastItems() === JSON.stringify([{desc:'Prior-period UPL', amt:20000}]), pastItems());

  console.log('\n-- zeroing the amount backs it out entirely --');
  X.state.adjustments[0].amt = '0';
  await X.saveToRolling();
  chk('payor figure restored to 100000', pastPayor() === 100000, pastPayor());
  chk('item row removed', pastItems() === '[]', pastItems());

  console.log('\n-- an exactly-cancelling adjustment blanks the bucket --');
  X.state.adjustments = [{ id:'a2', desc:'Full clawback', cat:'CIGNA', amt:'8000', week: PAST }];
  await X.saveToRolling();
  chk('CIGNA bucket removed for that week', !X.state.rolling.weeks.get(X.dateKey(PAST)).payors.has('CIGNA'));
  X.state.adjustments[0].amt = '0';
  await X.saveToRolling();
  chk('and comes back when zeroed', X.state.rolling.weeks.get(X.dateKey(PAST)).payors.get('CIGNA') === 8000);

  console.log('\n-- adjIssues --');
  const rk = X.refreshAdjWeekKeys();
  const iss = a => X.adjIssues(a, rk).map(i => i.level + ':' + i.msg.slice(0, 26));
  chk('good in-report row is clean',
      X.adjIssues({desc:'Settlement', cat:'MEDICARE', amt:'500', week: REPORTWK}, rk).length === 0);
  chk('zero amount blocks',
      X.adjBlocked({desc:'X', cat:'MEDICARE', amt:'00', week: REPORTWK}, rk).length === 1, iss({desc:'X',cat:'MEDICARE',amt:'00',week:REPORTWK}));
  chk('week in neither report nor ledger blocks',
      X.adjBlocked({desc:'X', cat:'MEDICARE', amt:'5', week:'01/02/99-01/08/99'}, rk).length === 1);
  chk('earlier ledger week is allowed, flagged info',
      X.adjBlocked({desc:'X', cat:'MEDICARE', amt:'5', week: PAST}, rk).length === 0 &&
      X.adjLevel(X.adjIssues({desc:'X', cat:'MEDICARE', amt:'5', week: PAST}, rk)) === 'info');
  chk('blank description warns but does not block',
      X.adjBlocked({desc:'', cat:'MEDICARE', amt:'5', week: REPORTWK}, rk).length === 0 &&
      X.adjLevel(X.adjIssues({desc:'', cat:'MEDICARE', amt:'5', week: REPORTWK}, rk)) === 'warn');

  console.log('\n-- blocked rows are skipped, not silently --');
  X.state.adjustments = [{ id:'a3', desc:'Ghost', cat:'MEDICARE', amt:'999', week:'01/02/99-01/08/99' }];
  const beforeSize = X.state.rolling.weeks.size;
  await X.saveToRolling();
  chk('no phantom week created', X.state.rolling.weeks.size === beforeSize, X.state.rolling.weeks.size);

  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
