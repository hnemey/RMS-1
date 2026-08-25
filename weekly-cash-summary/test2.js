const X = require('./harness.js');
let fails = 0;
const chk = (n, c, e) => { if (c) console.log('  ok   ' + n); else { fails++; console.log('  FAIL ' + n + (e ? ' — ' + e : '')); } };
const near = (a,b) => Math.abs(a-b) < 0.005;
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);

const p2 = n => String(n).padStart(2,'0');
const fmt = d => p2(d.getUTCMonth()+1)+'/'+p2(d.getUTCDate())+'/'+String(d.getUTCFullYear()).slice(-2);
const DAY = 86400000;
const wkRange = (endMs) => fmt(new Date(endMs - 6*DAY)) + '-' + fmt(new Date(endMs));

(async () => {
  // ---------- build a ledger directly ----------
  const led = X.ledgerNew();
  let end = Date.UTC(2026, 6, 9);
  const ranges = [];
  for (let i = 59; i >= 0; i--) { const r = wkRange(end - i*7*DAY); ranges.push(r); }   // 60 weeks > 52 window
  ranges.forEach((r, i) => {
    X.ledgerUpsertWeek(led, r,
      new Map([['NORIDIANAKID', 1000 + i], ['CIGNA', 500 + i], ['ALLOTHERPAYORS', 250.55]]),
      i === 59 ? [{ desc: 'UPL PAYMENT', amt: 4321.5 }] : []);
  });
  X.state.rolling = led;

  console.log('\n-- window --');
  chk('ledger holds all 60 weeks', led.weeks.size === 60, led.weeks.size);
  chk('window trims to 52', X.ledgerWindow(led, 52).length === 52);
  chk('window is the newest 52', X.ledgerWindow(led,52)[51].range === ranges[59]);
  chk('oldest week still stored', led.weeks.has(X.dateKey(ranges[0])));

  console.log('\n-- xlsx round-trip --');
  const bytes = X.writeRollingXlsx(led);
  const led2 = await X.parseRolling(buf(bytes));
  chk('re-parsed week count', led2.weeks.size === 60, led2.weeks.size);
  const d1 = X.computeRollingDash(led), d2 = X.computeRollingDash(led2);
  chk('labels identical', JSON.stringify(d1.labels) === JSON.stringify(d2.labels));
  chk('avgTotal identical', near(d1.avgTotal, d2.avgTotal), d1.avgTotal + ' vs ' + d2.avgTotal);
  chk('totals identical', d1.totals.every((v,i) => near(v, d2.totals[i])));
  chk('MEDICARE avg identical', near(d1.cats.MEDICARE.avg, d2.cats.MEDICARE.avg));
  chk('special item survives', JSON.stringify([...led2.weeks.values()].pop().items) === JSON.stringify([{desc:'UPL PAYMENT', amt:4321.5}]));
  chk('written file is a single Ledger sheet', (await X.readSheetGrids(buf(bytes))).map(s=>s.name).join() === 'Ledger');

  console.log('\n-- upsert replaces a week whole (the stale-adjustment bug) --');
  const before = X.ledgerWindow(led, 52).length;
  X.ledgerUpsertWeek(led, ranges[59], new Map([['CIGNA', 99]]), []);
  const last = led.weeks.get(X.dateKey(ranges[59]));
  chk('old payor buckets gone', !last.payors.has('NORIDIANAKID') && last.payors.get('CIGNA') === 99);
  chk('old special item gone', last.items.length === 0);
  chk('week count unchanged', X.ledgerWindow(led,52).length === before);

  console.log('\n-- adding a new week never drops history --');
  const newRange = wkRange(end + 7*DAY);
  X.ledgerUpsertWeek(led, newRange, new Map([['CIGNA', 1]]), []);
  chk('stored weeks grew to 61', led.weeks.size === 61, led.weeks.size);
  chk('window still 52', X.ledgerWindow(led,52).length === 52);
  chk('oldest week STILL stored', led.weeks.has(X.dateKey(ranges[0])));
  chk('newest week is in the window', X.ledgerWindow(led,52)[51].range === newRange);

  console.log('\n-- out-of-order week lands in date order, not append order --');
  const strayEnd = Date.UTC(2026, 0, 1);
  X.ledgerUpsertWeek(led, wkRange(strayEnd), new Map([['CIGNA', 7]]), []);
  const keys = X.ledgerSorted(led).map(w => w.key);
  chk('ledger stays date-sorted', keys.every((k,i) => i === 0 || keys[i-1] <= k));

  console.log('\n-- blank-week averaging matches the workbook rule --');
  const led3 = X.ledgerNew();
  X.ledgerUpsertWeek(led3, ranges[57], new Map([['CIGNA', 100]]), []);
  X.ledgerUpsertWeek(led3, ranges[58], new Map(), []);                      // CIGNA blank
  X.ledgerUpsertWeek(led3, ranges[59], new Map([['CIGNA', 200]]), []);
  const d3 = X.computeRollingDash(led3);
  chk('blank week excluded from the average (150, not 100)', near(d3.cats.CIGNA.avg, 150), d3.cats.CIGNA.avg);
  chk('blank week is 0 in the series', d3.cats.CIGNA.series[1] === 0);

  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
