const X = require('./harness.js'), fs = require('fs');
const buf = u8 => u8.buffer.slice(u8.byteOffset, u8.byteOffset+u8.byteLength);
(async () => {
  const a = await X.parseRolling(buf(new Uint8Array(fs.readFileSync('source.xlsx'))));
  const b = await X.parseRolling(buf(new Uint8Array(fs.readFileSync('Rolling_Ave_Ledger.xlsx'))));
  X.state.rolling = b;
  const da = X.computeRollingDash(a), db = X.computeRollingDash(b);
  const near=(x,y)=>Math.abs(x-y)<0.005; let bad=0;
  if (a.weeks.size !== b.weeks.size) { bad++; console.log('week count differs', a.weeks.size, b.weeks.size); }
  for (const [k,w] of a.weeks) {
    const w2 = b.weeks.get(k);
    if (!w2) { bad++; console.log('missing week', k); continue; }
    if (w.payors.size !== w2.payors.size) { bad++; console.log('payor count', k); }
    for (const [p,v] of w.payors) if (!near(v, w2.payors.get(p))) { bad++; console.log('payor', k, p, v, w2.payors.get(p)); }
    if (JSON.stringify(w.items) !== JSON.stringify(w2.items)) { bad++; console.log('items differ', k); }
  }
  if (!near(da.avgTotal, db.avgTotal)) { bad++; console.log('avgTotal', da.avgTotal, db.avgTotal); }
  console.log(bad ? bad + ' DIFFERENCES' : 'Re-read of the written file is byte-for-byte equivalent in content.');
  console.log('re-read: ' + b.weeks.size + ' weeks · legacy flag = ' + b.legacy + ' · window = ' + b.windowWeeks);
  process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1);});
