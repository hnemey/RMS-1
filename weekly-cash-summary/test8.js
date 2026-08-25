// Loads the REAL migrated ledger into the full app and renders the rolling tab.
const fs = require('fs'); const { JSDOM, VirtualConsole } = require('jsdom');
let html = fs.readFileSync('Weekly_Cash_Summary_ledger.html','utf8');
const EXPORTS = '<script>window.__ok=1;window.state=state;window.parseRolling=parseRolling;' +
  'window.renderRolling=renderRolling;window.computeRollingDash=computeRollingDash;' +
  'window.rollingSpecialItems=rollingSpecialItems;window.ledgerSorted=ledgerSorted;<\/script>';
const at = html.lastIndexOf('</body>'); html = html.slice(0,at) + EXPORTS + html.slice(at);
const errs = [];
const dom = new JSDOM(html, { runScripts:'dangerously', url:'https://localhost/', pretendToBeVisual:true,
  virtualConsole: new VirtualConsole().on('jsdomError', e => errs.push(e.message)) });
const win = dom.window, doc = win.document;
let fails = 0;
const chk = (n,c,e) => { if (c) console.log('  ok   '+n); else { fails++; console.log('  FAIL '+n+(e?' — '+e:'')); } };

setTimeout(async () => {
  // jsdom has no working DecompressionStream/Blob.stream pipeline; the app runs in
  // Chromium where these are native, so lend the page Node's real implementations
  win.DecompressionStream = DecompressionStream;
  win.CompressionStream = CompressionStream;
  win.Response = Response;
  win.Blob = Blob;
  const bytes = new Uint8Array(fs.readFileSync('Rolling_Ave_Ledger.xlsx'));
  const ab = new win.ArrayBuffer(bytes.length);
  new win.Uint8Array(ab).set(bytes);
  const t0 = Date.now();
  win.state.rolling = await win.parseRolling(ab);
  const tParse = Date.now() - t0;
  chk('real ledger parsed in the browser context', win.state.rolling.weeks.size === 238, win.state.rolling.weeks.size);
  console.log('       parse time: ' + tParse + 'ms · ' + errs.length + ' page errors');

  const t1 = Date.now(); win.renderRolling(); const tRender = Date.now() - t1;
  const tbl = doc.getElementById('rollTable').innerHTML;
  const kpis = doc.getElementById('rollKpis').innerHTML;
  const trend = doc.getElementById('rollTrend').innerHTML;
  chk('rolling table rendered', tbl.length > 500 && /<table/.test(tbl));
  chk('KPIs rendered', kpis.length > 100);
  chk('trend chart rendered', /<svg/.test(trend));
  chk('empty-state hidden', doc.getElementById('rollEmpty').classList.contains('hidden'));
  console.log('       render time: ' + tRender + 'ms');

  const d = win.computeRollingDash();
  chk('window is 51 weeks', d.labels.length === 51, d.labels.length);
  chk('last week is 08/07/26-08/13/26', d.labels[50] === '08/07/26-08/13/26', d.labels[50]);
  chk('avgTotal is the expected figure', Math.abs(d.avgTotal - 69079529.61) < 0.01, d.avgTotal);
  const cats = Object.keys(d.cats).sort();
  console.log('       categories: ' + cats.length + ' — ' + cats.join(', '));
  chk('special items visible', win.rollingSpecialItems().length === 27, win.rollingSpecialItems().length);
  if (errs.length) { console.log('  page errors:'); errs.slice(0,5).forEach(e => console.log('   ', String(e).split('\n')[0])); }
  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
}, 400);
