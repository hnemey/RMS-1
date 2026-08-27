// Real-DOM test of the adjustments editor: renders the app's own markup in jsdom
// and drives it through the actual input/change handlers.
const fs = require('fs');
const { JSDOM } = require('jsdom');
let fails = 0;
const chk = (n,c,e) => { if (c) console.log('  ok   '+n); else { fails++; console.log('  FAIL '+n+(e?' — '+e:'')); } };

let html = fs.readFileSync('Weekly_Cash_Summary_ledger.html', 'utf8');
// top-level const/let in a classic script are script-scoped, not window props —
// but a later classic script shares that lexical environment, so this reaches them
const EXPORTS = '<script>window.__ok=1;window.state=state;' +
  'window.renderAdjEditor=renderAdjEditor;window.ledgerNew=ledgerNew;' +
  'window.ledgerUpsertWeek=ledgerUpsertWeek;window.setAdjEditorOpen=v=>{adjEditorOpen=v};<\/script>';
// the app's own source contains "</body>" inside JS string literals, so anchor on the last one
const bodyAt = html.lastIndexOf('</body>');
html = html.slice(0, bodyAt) + EXPORTS + html.slice(bodyAt);
const errs = [];
const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'https://localhost/', pretendToBeVisual: true,
  virtualConsole: new (require('jsdom').VirtualConsole)().on('jsdomError', e => errs.push(e.message)) });
const win = dom.window, doc = win.document;

const D = (y,m,d) => new win.Date(win.Date.UTC(y,m-1,d));
const fire = (el, type) => el.dispatchEvent(new win.Event(type, { bubbles: true }));

setTimeout(() => {
  const S = win.state;
  chk('app booted in jsdom', !!win.__ok && !!S && typeof win.renderAdjEditor === 'function');
  if (errs.length) { console.log('  page errors:'); errs.forEach(e => console.log('   ', String(e).split('\n')[0])); }
  if (!S) { process.exit(1); }

  // a loaded report for 07/03/26-07/09/26 and a ledger holding one earlier week
  S.rows = [{ _id:'r1', BankDetailSID:1, LockBoxDate: D(2026,7,6), Payer:'NORIDIAN AK-ID',
    PayorName:'MEDICARE', Comments:'NORIDIAN AK-ID', DetailLineItemDepositAmount:250000,
    BatchAmount:250000, BankBatchNo:'B1', BankSID:1 }];
  const ledger = win.ledgerNew();
  win.ledgerUpsertWeek(ledger, '04/17/26-04/23/26', new Map([['NORIDIANAKID', 100000]]), []);
  S.rolling = ledger;

  S.adjustments = [
    { id:'g', desc:'FY25 Cost Report settlement', cat:'MEDICARE', amt:'50000', week:'07/03/26-07/09/26' },
    { id:'z', desc:'Zero amount',   cat:'MEDICARE', amt:'00',  week:'07/03/26-07/09/26' },
    { id:'n', desc:'',              cat:'MEDICARE', amt:'500', week:'07/03/26-07/09/26' },
    { id:'p', desc:'Prior UPL',     cat:'MEDICARE', amt:'900', week:'04/17/26-04/23/26' },
    { id:'x', desc:'Nowhere week',  cat:'MEDICARE', amt:'700', week:'01/02/99-01/08/99' },
  ];
  win.setAdjEditorOpen(true);
  win.renderAdjEditor();

  const rowOf = id => doc.querySelector('#adjRows tr[data-id="' + id + '"]');

  console.log('\n-- the editor stays quiet: no status line, no tinting --');
  chk('no status line is rendered anywhere', doc.querySelectorAll('#adjRows .adjmsg').length === 0);
  chk('no row is tinted', doc.querySelectorAll('#adjRows tr[class]').length === 0,
      [...doc.querySelectorAll('#adjRows tr[class]')].map(t=>t.className).join(','));
  chk('a blank new row is not flagged', rowOf('n').className === '', rowOf('n').className);
  chk('the description cell holds only the input',
      rowOf('z').children[0].children.length === 1 &&
      rowOf('z').children[0].children[0].className === 'desc');

  console.log('\n-- amount normalisation on commit --');
  const amtIn = rowOf('z').querySelector('.amt');
  amtIn.value = '00';
  fire(amtIn, 'change');
  chk('"00" -> empty in the field', amtIn.value === '', JSON.stringify(amtIn.value));
  chk('"00" -> empty in state', S.adjustments.find(a=>a.id==='z').amt === '', JSON.stringify(S.adjustments.find(a=>a.id==='z').amt));

  amtIn.value = '007.500';
  fire(amtIn, 'change');
  chk('"007.500" -> "7.5"', amtIn.value === '7.5', amtIn.value);

  amtIn.value = '1234.567';
  fire(amtIn, 'change');
  chk('rounds to cents', amtIn.value === '1234.57', amtIn.value);

  console.log('\n-- normalised on commit, never mid-typing --');
  const mid = rowOf('g').querySelector('.amt');
  mid.value = '0050';
  fire(mid, 'input');
  chk('input event leaves the field alone', mid.value === '0050', mid.value);
  chk('  …and state tracks it verbatim', S.adjustments.find(a=>a.id==='g').amt === '0050');
  fire(mid, 'change');
  chk('change event normalises to "50"', mid.value === '50', mid.value);
  chk('  …and state follows', S.adjustments.find(a=>a.id==='g').amt === '50');

  console.log('\n-- editing does not re-render the row (focus is kept) --');
  const before = rowOf('z');
  const descIn = rowOf('n').querySelector('.desc');
  descIn.value = 'Now described';
  fire(descIn, 'input');
  chk('row element was not replaced', rowOf('z') === before);
  chk('state took the edit', S.adjustments.find(a=>a.id==='n').desc === 'Now described');
  chk('still no status line after editing', doc.querySelectorAll('#adjRows .adjmsg').length === 0);

  console.log('\n-- the note under the table no longer over-promises --');
  const note = doc.querySelector('#adjEditor .note').textContent;
  chk('claims about "any week" are qualified', /as long as the ledger already holds it/.test(note));
  chk('no longer says it does not have to be in the report, unqualified',
      !/it doesn.t have to be a week in the current report\./.test(note));

  console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nALL PASS');
  process.exit(fails ? 1 : 0);
}, 300);
