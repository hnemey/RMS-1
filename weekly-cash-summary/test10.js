// Reproduces the reported window bug end-to-end in a real DOM.
const fs = require('fs'); const { JSDOM, VirtualConsole } = require('jsdom');
let html = fs.readFileSync('Weekly_Cash_Summary_ledger.html','utf8');
const EX = '<script>window.state=state;window.parseRolling=parseRolling;window.renderRolling=renderRolling;' +
  'window.computeRollingDash=computeRollingDash;window.ledgerUpsertWeek=ledgerUpsertWeek;' +
  'window.ledgerSorted=ledgerSorted;window.writeRollingXlsx=writeRollingXlsx;window.switchTab=null;<\/script>';
const at = html.lastIndexOf('</body>'); html = html.slice(0,at)+EX+html.slice(at);
const errs=[]; const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://localhost/',pretendToBeVisual:true,
  virtualConsole:new VirtualConsole().on('jsdomError',e=>errs.push(e.message))});
const win=dom.window, doc=win.document;
let fails=0; const chk=(n,c,e)=>{ if(c) console.log('  ok   '+n); else {fails++; console.log('  FAIL '+n+(e?' — '+e:''));} };
const fire=(el,t)=>el.dispatchEvent(new win.Event(t,{bubbles:true}));

setTimeout(async () => {
  win.DecompressionStream=DecompressionStream; win.CompressionStream=CompressionStream;
  win.Response=Response; win.Blob=Blob;
  const b=new Uint8Array(fs.readFileSync('Rolling_Ave_Ledger.xlsx'));
  const ab=new win.ArrayBuffer(b.length); new win.Uint8Array(ab).set(b);
  win.state.rolling = await win.parseRolling(ab);

  console.log('\n-- the v1 file says 51; that value came from a bug, so it is ignored --');
  chk('window resolves to 52', win.state.rolling.windowWeeks === 52, win.state.rolling.windowWeeks);

  // the week they saved this morning
  const all = win.ledgerSorted(win.state.rolling);
  win.ledgerUpsertWeek(win.state.rolling, '08/14/26-08/20/26', new Map(all[all.length-1].payors), []);

  console.log('\n-- with 08/14/26-08/20/26 saved on top --');
  let d = win.computeRollingDash();
  console.log('       ' + d.labels.length + ' weeks: ' + d.labels[0] + ' -> ' + d.labels[d.labels.length-1]);
  chk('52 weeks in the window', d.labels.length === 52, d.labels.length);
  chk('starts at 08/22/25-08/28/25', d.labels[0] === '08/22/25-08/28/25', d.labels[0]);
  chk('ends at 08/14/26-08/20/26', d.labels[d.labels.length-1] === '08/14/26-08/20/26', d.labels[d.labels.length-1]);
  chk('the oldest week is NOT dropped', d.labels.includes('08/22/25-08/28/25'));

  console.log('\n-- the week list is on screen and matches --');
  win.renderRolling();
  const rows = [...doc.querySelectorAll('#rollWeeks tbody tr')].map(tr => tr.children[1].textContent);
  console.log('       listed ' + rows.length + ' weeks, first "' + rows[0] + '", last "' + rows[rows.length-1] + '"');
  chk('list length matches the window', rows.length === d.labels.length);
  chk('list matches the window exactly', rows.join('|') === d.labels.join('|'));
  chk('window input shows 52', doc.getElementById('rollWin').value === '52', doc.getElementById('rollWin').value);
  const note = doc.getElementById('rollWeeks').textContent;
  chk('header states the count and span', /52 weeks/.test(note) && /08\/22\/25-08\/28\/25/.test(note));
  chk('states how many weeks are stored', /239 weeks stored/.test(note), note.slice(0,120));

  console.log('\n-- the control actually changes the number --');
  const a52 = win.computeRollingDash().avgTotal;
  const inp = doc.getElementById('rollWin');
  inp.value = '51'; fire(inp, 'change');
  const a51 = win.computeRollingDash().avgTotal;
  console.log('       52 weeks: $' + a52.toFixed(2) + '   51 weeks: $' + a51.toFixed(2) +
              '   delta $' + (a51-a52).toFixed(2));
  chk('51 gives a different (higher) figure', a51 > a52);
  chk('51 drops the oldest week', !win.computeRollingDash().labels.includes('08/22/25-08/28/25'));
  doc.getElementById('btnRollWinReset').click();
  chk('Reset to 52 restores it', win.computeRollingDash().labels.length === 52);
  chk('  …and the oldest week is back', win.computeRollingDash().labels[0] === '08/22/25-08/28/25');

  console.log('\n-- Save writes the window back as v2 --');
  const out = await win.writeRollingXlsx(win.state.rolling);
  const buf = out.buffer.slice(out.byteOffset, out.byteOffset+out.byteLength);
  const back = await win.parseRolling(buf);
  chk('re-read keeps 52', back.windowWeeks === 52, back.windowWeeks);
  chk('no page errors', errs.length === 0, errs.slice(0,2).join(' | '));
  console.log(fails ? '\n'+fails+' FAILURE(S)' : '\nALL PASS');
  process.exit(fails?1:0);
}, 400);
