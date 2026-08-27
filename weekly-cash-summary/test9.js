// Builds the leader report from the real ledger and checks the Fiscal YTD view.
const fs = require('fs'); const { JSDOM, VirtualConsole } = require('jsdom');
let html = fs.readFileSync('Weekly_Cash_Summary_ledger.html','utf8');
const EX = '<script>window.__ok=1;window.state=state;window.parseRolling=parseRolling;' +
  'window.buildLeaderReport=buildLeaderReport;window.rollingMonthsData=rollingMonthsData;' +
  'window.fiscalYearOf=fiscalYearOf;window.MONTH_ABBR=MONTH_ABBR;<\/script>';
const at = html.lastIndexOf('</body>'); html = html.slice(0,at)+EX+html.slice(at);
const errs=[]; const dom = new JSDOM(html,{runScripts:'dangerously',url:'https://localhost/',pretendToBeVisual:true,
  virtualConsole:new VirtualConsole().on('jsdomError',e=>errs.push(e.message))});
const win=dom.window;
let fails=0; const chk=(n,c,e)=>{ if(c) console.log('  ok   '+n); else {fails++; console.log('  FAIL '+n+(e?' — '+e:''));} };

setTimeout(async () => {
  win.DecompressionStream=DecompressionStream; win.CompressionStream=CompressionStream;
  win.Response=Response; win.Blob=Blob;
  const bytes=new Uint8Array(fs.readFileSync('Rolling_Ave_Ledger.xlsx'));
  const ab=new win.ArrayBuffer(bytes.length); new win.Uint8Array(ab).set(bytes);
  win.state.rolling = await win.parseRolling(ab);

  const rm = win.rollingMonthsData();
  console.log('\n-- fiscal-year window --');
  console.log('       months: ' + rm.months.map(m=>win.MONTH_ABBR[m.mo-1]+'-'+m.yr).join(', '));
  console.log('       fy=' + rm.fy + ' · nMonths=' + rm.nMonths + ' · label="' + rm.ytdLabel + '"');
  chk('starts in October', rm.months[0].mo === 10, rm.months[0].mo);
  chk('every month is in one fiscal year',
      rm.months.every(m => win.fiscalYearOf(m.yr,m.mo) === rm.fy));
  chk('no month from the prior FY leaked in', !rm.months.some(m => m.yr===2025 && m.mo<10));
  chk('label ends in YTD', /YTD$/.test(rm.ytdLabel), rm.ytdLabel);
  chk('label starts with October', /^October/.test(rm.ytdLabel), rm.ytdLabel);

  const rep = win.buildLeaderReport();
  const d2 = new JSDOM(rep).window.document;
  const tabs = [...d2.querySelectorAll('.rtabs button')].map(b=>b.textContent);
  console.log('\n-- report tabs --');
  console.log('       ' + tabs.join(' | '));
  chk('YTD tab is last', /YTD$/.test(tabs[tabs.length-1]), tabs[tabs.length-1]);
  chk('the report opens on the most recent month',
      d2.querySelector('.rtabs button.on').textContent === tabs[0],
      d2.querySelector('.rtabs button.on').textContent);
  chk('and that is the visible section', d2.querySelector('section.rsec.on').dataset.i === '0');
  chk('exactly one tab is active', d2.querySelectorAll('.rtabs button.on').length === 1);
  chk('month tabs run newest first', tabs[0] === "Aug ’26" && tabs[1] === "Jul ’26", tabs.slice(0,3).join(','));
  chk('oldest month tab is last before YTD', tabs[tabs.length-2] === "Oct ’25", tabs[tabs.length-2]);

  const ytdSec = d2.querySelector('section.rsec[data-i="' + (tabs.length-1) + '"]');
  const hdrs = [...ytdSec.querySelectorAll('table.wcs thead th')].map(t=>t.textContent);
  console.log('\n-- Fiscal YTD table header --');
  console.log('       ' + hdrs.join(' | '));
  chk('newest month column is leftmost', hdrs[1] === 'Aug-2026', hdrs[1]);
  chk('October is the last month column', hdrs[hdrs.length-3] === 'Oct-2025', hdrs[hdrs.length-3]);
  chk('YTD column sits after the months', hdrs[hdrs.length-2] === 'YTD', hdrs[hdrs.length-2]);
  chk('Mix is last', hdrs[hdrs.length-1] === 'Mix');

  // the YTD column must still equal the sum of the month cells on each row
  const rows = [...ytdSec.querySelectorAll('table.wcs tbody tr')];
  const money = t => Number(String(t).replace(/[^0-9.\-]/g,'')) || 0;
  let sumOk = 0, sumBad = 0;
  for (const tr of rows) {
    const tds = [...tr.querySelectorAll('td')];
    if (tds.length < 4) continue;
    const months = tds.slice(1, tds.length-2).map(t=>money(t.textContent));
    const ytd = money(tds[tds.length-2].textContent);
    const s = months.reduce((a,b)=>a+b,0);
    if (Math.abs(s-ytd) <= 1) sumOk++; else { sumBad++; console.log('   row mismatch:', tds[0].textContent, s, 'vs', ytd); }
  }
  chk('every row: months sum to the YTD column (' + sumOk + ' rows)', sumBad === 0);

  console.log('\n-- month-over-month shading still points at the real neighbour --');
  // Aug-2026 sits leftmost but its prior month is Jul-2026, the column to its RIGHT
  const augCell = [...ytdSec.querySelectorAll('table.wcs tbody tr')][0].querySelectorAll('td')[1];
  const julCell = [...ytdSec.querySelectorAll('table.wcs tbody tr')][0].querySelectorAll('td')[2];
  const priorFromTitle = t => { const m = /prior mo ([^\u00b7]+)\u00b7/.exec(t || ''); return m ? money(m[1]) : null; };
  const augPrior = priorFromTitle(augCell.getAttribute('title'));
  console.log('       Aug cell title:', (augCell.getAttribute('title')||'(none)').slice(0,64));
  console.log('       Jul cell value:', julCell.textContent);
  chk('leftmost column compares against the column to its right (not its left)',
      augPrior === null || Math.abs(augPrior - money(julCell.textContent)) <= 1,
      augPrior + ' vs ' + money(julCell.textContent));

  console.log('\n-- October on its own --');
  const led = win.state.rolling;
  // a week belongs to the month its Thursday END date falls in, so filter on that
  for (const [k, w] of [...led.weeks]) {
    const e = w.end;
    if (!(e && e.getUTCFullYear() === 2025 && e.getUTCMonth() === 9)) led.weeks.delete(k);
  }
  const octOnly = win.rollingMonthsData();
  console.log('       months:', octOnly.months.map(m=>win.MONTH_ABBR[m.mo-1]+'-'+m.yr).join(', '));
  chk('label is exactly "October YTD"', octOnly.ytdLabel === 'October YTD', octOnly.ytdLabel);
  chk('no dash, no year, no "last N months"', !/\u2013|20\d\d|last /.test(octOnly.ytdLabel), octOnly.ytdLabel);
  const rep2 = win.buildLeaderReport();
  const d3 = new JSDOM(rep2).window.document;
  const tabs2 = [...d3.querySelectorAll('.rtabs button')].map(b=>b.textContent);
  console.log('       tabs:', tabs2.join(' | '));
  chk('tabs are just the month then October YTD', tabs2.join('|') === "Oct \u201925|October YTD", tabs2.join('|'));
  chk('opens on the month, not the YTD tab', d3.querySelector('.rtabs button.on').textContent === "Oct \u201925");
  const h2 = [...d3.querySelector('section.rsec[data-i="' + (tabs2.length-1) + '"]').querySelectorAll('table.wcs thead th')].map(t=>t.textContent);
  console.log('       header:', h2.join(' | '));
  chk('single month column, then YTD', h2.join('|') === 'Payors|Oct-2025|YTD|Mix', h2.join('|'));

  chk('no page errors', errs.length === 0, errs.slice(0,2).join(' | '));
  console.log(fails ? '\n'+fails+' FAILURE(S)' : '\nALL PASS');
  process.exit(fails?1:0);
}, 400);
