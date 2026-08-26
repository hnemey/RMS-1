const fs = require('fs'); const { JSDOM, VirtualConsole } = require('jsdom');
let html = fs.readFileSync('Weekly_Cash_Summary_ledger.html','utf8');
const EX = '<script>window.state=state;window.parseRolling=parseRolling;window.buildLeaderReport=buildLeaderReport;<\/script>';
const at = html.lastIndexOf('</body>'); html = html.slice(0,at)+EX+html.slice(at);
const dom = new JSDOM(html,{runScripts:'dangerously',url:'https://localhost/',pretendToBeVisual:true,
  virtualConsole:new VirtualConsole()});
const win=dom.window;
setTimeout(async () => {
  win.DecompressionStream=DecompressionStream; win.CompressionStream=CompressionStream;
  win.Response=Response; win.Blob=Blob;
  const b=new Uint8Array(fs.readFileSync('Rolling_Ave_Ledger.xlsx'));
  const ab=new win.ArrayBuffer(b.length); new win.Uint8Array(ab).set(b);
  win.state.rolling = await win.parseRolling(ab);
  fs.writeFileSync('Leader_Report_FiscalYTD.html', win.buildLeaderReport());
  console.log('wrote Leader_Report_FiscalYTD.html');
  process.exit(0);
}, 400);
