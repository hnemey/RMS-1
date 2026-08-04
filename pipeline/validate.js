const XLSX=require('xlsx');
const P=require('./pipeline.js'); P.XLSXref(XLSX);
const H=P.helpers;
const MASTER='/root/.claude/uploads/320d79b8-a16b-541b-b4c2-0a388ad47e5f/d5d6ecd5-2026_Full_Report_MASTER_COPY.xlsm';
const TGT=process.env.TGT||'2026-06';   // recompute this month and compare its single column
function load(){ return XLSX.readFile(MASTER,{cellDates:true,cellNF:false,cellText:false}); }
function rangeRows(a,b){ const o=[]; for(let r=a;r<=b;r++)o.push(r); return o; }
function colFor(S,headerRow){ return H.dateColMap(S,headerRow)[TGT]; }

function isFormula(S,r,col){ const c=S[H.addr(r,col)]; return !!(c&&c.f); }
function snap(S,rows,col){ const m={}; rows.forEach(r=>m[r]={v:H.getCell(S,r,col),f:isFormula(S,r,col)}); return m; }
function clearCol(S,rows,col){ rows.forEach(r=>{ if(!isFormula(S,r,col)) H.setCell(S,r,col,''); }); }

function cmp(orig,S,rows,col,label){
  let match=0,blankBoth=0,diffs=[];
  rows.forEach(r=>{
    if(orig[r].f) return;               // skip formula cells (macro never writes them)
    const o=orig[r].v, n=H.getCell(S,r,col);
    const on=H.num(o),nn=H.num(n);
    const oB=(o===null||o===''||o===undefined), nB=(n===null||n===''||n===undefined);
    if(oB&&nB){blankBoth++;return;}
    let ok = (on!==null&&nn!==null) ? (Math.abs(on-nn)<1e-6||(Math.abs(on)>1&&Math.abs(on-nn)/Math.abs(on)<1e-9)) : String(o)===String(n);
    if(ok)match++; else diffs.push({cell:H.addr(r,col),lbl:H.getCell(S,r,1),o,n});
  });
  console.log(`\n=== ${label} (col ${H.colLetter(col)}, ${TGT}) ===  match:${match} MISMATCH:${diffs.length} blankBoth:${blankBoth}`);
  diffs.slice(0,20).forEach(d=>console.log(`   ✗ ${d.cell} [${d.lbl}] orig=${JSON.stringify(d.o)} got=${JSON.stringify(d.n)}`));
  if(diffs.length>20)console.log(`   ... +${diffs.length-20} more`);
  return diffs.length;
}

let tot=0;
// WQ
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const rW=[...rangeRows(8,52),...rangeRows(56,100),...rangeRows(104,148),151,154,157];
  const rP=[...rangeRows(361,396),...rangeRows(400,434),...rangeRows(438,472)];
  const cW=colFor(S,6),cP=colFor(S,360);
  const oW=snap(S,rW,cW),oP=snap(S,rP,cP); clearCol(S,rW,cW); clearCol(S,rP,cP);
  P.updateWQTable(wb,TGT);
  tot+=cmp(oW,S,rW,cW,'EPIC WQ'); tot+=cmp(oP,S,rP,cP,'EPIC Productivity');
}
// IVR
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const r=[161,164,165,170,171,172,176,177,178,179]; const c=colFor(S,160);
  const o=snap(S,r,c); clearCol(S,r,c); P.updateIVRTables(wb,TGT); tot+=cmp(o,S,r,c,'RevSpring IVR');
}
// FlowOut
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const r=[302,303]; const c=colFor(S,300);
  const o=snap(S,r,c); clearCol(S,r,c); P.updateFlowOut(wb,TGT); tot+=cmp(o,S,r,c,'Call Flow Out');
}
// Collections
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const r=[308,309,...rangeRows(738,772),...rangeRows(849,883)];
  // three header rows (307/737/848) share the same month cols; use 307
  const c=colFor(S,307);
  const o=snap(S,r,c); clearCol(S,r,c); P.updateCollections(wb,TGT); tot+=cmp(o,S,r,c,'Collections');
}
// Hospital
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const r=[...rangeRows(185,192),...rangeRows(194,203),...rangeRows(205,219),...rangeRows(705,722)];
  const c=colFor(S,183);
  const o=snap(S,r,c); clearCol(S,r,c); P.updateHospital(wb,TGT); tot+=cmp(o,S,r,c,'POS Hospital');
}
// Messages
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const r=[...rangeRows(224,227),...rangeRows(230,234),...rangeRows(237,239),...rangeRows(242,257),...rangeRows(726,734)];
  const c=colFor(S,222);
  const o=snap(S,r,c); clearCol(S,r,c); P.updateMessages(wb,TGT); tot+=cmp(o,S,r,c,'MyChart & Email');
}
// CMS
{ const wb=load(),S=wb.Sheets['Data Tables'];
  const rT=[...rangeRows(264,267),...rangeRows(270,272),...rangeRows(275,277),...rangeRows(280,282),...rangeRows(285,287),...rangeRows(290,292),...rangeRows(295,297)];
  const rR=[...rangeRows(476,511),...rangeRows(514,549),...rangeRows(552,587),...rangeRows(590,625),...rangeRows(628,663),...rangeRows(666,701)];
  const cT=colFor(S,260), cR=colFor(S,475);
  const oT=snap(S,rT,cT), oR=snap(S,rR,cR); clearCol(S,rT,cT); clearCol(S,rR,cR);
  P.updateCMS(wb,TGT); tot+=cmp(oT,S,rT,cT,'CMS Team'); tot+=cmp(oR,S,rR,cR,'CMS Rep');
}
console.log(`\n########  TOTAL MISMATCHES (${TGT}): ${tot}  ########`);
