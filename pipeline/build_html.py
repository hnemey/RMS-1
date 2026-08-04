#!/usr/bin/env python3
import io, sys, os

SRC='/root/.claude/uploads/320d79b8-a16b-541b-b4c2-0a388ad47e5f/1b86e179-030b8f98Customer_Service_Report_Builder_10_1.html'
OUT='/home/user/RMS-1/Customer_Service_Report_Builder.html'
SCR=os.path.dirname(os.path.abspath(__file__))

html=open(SRC,encoding='utf-8').read()
pipeline=open(os.path.join(SCR,'pipeline.js'),encoding='utf-8').read()
stage_a=open(os.path.join(SCR,'stage_a.js'),encoding='utf-8').read()

def repl(old,new,count=1):
    global html
    n=html.count(old)
    if n<count:
        raise SystemExit(f'PATTERN NOT FOUND ({n}x): {old[:80]!r}')
    html=html.replace(old,new,count)

# ---------------------------------------------------------------- 1. buttons
repl(
'  <button id="btnPPTX" class="act" disabled>⬇ Export PPTX</button>',
'  <button id="btnPPTX" class="act" disabled>⬇ Export PPTX</button>\n'
'  <button id="btnWB" class="act" disabled title="Download the 12-month workbook with the new month added">⬇ Download 12-Month Data</button>')

# ---------------------------------------------------------------- 2. drop zone text
repl('    <div class="big">Drop the report workbook here, or click to choose</div>',
     '    <div class="big">Drop your files here, or click to choose</div>')
repl(
'    <div class="small">Accepts the monthly master workbook (.xlsm / .xlsx) — e.g. <b>2026_Full_Report_MASTER_COPY.xlsm</b></div>',
'    <div class="small">Drop <b>two files</b>: the <b>12-month master workbook</b> (.xlsm / .xlsx) <u>and</u> this month’s <b>raw-data file</b> (.xlsx). '
'The new month is calculated automatically — the dashboard rebuilds and you can download the updated 12-month workbook.<br>'
'You can still drop just the master on its own to view the existing dashboard.</div>')

# allow multiple files
repl('  <input type="file" id="file" accept=".xlsm,.xlsx,.xls">',
     '  <input type="file" id="file" accept=".xlsm,.xlsx,.xls" multiple>')

# status panel + styles right after the file input
repl('  <input type="file" id="file" accept=".xlsm,.xlsx,.xls" multiple>',
'''  <input type="file" id="file" accept=".xlsm,.xlsx,.xls" multiple>
  <div id="pipeStatus" style="display:none;margin:8px 2px 4px;border:1px solid #dfe3ea;border-radius:10px;background:#fff;padding:12px 14px;box-shadow:0 1px 6px rgba(20,30,60,.06);font-size:13px;color:#31424f"></div>''')

# ---------------------------------------------------------------- 3. inject pipeline libs before the app script
marker='   Customer Service Monthly Report Builder'
pos=html.index(marker)
si=html.rindex('<script>',0,pos)
inject=('<script>/* ==== in-browser data pipeline (raw -> Data Tables) ==== */\n'
        + pipeline + '\n</script>\n'
        + '<script>/* ==== Stage-A ingestion (monthly raw exports -> master format) ==== */\n'
        + stage_a + '\n</script>\n')
html=html[:si]+inject+html[si:]

# ---------------------------------------------------------------- 4. refactor loadWorkbookBytes -> classify + reveal
old_load='''function loadWorkbookBytes(buf, name){
  FILE_BYTES=buf; FILE_NAME=name||"workbook.xlsx";
  var wb=XLSX.read(buf,{type:"array",cellDates:true});
  parseWorkbook(wb);
  EDITED={}; ROWNOTES={}; CELLEDITS={}; GENNOTES={};
  document.getElementById("empty").style.display="none";
  document.getElementById("topnav").style.display="flex";
  document.getElementById("btnHTML").disabled=false;
  document.getElementById("btnPPTX").disabled=false;
  document.getElementById("monthBadge").textContent=REPORT_MONTH;
  document.getElementById("monthBadge").style.display="inline-block";
  buildPages();
  showView("report");
  goToPage(0);
  reflectLink();
  setStatus("Loaded "+FILE_NAME+" — "+ACTIVE.length+" month(s) through "+REPORT_MONTH+(FILE_HANDLE?" · linked for save":"")+".");
}'''
new_load='''function revealWorkbook(wb){
  parseWorkbook(wb);
  EDITED={}; ROWNOTES={}; CELLEDITS={}; GENNOTES={};
  document.getElementById("empty").style.display="none";
  document.getElementById("topnav").style.display="flex";
  document.getElementById("btnHTML").disabled=false;
  document.getElementById("btnPPTX").disabled=false;
  document.getElementById("monthBadge").textContent=REPORT_MONTH;
  document.getElementById("monthBadge").style.display="inline-block";
  buildPages();
  showView("report");
  goToPage(0);
  reflectLink();
}
/* ---- two-file data pipeline state ---- */
var BASE_BYTES=null, BASE_NAME="", RAW_WB=null, RAW_NAME="", WORKING_WB=null, PIPE_MONTH="";
function loadWorkbookBytes(buf, name){ classifyAndStore(buf, name); }
function classifyAndStore(buf, name){
  var wb;
  try{ wb=XLSX.read(buf,{type:"array",cellDates:true}); }
  catch(err){ setStatus("Could not read "+name+": "+err.message); return; }
  if(wb.Sheets && wb.Sheets["Data Tables"]){ BASE_BYTES=buf; BASE_NAME=name; FILE_BYTES=buf; FILE_NAME=name; }
  else { RAW_WB=wb; RAW_NAME=name; }
  processPipeline();
}
var MONTH_NAMES=["January","February","March","April","May","June","July","August","September","October","November","December"];
function detectReportMonth(rawWb){
  var H=ReportPipeline.helpers, counts={};
  function scan(sheet,col,startRow){
    var ws=rawWb.Sheets[sheet]; if(!ws) return;
    var last=H.lastRowOfCol(ws,col), cap=Math.min(last,startRow+6000);
    for(var r=startRow;r<=cap;r++){ var k=H.monthKey(H.getCell(ws,r,col)); if(k) counts[k]=(counts[k]||0)+1; }
  }
  scan("IVR Raw Data",8,2); scan("Hospital POS Survey Raw Data",3,2); scan("Correspondence Survey Raw Data",3,2);
  scan("CMS Raw Flow Out Data",1,2);
  var best=null,bn=-1; for(var k in counts){ if(counts[k]>bn){bn=counts[k];best=k;} }
  return best;
}
function processPipeline(){
  if(BASE_BYTES && RAW_WB){
    setStatus("Building the report — processing raw data …");
    try{
      var wb=XLSX.read(BASE_BYTES,{type:"array",cellDates:true});   // fresh copy so re-runs are clean
      var tgt=detectReportMonth(RAW_WB);
      if(!tgt){ setStatus("Could not find any dated rows in "+RAW_NAME+"."); return; }
      var p=tgt.split("-"), md=new Date(+p[0], +p[1]-1, 1);
      var repA=StageA.ingest(ReportPipeline.helpers, wb, RAW_WB, md);
      var repB=ReportPipeline.runAll(wb, tgt);
      WORKING_WB=wb; PIPE_MONTH=tgt; FILE_HANDLE=null; FILE_BYTES=null; FILE_NAME=BASE_NAME;
      revealWorkbook(wb);
      document.getElementById("btnWB").disabled=false;
      renderPipeReport(repA,repB,tgt);
      setStatus("Built "+MONTH_NAMES[(+p[1])-1]+" "+p[0]+" — "+ACTIVE.length+" month(s) through "+REPORT_MONTH+".");
    }catch(err){ setStatus("Pipeline error: "+err.message); console.error(err); }
  } else if(BASE_BYTES){
    try{
      var wb2=XLSX.read(BASE_BYTES,{type:"array",cellDates:true});
      WORKING_WB=wb2; revealWorkbook(wb2);
      document.getElementById("btnWB").disabled=false;
      renderPipeReport(null,null,null);
      setStatus("Loaded "+BASE_NAME+" — "+ACTIVE.length+" month(s) through "+REPORT_MONTH+". Drop this month’s raw-data file to add a new month.");
    }catch(err){ setStatus("Error: "+err.message); }
  } else if(RAW_WB){
    setStatus("Raw file “"+RAW_NAME+"” loaded. Now drop the 12-month master workbook (contains the “Data Tables” sheet).");
    renderPipeReport(null,null,null);
  }
}
function downloadWorkbook(){
  if(!WORKING_WB){ return; }
  setStatus("Preparing the updated 12-month workbook …");
  try{
    var out=XLSX.write(WORKING_WB,{bookType:"xlsx",type:"array",cellDates:true,compression:true});
    var blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    var a=document.createElement("a"); a.href=URL.createObjectURL(blob);
    var base=(BASE_NAME||"12_Month_Report").replace(/\\.(xlsm|xlsx|xls)$/i,"");
    a.download=base+(PIPE_MONTH?("_through_"+PIPE_MONTH):"_updated")+".xlsx";
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },1500);
    setStatus("Downloaded the updated 12-month workbook.");
  }catch(err){ setStatus("Could not build workbook: "+err.message); console.error(err); }
}
function renderPipeReport(repA,repB,tgt){
  var el=document.getElementById("pipeStatus"); if(!el) return;
  el.style.display="block";
  var rows=[];
  rows.push('<div style="font-weight:700;color:#00428D;margin-bottom:6px">Data files</div>');
  rows.push('<div>'+(BASE_BYTES?'✅':'⬜')+' 12-month master: <b>'+(BASE_NAME||'not loaded yet')+'</b></div>');
  rows.push('<div>'+(RAW_WB?'✅':'⬜')+' Raw-data file: <b>'+(RAW_NAME||'not loaded yet')+'</b></div>');
  if(repA||repB){
    var mp=tgt?tgt.split("-"):null;
    rows.push('<div style="font-weight:700;color:#00428D;margin:10px 0 6px">New month added: '+(mp?(MONTH_NAMES[(+mp[1])-1]+" "+mp[0]):"")+'</div>');
    if(repA&&repA.domains){
      for(var k in repA.domains){ rows.push('<div style="color:#5c6785">• '+k+': '+repA.domains[k]+'</div>'); }
    }
    if(repA&&repA.collectionsMissing){
      rows.push('<div style="margin-top:8px;padding:8px 10px;border-left:4px solid #e8a13a;background:#fff7ea;border-radius:4px">'+
        '<b>Collections data is missing for this month.</b> The raw file’s “Rep Collection Data” tab is empty and the monthly goal/actual isn’t present, '+
        'so the collection charts will be blank for the new month. Add per-rep collections to that tab (Date, Employee Name, Total Collections, Transaction Count, Name) '+
        'and the monthly goal/actual to “Monthly Collections Data”, then re-drop the files.</div>');
    }
    if(repA&&repA.notes&&repA.notes.length){
      rows.push('<div style="margin-top:6px;color:#7a8196;font-size:12px">Note: '+repA.notes.join(" ")+'</div>');
    }
    rows.push('<div style="margin-top:6px;color:#7a8196;font-size:12px">The downloadable workbook keeps every sheet and updates the “Data Tables” grid; '+
      'cell formatting and the “Data Graphs” charts inside the file are not preserved (the dashboard above replaces them).</div>');
  }
  el.innerHTML=rows.join("");
}'''
repl(old_load,new_load)

# ---------------------------------------------------------------- 5. handleFile stays (reads -> classifyAndStore via loadWorkbookBytes) - no change needed

# ---------------------------------------------------------------- 6. init wiring: multiple files + btnWB
repl('  input.addEventListener("change",function(){ if(input.files[0]) handleFile(input.files[0]); });',
     '  input.addEventListener("change",function(){ var fs=input.files; for(var i=0;i<fs.length;i++) handleFile(fs[i]); });')

old_drop='''  drop.addEventListener("drop",function(e){
    var items=e.dataTransfer.items;
    if(items&&items[0]&&items[0].getAsFileSystemHandle){
      items[0].getAsFileSystemHandle().then(function(h){ if(h&&h.kind==="file") loadFromHandle(h);
        else { var f=e.dataTransfer.files[0]; if(f) handleFile(f); } })
        .catch(function(){ var f=e.dataTransfer.files[0]; if(f) handleFile(f); });
    } else { var f=e.dataTransfer.files[0]; if(f) handleFile(f); }
  });'''
new_drop='''  drop.addEventListener("drop",function(e){
    var files=e.dataTransfer.files;
    if(files&&files.length){ for(var i=0;i<files.length;i++) handleFile(files[i]); }
  });'''
repl(old_drop,new_drop)

repl('  document.getElementById("btnPPTX").addEventListener("click",exportPPTX);',
     '  document.getElementById("btnPPTX").addEventListener("click",exportPPTX);\n'
     '  var bwb=document.getElementById("btnWB"); if(bwb) bwb.addEventListener("click",downloadWorkbook);')

# also make the file picker allow multiple + read each chosen file
repl('window.showOpenFilePicker({multiple:false,types:',
     'window.showOpenFilePicker({multiple:true,types:')
repl('      .then(function(hs){ return loadFromHandle(hs[0]); })',
     '      .then(function(hs){ hs.forEach(function(h){ h.getFile().then(function(f){ handleFile(f); }); }); })')

open(OUT,'w',encoding='utf-8').write(html)
print('WROTE',OUT, os.path.getsize(OUT),'bytes')
