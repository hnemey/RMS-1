/* ============================================================================
   Customer Service Report — in-browser data pipeline
   Ports the Excel VBA macros (raw sheets -> "Data Tables") to JavaScript.
   Works in Node (validation) and the browser (integration) via SheetJS (XLSX).

   Design:
   - Operates directly on a SheetJS worksheet object using ABSOLUTE cell
     addressing (e.g. sheet['A307']), which aligns 1:1 with the VBA row/col
     constants. This lets each macro be ported almost line-for-line.
   - All month matching is keyed by "YYYY-MM" (every report date is the 1st of
     a month), which is equivalent to the VBA's CLng(date) / "mmmm - yy" keys
     but far more robust than serial-number arithmetic across engines.
   ========================================================================== */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.ReportPipeline = factory();
})(typeof self !== "undefined" ? self : this, function () {
"use strict";

/* ---- SheetJS handle: injected in browser, required() in Node ---- */
var XLSX = (typeof window !== "undefined" && window.XLSX) ? window.XLSX
         : (typeof require === "function" ? require("xlsx") : null);

/* ---------------- low-level cell helpers ---------------- */
function colLetter(c){ // 1-based -> letters
  var s=""; while(c>0){ var m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=(c-m-1)/26; } return s;
}
function addr(r,c){ return colLetter(c)+r; }

function getCell(ws,r,c){
  var cell=ws[addr(r,c)];
  return cell?cell.v:null;
}
function setCell(ws,r,c,v){
  var a=addr(r,c);
  if(v===null||v===undefined||v===""){ delete ws[a]; return; }
  var t = (typeof v==="number")?"n":(v instanceof Date?"d":"s");
  ws[a]={t:t,v:v};
  bumpRange(ws,r,c);
}
function bumpRange(ws,r,c){
  var ref=ws["!ref"]||("A1:"+addr(r,c));
  var range=XLSX.utils.decode_range(ref);
  if(r-1>range.e.r) range.e.r=r-1;
  if(c-1>range.e.c) range.e.c=c-1;
  if(r-1<range.s.r) range.s.r=r-1;
  if(c-1<range.s.c) range.s.c=c-1;
  ws["!ref"]=XLSX.utils.encode_range(range);
}

/* ---------------- value coercion ---------------- */
function isDateVal(v){ return v instanceof Date && !isNaN(v); }
function toDate(v){
  if(isDateVal(v)) return v;
  if(typeof v==="number"){ var d=XLSX.SSF? XLSX.SSF.parse_date_code(v):null; if(d) return new Date(Date.UTC(d.y,d.m-1,d.d)); }
  if(typeof v==="string"){ var t=Date.parse(v); if(!isNaN(t)) return new Date(t); }
  return null;
}
function monthKey(v){ // -> "YYYY-MM" or null
  var d=toDate(v); if(!d) return null;
  var y=d.getUTCFullYear? d.getUTCFullYear():d.getFullYear();
  var m=d.getUTCMonth? d.getUTCMonth():d.getMonth();
  // SheetJS cellDates gives local-ish Dates; use whichever is consistent.
  y=d.getFullYear(); m=d.getMonth();
  return y+"-"+String(m+1).padStart(2,"0");
}
function num(v){
  if(v===null||v===undefined||v==="") return null;
  if(typeof v==="number") return isFinite(v)?v:null;
  var s=String(v).replace(/[$,%\s]/g,"");
  if(s===""||s==="-") return null;
  var f=parseFloat(s); return isNaN(f)?null:f;
}
function isNumericVal(v){ return num(v)!==null && !(typeof v==="string" && !/^[-+]?[\d.,$%\s]+$/.test(v)); }

/* normalize a person/skill name for matching (mirrors VBA LCase/Trim, strips NBSP) */
function normName(s){ return String(s==null?"":s).replace(/ /g,"").replace(/\s+/g," ").trim(); }
function lc(s){ return normName(s).toLowerCase(); }
function uc(s){ return normName(s).toUpperCase(); }
/* VBA WorksheetFunction.Proper */
function proper(s){ return String(s==null?"":s).toLowerCase().replace(/\b\w/g,function(ch){return ch.toUpperCase();}); }

/* ---------------- Data Tables lookups ---------------- */
// month-col map for a header row: "YYYY-MM" -> column number
function dateColMap(ws,headerRow){
  var map={}, last=lastColOfRow(ws,headerRow);
  for(var c=2;c<=last;c++){
    var k=monthKey(getCell(ws,headerRow,c));
    if(k) map[k]=c;
  }
  return map;
}
// name-row map for a range (col A): normalizedKey -> row
function nameRowMap(ws,startRow,endRow,keyfn){
  keyfn=keyfn||lc;
  var map={};
  for(var r=startRow;r<=endRow;r++){
    var v=getCell(ws,r,1);
    var k=keyfn(v);
    if(k!=="") map[k]=r;
  }
  return map;
}
function lastColOfRow(ws,row){
  if(!ws||!ws["!ref"]) return 0;
  var range=XLSX.utils.decode_range(ws["!ref"]);
  var last=1;
  for(var c=1;c<=range.e.c+1;c++){ if(getCell(ws,row,c)!=null) last=c; }
  return last;
}
function lastRowOfCol(ws,col){
  if(!ws||!ws["!ref"]) return 0;
  var range=XLSX.utils.decode_range(ws["!ref"]);
  var last=0;
  for(var r=1;r<=range.e.r+1;r++){ if(getCell(ws,r,col)!=null) last=r; }
  return last;
}

/* =====================================================================
   DOMAIN 1 — EPIC Workqueues + Productivity  (VBA: UpdateWQTable)
   Raw: "EPIC WQ Raw Data" (A=Month,C=Name,F=EndCnt,H=EndAmt,I=AmtChng)
        "EPIC Productivity Data" (A=Month,B=Staff,E=UniqAcct,F=UniqGuar,J=Worked)
   ===================================================================== */
var WQ_ALIASES = lcMap({
  "Marilyn T Sprague, RRT":"Marilyn Sprague","Cat Hall":"Catherine Hall",
  "Karina Moreno Villafana":"Maria Karina Moreno Villafana","Kim Elliott":"Kimberly Elliott",
  "Vilma Arcos":"Magali Arcos"
});
function lcMap(o){ var m={}; for(var k in o) m[lc(k)]=o[k]; return m; }
function aliasName(map,name){ var k=lc(name); return map[k]!==undefined? map[k] : normName(name); }

function updateWQTable(wb, tgt){
  var wsData=wb.Sheets["EPIC WQ Raw Data"], wsProd=wb.Sheets["EPIC Productivity Data"], S=wb.Sheets["Data Tables"];
  var WQ_HEADER=6, PROD_HEADER=360;
  var T=[[8,52],[56,100],[104,148]];            // WQ tables: EndCnt(F) / EndAmt(H) / AmtChng(I)
  var P=[[361,396],[400,434],[438,472]];        // Prod tables: UniqAcct(F) / UniqGuar(E) / Worked(J)
  var TEAM_ROW=[151,154,157];                    // team totals: F,E,J
  var colWQ=dateColMap(S,WQ_HEADER)[tgt], colProd=dateColMap(S,PROD_HEADER)[tgt];
  var nm=T.map(function(t){return nameRowMap(S,t[0],t[1]);});
  var pm=P.map(function(t){return nameRowMap(S,t[0],t[1]);});

  if(colWQ){
    var lastR=lastRowOfCol(wsData,3);
    for(var i=2;i<=lastR;i++){
      if(monthKey(getCell(wsData,i,1))!==tgt) continue;
      var name=lc(aliasName(WQ_ALIASES,getCell(wsData,i,3)));
      writeIf(S,nm[0][name],colWQ,getCell(wsData,i,6));
      writeIf(S,nm[1][name],colWQ,getCell(wsData,i,8));
      writeIf(S,nm[2][name],colWQ,getCell(wsData,i,9));
    }
  }
  if(colProd){
    var lastP=lastRowOfCol(wsProd,2);
    for(var j=2;j<=lastP;j++){
      if(monthKey(getCell(wsProd,j,1))!==tgt) continue;
      var raw=normName(getCell(wsProd,j,2));
      if(lc(raw)==="team total"){
        writeIf(S,TEAM_ROW[0],colProd,getCell(wsProd,j,6));
        writeIf(S,TEAM_ROW[1],colProd,getCell(wsProd,j,5));
        writeIf(S,TEAM_ROW[2],colProd,getCell(wsProd,j,10));
      } else {
        var nmk=lc(aliasName(WQ_ALIASES,raw));
        writeIf(S,pm[0][nmk],colProd,getCell(wsProd,j,6));
        writeIf(S,pm[1][nmk],colProd,getCell(wsProd,j,5));
        writeIf(S,pm[2][nmk],colProd,getCell(wsProd,j,10));
      }
    }
  }
}
function writeIf(S,row,col,val){
  if(row>0 && col>0 && val!==null && val!==undefined && val!==""){ setCell(S,row,col,val); }
}

/* =====================================================================
   DOMAIN 2 — RevSpring IVR  (VBA: RemoveDuplicates + UpdateIVRTables)
   Operates on master "IVR Raw Data" with derived cols M..S already present:
     A=CallerID? actually A=Response ID, but VBA uses A as callerID key for
     "unique callers" — mirror VBA exactly: col A.
   Cols used: A(id/uniq), G(amount paid), H(result desc), L(call term),
     M(auth status), O/P/Q(transfer sums), S(month date).
   ===================================================================== */
function ivrRemoveDuplicates(ws){
  // keep one row per Response ID (col A); prefer a row whose auth status(M)=="Successful"
  var last=lastRowOfCol(ws,1);
  var keep={}; // id -> rowIndex chosen
  for(var i=2;i<=last;i++){
    var id=String(getCell(ws,i,1));
    var authM=getCell(ws,i,13);
    if(keep[id]===undefined){ keep[id]=i; }
    else if(authM==="Successful"){
      if(getCell(ws,keep[id],14)!=="Successful") keep[id]=i; // VBA checks col N of kept row
    }
  }
  // delete rows not kept (collect survivor rows, rebuild)
  var survivors={}; for(var id2 in keep){ survivors[keep[id2]]=true; }
  return survivors; // set of row numbers to include
}
function updateIVRTables(wb, tgt){
  var ws=wb.Sheets["IVR Raw Data"], S=wb.Sheets["Data Tables"];
  var HEADER=160, R={uniq:161,pay:164,appr:165,authS:170,authA:171,authN:172,
                     tStd:176,tSucc:177,tAtt:178,tNot:179};
  var survivors=ivrRemoveDuplicates(ws);
  var last=lastRowOfCol(ws,19); // col S
  var std={},sSucc={},sAtt={},sNot={},paid={},appr={},callers={},auth={};
  for(var i=2;i<=last;i++){
    if(!survivors[i]) continue;
    var rawDate=getCell(ws,i,19); // col S = Month Date
    var mk=monthKey(rawDate); if(mk!==tgt) continue;
    if(getCell(ws,i,12)==="Standard Transfer"){ std[mk]=(std[mk]||0)+1; } // col L
    var o=num(getCell(ws,i,15)); if(o!==null) sSucc[mk]=(sSucc[mk]||0)+o;
    var p=num(getCell(ws,i,16)); if(p!==null) sAtt[mk]=(sAtt[mk]||0)+p;
    var q=num(getCell(ws,i,17)); if(q!==null) sNot[mk]=(sNot[mk]||0)+q;
    var cid=getCell(ws,i,1); if(cid!=null && cid!==""){ (callers[mk]||(callers[mk]={}))[cid]=1; }
    var tp=num(getCell(ws,i,7)); if(tp!==null) paid[mk]=(paid[mk]||0)+tp; // col G
    if(getCell(ws,i,8)==="Approved") appr[mk]=(appr[mk]||0)+1;           // col H
    var a=getCell(ws,i,13); var am=(auth[mk]||(auth[mk]={})); var akey=String(a); am[akey]=(am[akey]||0)+1;
  }
  var dm=dateColMap(S,HEADER);
  if(dm[tgt]){
    var mk2=tgt, c=dm[tgt];
    setCell(S,R.tStd,c, std[mk2]!==undefined?std[mk2]:"");
    setCell(S,R.tSucc,c, sSucc[mk2]!==undefined?sSucc[mk2]:"");
    setCell(S,R.tAtt,c, sAtt[mk2]!==undefined?sAtt[mk2]:"");
    setCell(S,R.tNot,c, sNot[mk2]!==undefined?sNot[mk2]:"");
    setCell(S,R.uniq,c, callers[mk2]?Object.keys(callers[mk2]).length:"");
    setCell(S,R.pay,c, paid[mk2]!==undefined?paid[mk2]:"");
    setCell(S,R.appr,c, appr[mk2]!==undefined?appr[mk2]:"");
    var am2=auth[mk2];
    setCell(S,R.authS,c, am2&&am2["Successful"]!==undefined?am2["Successful"]:"");
    setCell(S,R.authA,c, am2&&am2["Attempted"]!==undefined?am2["Attempted"]:"");
    setCell(S,R.authN,c, am2&&am2["Not Attempted"]!==undefined?am2["Not Attempted"]:"");
  }
}

/* =====================================================================
   DOMAIN 3 — CMS Flow Out  (VBA: UpdateCMSFlowOutTable)
   Raw "CMS Flow Out": A=Month, C=Inbound Calls, M=Aban Calls
   -> Data Tables row 302 (Total), 303 (Abandoned)
   ===================================================================== */
function updateFlowOut(wb, tgt){
  var ws=wb.Sheets["CMS Flow Out"], S=wb.Sheets["Data Tables"];
  var HEADER=300, TOTAL=302, ABAN=303;
  var c=dateColMap(S,HEADER)[tgt]; if(!c) return;
  var last=lastRowOfCol(ws,1);
  for(var i=2;i<=last;i++){
    if(monthKey(getCell(ws,i,1))!==tgt) continue;
    setCell(S,TOTAL,c,getCell(ws,i,3));
    setCell(S,ABAN,c,getCell(ws,i,13));
  }
}

/* helper: clear a target-column cell unless it holds a formula */
function clearCell(S,r,c){ var a=addr(r,c); if(S[a]&&S[a].f) return; delete S[a]; }
function incCell(S,r,c){ if(r>0&&c>0){ var v=num(getCell(S,r,c))||0; setCell(S,r,c,v+1); } }
// find matching label row in [start,end] (case-insensitive trim); increment target col
function incMatch(S,val,col,start,end){
  var t=String(val==null?"":val).trim().toLowerCase();
  if(t==="") return false;
  for(var r=start;r<=end;r++){
    if(String(getCell(S,r,1)==null?"":getCell(S,r,1)).trim().toLowerCase()===t){ incCell(S,r,col); return true; }
  }
  return false;
}

/* =====================================================================
   DOMAIN 4 — Collections  (VBA: UpdateCollectionTables)
   "Monthly Collections Data": B=Goal,C=Actual,D=Date -> rows 308/309 (hdr 307)
   "Rep Collections Data": A=Date,C=TotalCollections,D=TxnCount,E=Name
       -> collection rows 738-772 (hdr 737, value C)
       -> txn rows 849-883 (hdr 848, value D)
   ===================================================================== */
var COLL_ALIASES = (function(){ var m={};
  [["ESTES, AMANDA","Amanda Estes"],["MANDY ESTES","Amanda Estes"],
   ["MORENO VILLAFANA, KARINA","Maria Karina Moreno Villafana"],
   ["KARINA MORENO","Maria Karina Moreno Villafana"],
   ["KARINA MORENO VILLAFANA","Maria Karina Moreno Villafana"],
   ["Kim Elliott","Kimberly Elliott"],["Vilma Arcos","Magali Arcos"],
   ["Cat Hall","Catherine Hall"]].forEach(function(p){ m[p[0].toLowerCase()]=p[1]; });
  return m; })();
// VBA NormalizeName: "Last, First" -> Proper("First Last"); else Proper(name)
function normalizeCollName(name){
  name=String(name==null?"":name).trim();
  if(name.indexOf(",")>=0){ var p=name.split(","); return proper((p[1]||"").trim()+" "+(p[0]||"").trim()); }
  return proper(name);
}
function updateCollections(wb, tgt){
  var wsData=wb.Sheets["Monthly Collections Data"], wsRep=wb.Sheets["Rep Collections Data"], S=wb.Sheets["Data Tables"];
  var MON_HDR=307,GOAL=308,ACTUAL=309, COL_HDR=737,COL_S=738,COL_E=772, TXN_HDR=848,TXN_S=849,TXN_E=883;
  // monthly
  var mcol=dateColMap(S,MON_HDR)[tgt];
  if(mcol){
    var lastM=lastRowOfCol(wsData,4);
    for(var i=2;i<=lastM;i++){
      if(monthKey(getCell(wsData,i,4))!==tgt) continue;
      writeIf(S,GOAL,mcol,getCell(wsData,i,2));
      writeIf(S,ACTUAL,mcol,getCell(wsData,i,3));
    }
  }
  // rep collections + txn
  var ccol=dateColMap(S,COL_HDR)[tgt], tcol=dateColMap(S,TXN_HDR)[tgt];
  var repRows=nameRowMap(S,COL_S,COL_E,function(v){return normalizeCollName(v);});
  var txnRows=nameRowMap(S,TXN_S,TXN_E,function(v){return normalizeCollName(v);});
  var lastR=lastRowOfCol(wsRep,1);
  for(var j=2;j<=lastR;j++){
    if(monthKey(getCell(wsRep,j,1))!==tgt) continue;
    var nm=getCell(wsRep,j,5); if(nm==null||String(nm).trim()==="") continue;
    var key=normalizeCollName(nm);
    if(COLL_ALIASES[key.toLowerCase()]!==undefined) key=normalizeCollName(COLL_ALIASES[key.toLowerCase()]);
    if(ccol) writeIf(S,repRows[key],ccol,getCell(wsRep,j,3));
    if(tcol) writeIf(S,txnRows[key],tcol,getCell(wsRep,j,4));
  }
}

/* =====================================================================
   DOMAIN 5 — Hospital POS  (VBA: UpdateHospitalPOSRepData)
   "Hospital Rep Data": A=MonthDate,D=Rep,E=Location,F=TimeSpent,G=Reason
   Data Tables (hdr 183): Location 185-192, TimeSpent 194-203,
     Reason 205-218 (+other 219); Rep counts hdr 704, rows 705-722.
   ===================================================================== */
var HOSP_ALIASES = { "karina moreno":"Maria Karina Moreno Villafana",
  "maria moreno villafana":"Maria Karina Moreno Villafana","m moreno":"Maria Karina Moreno Villafana",
  "vilma arcos":"Magali Arcos" };
function updateHospital(wb, tgt){
  var ws=wb.Sheets["Hospital Rep Data"], S=wb.Sheets["Data Tables"];
  var DHDR=183, LOC=[185,192], TS=[194,203], RE=[205,218], REOTHER=219, RHDR=704, REP=[705,722];
  var col=dateColMap(S,DHDR)[tgt], rcol=dateColMap(S,RHDR)[tgt];
  if(col){
    // clear this month's count cells
    for(var r=LOC[0];r<=LOC[1];r++) clearCell(S,r,col);
    for(r=TS[0];r<=TS[1];r++) clearCell(S,r,col);
    for(r=RE[0];r<=RE[1];r++) clearCell(S,r,col);
    clearCell(S,REOTHER,col);
    var last=lastRowOfCol(ws,1);
    for(var i=2;i<=last;i++){
      if(monthKey(getCell(ws,i,1))!==tgt) continue;
      incMatch(S,String(getCell(ws,i,5)||"").trim(),col,LOC[0],LOC[1]);
      incMatch(S,String(getCell(ws,i,6)||"").trim(),col,TS[0],TS[1]);
      var reason=String(getCell(ws,i,7)||"").trim();
      if(reason!==""){
        var seen={};
        reason.split(";").forEach(function(rt){ rt=rt.trim(); if(rt===""||seen[rt])return; seen[rt]=1;
          if(!incMatch(S,rt,col,RE[0],RE[1])) incCell(S,REOTHER,col); });
      }
    }
  }
  if(rcol){
    for(var rr=REP[0];rr<=REP[1];rr++) clearCell(S,rr,rcol);
    var valid={}; for(rr=REP[0];rr<=REP[1];rr++){ var b=String(getCell(S,rr,1)||"").trim(); if(b!=="") valid[b]=rr; }
    var counts={}, last2=lastRowOfCol(ws,1);
    for(var k=2;k<=last2;k++){
      if(monthKey(getCell(ws,k,1))!==tgt) continue;
      var rn=String(getCell(ws,k,4)||"").trim(); if(rn==="") continue;
      if(HOSP_ALIASES[rn.toLowerCase()]!==undefined) rn=HOSP_ALIASES[rn.toLowerCase()];
      if(valid[rn]!==undefined) counts[rn]=(counts[rn]||0)+1;
    }
    for(var nm in counts) setCell(S,valid[nm],rcol,counts[nm]);
  }
}

/* =====================================================================
   DOMAIN 6 — MyChart In-Basket & Email  (VBA: UpdateInboxTables)
   "Message Data": A=date,D=rep,E=platform,F=sender,G=agency,H=status,I=reason
   Data Tables (hdr 222): platform 224-225, sender 226-227, agency 230-234,
     status 237-239, reason 242-256 (+other 257); rep rows 726-734.
   ===================================================================== */
function msgNorm(s){ return String(s==null?"":s).replace(/\u00a0/g,"").replace(/\s+/g," ").trim().toLowerCase(); }
var MSG_REP_ALIAS={ "maria moreno villafana":"maria karina moreno villafana","vilma arcos":"magali arcos" };
var MSG_SENDER_ALIAS={ "internal":"internal emails","external":"external emails" };
function buildNormDict(S,start,end){ var d={}; for(var r=start;r<=end;r++){ var k=msgNorm(getCell(S,r,1)); if(k!=="") d[k]=r; } return d; }
function updateMessages(wb, tgt){
  var ws=wb.Sheets["Message Data"], S=wb.Sheets["Data Tables"];
  var HDR=222, PLAT=[224,225],SEND=[226,227],AGE=[230,234],STAT=[237,239],REAS=[242,256],REOTHER=257,REP=[726,734];
  var c=dateColMap(S,HDR)[tgt]; if(!c) return;
  // clear this month's cells across all message table rows
  [PLAT,SEND,AGE,STAT,REAS,REP].forEach(function(rg){ for(var r=rg[0];r<=rg[1];r++) clearCell(S,r,c); });
  clearCell(S,REOTHER,c);
  var dP=buildNormDict(S,PLAT[0],PLAT[1]),dS=buildNormDict(S,SEND[0],SEND[1]),dA=buildNormDict(S,AGE[0],AGE[1]),
      dSt=buildNormDict(S,STAT[0],STAT[1]),dR=buildNormDict(S,REAS[0],REAS[1]),dRep=buildNormDict(S,REP[0],REP[1]);
  var last=lastRowOfCol(ws,1);
  for(var i=2;i<=last;i++){
    if(monthKey(getCell(ws,i,1))!==tgt) continue;
    var rep=msgNorm(getCell(ws,i,4)); if(rep==="") continue;
    if(MSG_REP_ALIAS[rep]!==undefined) rep=MSG_REP_ALIAS[rep];
    var plat=msgNorm(getCell(ws,i,5));
    var send=msgNorm(getCell(ws,i,6)); if(MSG_SENDER_ALIAS[send]!==undefined) send=MSG_SENDER_ALIAS[send];
    var age=msgNorm(getCell(ws,i,7)), stat=msgNorm(getCell(ws,i,8));
    if(plat!==""&&dP[plat]!==undefined) incCell(S,dP[plat],c);
    if(send!==""&&dS[send]!==undefined) incCell(S,dS[send],c);
    if(age!==""&&dA[age]!==undefined) incCell(S,dA[age],c);
    if(stat!==""&&dSt[stat]!==undefined) incCell(S,dSt[stat],c);
    var reasonVal=String(getCell(ws,i,9)==null?"":getCell(ws,i,9));
    if(reasonVal.length>0){
      var set={};
      reasonVal.split(";").forEach(function(r){ var nr=msgNorm(r); if(nr!=="") set[nr]=1; });
      for(var nr in set){ if(dR[nr]!==undefined) incCell(S,dR[nr],c); else incCell(S,REOTHER,c); }
    }
    if(dRep[rep]!==undefined) incCell(S,dRep[rep],c);
  }
}

/* =====================================================================
   DOMAIN 7 — CMS phone stats  (VBA: UpdateCMSTables)
   Team "CMS Raw Data" (hdr row3, data row4+): A=Month,B=Skill, derived cols.
   Rep  "CMS Rep Raw Data" (hdr row2, data row3+): A=Month,B=Agent, derived cols.
   ===================================================================== */
var CMS_ALIASES={ "karina moreno":"Maria Karina Moreno Villafana",
  "moreno villafana, karina":"Maria Karina Moreno Villafana","sanford, jessi":"Jessica R Sanford",
  "forbus, gunnar":"Gunner Forbus","karen":"Karen Lyon","kim elliott":"Kimberly Elliott",
  "mandy estes":"AMANDA ESTES","ronni macdonald":"Veronica Petty","sevy b.":"Severeen Barela",
  "shelby holzworth":"Shelby Yates","krista plumm":"Krista Plummer","nicole anken":"Nicole Ankenman",
  "abigail garcia":"Abigail Garcia Amaral" };
function writeTeam(S,skill,col,val,start,end){
  if(val===""||val===null||val===undefined) return;
  var t=String(skill).trim().toUpperCase();
  for(var r=start;r<=end;r++){ if(String(getCell(S,r,1)||"").trim().toUpperCase()===t){ setCell(S,r,col,val); return; } }
}
function updateCMS(wb, tgt){
  var wsT=wb.Sheets["CMS Raw Data"], wsR=wb.Sheets["CMS Rep Raw Data"], S=wb.Sheets["Data Tables"];
  var TH=260, T1=[264,267],T2=[270,272],T3=[275,277],T4=[280,282],T5=[285,287],T6=[290,292],T7=[295,297];
  var RH=475, R1=[476,511],R2=[514,549],R3=[552,587],R4=[590,625],R5=[628,663],R6=[666,701];
  var tcol=dateColMap(S,TH)[tgt];
  if(tcol){
    var lastT=lastRowOfCol(wsT,2);
    for(var i=4;i<=lastT;i++){
      if(monthKey(getCell(wsT,i,1))!==tgt) continue;
      var skill=String(getCell(wsT,i,2)||"").trim(), U=skill.toUpperCase();
      if(U==="SLHS PFS PAT REP 217"){
        writeTeam(S,"SLHS PFS PAT REP 217",tcol,getCell(wsT,i,5),T1[0],T1[1]);
        writeTeam(S,"SLHS PFS PAT REP 217 ABANDONED CALLS",tcol,getCell(wsT,i,8),T1[0],T1[1]);
      } else if(U==="PFS CS ESCALATED 229"){
        writeTeam(S,"PFS CS ESCALATED 229",tcol,getCell(wsT,i,5),T1[0],T1[1]);
        writeTeam(S,"PFS CS ESCALATED 229 ABANDONED CALLS",tcol,getCell(wsT,i,8),T1[0],T1[1]);
      }
      writeTeam(S,skill,tcol,getCell(wsT,i,19),T2[0],T2[1]); // S
      writeTeam(S,skill,tcol,getCell(wsT,i,23),T3[0],T3[1]); // W
      writeTeam(S,skill,tcol,getCell(wsT,i,24),T4[0],T4[1]); // X
      writeTeam(S,skill,tcol,getCell(wsT,i,20),T5[0],T5[1]); // T
      writeTeam(S,skill,tcol,getCell(wsT,i,21),T6[0],T6[1]); // U
      writeTeam(S,skill,tcol,getCell(wsT,i,25),T7[0],T7[1]); // Y
    }
  }
  var rcol=dateColMap(S,RH)[tgt];
  if(rcol){
    var maps=[R1,R2,R3,R4,R5,R6].map(function(t){return nameRowMap(S,t[0],t[1],function(v){return String(v==null?"":v).trim().toUpperCase();});});
    var cols=[3,10,19,21,30,26]; // C,J,S,U,AD,Z
    var lastR=lastRowOfCol(wsR,2);
    for(var j=3;j<=lastR;j++){
      if(monthKey(getCell(wsR,j,1))!==tgt) continue;
      var rep=String(getCell(wsR,j,2)||"").trim();
      if(CMS_ALIASES[rep.toLowerCase()]!==undefined) rep=CMS_ALIASES[rep.toLowerCase()];
      // NB: VBA skips names containing "TOTAL"; we keep them so the "Totals" summary
      // row (present in every historical column) is populated from the raw Totals row.
      var key=rep.toUpperCase();
      for(var m=0;m<6;m++){ var row=maps[m][key]; if(row>0){ var v=getCell(wsR,j,cols[m]); if(v!==null&&v!==undefined&&v!=="") setCell(S,row,rcol,v); } }
    }
  }
}

/* =====================================================================
   FORMULA RECALC for a single column of "Data Tables".
   The workbook's formula cells (SUM/ratios/IFERROR/header refs) don't
   recalc in the browser, so we evaluate them in JS for the target month's
   column. INDEX/MATCH formulas (appendix rep layout, not charted) are left
   for Excel to recompute on open.
   ===================================================================== */
function evalArith(expr){
  // recursive-descent over + - * / and parentheses on a numeric string
  var s=expr, i=0;
  function peek(){ return s[i]; }
  function parseExpr(){ var v=parseTerm(); while(s[i]==='+'||s[i]==='-'){ var op=s[i++]; var r=parseTerm(); v=(op==='+')?v+r:v-r; } return v; }
  function parseTerm(){ var v=parseFactor(); while(s[i]==='*'||s[i]==='/'){ var op=s[i++]; var r=parseFactor(); v=(op==='*')?v*r:v/r; } return v; }
  function parseFactor(){
    while(s[i]===' ')i++;
    if(s[i]==='('){ i++; var v=parseExpr(); if(s[i]===')')i++; return v; }
    if(s[i]==='-'){ i++; return -parseFactor(); }
    var start=i; while(i<s.length && /[0-9.eE]/.test(s[i])) i++;
    var num=parseFloat(s.slice(start,i));
    return isNaN(num)?0:num;
  }
  var out=parseExpr();
  return out;
}
function refValue(S,ref,col){ // ref like "H164" / "H$6" / "$H$6" (col letters ignored; use target col)
  var m=ref.match(/\$?([A-Z]+)\$?(\d+)/); if(!m) return null;
  var r=parseInt(m[2],10);
  return getCell(S,r,col);
}
function evalFormulaColumn(S,f,col){
  f=f.replace(/^=/,"");
  if(/INDEX|MATCH/i.test(f)) return undefined;            // skip — not charted; Excel recalcs
  // IFERROR(expr, fallback)
  var ie=f.match(/^IFERROR\((.*),\s*("(?:[^"]*)"|[^,]*)\)\s*$/i);
  if(ie){
    var inner=ie[1], fb=ie[2].replace(/^"|"$/g,"");
    var v=evalFormulaColumn(S,inner,col);
    if(v===undefined) return undefined;
    if(v===null||typeof v==="number"&&(!isFinite(v)||isNaN(v))) return fb==="-"||fb===""?fb:fb;
    return v;
  }
  // pure single ref (e.g. =H$6 header date) -> return raw value
  if(/^\$?[A-Z]+\$?\d+$/.test(f.trim())) return refValue(S,f.trim(),col);
  // SUM(a:b)
  f=f.replace(/SUM\(\s*(\$?[A-Z]+\$?\d+)\s*:\s*(\$?[A-Z]+\$?\d+)\s*\)/gi,function(_,a,b){
    var ra=parseInt(a.match(/\d+/)[0],10), rb=parseInt(b.match(/\d+/)[0],10), sum=0;
    for(var r=Math.min(ra,rb);r<=Math.max(ra,rb);r++){ var n=num(getCell(S,r,col)); if(n!=null) sum+=n; }
    return "("+sum+")";
  });
  // replace remaining cell refs with numeric values (blank -> 0)
  f=f.replace(/\$?[A-Z]+\$?\d+/g,function(ref){ var n=num(refValue(S,ref,col)); return "("+(n==null?0:n)+")"; });
  if(/[^0-9.+\-*/() eE]/.test(f)) return undefined;        // anything unexpected -> skip
  var val=evalArith(f);
  return (typeof val==="number"&&isFinite(val))?val:undefined;
}
function recalcColumn(wb, tgt){
  var S=wb.Sheets["Data Tables"];
  var col=dateColMap(S,6)[tgt]; if(!col) return;
  // collect formula cells in this column
  var cells=[]; var range=XLSX.utils.decode_range(S["!ref"]);
  for(var r=1;r<=range.e.r+1;r++){ var a=addr(r,col); var c=S[a]; if(c&&c.f) cells.push({r:r,a:a,f:c.f}); }
  // iterate to a fixpoint (handles cross-row dependencies)
  for(var pass=0;pass<12;pass++){
    var changed=false;
    for(var k=0;k<cells.length;k++){
      var cell=cells[k]; var v=evalFormulaColumn(S,cell.f,col);
      if(v===undefined) continue;
      var cur=S[cell.a]; var prev=cur?cur.v:null;
      var nv=(v===""||v==="-")?v:v;
      if(prev!==nv){
        var t=(typeof nv==="number")?"n":(nv instanceof Date?"d":"s");
        S[cell.a]={t:t,v:nv,f:cell.f};    // keep the formula for Excel
        changed=true;
      }
    }
    if(!changed) break;
  }
}

/* =====================================================================
   Public API — run selected domains, return a status report
   ===================================================================== */
function runAll(wb, tgt, opts){
  opts=opts||{};
  var report={domains:{}, warnings:[]};
  var domains=[
    ["EPIC Workqueues", updateWQTable],
    ["RevSpring IVR",   updateIVRTables],
    ["Call Flow Out",   updateFlowOut],
    ["Collections",     updateCollections],
    ["POS Hospital",    updateHospital],
    ["MyChart & Email", updateMessages],
    ["CMS Phone Stats", updateCMS]
  ];
  domains.forEach(function(d){
    try{ d[1](wb,tgt); report.domains[d[0]]="ok"; }
    catch(e){ report.domains[d[0]]="error: "+e.message; report.warnings.push(d[0]+": "+e.message); }
  });
  try{ recalcColumn(wb,tgt); }catch(e){ report.warnings.push("recalc: "+e.message); }
  return report;
}

return {
  XLSXref:function(x){ XLSX=x; },
  helpers:{colLetter:colLetter,addr:addr,getCell:getCell,setCell:setCell,monthKey:monthKey,
           num:num,normName:normName,lc:lc,uc:uc,proper:proper,dateColMap:dateColMap,
           nameRowMap:nameRowMap,lastColOfRow:lastColOfRow,lastRowOfCol:lastRowOfCol,toDate:toDate,
           aliasName:aliasName,lcMap:lcMap},
  updateWQTable:updateWQTable, updateIVRTables:updateIVRTables, updateFlowOut:updateFlowOut,
  updateCollections:updateCollections, updateHospital:updateHospital, updateMessages:updateMessages,
  updateCMS:updateCMS, recalcColumn:recalcColumn, evalFormulaColumn:evalFormulaColumn, runAll:runAll
};
});
