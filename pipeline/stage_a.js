/* Stage A — transform monthly raw exports into the master's raw-sheet format.
   Includes the derived-column formulas (ported from the worksheet formulas).
   Exposed as pure helpers so they can be validated against the master. */
(function (root, factory){
  if(typeof module==="object"&&module.exports) module.exports=factory();
  else root.StageA=factory();
})(typeof self!=="undefined"?self:this, function(){
"use strict";

/* ---- time helpers ----
   The CMS raw time cells arrive either as Excel time-serials (SheetJS Date
   objects, e.g. Max Delay = "16:42" == serial 1.6958) or as clock strings
   (e.g. ":56"). The worksheet formula is  TEXT(clock,"[hh]:mm:ss") * 1/60,
   which numerically equals  excelSerial(clock) / 60. */
var EXCEL_EPOCH = Date.UTC(1899,11,30,0,0,0,0);
function excelSerial(v){
  if(v instanceof Date && !isNaN(v)) return (v.getTime()-EXCEL_EPOCH)/86400000;
  if(typeof v==="number") return v;                 // already a serial
  var s=String(v==null?"":v).trim();
  if(s==="") return null;
  // mirror the sheet's leading-zero cleanup for ":ss" / single-colon strings
  if(s.charAt(0)===":" || (s.split(":").length-1)===1) s="0"+s;
  var p=s.split(":");
  if(p.length<2){ var f=parseFloat(s); return isNaN(f)?null:f; }
  var h=parseInt(p[0],10)||0, m=parseInt(p[1],10)||0, sec=p.length>=3?(parseInt(p[2],10)||0):0;
  return (h*3600+m*60+sec)/86400;
}
// SECONDS represented by a clock value (used where seconds are wanted)
function clockToSeconds(v){ var s=excelSerial(v); return s==null?null:Math.round(s*86400); }
// derived sheet value: excelSerial / 60
function clockToSheetMinuteSerial(v){ var s=excelSerial(v); return s==null?"":s/60; }

/* =========================================================================
   IVR derived columns M..S  (from base master layout)
   base cols: C=Account ID, G=Amount Paid, I=Call Date, L=Call Term Location
   M=Auth Status, N=Auth&Payment, O=SuccTransfer, P=AttTransfer,
   Q=NotAttTransfer, R=TransferCount, S=Month Date
   ========================================================================= */
function ivrAuthStatus(accountID){
  if(accountID==null||String(accountID).trim()==="") return "Not Attempted";
  if(typeof accountID==="number"||/^\d+$/.test(String(accountID).trim())) return "Successful";
  if(/account/i.test(String(accountID))) return "Attempted";
  return "Unknown Status";
}
function ivrDerive(o){ // o={accountID,amountPaid,callDate,callTerm}
  var M=ivrAuthStatus(o.accountID);
  var Lterm=String(o.callTerm||"").toLowerCase();
  var isStd=(Lterm==="standard transfer");
  var out={M:M};
  out.N=(isNum(o.amountPaid)&&/successful/i.test(M))?1:"";
  out.O=(M.toLowerCase()==="successful"&&isStd)?1:"";
  out.P=(M.toLowerCase()==="attempted"&&isStd)?1:"";
  out.Q=(M.toLowerCase()==="not attempted"&&isStd)?1:"";
  out.R=isStd?1:"";
  out.S=firstOfMonth(o.callDate);
  return out;
}

/* =========================================================================
   CMS Team derived P..Z  (base cols per master "CMS Raw Data")
   C=AvgSpeedAns D=AvgAbanTime E=ACDCalls F=AvgACDTime G=AvgACWTime H=AbanCalls
   I=MaxDelay M=AvgExtnOutTime N=%ACDTime O=%AnsCalls
   P=N/100 Q=O/100 R=E+H S=X+W T=1-Q
   U=clock(C) V=clock(D) W=clock(F) X=clock(G) Y=clock(I) Z=clock(M)
   ========================================================================= */
function cmsTeamDerive(row){ // row keyed by column letter -> value
  var P=numOr(row.N)/100, Q=numOr(row.O)/100;
  var U=clockToSheetMinuteSerial(row.C), V=clockToSheetMinuteSerial(row.D),
      W=clockToSheetMinuteSerial(row.F), X=clockToSheetMinuteSerial(row.G),
      Y=clockToSheetMinuteSerial(row.I), Z=clockToSheetMinuteSerial(row.M);
  var R=numOr(row.E)+numOr(row.H);
  var S=(numish(X)?X:0)+(numish(W)?W:0);
  var T=1-Q;
  return {P:P,Q:Q,R:R,S:S,T:T,U:U,V:V,W:W,X:X,Y:Y,Z:Z};
}

/* ---- small utils ---- */
function isNum(v){ return typeof v==="number"&&isFinite(v); }
function numish(v){ return typeof v==="number"&&isFinite(v); }
function numOr(v){ var n=(typeof v==="number")?v:parseFloat(String(v).replace(/[,%\s]/g,"")); return isFinite(n)?n:0; }
function firstOfMonth(d){
  var dt=(d instanceof Date)?d:(typeof d==="string"?new Date(d):null);
  if(!dt||isNaN(dt)) return "";
  return new Date(dt.getFullYear(), dt.getMonth(), 1);
}

/* =========================================================================
   INGEST — append this month's raw exports (mapped to master format) to the
   master workbook's raw sheets, computing derived columns. Then Stage B can
   aggregate the target month. `H` is ReportPipeline.helpers.
   ========================================================================= */
function firstOfMonthDate(d){ var v=firstOfMonth(d); return v===""?null:v; }
function coalesce(){ for(var i=0;i<arguments.length;i++){ var v=arguments[i]; if(v!=null&&String(v).trim()!=="") return v; } return null; }

function ingest(H, master, rawWb, monthDate){
  var report={month:monthDate, domains:{}, notes:[]};
  function rawSheet(name){ return rawWb.Sheets[name]; }
  function has(name){ return !!rawWb.Sheets[name]; }
  // append mapped rows to a master sheet; mapFn(rawRowIndex)-> array of cell values (1-based, index0 unused)
  function append(masterName, rawName, keyColRaw, startRawRow, mapFn){
    var mS=master.Sheets[masterName], rS=rawSheet(rawName);
    if(!mS||!rS) { report.domains[masterName]="skipped (sheet missing)"; return 0; }
    var lastRaw=H.lastRowOfCol(rS,keyColRaw);
    var dest=H.lastRowOfCol(mS, destKeyCol(masterName))+1;
    var n=0;
    for(var i=startRawRow;i<=lastRaw;i++){
      var vals=mapFn(rS,i); if(!vals) continue;
      for(var c=1;c<vals.length;c++){ if(vals[c]!==undefined && vals[c]!==null && vals[c]!=="") H.setCell(mS,dest,c,vals[c]); }
      dest++; n++;
    }
    report.domains[masterName]="added "+n+" rows";
    return n;
  }
  function destKeyCol(masterName){
    // a column that is populated on every data row (used to find append point)
    switch(masterName){
      case "CMS Raw Data": case "CMS Rep Raw Data": return 2;    // Skill / Agent
      case "Hospital Rep Data": case "Message Data": return 2;   // Completion time
      default: return 1;                                          // Month / RespID / etc.
    }
  }
  var md=monthDate;

  // ---- EPIC WQ ---- raw:1WQID 2Name 3MaxAge 4Start 5End 6StartAmt 7EndAmt 8AmtChng 9CntChng 10DefCnt 11DefAmt
  if(has("EPIC WQ Raw Data")){
    var keyMap=buildKeyTargetDays(H,master);
    append("EPIC WQ Raw Data","EPIC WQ Raw Data",1,2,function(rS,i){
      var wqid=H.getCell(rS,i,1); if(wqid==null||String(wqid).trim()==="") return null;
      var o=[]; o[1]=md; for(var c=1;c<=11;c++) o[c+1]=H.getCell(rS,i,c);
      o[13]=keyMap[String(wqid).trim()]!==undefined?keyMap[String(wqid).trim()]:"NONE";
      return o;
    });
  }
  // ---- EPIC Productivity ---- raw:1Staff 2AvgTime 3Gap 4UniqAcct 5UniqGuar 6Activities 7Score 8Recovered 9Worked
  if(has("EPIC Rep Prod Raw Data")){
    append("EPIC Productivity Data","EPIC Rep Prod Raw Data",1,2,function(rS,i){
      var staff=H.getCell(rS,i,1); if(staff==null||String(staff).trim()==="") return null;
      var o=[]; o[1]=md; for(var c=1;c<=9;c++) o[c+1]=H.getCell(rS,i,c); return o;
    });
  }
  // ---- CMS Team ---- raw:1Skill 2AvgSpeed 3AvgAban 4ACDCalls 5AvgACD 6AvgACW 7Aban 8MaxDelay 9FlowIn 10FlowOut 11ExtnOutCalls 12AvgExtnOut 13%ACD 14%Ans
  if(has("CMS Raw Team Data")){
    append("CMS Raw Data","CMS Raw Team Data",1,2,function(rS,i){
      var skill=H.getCell(rS,i,1); if(skill==null||String(skill).trim()==="") return null;
      var o=[]; o[1]=md; for(var c=1;c<=14;c++) o[c+1]=H.getCell(rS,i,c);
      // derived P..Z from base cols C..O (master cols 3..15)
      var row={C:o[3],D:o[4],E:o[5],F:o[6],G:o[7],H:o[8],I:o[9],M:o[13],N:o[14],O:o[15]};
      var d=cmsTeamDerive(row);
      o[16]=d.P;o[17]=d.Q;o[18]=d.R;o[19]=d.S;o[20]=d.T;o[21]=d.U;o[22]=d.V;o[23]=d.W;o[24]=d.X;o[25]=d.Y;o[26]=d.Z;
      return o;
    });
  }
  // ---- CMS Flow Out ---- raw already in master layout (1..24), col1=Month
  if(has("CMS Raw Flow Out Data")){
    append("CMS Flow Out","CMS Raw Flow Out Data",1,2,function(rS,i){
      var m0=H.getCell(rS,i,1); if(m0==null) return null;
      var o=[]; o[1]=md; for(var c=2;c<=24;c++) o[c]=H.getCell(rS,i,c); return o;
    });
  }
  // ---- IVR ---- raw:1Resp 2Caller 3Acct 4AuthSucc 5Bal 6Paid 7Result 8Date 9Dur 10Xfer 11Term
  if(has("IVR Raw Data")){
    append("IVR Raw Data","IVR Raw Data",1,2,function(rS,i){
      var resp=H.getCell(rS,i,1); if(resp==null||String(resp).trim()==="") return null;
      var o=[]; o[1]=resp;o[2]=H.getCell(rS,i,2);o[3]=H.getCell(rS,i,3);o[4]="";o[5]="";
      o[6]=H.getCell(rS,i,5);o[7]=H.getCell(rS,i,6);o[8]=H.getCell(rS,i,7);o[9]=H.getCell(rS,i,8);
      o[10]=H.getCell(rS,i,9);o[11]=H.getCell(rS,i,10);o[12]=H.getCell(rS,i,11);
      var d=ivrDerive({accountID:o[3],amountPaid:o[7],callDate:o[9],callTerm:o[12]});
      o[13]=d.M;o[14]=d.N;o[15]=d.O;o[16]=d.P;o[17]=d.Q;o[18]=d.R;o[19]=d.S;
      return o;
    });
  }
  // ---- Hospital ---- survey:1Id 2Start 3Completion 4Email 5Name 6Location 7TimeSpent 8DidFCA 9ReasonNoFCA 10ReasonFCA
  if(has("Hospital POS Survey Raw Data")){
    append("Hospital Rep Data","Hospital POS Survey Raw Data",3,2,function(rS,i){
      var comp=H.getCell(rS,i,3); if(comp==null) return null;
      var o=[]; o[1]=firstOfMonthDate(comp); o[2]=comp; o[3]=H.getCell(rS,i,4); o[4]=H.getCell(rS,i,5);
      o[5]=H.getCell(rS,i,6); o[6]=H.getCell(rS,i,7); o[7]=coalesce(H.getCell(rS,i,9),H.getCell(rS,i,10));
      return o;
    });
  }
  // ---- Messages ---- survey:1Id 2Start 3Completion 4Email 5Name 6Where 7WhoSent 8Agency 9Status 10Reason
  if(has("Correspondence Survey Raw Data")){
    append("Message Data","Correspondence Survey Raw Data",3,2,function(rS,i){
      var comp=H.getCell(rS,i,3); if(comp==null) return null;
      var o=[]; o[1]=firstOfMonthDate(comp); o[2]=comp; o[3]=H.getCell(rS,i,4); o[4]=H.getCell(rS,i,5);
      o[5]=H.getCell(rS,i,6); o[6]=H.getCell(rS,i,7); o[7]=H.getCell(rS,i,8); o[8]=H.getCell(rS,i,9); o[9]=H.getCell(rS,i,10);
      return o;
    });
  }
  // ---- Collections (from raw file's own tab, if populated) ----
  ingestCollections(H, master, rawWb, md, report);
  // CMS Rep detail (appendix only, no chart) — not auto-computed
  report.notes.push("CMS per-rep appendix detail is not auto-computed (source time format is ambiguous and no dashboard chart uses it).");
  return report;
}

/* Collections come from a separate source; the standard raw export tab is
   usually empty. Populate from the raw file's "Rep Collection Data" tab if it
   has data, matching the master's "Rep Collections Data" layout. */
function ingestCollections(H, master, rawWb, md, report){
  var rs=rawWb.Sheets["Rep Collection Data"] || rawWb.Sheets["Rep Collections Data"];
  var last = rs? H.lastRowOfCol(rs,1) : 0;
  if(!rs || last<2){ report.domains["Collections"]="no collection data in raw file (add it to the 'Rep Collection Data' tab or drop a collections file)"; report.collectionsMissing=true; return; }
  // Expect columns matching master Rep Collections Data: Date, Employee Name, Total Collections, Transaction Count, Name
  var mS=master.Sheets["Rep Collections Data"]; var dest=H.lastRowOfCol(mS,1)+1, n=0;
  for(var i=2;i<=last;i++){
    var dt=H.getCell(rs,i,1); if(dt==null) continue;
    for(var c=1;c<=5;c++){ var v=H.getCell(rs,i,c); if(v!==null&&v!==undefined&&v!=="") H.setCell(mS,dest,c,v); }
    dest++; n++;
  }
  report.domains["Collections"]="added "+n+" rep rows (monthly goal/actual must be present in 'Monthly Collections Data')";
}

/* Key sheet: WQ ID (A3:A50) -> Target Days (C3:C50) */
function buildKeyTargetDays(H,master){
  var ks=master.Sheets["Key"]; var map={}; if(!ks) return map;
  for(var r=3;r<=50;r++){ var id=H.getCell(ks,r,1); if(id!=null&&String(id).trim()!=="") map[String(id).trim()]=H.getCell(ks,r,3); }
  return map;
}

return { clockToSeconds:clockToSeconds, clockToSheetMinuteSerial:clockToSheetMinuteSerial,
         ivrAuthStatus:ivrAuthStatus, ivrDerive:ivrDerive, cmsTeamDerive:cmsTeamDerive,
         firstOfMonth:firstOfMonth, ingest:ingest };
});
