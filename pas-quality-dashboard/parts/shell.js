/* ================= RMS Bypassed Warnings — app shell (import + column config) ================= */
(function () {
  "use strict";
  var sha512 = window.sha512;
  var XLSX = window.XLSX;
  var rawRecords = [];         // rows keyed by original header
  var headers = [];            // union of headers, in first-seen order
  var fileNames = [];
  var deptKey = null, empKey = null, colMap = {};
  var copayRecords = [], copayHeaders = [], copayFileNames = [];
  var regEntries = [], regMonth = "", regFileNames = [];   // Detailed-View registrations (staff -> count)
  var coverageRecords = [], coverageHeaders = [], coverageFileNames = [];  // May.xlsx PB (coverage verification) rows
  var activeModule = "bypass"; // "bypass" | "copays" | "coverage"
  var gState = { dept: "", emp: "", period: "", gran: "month" };   // global Department / Employee / Period filters (shown above the tabs, shared across modules). period "" = all periods; gran = "month" | "week"
  var copayPeriod = "all";              // remembered month selection for the Copays tab (standalone exports only)
  var periodInitialized = false;        // once true, respect the user's Month pick; until then, default to the most recent month
  // On first render after an import, open on the MOST RECENT month (leaders want the latest numbers by
  // default). The month-over-month sections still show every month regardless of this pick.
  function ensurePeriodDefault() {
    if (periodInitialized) return;
    var ms = historyMonths();
    if (ms.length >= 2) gState.period = ms[ms.length - 1];   // newest period; a single period stays combined
    periodInitialized = true;
  }
  /* Flipping Week/Month invalidates the current pick (a "2026-04" month key means nothing in a list of
     week keys), so the selection is reset and re-defaulted to the newest bucket at the new granularity. */
  function setGranularityG(g) {
    g = (g === "week") ? "week" : "month";
    if (g === gState.gran) return;
    gState.gran = g; gState.period = ""; periodInitialized = false;
    if (window.RMSViewer && RMSViewer.setGranularity) RMSViewer.setGranularity(g);
    ensurePeriodDefault();
  }
  var ruleInclude = null, ruleExclude = null, ruleIncludeRaw = [], ruleExcludeRaw = [], ruleFileName = "", applyRules = true;
  function ruleKey(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function escText(s) { return String(s == null ? "" : s).replace(/[&<>]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]; }); }
  function setRules(rules) {
    ruleIncludeRaw = rules.include || []; ruleExcludeRaw = rules.exclude || [];
    ruleInclude = new Set(ruleIncludeRaw.map(ruleKey)); ruleExclude = new Set(ruleExcludeRaw.map(ruleKey));
  }

  /* ---------- byte utils ---------- */
  function b64ToBytes(b64) { var bin = atob(b64), a = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; }
  function u32le(n) { var a = new Uint8Array(4); a[0] = n & 255; a[1] = (n >>> 8) & 255; a[2] = (n >>> 16) & 255; a[3] = (n >>> 24) & 255; return a; }
  function concat(list) { var len = 0, i; for (i = 0; i < list.length; i++) len += list[i].length; var out = new Uint8Array(len), off = 0; for (i = 0; i < list.length; i++) { out.set(list[i], off); off += list[i].length; } return out; }
  function sha(buf) { return new Uint8Array(sha512.arrayBuffer(buf)); }

  /* ---------- WebCrypto AES-CBC no-padding decrypt (append-trick) ---------- */
  async function aesNoPad(keyBytes, iv, data) {
    var key = await crypto.subtle.importKey("raw", keyBytes, { name: "AES-CBC" }, false, ["encrypt", "decrypt"]);
    var n = data.length, Cn = data.slice(n - 16, n), input = new Uint8Array(16);
    for (var i = 0; i < 16; i++) input[i] = 0x10 ^ Cn[i];
    var enc = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-CBC", iv: new Uint8Array(16) }, key, input));
    var ext = new Uint8Array(n + 16); ext.set(data, 0); ext.set(enc.slice(0, 16), n);
    return new Uint8Array(await crypto.subtle.decrypt({ name: "AES-CBC", iv: iv.slice(0, 16) }, key, ext));
  }

  /* ---------- ECMA-376 Agile decryption ---------- */
  var BK_KEY = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
  async function decryptAgile(arrayBuffer, password, onProgress) {
    var cfb = XLSX.CFB.read(new Uint8Array(arrayBuffer), { type: "array" });
    function stream(name) { var e = XLSX.CFB.find(cfb, name); return e ? new Uint8Array(e.content) : null; }
    var info = stream("EncryptionInfo") || stream("/EncryptionInfo");
    var pkg = stream("EncryptedPackage") || stream("/EncryptedPackage");
    if (!info || !pkg) throw new Error("Not a recognized encrypted Office file.");
    var xml = new TextDecoder("utf-8").decode(info.subarray(8));
    if (xml.indexOf("cipherAlgorithm=\"AES\"") < 0 || xml.indexOf("SHA512") < 0) throw new Error("Unsupported encryption (need AES-256 / SHA-512 agile).");
    function attr(seg, a) { var m = seg.match(new RegExp(a + '="([^"]*)"')); return m ? m[1] : null; }
    // Match the opening tag + attributes regardless of namespace prefix (p:, or none) or whether
    // the tag is self-closing. [^>]* captures every attribute up to the first '>', and base64
    // salt/key values never contain '>'. Guard against null so a format we can't parse gives a
    // clear message instead of "Cannot read properties of null (reading '0')".
    var kdMatch = xml.match(/<(?:[A-Za-z0-9]+:)?keyData\b[^>]*>/);
    if (!kdMatch) throw new Error("Encrypted file: could not find keyData (unsupported encryption layout).");
    var kd = kdMatch[0];
    // A file may carry several key encryptors (e.g. password + certificate). Pick the password one,
    // identified by spinCount + encryptedKeyValue; fall back to the first if only one is present.
    var keCandidates = xml.match(/<(?:[A-Za-z0-9]+:)?encryptedKey\b[^>]*>/g) || [];
    var ke = null;
    for (var _k = 0; _k < keCandidates.length; _k++) {
      if (/spinCount=/.test(keCandidates[_k]) && /encryptedKeyValue=/.test(keCandidates[_k])) { ke = keCandidates[_k]; break; }
    }
    if (!ke) ke = keCandidates[0];
    if (!ke) throw new Error("Encrypted file: could not find the password key (unsupported encryption layout).");
    var kdSalt = b64ToBytes(attr(kd, "saltValue"));
    var spin = +attr(ke, "spinCount"), keyBytes = (+attr(ke, "keyBits")) / 8, bs = +attr(ke, "blockSize") || 16;
    var pwSalt = b64ToBytes(attr(ke, "saltValue")), encKeyVal = b64ToBytes(attr(ke, "encryptedKeyValue"));
    var pw = new Uint8Array(password.length * 2);
    for (var i = 0; i < password.length; i++) { var c = password.charCodeAt(i); pw[i * 2] = c & 255; pw[i * 2 + 1] = (c >>> 8) & 255; }
    var H = sha(concat([pwSalt, pw]));
    for (var j = 0; j < spin; j++) { H = sha(concat([u32le(j), H])); if (onProgress && (j & 8191) === 0) onProgress(j / spin); }
    var keyEncKey = sha(concat([H, BK_KEY])).slice(0, keyBytes);
    var secretKey = (await aesNoPad(keyEncKey, pwSalt, encKeyVal)).slice(0, keyBytes);
    var total = pkg[0] | (pkg[1] << 8) | (pkg[2] << 16) | (pkg[3] * 16777216);
    var enc = pkg.subarray(8), SEG = 4096, out = new Uint8Array(enc.length);
    for (var seg = 0, off = 0; off < enc.length; off += SEG, seg++) {
      var iv = sha(concat([kdSalt, u32le(seg)])).slice(0, bs);
      var chunk = enc.subarray(off, Math.min(off + SEG, enc.length));
      out.set((await aesNoPad(secretKey, iv, chunk)).subarray(0, chunk.length), off);
      if (onProgress) onProgress(0.5 + 0.5 * (off / enc.length));
    }
    return out.subarray(0, total);
  }

  /* ---------- header role detection ---------- */
  function norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function detectRoles(hs) {
    var N = hs.map(norm);
    function find(pred) { for (var i = 0; i < hs.length; i++) if (pred(N[i], hs[i])) return hs[i]; return null; }
    var cm = {};
    cm.checkinUser = find(function (n) { return n.indexOf("checkin") >= 0; });
    cm.user = find(function (n, h) { return n.indexOf("bypass") >= 0 && n.indexOf("user") >= 0; }) || find(function (n, h) { return n === "user" && h !== cm.checkinUser; });
    cm.loginDept = find(function (n) { return n === "logindept" || (n.indexOf("login") >= 0 && n.indexOf("dept") >= 0); });
    cm.department = find(function (n, h) { return h !== cm.loginDept && (n === "department" || (n.indexOf("department") >= 0 && n.indexOf("login") < 0)); });
    cm.encDate = find(function (n) { return n.indexOf("encounter") >= 0; });
    cm.date = find(function (n, h) { return h !== cm.encDate && (n === "date" || (n.indexOf("date") >= 0 && n.indexOf("encounter") < 0)); });
    cm.time = find(function (n) { return n.indexOf("time") >= 0; });
    cm.messageType = find(function (n) { return n.indexOf("messagetype") >= 0 || (n.indexOf("message") >= 0 && n.indexOf("type") >= 0); });
    cm.errorText = find(function (n) { return n.indexOf("errortext") >= 0 || (n.indexOf("error") >= 0 && n.indexOf("text") >= 0); });
    cm.errorDetail = find(function (n, h) { return h !== cm.errorText && (n.indexOf("errortypedetail") >= 0 || n.indexOf("errordetail") >= 0 || n.indexOf("errortype") >= 0 || (n.indexOf("error") >= 0 && n.indexOf("detail") >= 0)); });
    cm.record = find(function (n) { return n.indexOf("record") >= 0; });
    cm.patient = find(function (n) { return n.indexOf("patient") >= 0; });
    cm.workflow = find(function (n) { return n.indexOf("workflow") >= 0; });
    return cm;
  }

  /* ---------- parse a file into {headers, rows} ---------- */
  function isOLE(b) { return b.length > 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0; }
  function isZip(b) { return b.length > 4 && b[0] === 0x50 && b[1] === 0x4B; }
  // Defensive worksheet -> array-of-rows reader. Replaces XLSX.utils.sheet_to_json({header:1}),
  // which in SheetJS 0.18.5 can throw "Cannot read properties of null (reading '0')" for the whole
  // sheet if a single cell/row can't be formatted. This reads every cell individually and guards
  // each access, so a bad cell yields "" instead of aborting the entire file. Column indices are
  // absolute (index 0 = column A), matching header:1 output. Fully-blank rows are dropped
  // (blankrows:false semantics).
  function sheetToMatrix(ws) {
    if (!ws || ws["!ref"] == null) return [];
    var range;
    try { range = XLSX.utils.decode_range(ws["!ref"]); } catch (e) { return []; }
    if (!range || !range.s || !range.e) return [];
    var out = [];
    for (var R = range.s.r; R <= range.e.r; R++) {
      var row = [], any = false;
      for (var C = range.s.c; C <= range.e.c; C++) {
        var v = "";
        try {
          var cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
          if (cell != null) {
            if (cell.w != null) v = cell.w;
            else if (cell.v != null) v = cell.v instanceof Date ? cell.v.toISOString() : cell.v;
          }
        } catch (e) { v = ""; }
        row[C] = v;
        if (v !== "" && v != null) any = true;
      }
      if (any) out.push(row);
    }
    return out;
  }
  function findHeaderRowIdx(matrix) {
    // Score each candidate row by how many cells look like known header names; pick the best.
    // Robust to title / filter rows above the header (common in raw Epic exports).
    var TOK = ["checkin", "messagetype", "errortext", "errortype", "logindept", "workflow", "patient",
               "copaydue", "copaypaid", "copay", "mrn", "apptstatus", "visitdate", "provider", "encounter",
               "associatedrecord", "bypassed", "payer", "payor", "date", "time", "user", "dept"];
    var best = -1, bestScore = 0;
    for (var i = 0; i < Math.min(matrix.length, 15); i++) {
      var ns = (matrix[i] || []).map(norm).filter(function (n) { return n !== ""; });
      if (ns.length < 3) continue;
      var score = 0; ns.forEach(function (n) { if (TOK.some(function (t) { return n.indexOf(t) >= 0; })) score++; });
      if (score > bestScore) { bestScore = score; best = i; }
    }
    if (best >= 0 && bestScore >= 2) return best;
    // fallback: first row with >=3 non-empty cells
    for (var k = 0; k < matrix.length; k++) { var r = matrix[k] || []; if (r.filter(function (v) { return v != null && String(v).trim() !== ""; }).length >= 3) return k; }
    return 0;
  }
  async function fileToTable(file, password, onProgress) {
    var buf = await file.arrayBuffer(), bytes = new Uint8Array(buf), wbInput, type;
    if (isOLE(bytes)) { wbInput = await decryptAgile(buf, password || "", onProgress); type = "array"; }
    else if (isZip(bytes)) { wbInput = bytes; type = "array"; }
    else { wbInput = new TextDecoder("utf-8").decode(bytes); type = "string"; }
    var wb = XLSX.read(wbInput, { type: type, cellDates: false, raw: false });
    // Rules file? (NEW PQ "Include and Exclude" sheet: col A = include, col B = exclude)
    var ruleSheetName = wb.SheetNames.filter(function (name) { var n = norm(name); return n.indexOf("include") >= 0 && n.indexOf("exclude") >= 0; })[0];
    if (ruleSheetName) {
      var rm = sheetToMatrix(wb.Sheets[ruleSheetName]);
      var include = [], exclude = [];
      for (var ri = 1; ri < rm.length; ri++) {
        var a = rm[ri][0], b = rm[ri][1];
        if (a != null && String(a).trim() !== "") include.push(String(a).trim());
        if (b != null && String(b).trim() !== "") exclude.push(String(b).trim());
      }
      return { headers: [], rows: [], rules: { include: include, exclude: exclude, sheet: ruleSheetName } };
    }
    var best = null, bestScore = -1, cov = null, reg = null;
    wb.SheetNames.forEach(function (name) {
      var m = sheetToMatrix(wb.Sheets[name]);
      var hi = findHeaderRowIdx(m), hs = (m[hi] || []).map(function (v) { return String(v == null ? "" : v).trim(); });
      var hn = hs.map(norm);
      // Coverage source sheet (e.g. "PB"): the raw table with VERIF_STATUS + CHECK_IN_USER columns.
      // Prefer it outright so pivot/summary tabs (whose label cells merely echo CHECK_IN_USER) can't win.
      if (hn.some(function (n) { return n.indexOf("verifstatus") >= 0; }) && hn.some(function (n) { return n.indexOf("checkin") >= 0 && n.indexOf("user") >= 0; })) {
        if (!cov || m.length > cov.matrix.length) cov = { hi: hi, hs: hs, matrix: m };
      }
      // Registrations sheet (Detailed View: a "Staff" column + a monthly total). A workbook can carry
      // several tabs (e.g. a "Group Comparison Scorecard" dashboard) — find the row that actually starts
      // the Staff table and prefer it, so the scorecard tab can't win on row-count alone.
      var regHi = -1;
      for (var _r = 0; _r < Math.min(m.length, 15); _r++) {
        if ((m[_r] || []).map(norm).some(function (n) { return n.indexOf("staff") >= 0; })) { regHi = _r; break; }
      }
      if (regHi >= 0) {
        var regHs = (m[regHi] || []).map(function (v) { return String(v == null ? "" : v).trim(); });
        if (looksLikeRegistrations(regHs) && (!reg || (m.length - regHi) > (reg.matrix.length - reg.hi))) reg = { hi: regHi, hs: regHs, matrix: m };
      }
      var score = (hn.some(function (n) { return n.indexOf("checkin") >= 0; }) ? 100000 : 0) + m.length;
      if (score > bestScore) { bestScore = score; best = { hi: hi, hs: hs, matrix: m }; }
    });
    if (cov) best = cov;
    else if (reg) best = reg;
    if (!best) return { headers: [], rows: [] };
    // de-duplicate/clean headers
    var hs = best.hs.map(function (h, i) { return h || ("Column " + (i + 1)); });
    var seen = {}; hs = hs.map(function (h) { if (seen[h] == null) { seen[h] = 0; return h; } seen[h]++; return h + " (" + seen[h] + ")"; });
    var rows = [];
    for (var r = best.hi + 1; r < best.matrix.length; r++) {
      var row = best.matrix[r] || [], obj = {}, any = false;
      for (var c = 0; c < hs.length; c++) { var v = row[c]; v = v == null ? "" : String(v).trim(); obj[hs[c]] = v; if (v) any = true; }
      if (any) rows.push(obj);
    }
    return { headers: hs, rows: rows };
  }

  /* ---------- DOM ---------- */
  var $ = function (id) { return document.getElementById(id); };
  var drop = $("drop"), fileInput = $("file"), pwInput = $("pw"), status = $("status"),
      dashboard = $("dashboard"), empty = $("empty"), clearBtn = $("btn-clear"), config = $("config"),
      deptSel = $("deptcol"), empSel = $("empcol"), fileList = $("filelist"), importwrap = $("importwrap"),
      moduletabs = $("moduletabs"), rulebar = $("rulebar"), globalfilters = $("globalfilters"),
      appheaderactions = $("appheaderactions"), appheaderexports = $("appheaderexports");
  function showImport() { importwrap.style.display = ""; try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (e) { window.scrollTo(0, 0); } }
  function setStatus(msg, kind) { status.textContent = msg || ""; status.className = "status" + (kind ? " " + kind : ""); }
  function fmt(n) { return n.toLocaleString("en-US"); }

  /* ---------- registrations (Detailed View: col A = Staff, col B = monthly total) ---------- */
  function looksLikeRegistrations(hs) {
    if (!hs || !hs.length) return false;
    var hasStaff = hs.some(function (h) { return norm(h).indexOf("staff") >= 0; });
    if (!hasStaff) return false;
    var MON = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    var hasCount = hs.some(function (h) { var n = norm(h); if (n.indexOf("staff") >= 0) return false;
      return n.indexOf("registration") >= 0 || n.indexOf("total") >= 0 || MON.some(function (m) { return n === m || n.indexOf(m) >= 0; }); });
    return hasStaff && (hasCount || hs.length <= 3);
  }
  function parseRegistrations(hs, rows) {
    var MON = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    var staffH = null, countH = null, monthH = null;
    hs.forEach(function (h) { if (staffH == null && norm(h).indexOf("staff") >= 0) staffH = h; });
    if (staffH == null) staffH = hs[0];
    hs.forEach(function (h) { if (h === staffH || countH != null) return; var n = norm(h);
      if (MON.some(function (m) { return n === m || n.indexOf(m) >= 0; })) { countH = h; monthH = h; }
      else if (n.indexOf("registration") >= 0 || n.indexOf("total") >= 0) { countH = h; } });
    if (countH == null) { for (var i = 0; i < hs.length; i++) { if (hs[i] === staffH) continue;
      var numeric = rows.some(function (r) { var v = String(r[hs[i]] == null ? "" : r[hs[i]]).replace(/[,\s]/g, ""); return v !== "" && !isNaN(+v); });
      if (numeric) { countH = hs[i]; break; } } }
    var entries = [];
    rows.forEach(function (r) { var s = String(r[staffH] == null ? "" : r[staffH]).trim(); if (!s) return;
      if (/percentile|^total|average|median/i.test(s)) return;
      var raw = String(countH ? (r[countH] == null ? "" : r[countH]) : "").replace(/[,\s$]/g, "");
      var c = raw === "" ? 0 : +raw; if (isNaN(c)) c = 0;
      entries.push({ staff: s, count: c }); });
    var month = monthH ? String(monthH).trim() : (countH ? String(countH).trim() : "");
    return { entries: entries, month: month };
  }

  /* ---------- coverage accuracy (May.xlsx "Error by User" pivot: Count of CSN by VERIF_STATUS) ---------- */
  // The 11 VERIF_STATUS values selected in the spreadsheet's pivot (its "filter"); the rest are excluded.
  var COVERAGE_STATUSES = ["Contact Payor", "Content Error", "Data Mismatch", "E-Rejected", "Manually Verified",
    "New", "Other Additional Payor Returned", "Plan Mismatch", "Replacement Returned", "Verified by Card", "Verified by Fax"];
  function looksLikeCoverage(hs) {
    if (!hs || !hs.length) return false;
    var N = hs.map(norm);
    var hasStatus = N.some(function (n) { return n.indexOf("verifstatus") >= 0 || (n.indexOf("verif") >= 0 && n.indexOf("status") >= 0); });
    var hasUserOrCsn = N.some(function (n) { return (n.indexOf("checkin") >= 0 && n.indexOf("user") >= 0) || n === "csn" || n.indexOf("cvgverification") >= 0; });
    return hasStatus && hasUserOrCsn;
  }
  function covNameKey(s) { s = String(s == null ? "" : s).trim(); if (!s) return "";
    var last, first;
    if (s.indexOf(",") >= 0) { var p = s.split(","); last = p[0]; first = (p[1] || "").trim().split(/\s+/)[0] || ""; }
    else { var t = s.split(/\s+/); first = t[0] || ""; last = t.length > 1 ? t[t.length - 1] : ""; }
    return String(last).toLowerCase().replace(/[^a-z0-9]/g, "") + "|" + String(first).toLowerCase().replace(/[^a-z0-9]/g, ""); }
  // The exact record set the Bypassed Warnings tab shows: rawRecords with the
  // include/exclude rules + same-day + same-user Power Query logic applied (when a
  // rules file is loaded and rules are on). Keep this in lock-step with render().
  // All months (rule-filtered), NOT scoped to the selected month — used for month-over-month.
  function bypassDisplayRecordsBase() {
    var filterActive = !!(ruleInclude && applyRules);
    if (!filterActive) return rawRecords;
    var edKey = colMap.errorDetail, dK = colMap.date, edK = colMap.encDate, ciK = colMap.checkinUser, buK = colMap.user;
    var asDay = function (v) { if (v == null) return ""; var s = String(v).trim(); if (!s) return ""; var dt = new Date(s); return isNaN(dt) ? s : (dt.getFullYear() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate()); };
    var same = function (a, b) { return String(a == null ? "" : a).trim() === String(b == null ? "" : b).trim(); };
    return rawRecords.filter(function (r) {
      if (dK && edK && asDay(r[dK]) !== asDay(r[edK])) return false;   /* Encounter Date == Date */
      if (ciK && buK && !same(r[ciK], r[buK])) return false;          /* Check-in User == Bypassed Warning User */
      var k = ruleKey(r[edKey]);
      if (ruleExclude && ruleExclude.has(k)) return false;            /* not on Exclude list */
      return ruleInclude.has(k);                                      /* on Include list */
    });
  }
  function bypassDisplayRecords() {
    return scopeByPeriod(bypassDisplayRecordsBase(), bypassDateHeader());   // narrow to the globally-selected month (if any)
  }
  function coverageUserKeySet() {
    var set = {};
    function add(v) { var k = covNameKey(v); if (k && k !== "|") set[k] = 1; }
    // The union of everyone on the Bypassed Warnings sheet AND the copay sheet, so Coverage counts
    // encounters for any name that shows up on either — including copay check-in users who never
    // appear on the bypass sheet. (Matches the "every employee" filter on the other tabs.)
    if (rawRecords.length && empKey) bypassDisplayRecords().forEach(function (r) { add(r[empKey]); });
    if (copayRecords.length && window.RMSCopays) { var ccm = RMSCopays.detectCopayCols(copayHeaders);
      if (ccm.checkinUser) copayRecords.forEach(function (r) { add(r[ccm.checkinUser]); }); }
    return set;
  }
  // XLOOKUP each employee to the clinic/department they work in, per the bypass (and copay) sheets.
  // The coverage tab groups encounters by this looked-up clinic — NOT by the coverage file's own
  // DEPARTMENT column — so a person's coverage rows count under the clinic they show up in on bypass.
  function coverageEmpDeptMap() {
    var map = {};
    function add(nameV, deptV) { var k = covNameKey(nameV); if (!k || k === "|") return; var d = String(deptV == null ? "" : deptV).replace(/\s+/g, " ").trim(); if (!d) return; if (!map[k]) map[k] = d; }
    var bDept = deptKey || (colMap && (colMap.department || colMap.loginDept));
    // Use the same rule-filtered records the Bypassed Warnings tab groups by, so each
    // employee's clinic matches the clinic they appear under on that tab.
    if (bDept) bypassDisplayRecords().forEach(function (r) {
      if (empKey) add(r[empKey], r[bDept]);
      if (colMap && colMap.checkinUser) add(r[colMap.checkinUser], r[bDept]);
      if (colMap && colMap.user) add(r[colMap.user], r[bDept]);
    });
    if (copayRecords.length && window.RMSCopays) { var ccm = RMSCopays.detectCopayCols(copayHeaders);
      if (ccm.checkinUser && ccm.dept) copayRecords.forEach(function (r) { add(r[ccm.checkinUser], r[ccm.dept]); }); }
    return map;
  }
  function findCovHeader(hs, preds) { for (var i = 0; i < hs.length; i++) { var n = norm(hs[i]); for (var j = 0; j < preds.length; j++) if (preds[j](n)) return hs[i]; } return null; }
  function coverageMonthLabel(dateH) {
    if (!dateH) return "";
    var lo = null, hi = null;
    coverageRecords.forEach(function (r) { var s = String(r[dateH] == null ? "" : r[dateH]).trim(); if (!s) return; var d = new Date(s); if (isNaN(d)) return; var t = d.getTime(); if (lo == null || t < lo) lo = t; if (hi == null || t > hi) hi = t; });
    if (lo == null) return "";
    var M = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], a = new Date(lo), b = new Date(hi);
    if (a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth()) return M[a.getMonth()] + " " + a.getFullYear();
    return M[a.getMonth()] + " " + a.getFullYear() + " – " + M[b.getMonth()] + " " + b.getFullYear();
  }
  function ce(tag, cls, txt) { var e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; }
  var covState = { dept: "", emp: "" };   // Department / Employee filters on the Coverage Accuracy tab
  function coverageCols() {
    var hs = coverageHeaders;
    return {
      status: findCovHeader(hs, [function (n) { return n.indexOf("verifstatus") >= 0; }, function (n) { return n.indexOf("verif") >= 0 && n.indexOf("status") >= 0; }]),
      user: findCovHeader(hs, [function (n) { return n.indexOf("checkin") >= 0 && n.indexOf("user") >= 0; }]),
      csn: findCovHeader(hs, [function (n) { return n === "csn"; }, function (n) { return n.indexOf("csn") >= 0; }]),
      date: findCovHeader(hs, [function (n) { return n.indexOf("encounterdate") >= 0; }, function (n) { return n.indexOf("encounter") >= 0 && n.indexOf("date") >= 0; }]),
      dept: findCovHeader(hs, [function (n) { return n.indexOf("department") >= 0; }, function (n) { return n === "dept"; }])
    };
  }
  function coverageAggregate(restrict, applyState) {
    var c = coverageCols();
    var covRecs = scopeByPeriod(coverageRecords, c.date);   // narrow to the globally-selected month (if any)
    // restrict = explicit set of employee name-keys (e.g. one clinic's staff); otherwise every
    // employee on the bypass/copay sheets. applyState folds in the on-screen Dept/Employee pickers.
    var allowed = restrict || coverageUserKeySet(), haveFilter = Object.keys(allowed).length > 0;
    var CANON = {}; COVERAGE_STATUSES.forEach(function (s) { CANON[s.toLowerCase()] = s; });
    var empDeptMap = coverageEmpDeptMap();
    function dpOf(r) { var d = c.user ? empDeptMap[covNameKey(r[c.user])] : null; return d || "(no clinic)"; }
    function unOf(r) { return c.user ? cleanDisplayName(r[c.user]) : ""; }
    function baseEligible(r) {
      var st = String(r[c.status] == null ? "" : r[c.status]).trim(); if (!CANON[st.toLowerCase()]) return false;
      if (c.user && haveFilter && !allowed[covNameKey(r[c.user])]) return false;
      if (c.csn && String(r[c.csn] == null ? "" : r[c.csn]).trim() === "") return false;
      return true;
    }
    var deptSet = {}, empByDept = { "": {} };
    covRecs.forEach(function (r) { if (!baseEligible(r)) return;
      var dp = dpOf(r); deptSet[dp] = 1;
      if (c.user) { var un = unOf(r); empByDept[""][un] = 1; if (!empByDept[dp]) empByDept[dp] = {}; empByDept[dp][un] = 1; } });
    if (applyState && covState.dept && !deptSet[covState.dept]) { covState.dept = ""; covState.emp = ""; }
    var empPool = empByDept[(applyState ? covState.dept : "") || ""] || {};
    // Keep covState.emp even if this employee has no coverage rows — the dropdown lists ALL employees and
    // filtering matches by normalized name, so a selected employee with no errors just shows zeros.
    var counts = {}, deptCounts = {}, grand = 0, matched = {}, excluded = {}, empCounts = {};
    var empErrors = {};   // Count of CSN per check-in user, under the SAME filters as `grand`
    var byMonth = {}, empMonth = {};   // month-over-month: month -> Count of CSN, and employee -> {month -> count}
    covRecs.forEach(function (r) {
      var st = String(r[c.status] == null ? "" : r[c.status]).trim(); var canon = CANON[st.toLowerCase()]; if (!canon) return;
      if (c.user && haveFilter) { var uk = covNameKey(r[c.user]); if (!allowed[uk]) { if (uk && uk !== "|") excluded[uk] = 1; return; } }
      if (c.csn && String(r[c.csn] == null ? "" : r[c.csn]).trim() === "") return;
      if (applyState && covState.dept && dpOf(r) !== covState.dept) return;
      if (c.user) { var _un = unOf(r); empCounts[_un] = (empCounts[_un] || 0) + 1; }   // per-employee count for the dropdown (before the employee filter)
      if (applyState && covState.emp && c.user && covNameKey(unOf(r)) !== covNameKey(covState.emp)) return;
      counts[canon] = (counts[canon] || 0) + 1; grand++;
      if (c.user) { var _ue = unOf(r); empErrors[_ue] = (empErrors[_ue] || 0) + 1; }
      var dpv = dpOf(r); deptCounts[dpv] = (deptCounts[dpv] || 0) + 1;
      if (c.user) matched[covNameKey(r[c.user])] = 1;
    });
    // Month-over-month is built from EVERY month (ignores the global month filter) so the comparison stays
    // visible even when the tab is scoped to one month. Same eligibility as above, minus the period scope.
    coverageRecords.forEach(function (r) {
      var st2 = String(r[c.status] == null ? "" : r[c.status]).trim(); if (!CANON[st2.toLowerCase()]) return;
      if (c.user && haveFilter && !allowed[covNameKey(r[c.user])]) return;
      if (c.csn && String(r[c.csn] == null ? "" : r[c.csn]).trim() === "") return;
      if (applyState && covState.dept && dpOf(r) !== covState.dept) return;
      if (applyState && covState.emp && c.user && covNameKey(unOf(r)) !== covNameKey(covState.emp)) return;
      var mk = c.date ? gPeriodFrom(r[c.date]) : ""; if (!mk) return;
      byMonth[mk] = (byMonth[mk] || 0) + 1;
      if (c.user) { var un2 = unOf(r); if (!empMonth[un2]) empMonth[un2] = {}; empMonth[un2][mk] = (empMonth[un2][mk] || 0) + 1; }
    });
    var statusObj = {}; COVERAGE_STATUSES.forEach(function (s) { if (counts[s]) statusObj[s] = counts[s]; });
    return { cols: c, haveFilter: haveFilter, statusObj: statusObj, deptCounts: deptCounts, grand: grand, matched: matched, excluded: excluded, empCounts: empCounts,
      byMonth: byMonth, empMonth: empMonth, empErrors: empErrors,
      depts: Object.keys(deptSet).sort(), emps: Object.keys(empPool).sort(), month: gPeriodActive() ? periodLabelLongG(gState.period) : coverageMonthLabel(c.date) };
  }
  function coverageDeptTable(deptCounts, grand, onPick, firstColLabel) {
    var rows = Object.keys(deptCounts).map(function (k) { return { name: k, n: deptCounts[k] }; }).sort(function (a, b) { return b.n - a.n || String(a.name).localeCompare(String(b.name)); });
    var tw = ce("div", "rms-tablewrap"), sc = ce("div", "rms-scroll"), t = ce("table", "rms-table");
    var thead = ce("thead"), htr = ce("tr"); htr.appendChild(ce("th", null, firstColLabel || "Department (clinic)")); htr.appendChild(ce("th", "num", "Count of CSN")); thead.appendChild(htr); t.appendChild(thead);
    var tb = ce("tbody");
    rows.forEach(function (r) { var tr = ce("tr");
      var td0 = ce("td");
      if (onPick) { var b = ce("button", "rms-linkbtn", r.name); b.addEventListener("click", function () { onPick(r.name); }); td0.appendChild(b); } else td0.textContent = r.name;
      tr.appendChild(td0);
      var td = ce("td", "num"); td.appendChild(ce("span", "pill", fmt(r.n))); tr.appendChild(td); tb.appendChild(tr); });
    var gtr = ce("tr", "rms-totalrow"); gtr.setAttribute("style", "border-top:2px solid var(--border-2);font-weight:600");
    var g1 = ce("td"); g1.appendChild(ce("b", null, "Grand Total")); gtr.appendChild(g1);
    var g2 = ce("td", "num"); g2.appendChild(ce("b", null, fmt(grand))); gtr.appendChild(g2); tb.appendChild(gtr);
    t.appendChild(tb); sc.appendChild(t); tw.appendChild(sc); return tw;
  }
  function coverageStatusTable(statusObj, grand) {
    var rowsData = COVERAGE_STATUSES.filter(function (s) { return statusObj[s]; }).map(function (s) { return { status: s, n: statusObj[s] }; });
    var tw = ce("div", "rms-tablewrap"), sc = ce("div", "rms-scroll"), t = ce("table", "rms-table");
    var thead = ce("thead"), htr = ce("tr"); htr.appendChild(ce("th", null, "VERIF_STATUS")); htr.appendChild(ce("th", "num", "Count of CSN")); thead.appendChild(htr); t.appendChild(thead);
    var tb = ce("tbody");
    rowsData.forEach(function (r) { var tr = ce("tr"); tr.appendChild(ce("td", null, r.status)); var td = ce("td", "num"); td.appendChild(ce("span", "pill", fmt(r.n))); tr.appendChild(td); tb.appendChild(tr); });
    var gtr = ce("tr", "rms-totalrow"); gtr.setAttribute("style", "border-top:2px solid var(--border-2);font-weight:600");
    var g1 = ce("td"); g1.appendChild(ce("b", null, "Grand Total")); gtr.appendChild(g1);
    var g2 = ce("td", "num"); g2.appendChild(ce("b", null, fmt(grand))); gtr.appendChild(g2); tb.appendChild(gtr);
    t.appendChild(tb); sc.appendChild(t); tw.appendChild(sc); return tw;
  }
  // Month-over-month for coverage: a simple vertical-bar chart (Count of CSN per month) matching the
  // shared .chart/.bar look, and a Check-in User × month grid. Both live at controller level.
  function covMonthBar(items, flagIdx) {
    var NS = "http://www.w3.org/2000/svg";
    function s(tag, attrs) { var e = document.createElementNS(NS, tag); if (attrs) for (var k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]); return e; }
    var wrap = ce("div", "chart");
    if (!items.length) { wrap.appendChild(ce("div", "empty-note", "No dated records")); return wrap; }
    var max = Math.max.apply(null, items.map(function (i) { return i.value; })) || 1;
    var W = 640, H = 190, padB = 26, padT = 16, padX = 6, n = items.length, slot = (W - padX * 2) / n, bw = Math.min(46, slot * 0.62);
    var svg = s("svg", { viewBox: "0 0 " + W + " " + H, role: "img" }), base = H - padB;
    svg.appendChild(s("line", { x1: 0, y1: base + 0.5, x2: W, y2: base + 0.5, stroke: "var(--border)" }));
    items.forEach(function (it, i) {
      var cx = padX + slot * i + slot / 2, bh = (it.value / max) * (base - padT), x = cx - bw / 2, y = base - bh;
      svg.appendChild(s("rect", { x: x, y: y, width: bw, height: Math.max(bh, it.value ? 2 : 0), rx: 4, "class": "bar" + (flagIdx === i ? " flag" : "") }));
      if (it.value) { var tl = s("text", { x: cx, y: y - 5, "text-anchor": "middle", "class": "barlabel" }); tl.textContent = fmt(it.value); svg.appendChild(tl); }
      var tc = s("text", { x: cx, y: base + 16, "text-anchor": "middle", "class": "barcat", "font-size": "11" }); tc.textContent = it.label; svg.appendChild(tc);
    });
    wrap.appendChild(svg); return wrap;
  }
  function covMonthMatrix(empMonth, mkeys, onPick) {
    var rows = Object.keys(empMonth).map(function (nm) { var m = empMonth[nm], total = 0; mkeys.forEach(function (k) { total += (m[k] || 0); }); return { name: nm, m: m, total: total }; })
      .sort(function (a, b) { return b.total - a.total || String(a.name).localeCompare(String(b.name)); });
    var tw = ce("div", "rms-tablewrap"), sc = ce("div", "rms-scroll"), t = ce("table", "rms-table");
    var thead = ce("thead"), htr = ce("tr"); htr.appendChild(ce("th", null, "Check-in User"));
    mkeys.forEach(function (k) { htr.appendChild(ce("th", "num", periodLabelG(k))); });
    htr.appendChild(ce("th", "num", "Total")); thead.appendChild(htr); t.appendChild(thead);
    var tb = ce("tbody"), colTot = {}, grand = 0; mkeys.forEach(function (k) { colTot[k] = 0; });
    rows.forEach(function (r) {
      var tr = ce("tr"), td0 = ce("td");
      if (onPick) { var b = ce("button", "rms-linkbtn", r.name); b.addEventListener("click", function () { onPick(r.name); }); td0.appendChild(b); } else td0.textContent = r.name;
      tr.appendChild(td0);
      mkeys.forEach(function (k) { var v = r.m[k] || 0; colTot[k] += v; tr.appendChild(ce("td", "num", fmt(v))); });
      var tt = ce("td", "num"); tt.appendChild(ce("b", null, fmt(r.total))); tr.appendChild(tt); grand += r.total; tb.appendChild(tr);
    });
    if (rows.length) { var gtr = ce("tr", "rms-totalrow"); gtr.setAttribute("style", "border-top:2px solid var(--border-2);font-weight:600");
      var g0 = ce("td"); g0.appendChild(ce("b", null, "TOTALS")); gtr.appendChild(g0);
      mkeys.forEach(function (k) { var gc = ce("td", "num"); gc.appendChild(ce("b", null, fmt(colTot[k]))); gtr.appendChild(gc); });
      var gt = ce("td", "num"); gt.appendChild(ce("b", null, fmt(grand))); gtr.appendChild(gt); tb.appendChild(gtr); }
    t.appendChild(tb); sc.appendChild(t); tw.appendChild(sc); return tw;
  }
  function coverageBuild(container, interactive, restrict, hidePickers) {
    container.classList.add("rms");
    var agg = coverageAggregate(interactive ? null : (restrict || null), interactive);
    var selLabel = interactive ? ((covState.dept ? " • " + covState.dept : "") + (covState.emp ? " • " + covState.emp : "")) : "";
    var wrap = ce("div", "rms-wrap"), head = ce("div", "rms-head2"), tbox = ce("div", "rms-titlebox");
    tbox.appendChild(ce("h1", null, "Coverage Accuracy" + (agg.month ? " — " + agg.month : "")));
    tbox.appendChild(ce("div", "rms-subtitle", "Count of CSN by verification status" + (agg.haveFilter ? (rawRecords.length && copayRecords.length ? " • employees on the bypass + copay sheets" : rawRecords.length ? " • employees on the Bypassed Warnings tab" : " • check-in users on the copay sheet") : "")));
    head.appendChild(tbox);
    if (interactive && !hidePickers) {
      var right = ce("div", "rms-headright");
      function picker(labelText, sel) { var p = ce("div", "rms-picker"); p.appendChild(ce("label", null, labelText)); p.appendChild(sel); return p; }
      if (agg.depts.length) {
        var dsel = ce("select", "rms-select");
        var o0 = ce("option", null, "▸ All departments"); o0.value = ""; dsel.appendChild(o0);
        agg.depts.forEach(function (d) { var o = ce("option", null, d.length > 42 ? d.slice(0, 41) + "…" : d); o.value = d; if (d === covState.dept) o.selected = true; dsel.appendChild(o); });
        dsel.addEventListener("change", function () { covState.dept = dsel.value; covState.emp = ""; renderCoverage(); });
        right.appendChild(picker("Department", dsel));
      }
      if (agg.cols.user && agg.haveFilter && agg.emps.length) {
        var esel = ce("select", "rms-select");
        var e0 = ce("option", null, "▸ All employees"); e0.value = ""; esel.appendChild(e0);
        agg.emps.forEach(function (e2) { var o = ce("option", null, e2.length > 34 ? e2.slice(0, 33) + "…" : e2); o.value = e2; if (e2 === covState.emp) o.selected = true; esel.appendChild(o); });
        esel.addEventListener("change", function () { covState.emp = esel.value; renderCoverage(); });
        right.appendChild(picker("Employee", esel));
      }
      head.appendChild(right);
    }
    wrap.appendChild(head);
    var body = ce("div");
    var k = ce("div", "rms-kpis");
    function kpi(lbl, val, sub) { var d = ce("div", "kpi"); d.appendChild(ce("div", "k-label", lbl)); d.appendChild(ce("div", "k-val", val)); d.appendChild(ce("div", "k-sub", sub || "")); return d; }
    k.appendChild(kpi("Encounters (Count of CSN)", fmt(agg.grand), "across " + Object.keys(agg.statusObj).length + " statuses"));
    k.appendChild(agg.haveFilter ? kpi("Users matched", fmt(Object.keys(agg.matched).length), rawRecords.length ? "on Bypassed Warnings tab" : "on copay sheet") : kpi("User filter", "off", "load copay or bypass to filter"));
    k.appendChild(kpi("Clinics", fmt(Object.keys(agg.deptCounts).length), "by bypass / copay lookup"));
    body.appendChild(k);
    /* Period over period — the "All … (combined)" view only, and only when the coverage records span
       ≥2 periods. Selecting a single period up top takes the trend sections off the page. Windowed to
       the most recent 13 at week granularity so the grid stays readable. */
    var covMkAll = Object.keys(agg.byMonth || {}).sort();
    var covMk = covMkAll.slice(gState.gran === "week" ? -13 : -24);
    var covDropped = covMkAll.length - covMk.length;
    if (!gPeriodActive() && covMk.length >= 2) {
      body.appendChild(ce("div", "rms-section", gState.gran === "week" ? "Week over week" : "Month over month"));
      var peak = 0; covMk.forEach(function (kk, i) { if (agg.byMonth[kk] > agg.byMonth[covMk[peak]]) peak = i; });
      var citems = covMk.map(function (kk) { var pp = kk.split("-");
        return { label: isWeekKeyG(kk) ? MON_G[+pp[1]] + " " + (+pp[2]) : MON_G[+pp[1]], value: agg.byMonth[kk] }; });
      var cf = covMk[0], cl = covMk[covMk.length - 1];
      var cardEl = ce("div", "card");
      cardEl.appendChild(ce("h3", null, "Encounters (Count of CSN) by " + periodNounG() + " — whole clinic"));
      var cnoteTxt = periodLabelG(cf) + ": " + fmt(agg.byMonth[cf]) + " → " + periodLabelG(cl) + ": " + fmt(agg.byMonth[cl]) + " encounters";
      if (covDropped > 0) cnoteTxt += "  • latest " + covMk.length + " " + periodNounG(true) + " (" + covDropped + " earlier not shown)";
      cardEl.appendChild(ce("p", "c-note", cnoteTxt));
      cardEl.appendChild(covMonthBar(citems, peak));
      var grid = ce("div", "rms-grid one"); grid.appendChild(cardEl); body.appendChild(grid);
      if (agg.cols.user && Object.keys(agg.empMonth || {}).length) {
        body.appendChild(ce("div", "rms-section", "Coverage encounters by Check-in User per " + periodNounG()));
        body.appendChild(covMonthMatrix(agg.empMonth, covMk, null));
      }
    }
    if (!agg.cols.status) { body.appendChild(ce("div", "empty-note", "No VERIF_STATUS column found in the coverage file.")); wrap.appendChild(body); container.appendChild(wrap); return; }
    if (Object.keys(agg.deptCounts).length) {
      body.appendChild(ce("div", "rms-section", "Coverage by clinic" + selLabel));
      body.appendChild(coverageDeptTable(agg.deptCounts, agg.grand, null));
    }
    // Total coverage errors per employee, with the clinic's total on the Grand Total row. The combined
    // view carries these totals in the Check-in User x month grid above, so this is the single-month page's
    // equivalent — same numbers the status table below is built from.
    if (gPeriodActive() && agg.cols.user && Object.keys(agg.empErrors || {}).length) {
      body.appendChild(ce("div", "rms-section", "Coverage errors by employee" + selLabel));
      body.appendChild(coverageDeptTable(agg.empErrors, agg.grand, null, "Check-in User"));
    }
    body.appendChild(ce("div", "rms-section", "Coverage errors by verification status" + selLabel));
    body.appendChild(coverageStatusTable(agg.statusObj, agg.grand));
    body.appendChild(ce("div", "rms-subtitle", "Kept the spreadsheet's filters: " + COVERAGE_STATUSES.length + " selected VERIF_STATUS values" +
      (agg.haveFilter ? "; users limited to " + (rawRecords.length && copayRecords.length ? "the employees on the bypass + copay sheets" : rawRecords.length ? "the employees on the Bypassed Warnings tab" : "the copay sheet's check-in users") + " (" + fmt(Object.keys(agg.excluded).length) + " other check-in users excluded)" : "; no copay/bypass loaded, so all users are shown") +
      (selLabel ? "; filtered to" + selLabel : "") + "."));
    wrap.appendChild(body); container.appendChild(wrap);
  }
  function renderCoverage() { covState.dept = gState.dept; covState.emp = gState.emp; dashboard.innerHTML = ""; coverageBuild(dashboard, true, null, true); }
  function coverageSnapshotHTML(restrict) { if (!coverageRecords.length) return null; var d = document.createElement("div"); coverageBuild(d, false, restrict); return d.innerHTML; }
  // Build a restrict-set of employee name-keys from a list of display names (bypass/copay format).
  function coverageKeysFor(names) { if (!names || !names.length) return null; var set = {}; names.forEach(function (nm) { var k = covNameKey(nm); if (k && k !== "|") set[k] = 1; }); return set; }

  /* ================= Historical data archive (one rolling Excel across months) =================
     Import each month's export(s) together with the prior "PAS Historical Data" .xlsx and the app
     merges the new month into the running history (duplicate rows collapse), then re-downloads an
     updated archive that now spans every month loaded. The archive is a plain .xlsx with one sheet
     per data type (Bypassed Warnings / Copays / Coverage / Registrations), each carrying a Period
     column, plus a PAS_History manifest sheet that marks the file as an archive. ------------------ */
  var HIST_MANIFEST = "PAS_History";          // manifest sheet name (identifies the file as an archive)
  var HIST_SENTINEL = "PAS_QUALITY_HISTORY";   // sentinel token stored in the manifest header
  var HIST_PERIOD_COL = "Period";              // per-row month column added to each data sheet
  var HIST_SHEETS = { bypass: "Bypassed Warnings", copays: "Copays", coverage: "Coverage", registrations: "Registrations" };

  function periodFromDate(v) {
    if (v == null) return "";
    var s = String(v).trim(); if (!s) return "";
    var m = s.match(/^(\d{4})-(\d{2})/); if (m) return m[1] + "-" + m[2];   // ISO-ish, avoid TZ drift
    var d = new Date(s); if (isNaN(d)) return "";
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  }
  function bypassDateHeader() { return colMap && (colMap.date || colMap.encDate); }
  function copayColsForHist() { return (window.RMSCopays && copayHeaders.length) ? RMSCopays.detectCopayCols(copayHeaders) : {}; }

  // Collapse rows that are identical across their original data columns (ignores the injected
  // Period / __PAS_* keys). Re-importing the same archive, or an archive alongside the raw file it
  // was built from, therefore counts each row once.
  function histSig(r, hdrs) { var s = ""; for (var j = 0; j < hdrs.length; j++) { var v = r[hdrs[j]]; s += (v == null ? "" : String(v)) + ""; } return s; }
  function dedupByHeaders(records, hdrs) {
    if (!hdrs || !hdrs.length) return records;
    var seen = Object.create(null), out = [];
    for (var i = 0; i < records.length; i++) { var k = histSig(records[i], hdrs); if (seen[k]) continue; seen[k] = 1; out.push(records[i]); }
    return out;
  }
  function dedupRegEntries(entries) {
    var seen = Object.create(null), out = [];
    entries.forEach(function (e) { var k = (e.__period || "") + "" + String(e.staff == null ? "" : e.staff).toLowerCase().trim(); if (seen[k]) return; seen[k] = 1; out.push(e); });
    return out;
  }

  // A clean sheet (header row = row 0) -> {headers, rows}. Used only for archives we generate.
  function matrixToTable(m) {
    if (!m || !m.length) return { headers: [], rows: [] };
    var hs = (m[0] || []).map(function (v) { return String(v == null ? "" : v).trim(); });
    var rows = [];
    for (var r = 1; r < m.length; r++) {
      var row = m[r] || [], obj = {}, any = false;
      for (var c = 0; c < hs.length; c++) { if (!hs[c]) continue; var v = row[c]; v = v == null ? "" : String(v).trim(); obj[hs[c]] = v; if (v) any = true; }
      if (any) rows.push(obj);
    }
    return { headers: hs.filter(Boolean), rows: rows };
  }
  // If the dropped file is one of our historical archives, return its workbook; else null.
  async function readHistoryWorkbook(file) {
    var buf, bytes;
    try { buf = await file.arrayBuffer(); bytes = new Uint8Array(buf); } catch (e) { return null; }
    if (isOLE(bytes)) return null;                    // encrypted / legacy .xls -> not one of ours
    var wb;
    try { wb = XLSX.read(isZip(bytes) ? bytes : new TextDecoder("utf-8").decode(bytes), { type: isZip(bytes) ? "array" : "string", cellDates: false, raw: false }); }
    catch (e) { return null; }
    if (!wb || !wb.SheetNames || wb.SheetNames.indexOf(HIST_MANIFEST) < 0) return null;
    var mm = matrixToTable(sheetToMatrix(wb.Sheets[HIST_MANIFEST]));
    if (mm.headers.indexOf(HIST_SENTINEL) < 0) return null;
    return wb;
  }
  // Merge every data sheet from an archive workbook into the live buckets.
  function mergeHistoryWorkbook(wb) {
    var summary = { bypass: 0, copays: 0, coverage: 0, reg: 0 };
    function tableFor(name) { return wb.SheetNames.indexOf(name) < 0 ? null : matrixToTable(sheetToMatrix(wb.Sheets[name])); }
    var b = tableFor(HIST_SHEETS.bypass);
    if (b && b.rows.length) { b.headers.forEach(function (h) { if (h !== HIST_PERIOD_COL && headers.indexOf(h) < 0) headers.push(h); }); rawRecords = rawRecords.concat(b.rows); summary.bypass = b.rows.length; }
    var c = tableFor(HIST_SHEETS.copays);
    if (c && c.rows.length) { c.headers.forEach(function (h) { if (h !== HIST_PERIOD_COL && copayHeaders.indexOf(h) < 0) copayHeaders.push(h); }); copayRecords = copayRecords.concat(c.rows); summary.copays = c.rows.length; }
    var v = tableFor(HIST_SHEETS.coverage);
    if (v && v.rows.length) { v.headers.forEach(function (h) { if (h !== HIST_PERIOD_COL && coverageHeaders.indexOf(h) < 0) coverageHeaders.push(h); }); coverageRecords = coverageRecords.concat(v.rows); summary.coverage = v.rows.length; }
    var g = tableFor(HIST_SHEETS.registrations);
    if (g && g.rows.length) {
      g.rows.forEach(function (r) {
        var staff = r.Staff != null ? r.Staff : (r.staff != null ? r.staff : r[g.headers[0]]);
        staff = String(staff == null ? "" : staff).trim(); if (!staff) return;
        var raw = String(r.Count != null ? r.Count : (r.count != null ? r.count : "")).replace(/[,\s$]/g, "");
        var n = raw === "" ? 0 : +raw; if (isNaN(n)) n = 0;
        var per = r[HIST_PERIOD_COL] || r.Period || "";
        regEntries.push({ staff: staff, count: n, __period: per });
        if (per) regMonth = per;
        summary.reg++;
      });
    }
    return summary;
  }
  function applyHistoryDedup() {
    rawRecords = dedupByHeaders(rawRecords, headers);
    copayRecords = dedupByHeaders(copayRecords, copayHeaders);
    coverageRecords = dedupByHeaders(coverageRecords, coverageHeaders);
    normalizeRegPeriods();   // one month can arrive as "MAR" and as "2026-03" — same period, one entry
    regEntries = dedupRegEntries(regEntries);
  }

  /* ---- view-layer period keys ----
     `periodFromDate` stays MONTH-only on purpose: the historical archive's Period column and the
     executive rollup's clinic -> month partitions are file formats other tools already read, and a
     week key in either would break the Executive View's merge (it de-duplicates a clinic by replacing
     its months, so mixed month/week keys would double-count). Everything on screen goes through
     `gPeriodFrom` instead, which honours the Week/Month setting. Week keys are the Monday as
     "YYYY-MM-DD" — sortable, never confusable with "YYYY-MM", and free of ISO week-year edge cases. */
  function gWeekKeyFromDate(v) {
    if (v == null || v === "") return "";
    var s = String(v).trim(), d = null, m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) d = new Date(+m[1], +m[2] - 1, +m[3]);
    else { m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
      if (m) { var y = +m[3]; if (y < 100) y += 2000; d = new Date(y, +m[1] - 1, +m[2]); }
      else { var dd = new Date(s); if (!isNaN(dd)) d = new Date(dd.getFullYear(), dd.getMonth(), dd.getDate()); } }
    if (!d || isNaN(d)) return "";
    d.setDate(d.getDate() - ((d.getDay() + 6) % 7));          // back up to Monday
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }
  function gPeriodFrom(v) { return gState.gran === "week" ? gWeekKeyFromDate(v) : periodFromDate(v); }
  function isWeekKeyG(p) { return /^\d{4}-\d{2}-\d{2}$/.test(String(p)); }
  function isPeriodKeyG(p) { return /^\d{4}-\d{2}$/.test(String(p)) || isWeekKeyG(p); }
  function periodNounG(plural) { return gState.gran === "week" ? (plural ? "weeks" : "week") : (plural ? "months" : "month"); }

  // The sorted set of periods present in the dated record sets (bypass / copay / coverage).
  function dataMonths() {
    var set = {};
    function addFrom(records, dateH) { if (!dateH) return; records.forEach(function (r) { var p = gPeriodFrom(r[dateH]); if (p) set[p] = 1; }); }
    addFrom(rawRecords, bypassDateHeader());
    addFrom(copayRecords, copayColsForHist().date);
    addFrom(coverageRecords, coverageHeaders.length ? coverageCols().date : null);
    return Object.keys(set).sort();
  }
  /* The sorted set of periods present across all loaded data. Registration entries carry a month and
     nothing finer, so they only contribute a bucket at month granularity — at week granularity the
     period list comes from the dated rows alone (registrations still get their own monthly section). */
  function historyMonths() {
    var set = {};
    dataMonths().forEach(function (p) { set[p] = 1; });
    if (gState.gran !== "week") regEntries.forEach(function (e) { if (e.__period && /^\d{4}-\d{2}$/.test(e.__period)) set[e.__period] = 1; });
    return Object.keys(set).sort();
  }
  // Every MONTH present, ignoring the Week/Month setting — for the archive and the executive rollup.
  function historyMonthsMonthly() {
    var set = {};
    function addFrom(records, dateH) { if (!dateH) return; records.forEach(function (r) { var p = periodFromDate(r[dateH]); if (p) set[p] = 1; }); }
    addFrom(rawRecords, bypassDateHeader());
    addFrom(copayRecords, copayColsForHist().date);
    addFrom(coverageRecords, coverageHeaders.length ? coverageCols().date : null);
    regEntries.forEach(function (e) { if (e.__period && /^\d{4}-\d{2}$/.test(e.__period)) set[e.__period] = 1; });
    return Object.keys(set).sort();
  }
  /* ---- registration months ----
     A registrations file names its month in the count column's header ("MAR", "Mar 2026", "3/2026")
     and covers that month only. Resolve that text to a YYYY-MM period so the Month filter, the
     archive sheet, and the per-registration tables all agree on which month the counts belong to —
     without it, March registrations get divided into every loaded month's warnings. */
  var REG_MON_N = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function pad2(n) { return String(n).length < 2 ? "0" + n : String(n); }
  function regYearFor(mo) {
    var ms = dataMonths(), hit = ms.filter(function (p) { return +p.split("-")[1] === mo; });
    if (hit.length) return hit[hit.length - 1].split("-")[0];          // same month in the loaded data
    if (ms.length) return ms[ms.length - 1].split("-")[0];             // otherwise the newest loaded year
    return "";
  }
  function regPeriodKey(text) {
    var s = String(text == null ? "" : text).trim(); if (!s) return "";
    if (/^\d{4}-\d{1,2}$/.test(s)) { var a = s.split("-"); return a[0] + "-" + pad2(+a[1]); }
    var m = s.match(/^(\d{4})[\/.](\d{1,2})$/); if (m) return m[1] + "-" + pad2(+m[2]);
    m = s.match(/^(\d{1,2})[\/.-](\d{4})$/); if (m) return m[2] + "-" + pad2(+m[1]);
    // a spreadsheet reader turns a "Aug 2026" header cell into a date ("8/1/26") — take month + year from it
    m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/); if (m) { var yy = +m[3]; if (yy < 100) yy += 2000; return yy + "-" + pad2(+m[1]); }
    var n = s.toLowerCase().match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/); if (!n) return "";
    var mo = REG_MON_N[n[1]], y = (s.match(/(?:19|20)\d{2}/) || [])[0] || regYearFor(mo);
    return y ? y + "-" + pad2(mo) : "";
  }
  // Rewrite every registration entry's month text to its YYYY-MM period (leaving entries whose month
  // can't be resolved alone). Cheap and idempotent, so it can run before any read of regEntries.
  function normalizeRegPeriods() {
    regEntries.forEach(function (e) { var k = regPeriodKey(e.__period); if (k) e.__period = k; });
    var rk = regPeriodKey(regMonth); if (rk) regMonth = periodLabelG(rk);
  }
  // ---- Global Month filter helpers (drives every tab from the top filter bar) ----
  var MON_G = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function gPeriodActive() { return !!(gState.period && isPeriodKeyG(gState.period)); }
  function periodLabelG(p) {
    if (!p || !isPeriodKeyG(p)) return gState.gran === "week" ? "All weeks" : "All months";
    var s = String(p).split("-");
    return isWeekKeyG(p) ? "Wk " + MON_G[+s[1]] + " " + (+s[2]) : MON_G[+s[1]] + " " + s[0];
  }
  // Long form for a title or a sentence ("Week of Apr 6, 2026").
  function periodLabelLongG(p) {
    if (!p || !isPeriodKeyG(p)) return gState.gran === "week" ? "All weeks" : "All months";
    var s = String(p).split("-");
    return isWeekKeyG(p) ? "Week of " + MON_G[+s[1]] + " " + (+s[2]) + ", " + s[0] : MON_G[+s[1]] + " " + s[0];
  }
  // Narrow a record set to the globally-selected period (no-op when "All …" is selected or no date column).
  function scopeByPeriod(records, dateH) {
    if (!gPeriodActive() || !dateH) return records;
    return records.filter(function (r) { return gPeriodFrom(r[dateH]) === gState.period; });
  }
  // Combined per-period row counts across all loaded data (for the Period dropdown labels).
  function periodRowCounts() {
    var m = {};
    function addF(records, dateH) { if (!dateH) return; records.forEach(function (r) { var p = gPeriodFrom(r[dateH]); if (p) m[p] = (m[p] || 0) + 1; }); }
    addF(rawRecords, bypassDateHeader());
    addF(copayRecords, copayColsForHist().date);
    addF(coverageRecords, coverageHeaders.length ? coverageCols().date : null);
    return m;
  }
  function buildHistoryWorkbook() {
    normalizeRegPeriods();   // archive the registrations' resolved YYYY-MM period, not the file's header text
    var wb = XLSX.utils.book_new(), months = historyMonthsMonthly();   // the archive is always month-keyed
    var manifest = [
      [HIST_SENTINEL, "Generated", "Months", "Bypass rows", "Copay rows", "Coverage rows", "Registration rows"],
      ["PAS Quality historical archive — re-import this file next month to keep building history. Do not edit by hand.",
       nowLabel(), months.join(", "), rawRecords.length, copayRecords.length, coverageRecords.length, regEntries.length]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(manifest), HIST_MANIFEST);
    function addSheet(name, records, hdrs, dateH) {
      if (!records.length || !hdrs.length) return;
      var aoa = [hdrs.concat([HIST_PERIOD_COL])];
      records.forEach(function (r) {
        var row = hdrs.map(function (h) { var val = r[h]; return val == null ? "" : val; });
        row.push(dateH ? periodFromDate(r[dateH]) : (r[HIST_PERIOD_COL] || ""));
        aoa.push(row);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    }
    addSheet(HIST_SHEETS.bypass, rawRecords, headers, bypassDateHeader());
    addSheet(HIST_SHEETS.copays, copayRecords, copayHeaders, copayColsForHist().date);
    addSheet(HIST_SHEETS.coverage, coverageRecords, coverageHeaders, coverageHeaders.length ? coverageCols().date : null);
    if (regEntries.length) {
      var raoa = [["Staff", "Count", HIST_PERIOD_COL]];
      regEntries.forEach(function (e) { raoa.push([e.staff, e.count, e.__period || regMonth || ""]); });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(raoa), HIST_SHEETS.registrations);
    }
    return { wb: wb, months: months };
  }
  function historyFileName(months) {
    var ym = (months || []).filter(function (m) { return /^\d{4}-\d{2}$/.test(m); });
    var span = ym.length ? (ym[0] === ym[ym.length - 1] ? ym[0] : ym[0] + " to " + ym[ym.length - 1]) : "current";
    return "PAS Historical Data (" + span + ").xlsx";
  }
  function hasAnyData() { return !!(rawRecords.length || copayRecords.length || coverageRecords.length || regEntries.length); }
  function downloadHistory(auto) {
    if (!hasAnyData()) { if (!auto) setHeaderXStatus("No data loaded yet — import your export files first."); return; }
    var built = buildHistoryWorkbook(), fname = historyFileName(built.months);
    XLSX.writeFile(built.wb, fname);
    var mn = built.months.length, tag = mn ? (mn + " month" + (mn === 1 ? "" : "s")) : "current data";
    setHeaderXStatus((auto ? "Saved historical data → " : "Downloaded ") + fname + " (" + tag + ").");
    return fname;
  }

  async function handleFiles(list) {
    var files = Array.from(list); if (!files.length) return;
    var pw = pwInput.value, added = 0, errs = [], ruleNote = "", regNote = "", covNote = "", histNote = "", histImported = false;
    for (var i = 0; i < files.length; i++) {
      var f = files[i]; setStatus("Reading " + f.name + " …", "busy");
      try {
        var histWb = null; try { histWb = await readHistoryWorkbook(f); } catch (e) { histWb = null; }
        if (histWb) {
          var sm = mergeHistoryWorkbook(histWb); histImported = true;
          added += sm.bypass + sm.copays + sm.coverage + sm.reg;
          histNote = "Historical archive loaded from " + f.name + ": " + fmt(sm.bypass) + " bypass · " + fmt(sm.copays) + " copay · " + fmt(sm.coverage) + " coverage · " + fmt(sm.reg) + " registration rows.";
          continue;
        }
        var t = await fileToTable(f, pw, function (p) { setStatus("Decrypting " + f.name + " … " + Math.round(p * 100) + "%", "busy"); });
        if (t.rules) {
          setRules(t.rules); ruleFileName = f.name;
          ruleNote = "Rules loaded from " + f.name + ": " + t.rules.include.length + " include / " + t.rules.exclude.length + " exclude.";
        } else if (window.RMSCopays && RMSCopays.looksLikeCopays(t.headers)) {
          t.headers.forEach(function (h) { if (copayHeaders.indexOf(h) < 0) copayHeaders.push(h); });
          copayRecords = copayRecords.concat(t.rows); copayFileNames.push(f.name + " (" + t.rows.length + ")"); added += t.rows.length;
        } else if (looksLikeRegistrations(t.headers)) {
          var pr = parseRegistrations(t.headers, t.rows);
          pr.entries.forEach(function (e) { e.__period = pr.month || ""; });
          regEntries = regEntries.concat(pr.entries);
          if (pr.month) regMonth = pr.month;
          regFileNames.push(f.name + " (" + pr.entries.length + ")");
          regNote = "Registrations loaded from " + f.name + ": " + pr.entries.length + " staff" + (pr.month ? " (" + pr.month + ")" : "") + ".";
        } else if (looksLikeCoverage(t.headers)) {
          t.headers.forEach(function (h) { if (coverageHeaders.indexOf(h) < 0) coverageHeaders.push(h); });
          coverageRecords = coverageRecords.concat(t.rows); coverageFileNames.push(f.name + " (" + t.rows.length + ")");
          covNote = "Coverage data loaded from " + f.name + ": " + fmt(t.rows.length) + " encounters.";
        } else {
          t.headers.forEach(function (h) { if (headers.indexOf(h) < 0) headers.push(h); });
          rawRecords = rawRecords.concat(t.rows); fileNames.push(f.name + " (" + t.rows.length + ")"); added += t.rows.length;
        }
      } catch (e) { errs.push(f.name + ": " + (e && e.message || e)); }
    }
    normalizeRegPeriods();                   // resolve each registrations file's month against the loaded data
    if (histImported) applyHistoryDedup();   // collapse rows shared between the archive and this month's files
    var notes = (ruleNote ? " " + ruleNote : "") + (regNote ? " " + regNote : "") + (covNote ? " " + covNote : "") + (histNote ? " " + histNote : "");
    if (errs.length) setStatus("Loaded " + fmt(added) + " new records. Issues → " + errs.join(" | ") + notes, "warn");
    else setStatus("Loaded " + fmt(added) + " records from " + files.length + " file" + (files.length === 1 ? "" : "s") + ". Total: " + fmt(rawRecords.length) + "." + notes, "ok");
    if (!rawRecords.length && copayRecords.length) activeModule = "copays";
    if (!rawRecords.length && !copayRecords.length && coverageRecords.length) activeModule = "coverage";
    periodInitialized = false;   // a fresh import re-opens on the newest month now loaded
    refreshConfig(); render();
    // After every import, roll the running history forward and re-download the updated archive.
    if (added > 0 && hasAnyData()) {
      try {
        var fn = downloadHistory(true);
        if (fn) setStatus(status.textContent + "  ⬇ Historical data saved to " + fn + (histImported ? "" : " — re-import it next month to keep building history."), errs.length ? "warn" : "ok");
      } catch (e) {}
    }
  }

  function refreshConfig() {
    if (!headers.length) { config.style.display = "none"; return; }
    colMap = detectRoles(headers);
    if (!deptKey || headers.indexOf(deptKey) < 0) deptKey = colMap.department || colMap.loginDept || firstNonPeople();
    if (!empKey || headers.indexOf(empKey) < 0) empKey = colMap.checkinUser || colMap.user || headers[0];
    fillSelect(deptSel, headers, deptKey);
    fillSelect(empSel, headers, empKey);
    fileList.textContent = fileNames.length ? "Files: " + fileNames.join(" · ") : "";
    config.style.display = "";
  }
  function firstNonPeople() { for (var i = 0; i < headers.length; i++) { var n = norm(headers[i]); if (n.indexOf("user") < 0 && n.indexOf("patient") < 0) return headers[i]; } return headers[0]; }
  function fillSelect(sel, opts, cur) { sel.innerHTML = ""; opts.forEach(function (h) { var o = document.createElement("option"); o.value = h; o.textContent = h; if (h === cur) o.selected = true; sel.appendChild(o); }); }

  function nowLabel() { return new Date().toLocaleString("en-US", { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); }
  function moduleList() {
    var mods = [];
    if (rawRecords.length) mods.push(["bypass", "Bypassed Warnings"]);
    if (copayRecords.length) mods.push(["copays", "Copays"]);
    if (coverageRecords.length) mods.push(["coverage", "Coverage Accuracy"]);
    return mods;
  }
  function refreshTabs() {
    if (!moduletabs) return;
    var mods = moduleList();
    if (mods.length < 2) { moduletabs.style.display = "none"; moduletabs.innerHTML = ""; return; }
    moduletabs.style.display = ""; moduletabs.innerHTML = "";
    mods.forEach(function (t) {
      var b = document.createElement("button");
      b.className = "module-tab" + (activeModule === t[0] ? " active" : "");
      b.type = "button"; b.textContent = t[1];
      b.addEventListener("click", function () { if (activeModule !== t[0]) { activeModule = t[0]; render(); try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (e) {} } });
      moduletabs.appendChild(b);
    });
  }
  /* ---------- global filter bar (Department + Employee, shown above the tabs) ---------- */
  function trimVal(v) { var s = String(v == null ? "" : v).replace(/\s+/g, " ").trim(); return s; }
  function distinctSorted(arr) { var s = {}; arr.forEach(function (v) { if (v !== "" && v != null) s[v] = 1; }); return Object.keys(s).sort(function (a, b) { return String(a).localeCompare(String(b)); }); }
  function bypassDeptHeader() { return (deptSel && deptSel.value) || deptKey || (colMap && (colMap.department || colMap.loginDept)) || null; }
  // Names shown to users: keep letters, spaces and the Last, First comma — drop IDs, brackets, digits, other punctuation.
  function cleanDisplayName(s) { s = String(s == null ? "" : s); if (s === "(blank)") return s;
    var c = s.replace(/[^A-Za-z, ]+/g, " ").replace(/\s+/g, " ").replace(/\s*,\s*/g, ", ").replace(/^[,\s]+|[,\s]+$/g, "").trim();
    return c === "" ? "(blank)" : c; }
  var PAS_EMP = "__PAS_EMP__";   // injected cleaned employee-name column used for grouping + display
  function bypassEmpOf(r) { if (!empKey) return "(blank)"; return cleanDisplayName(r[empKey]); }
  function copayColsNow() { return (window.RMSCopays && copayHeaders.length) ? RMSCopays.detectCopayCols(copayHeaders) : {}; }
  function copayDeptOf(r, ccm) { return ccm.dept ? trimVal(r[ccm.dept]) : ""; }
  function copayEmpOf(r, ccm) { return cleanDisplayName(RMSCopays.cleanName(r[ccm.checkinUser])); }
  // Every employee across the bypass + copay sheets (display names). Used so the Coverage tab lists
  // everyone — even people with zero coverage errors — instead of only those who have errors.
  function allEmployeeDisplayNames() {
    var s = {};
    if (rawRecords.length && empKey) bypassDisplayRecords().forEach(function (r) { s[cleanDisplayName(r[empKey])] = 1; });
    if (copayRecords.length && window.RMSCopays) { var ccm = copayColsNow();
      if (ccm.checkinUser) copayRecords.forEach(function (r) { s[copayEmpOf(r, ccm)] = 1; }); }
    return distinctSorted(Object.keys(s));
  }
  // Map each employee (normalized name) to the department they sit in on the Bypassed Warnings sheet, so
  // the Copays tab can share the SAME department filter as Bypass (one clinic) and narrow copay by it.
  function bypassEmpDeptMap() {
    var m = {}, dH = bypassDeptHeader();
    if (!dH || !empKey) return m;
    bypassDisplayRecords().forEach(function (r) { var k = covNameKey(r[empKey]); if (k && !(k in m)) m[k] = trimVal(r[dH]) || "(blank)"; });
    return m;
  }
  // Employee display strings differ between tabs (Bypass keeps the raw "LAST, FIRST [id]"; Copays/Coverage
  // clean it to "LAST, FIRST"). Match the selected employee by a normalized name key so the filter carries
  // across tabs; remap it to THIS tab's spelling when found, and keep it as-is if the tab has no such person.
  function reconcileEmp(emps, target) {
    if (!target) return "";
    if (emps.indexOf(target) >= 0) return target;
    var tk = covNameKey(target);
    for (var i = 0; i < emps.length; i++) { if (covNameKey(emps[i]) === tk) return emps[i]; }
    return target;
  }

  function countMap(records, keyFn) { var m = {}; records.forEach(function (r) { var k = keyFn(r); if (k != null && k !== "") m[k] = (m[k] || 0) + 1; }); return m; }
  function renderGlobalFilters() {
    if (!globalfilters) return;
    var depts = [], emps = [], deptCounts = {}, empCounts = {};
    if (activeModule === "coverage") {
      // clinics + users come from the coverage aggregation (respects the current covState = gState)
      var agg = coverageAggregate(null, true);
      gState.dept = covState.dept; gState.emp = covState.emp;   // coverageAggregate may have sanitized invalid picks
      depts = agg.depts || [];
      // List EVERY employee (bypass + copay), not just those with coverage errors, so a person with
      // zero coverage rows can still be selected and simply shows zeros.
      emps = agg.cols.user ? allEmployeeDisplayNames() : [];
      deptCounts = agg.deptCounts || {}; empCounts = agg.empCounts || {};
    } else if (activeModule === "copays") {
      // Department filter mirrors the Bypassed Warnings tab (the one clinic) — use the BYPASS departments,
      // not the copay sheet's own. Employees come from the copay sheet, narrowed to the picked bypass dept.
      var ccm = copayColsNow();
      var dHc = bypassDeptHeader();
      var recsC = bypassDisplayRecords();
      depts = dHc ? distinctSorted(recsC.map(function (r) { var v = trimVal(r[dHc]); return v === "" ? "(blank)" : v; })) : [];
      if (gState.dept && depts.indexOf(gState.dept) < 0) gState.dept = "";
      var edm = bypassEmpDeptMap();
      // Employees: EVERYONE on the bypass + copay sheets (so a bypass employee with no copays still shows,
      // as (0), and selecting them shows zeros), narrowed to the picked bypass department.
      var allE = allEmployeeDisplayNames();
      emps = gState.dept ? allE.filter(function (nm) { return edm[covNameKey(nm)] === gState.dept; }) : allE;
      gState.emp = reconcileEmp(emps, gState.emp);
      var poolC = copayRecords.filter(function (r) { return !gState.dept || edm[covNameKey(r[ccm.checkinUser])] === gState.dept; });
      deptCounts = countMap(copayRecords, function (r) { return edm[covNameKey(r[ccm.checkinUser])]; });
      empCounts = countMap(poolC, function (r) { return copayEmpOf(r, ccm); });
    } else { // bypass
      var dH = bypassDeptHeader();
      var recs0 = bypassDisplayRecords();
      depts = dH ? distinctSorted(recs0.map(function (r) { var v = trimVal(r[dH]); return v === "" ? "(blank)" : v; })) : [];
      if (gState.dept && depts.indexOf(gState.dept) < 0) { gState.dept = ""; gState.emp = ""; }
      var poolB = recs0.filter(function (r) { if (!gState.dept || !dH) return true; var v = trimVal(r[dH]); if (v === "") v = "(blank)"; return v === gState.dept; });
      // Employees: EVERYONE on the bypass + copay sheets (so a copay-only employee still shows, as (0),
      // and selecting them shows zeros), narrowed to the picked department.
      var edmB = bypassEmpDeptMap();
      var allEB = allEmployeeDisplayNames();
      emps = gState.dept ? allEB.filter(function (nm) { return edmB[covNameKey(nm)] === gState.dept; }) : allEB;
      gState.emp = reconcileEmp(emps, gState.emp);
      if (dH) deptCounts = countMap(recs0, function (r) { var v = trimVal(r[dH]); return v === "" ? "(blank)" : v; });
      empCounts = countMap(poolB, bypassEmpOf);
    }
    buildFilterBar(depts, emps, deptCounts, empCounts);
  }
  function buildFilterBar(depts, emps, deptCounts, empCounts) {
    globalfilters.innerHTML = "";
    var inner = document.createElement("div"); inner.className = "gf-inner";
    var title = document.createElement("div"); title.className = "gf-title"; title.textContent = "Filters"; inner.appendChild(title);
    function field(labelText, options, current, allText, onChange, counts) {
      var f = document.createElement("div"); f.className = "gf-field";
      var lab = document.createElement("label"); lab.textContent = labelText; f.appendChild(lab);
      var sel = document.createElement("select");
      var o0 = document.createElement("option"); o0.value = ""; o0.textContent = allText; sel.appendChild(o0);
      options.forEach(function (v) {
        var nm = v.length > 44 ? v.slice(0, 43) + "…" : v;
        var label = counts ? nm + " (" + (counts[v] || 0).toLocaleString("en-US") + ")" : nm;
        var o = document.createElement("option"); o.value = v; o.textContent = label; if (v === current) o.selected = true; sel.appendChild(o);
      });
      sel.addEventListener("change", function () { onChange(sel.value); });
      f.appendChild(sel); return f;
    }
    /* Group by — Month (the default, and what every clinic past onboarding gets) or Week. Week exists
       for the first month with a new clinic, where a single monthly bar cannot show anyone whether
       the week's coaching landed. It is a view setting: it re-buckets what is on screen and is baked
       into any file exported while it is on, and it never reaches the historical archive or the
       executive rollup, both of which stay month-keyed. */
    var gf = document.createElement("div"); gf.className = "gf-field";
    var glab = document.createElement("label"); glab.textContent = "Group by"; gf.appendChild(glab);
    var gsel = document.createElement("select");
    [["month", "Month"], ["week", "Week (onboarding)"]].forEach(function (o) {
      var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1];
      if (o[0] === gState.gran) op.selected = true; gsel.appendChild(op);
    });
    gsel.addEventListener("change", function () { setGranularityG(gsel.value); render(); });
    gf.appendChild(gsel); inner.appendChild(gf);

    // Period filter — every period present across the loaded data. "All …" combines them; picking one
    // scopes EVERY tab (Bypassed Warnings, Copays, Coverage) to that period.
    var months = historyMonths();
    if (months.length > 1) {
      if (gState.period && months.indexOf(gState.period) < 0) gState.period = "";   // period no longer present after a re-import
      var pcounts = periodRowCounts();
      var mf = document.createElement("div"); mf.className = "gf-field";
      var mlab = document.createElement("label"); mlab.textContent = gState.gran === "week" ? "Week" : "Month"; mf.appendChild(mlab);
      var msel = document.createElement("select");
      var m0 = document.createElement("option"); m0.value = "";
      m0.textContent = "All " + periodNounG(true) + " (combined)"; msel.appendChild(m0);
      months.forEach(function (p) {
        var o = document.createElement("option"); o.value = p;
        o.textContent = periodLabelG(p) + (pcounts[p] ? " (" + pcounts[p].toLocaleString("en-US") + ")" : "");
        if (p === gState.period) o.selected = true; msel.appendChild(o);
      });
      msel.addEventListener("change", function () { gState.period = msel.value; render(); });
      mf.appendChild(msel); inner.appendChild(mf);
    } else {
      gState.period = "";
    }
    inner.appendChild(field("Department", depts, gState.dept, "All Departments", function (v) { gState.dept = v; gState.emp = ""; render(); }, deptCounts));
    inner.appendChild(field("Employee", emps, gState.emp, "All Employees", function (v) { gState.emp = v; render(); }, empCounts));
    var spacer = document.createElement("div"); spacer.className = "gf-spacer";
    if (gState.dept || gState.emp || gState.period) {
      var cb = document.createElement("button"); cb.className = "gf-clear"; cb.type = "button"; cb.textContent = "✕ Clear filters";
      cb.addEventListener("click", function () { gState.dept = ""; gState.emp = ""; gState.period = ""; render(); });
      spacer.appendChild(cb);
    }
    inner.appendChild(spacer);
    globalfilters.appendChild(inner);
    globalfilters.style.display = "";
  }
  function clearHeaderExports() { if (appheaderexports) appheaderexports.innerHTML = ""; }
  function renderHeaderExports(specs, controls) {
    if (!appheaderexports) return;
    appheaderexports.innerHTML = "";
    (controls || []).forEach(function (c) { if (c) appheaderexports.appendChild(c); });
    (specs || []).forEach(function (s) {
      var btn = document.createElement("button"); btn.className = "hx-btn"; btn.type = "button"; btn.textContent = s.label;
      btn.addEventListener("click", function () { try { s.on(); } catch (e) {} });
      appheaderexports.appendChild(btn);
    });
    appheaderexports.appendChild(el2("span", "hx-status", "", "rms-xstatus"));
  }
  function el2(tag, cls, txt, id) { var e = document.createElement(tag); if (cls) e.className = cls; if (id) e.id = id; if (txt != null) e.textContent = txt; return e; }
  function renderHeaderActions() {
    if (!appheaderactions) return;
    appheaderactions.innerHTML = "";
    var tb = document.createElement("button"); tb.className = "gf-clear"; tb.type = "button";
    function isDark() { var d = document.documentElement.getAttribute("data-theme"); return d === "dark" || (d == null && window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches); }
    function paint() { tb.textContent = isDark() ? "☀︎ Light" : "☾ Dark"; }
    tb.addEventListener("click", function () { document.documentElement.setAttribute("data-theme", isDark() ? "light" : "dark"); paint(); });
    paint(); appheaderactions.appendChild(tb);
    if (rawRecords.length || copayRecords.length || coverageRecords.length) {
      var nb = document.createElement("button"); nb.className = "gf-clear"; nb.type = "button"; nb.textContent = "↺ New import";
      nb.title = "Load different files"; nb.addEventListener("click", function () { showImport(); });
      appheaderactions.appendChild(nb);
    }
  }

  /* ---------- exports (rendered in the frozen header on EVERY tab; build the whole dashboard
     — Bypass + Copays + Coverage — scoped to the current Department/Employee filters) ---------- */
  var PAS_SYN = "__PAS_DEPT__";
  function shellDownload(html, name) {
    var blob = new Blob([html], { type: "text/html" }), a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function exportFileName(name) { return String(name || "export").replace(/[\\/:*?"<>|\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 90) + ".html"; }
  function setHeaderXStatus(msg) { var e = document.getElementById("rms-xstatus"); if (e) e.textContent = msg || ""; }
  /* Records for an export: every loaded month, NOT just the month on screen. The downloaded file
     carries the whole set plus a Data picker, so a leader can flip between one month and all months
     combined; the Month filter here only decides which of those the file opens on. */
  function bypassExportScope() {
    var dH = bypassDeptHeader(), recs0 = bypassDisplayRecordsBase();
    var recs1 = (gState.dept && dH) ? recs0.filter(function (r) { var v = trimVal(r[dH]); if (v === "") v = "(blank)"; return v === gState.dept; }) : recs0;
    var deptDisplay = gState.dept || "All Departments";
    recs1.forEach(function (r) { r[PAS_SYN] = deptDisplay; r[PAS_EMP] = cleanDisplayName(r[empKey]); });
    return { recs: recs1, deptDisplay: deptDisplay };
  }
  function uniqEmpNames(subset) { var s = {}; subset.forEach(function (r) { s[cleanDisplayName(r[empKey])] = 1; }); return Object.keys(s); }
  // Copay rows for an export: EVERY loaded period (the copay module buckets them itself and opens on
  // `period`), so a downloaded file keeps its trend section instead of one flat period.
  function copayForNames(names) {
    if (!copayRecords.length) return null;
    var ccm = copayColsNow(), recsC = copayRecords;
    if (names && names.length && ccm.checkinUser) {
      var want = {}; names.forEach(function (nm) { var k = covNameKey(nm); if (k && k !== "|") want[k] = 1; });
      if (Object.keys(want).length) recsC = recsC.filter(function (r) { return want[covNameKey(r[ccm.checkinUser])]; });
    }
    return { records: recsC, headers: copayHeaders, title: "PAS Quality — Copays", generatedAt: nowLabel(), groupKey: "checkin",
      granularity: gState.gran,
      period: gPeriodActive() ? gState.period : "all", hideMonthPicker: historyMonths().length > 1 };
  }
  // Every month's registration counts, each entry tagged with its own YYYY-MM period. The viewer lines
  // them up with the months it is actually showing, so a single month's file is never divided into a
  // multi-month warning count (and never labelled with the wrong month).
  function scopedRegData() {
    if (!regEntries.length) return null;
    normalizeRegPeriods();
    return { entries: regEntries, month: regMonth };   // per-month scoping is by each entry's __period
  }
  /* All-month companion records for an export's Month-over-month section, collapsed to one row per
     month × employee (the count rides along in __n). Keeps the trend in every downloaded file without
     carrying every month's rows into it. */
  /* The true first/last dates behind the collapsed rows, so a downloaded file can still tell which of
     its periods are only partly covered. Set as a side effect of the call below and read straight after. */
  var momWindow = null;
  function momDateWindow() {
    if (!momWindow) return null;
    function iso(t) { var d = new Date(t); return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
    return [iso(momWindow[0]), iso(momWindow[1])];
  }
  function momRecordsFor(deptDisplay, empName) {
    var dH = bypassDateHeader(); if (!dH) return null;
    var dHd = bypassDeptHeader(), base = bypassDisplayRecordsBase();
    if (gState.dept && dHd) base = base.filter(function (r) { var v = trimVal(r[dHd]); if (v === "") v = "(blank)"; return v === gState.dept; });
    if (empName) base = base.filter(function (r) { return cleanDisplayName(r[empKey]) === empName; });
    /* One row per period × employee, carrying its count in __n. The synthetic date has to land inside
       the bucket it stands for, because the viewer re-derives the period from it: the 1st for a month,
       the Monday itself for a week. */
    var buckets = {}, out = [];
    momWindow = null;
    base.forEach(function (r) {
      var rawD = new Date(String(r[dH] == null ? "" : r[dH]).trim());
      if (!isNaN(rawD)) { var t = rawD.getTime();
        if (!momWindow) momWindow = [t, t];
        else { if (t < momWindow[0]) momWindow[0] = t; if (t > momWindow[1]) momWindow[1] = t; } }
      var p = gPeriodFrom(r[dH]); if (!p) return;
      var emp = cleanDisplayName(r[empKey]), k = p + "|" + emp;
      if (!buckets[k]) { var a = p.split("-"), row = {};
        row[dH] = a[1].replace(/^0/, "") + "/" + (isWeekKeyG(p) ? a[2].replace(/^0/, "") : "1") + "/" + a[0];
        row[PAS_SYN] = deptDisplay; row[PAS_EMP] = emp; row.__n = 0; buckets[k] = row; out.push(row); }
      buckets[k].__n++;
    });
    return out;
  }
  // Every loaded month, oldest first, for the export's Data picker. Fewer than 2 → no picker.
  function exportPeriods() { return historyMonths().map(function (k) { return { key: k, label: periodLabelG(k) }; }); }
  /* An exported file is locked to the granularity chosen here: a clinic in its first month gets the
     weekly file, everyone else gets the monthly one, and neither sees a control for the other. */
  function bypassPayload(level, dept, emp, subset, ro, cpay) {
    return { level: level, dept: dept, emp: emp, records: subset, headers: headers, colMap: colMap,
      granularity: gState.gran,
      periods: exportPeriods(), period: gPeriodActive() ? gState.period : "",
      deptKey: PAS_SYN, empKey: PAS_EMP, deptLabel: "department", empLabel: labelFromHeader(empKey, "employee"),
      title: "PAS Quality — Bypassed Warnings", generatedAt: nowLabel(),
      monthRecords: momRecordsFor(gState.dept || "All Departments", emp || null), monthWindow: momDateWindow(),
      regData: scopedRegData(), readOnly: !!ro, copayData: cpay || null };
  }
  /* Coverage Accuracy is rendered to static HTML at export time, so build one snapshot per period
     (plus the combined one, keyed "") and let the file's Data picker swap between them. */
  function covSnap(names) {
    if (!coverageRecords.length) return null;
    var restrict = names ? coverageKeysFor(names) : null, out = {}, keys = [""].concat(historyMonths());
    keys.forEach(function (k) { withPeriod(k, function () { out[k] = coverageSnapshotHTML(restrict); }); });
    return out;
  }
  // What a downloaded file opens on, in words, for the status line under the buttons.
  function exportScopeNote() {
    if (historyMonths().length < 2) return "";
    return " (opens on " + (gPeriodActive() ? periodLabelLongG(gState.period) : "all " + periodNounG(true) + " combined") +
      ", grouped by " + periodNounG() + " — the reader can switch " + periodNounG(true) + ")";
  }
  /* ================= Executive rollup (.pasq.json) =================
     "Export for leadership" used to write a self-contained HTML with every raw row baked in, so two
     clinics' exports could never be added together. It now writes a small pre-aggregated JSON rollup:
     counts only, partitioned clinic -> month, with no patient identifiers in the file. Each clinic
     drops its rollup into one shared folder; "PAS Executive View.html" links that folder and merges
     every rollup it finds. Because every number is bucketed under (clinic, month), a clinic dropping a
     newer file simply replaces its own months instead of double-counting them. ------------------- */
  var ROLLUP_FORMAT = "PAS_QUALITY_ROLLUP";
  var ROLLUP_VERSION = 1;
  var ROLLUP_NO_MONTH = "unknown";        // bucket for rows whose date will not parse
  var ROLLUP_NO_CLINIC = "(no clinic)";

  function rollInc(o, k, n) { k = trimVal(k); if (k === "") return; o[k] = (o[k] || 0) + (n == null ? 1 : n); }
  function rollAdd(o, k, n) { if (!n) n = 0; o[k] = (o[k] || 0) + n; }
  /* Date/hour parsing mirrors the viewer's (assets js) so a rollup's month, weekday and hour buckets
     land on exactly the same values the on-screen dashboard shows. */
  function rollDate(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim(), m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) { var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[1] - 1, +m[2]); }
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    var d = new Date(s); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  function rollHour(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim(), m = s.match(/(\d{1,2}):(\d{2})\s*([AaPp])/);
    if (m) { var h = +m[1] % 12; if (/[Pp]/.test(m[3])) h += 12; return h; }
    m = s.match(/^(\d{1,2}):(\d{2})/); return m ? +m[1] : null;
  }
  function rollPeriod(v) { var d = rollDate(v); return d ? d.getFullYear() + "-" + pad2(d.getMonth() + 1) : periodFromDate(v); }
  function rollDayKey(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }

  // covNameKey(staff name) -> clinic, learned from whichever sheet carries a department column.
  function rollEmpDeptMap() {
    var m = {}, dH = bypassDeptHeader();
    if (rawRecords.length && empKey && dH) bypassDisplayRecordsBase().forEach(function (r) {
      var k = covNameKey(r[empKey]); if (!k || k === "|" || (k in m)) return;
      m[k] = trimVal(r[dH]) || "(blank)";
    });
    if (copayRecords.length) {
      var ccm = copayColsNow();
      if (ccm.checkinUser && ccm.dept) copayRecords.forEach(function (r) {
        var k = covNameKey(r[ccm.checkinUser]); if (!k || k === "|" || (k in m)) return;
        m[k] = trimVal(r[ccm.dept]) || "(blank)";
      });
    }
    return m;
  }

  /* Builds { clinics: { <clinic>: { months: { "YYYY-MM": { bypass, copays, coverage, registrations } } } } }.
     Every leaf is a count or a dollar total. Patient identifiers are gathered into throwaway sets
     (`_pt`) purely to size a distinct count, then deleted before the file is written. */
  function rollupBuild() {
    normalizeRegPeriods();
    var edm = rollEmpDeptMap(), clinics = {};
    var only = gState.dept || null;          // a clinic picked on screen scopes the file to that clinic

    function clinicOf(v) { var s = trimVal(v); return s === "" ? ROLLUP_NO_CLINIC : s; }
    function deptForEmp(nameV, fallback) {
      var k = covNameKey(nameV), d = (k && k !== "|") ? edm[k] : null;
      return clinicOf(d || fallback);
    }
    function M(clinic, period) {
      if (only && clinic !== only) return null;
      var c = clinics[clinic] || (clinics[clinic] = { months: {} });
      var p = period || ROLLUP_NO_MONTH;
      return c.months[p] || (c.months[p] = {});
    }
    function sub(bucket, key, init) { return bucket[key] || (bucket[key] = init); }

    /* ---- Bypassed warnings ---- */
    var dH = bypassDeptHeader(), dateH = bypassDateHeader(), cm = colMap || {};
    if (rawRecords.length && empKey) {
      bypassDisplayRecordsBase().forEach(function (r) {
        var clinic = dH ? clinicOf(r[dH]) : deptForEmp(r[empKey]);
        var mb = M(clinic, dateH ? rollPeriod(r[dateH]) : ""); if (!mb) return;
        var b = sub(mb, "bypass", { n: 0, types: {}, details: {}, byWeekday: [0, 0, 0, 0, 0, 0, 0], byHour: {}, byDay: {}, employees: {}, _pt: {} });
        var emp = cleanDisplayName(r[empKey]) || "(blank)";
        var e = sub(b.employees, emp, { n: 0, types: {}, _pt: {} });
        b.n++; e.n++;
        rollInc(b.types, r[cm.messageType]); rollInc(e.types, r[cm.messageType]);
        rollInc(b.details, r[cm.errorDetail]);
        var pv = cm.patient ? trimVal(r[cm.patient]) : "";
        if (pv !== "") { b._pt[pv] = 1; e._pt[pv] = 1; }
        var dv = dateH ? rollDate(r[dateH]) : null;
        if (dv) { b.byWeekday[dv.getDay()]++; rollInc(b.byDay, rollDayKey(dv)); }
        var hv = (cm.time && trimVal(r[cm.time]) !== "") ? r[cm.time] : (dateH ? r[dateH] : null);
        var hr = rollHour(hv); if (hr != null && hr >= 0 && hr < 24) rollInc(b.byHour, String(hr));
      });
    }

    /* ---- Copays (same MyChart exclusion the Copays tab applies) ---- */
    if (copayRecords.length && window.RMSCopays) {
      var ccm = copayColsNow();
      copayRecords.forEach(function (r) {
        if (ccm.historyUser && RMSCopays.cleanName(r[ccm.historyUser]).toUpperCase() === "MYCHART, GENERIC") return;
        var raw = ccm.checkinUser ? r[ccm.checkinUser] : "";
        var clinic = ccm.dept ? clinicOf(trimVal(r[ccm.dept]) || deptForEmp(raw)) : deptForEmp(raw);
        var mb = M(clinic, ccm.date ? rollPeriod(r[ccm.date]) : ""); if (!mb) return;
        var c = sub(mb, "copays", { countDue: 0, countPaid: 0, sumDue: 0, sumPaid: 0, employees: {} });
        var emp = cleanDisplayName(RMSCopays.cleanName(raw)) || "(blank)";
        var e = sub(c.employees, emp, { countDue: 0, countPaid: 0, sumDue: 0, sumPaid: 0 });
        var due = r[ccm.copayDue], paid = r[ccm.copayPaid];
        if (due != null && String(due).trim() !== "") { c.countDue++; e.countDue++; var dm = RMSCopays.money(due); c.sumDue += dm; e.sumDue += dm; }
        if (paid != null && String(paid).trim() !== "") { c.countPaid++; e.countPaid++; var pm = RMSCopays.money(paid); c.sumPaid += pm; e.sumPaid += pm; }
      });
    }

    /* ---- Coverage verification (same status whitelist + CSN/staff filters as the Coverage tab) ---- */
    if (coverageRecords.length) {
      var cc = coverageCols(), CANON = {};
      COVERAGE_STATUSES.forEach(function (s) { CANON[s.toLowerCase()] = s; });
      var allowed = coverageUserKeySet(), haveFilter = Object.keys(allowed).length > 0;
      coverageRecords.forEach(function (r) {
        var canon = CANON[String(r[cc.status] == null ? "" : r[cc.status]).trim().toLowerCase()]; if (!canon) return;
        if (cc.csn && String(r[cc.csn] == null ? "" : r[cc.csn]).trim() === "") return;
        if (cc.user && haveFilter && !allowed[covNameKey(r[cc.user])]) return;
        var raw = cc.user ? r[cc.user] : "";
        var clinic = deptForEmp(raw, cc.dept ? trimVal(r[cc.dept]) : "");
        var mb = M(clinic, cc.date ? rollPeriod(r[cc.date]) : ""); if (!mb) return;
        var v = sub(mb, "coverage", { n: 0, statuses: {}, employees: {} });
        var emp = cleanDisplayName(raw) || "(blank)";
        var e = sub(v.employees, emp, { n: 0, statuses: {} });
        v.n++; e.n++; rollAdd(v.statuses, canon, 1); rollAdd(e.statuses, canon, 1);
      });
    }

    /* ---- Registrations (the denominator for warnings per 100 registrations) ---- */
    regEntries.forEach(function (en) {
      var clinic = deptForEmp(en.staff);
      var mb = M(clinic, en.__period || regMonth || ""); if (!mb) return;
      var g = sub(mb, "registrations", { n: 0, employees: {} });
      var emp = cleanDisplayName(en.staff) || "(blank)", n = +en.count || 0;
      g.n += n; rollAdd(g.employees, emp, n);
    });

    /* ---- Collapse the distinct-patient sets to counts, then drop the identifiers ---- */
    var monthSet = {};
    Object.keys(clinics).forEach(function (cn) {
      var ms = clinics[cn].months;
      Object.keys(ms).forEach(function (p) {
        if (p !== ROLLUP_NO_MONTH) monthSet[p] = 1;
        var b = ms[p].bypass; if (!b) return;
        b.patients = Object.keys(b._pt).length; delete b._pt;
        Object.keys(b.employees).forEach(function (en) {
          var e = b.employees[en]; e.patients = Object.keys(e._pt).length; delete e._pt;
        });
      });
    });

    return { format: ROLLUP_FORMAT, version: ROLLUP_VERSION,
      generatedAt: nowLabel(), generatedAtISO: new Date().toISOString(),
      source: (fileNames || []).concat(copayFileNames || [], coverageFileNames || [], regFileNames || []),
      months: Object.keys(monthSet).sort(), clinics: clinics };
  }

  function rollupFileName(label, months) {
    var ym = (months || []).filter(function (m) { return /^\d{4}-\d{2}$/.test(m); });
    var span = ym.length ? (ym[0] === ym[ym.length - 1] ? ym[0] : ym[0] + " to " + ym[ym.length - 1]) : "current";
    var safe = String(label || "All clinics").replace(/[\\/:*?"<>|\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
    return "PAS " + safe + " " + span + ".pasq.json";
  }

  function exportExecutiveData() {
    if (!hasAnyData()) { setHeaderXStatus("No data loaded yet — import your export files first."); return; }
    var roll = rollupBuild(), names = Object.keys(roll.clinics);
    if (!names.length) { setHeaderXStatus("Nothing to export for the current filters."); return; }
    var label = gState.dept || (names.length === 1 ? names[0] : "All clinics");
    var json = JSON.stringify(roll), fname = rollupFileName(label, roll.months);
    var blob = new Blob([json], { type: "application/json" }), a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = fname; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    var kb = Math.max(1, Math.round(json.length / 1024)), mn = roll.months.length;
    setHeaderXStatus("Saved " + fname + " — " + names.length + " clinic" + (names.length === 1 ? "" : "s") + ", " +
      mn + " month" + (mn === 1 ? "" : "s") + ", " + kb + " KB. Put it in the shared executive folder.");
  }

  function exportLeadership() {
    var sc = bypassExportScope(), cp = copayForNames(null), name = "Leadership — " + sc.deptDisplay;
    shellDownload(RMSViewer.buildHTML(bypassPayload("overview", null, null, sc.recs, true, cp), name, cp, covSnap(null)), exportFileName("Leadership - " + sc.deptDisplay));
    setHeaderXStatus("Downloaded leadership summary" + exportScopeNote() + ".");
  }
  function exportClinic() {
    var sc = bypassExportScope(), emps = uniqEmpNames(sc.recs), cp = copayForNames(emps);
    shellDownload(RMSViewer.buildHTML(bypassPayload("department", sc.deptDisplay, null, sc.recs, false, cp), sc.deptDisplay, cp, covSnap(emps)), exportFileName(sc.deptDisplay));
    setHeaderXStatus("Downloaded clinic export" + exportScopeNote() + " — open it to export by employee.");
  }
  function exportEmployee() {
    var emp = gState.emp, sc = bypassExportScope();
    var subset = sc.recs.filter(function (r) { return cleanDisplayName(r[empKey]) === emp; });
    var cp = copayForNames([emp]);
    shellDownload(RMSViewer.buildHTML(bypassPayload("employee", sc.deptDisplay, emp, subset, true, cp), emp, cp, covSnap([emp])), exportFileName(emp));
    setHeaderXStatus("Downloaded " + emp + exportScopeNote() + ".");
  }
  /* ---- Monthly vs. combined lives in the exported FILE, not here ----
     Every download carries all of the loaded months and a Data picker in its tab bar, so whoever opens
     it chooses one month or all months combined. The Month filter above only sets the view the file
     opens on. `withPeriod` swaps gState.period for one build — used to pre-render the Coverage
     Accuracy tab, which is static HTML, once per month. */
  function withPeriod(period, fn) {
    var saved = gState.period;
    gState.period = period || "";
    try { return fn(); } finally { gState.period = saved; }
  }

  function renderExportsBar() {
    var specs = [];
    if (rawRecords.length) {
      specs.push({ label: "⬇ Export for executive view (data)", on: exportExecutiveData });
      specs.push({ label: "⬇ Clinic export", on: exportClinic });
      if (gState.emp) specs.push({ label: "⬇ Export this employee", on: exportEmployee });
    }
    if (hasAnyData()) specs.push({ label: "⬇ Save historical data (Excel)", on: function () { downloadHistory(false); } });
    if (!specs.length) { clearHeaderExports(); return; }
    renderHeaderExports(specs);
  }

  function render() {
    renderHeaderActions();
    normalizeRegPeriods();   // registrations' month text → YYYY-MM (a file loaded before the bypass data resolves here)
    var hasBypass = rawRecords.length, hasCopay = copayRecords.length, hasCoverage = coverageRecords.length;
    if (!hasBypass && !hasCopay && !hasCoverage) {
      importwrap.style.display = ""; empty.style.display = ""; dashboard.style.display = "none";
      if (moduletabs) moduletabs.style.display = "none";
      if (globalfilters) { globalfilters.style.display = "none"; globalfilters.innerHTML = ""; }
      clearHeaderExports();
      return;
    }
    importwrap.style.display = "none"; empty.style.display = "none"; dashboard.style.display = "";
    ensurePeriodDefault();   // first paint after an import opens on the most recent month
    var present = { bypass: hasBypass, copays: hasCopay, coverage: hasCoverage };
    if (!present[activeModule]) activeModule = hasBypass ? "bypass" : hasCopay ? "copays" : "coverage";
    refreshTabs();
    renderRuleBar();
    if (activeModule === "coverage") { covState.dept = gState.dept; covState.emp = gState.emp; }
    renderGlobalFilters();
    renderExportsBar();   // exports live in the frozen header on every tab

    if (activeModule === "coverage") { renderCoverage(); return; }

    if (activeModule === "copays") {
      var ccm = copayColsNow();
      var edmC = bypassEmpDeptMap();   // narrow copay by the BYPASS department (same filter as the Bypass tab)
      // Pass ALL months (dept/employee-scoped). The copay module buckets by month internally: it shows
      // the selected month's numbers (via `period`) while its Month-over-month section spans every month.
      var subset = copayRecords.filter(function (r) {
        if (gState.dept && edmC[covNameKey(r[ccm.checkinUser])] !== gState.dept) return false;
        if (gState.emp && copayEmpOf(r, ccm) !== gState.emp) return false;
        return true;
      });
      RMSCopays.mount(dashboard, { records: subset, headers: copayHeaders, colMap: ccm,
        title: "PAS Quality — Copays", generatedAt: nowLabel(), groupKey: "checkin", granularity: gState.gran,
        externalFilters: true, period: gPeriodActive() ? gState.period : "all",
        onPeriod: function (p) { copayPeriod = p; } });
      return;
    }

    // bypass — the whole uploaded sheet is treated as one department (one sheet at a time)
    deptKey = deptSel.value || deptKey; empKey = empSel.value || empKey;
    var empLabel = labelFromHeader(empKey, "employee");
    var filterActive = !!(ruleInclude && applyRules);
    var dH = bypassDeptHeader();
    var recs0 = bypassDisplayRecords();   /* same rule-filtered set the coverage tab uses */
    var recs1 = (gState.dept && dH) ? recs0.filter(function (r) { var v = trimVal(r[dH]); if (v === "") v = "(blank)"; return v === gState.dept; }) : recs0;
    var deptDisplay = gState.dept || "All Departments";
    var SYN = "__PAS_DEPT__";
    // All-period records (same dept + employee scope, but NOT the selected period) for the trend section.
    var recsMoM = bypassDisplayRecordsBase();
    if (gState.dept && dH) recsMoM = recsMoM.filter(function (r) { var v = trimVal(r[dH]); if (v === "") v = "(blank)"; return v === gState.dept; });
    if (gState.emp) recsMoM = recsMoM.filter(function (r) { return cleanDisplayName(r[empKey]) === gState.emp; });
    recsMoM.forEach(function (r) { r[SYN] = deptDisplay; r[PAS_EMP] = cleanDisplayName(r[empKey]); });
    recs1.forEach(function (r) { r[SYN] = deptDisplay; r[PAS_EMP] = cleanDisplayName(r[empKey]); });
    var level = gState.emp ? "employee" : "department";
    RMSViewer.mount(dashboard, {
      level: level, dept: deptDisplay, emp: gState.emp || null, granularity: gState.gran,
      records: recs1, monthRecords: recsMoM, headers: headers, colMap: colMap,
      deptKey: SYN, empKey: PAS_EMP, deptLabel: "department", empLabel: empLabel,
      // the viewer appends the month it is showing to this subtitle — don't name it twice
      title: "PAS Quality — Bypassed Warnings" + (filterActive ? " (Include-only)" : ""), generatedAt: nowLabel(), onReset: showImport, externalFilters: true,
      copayData: copayRecords.length ? { records: copayRecords, headers: copayHeaders, title: "PAS Quality — Copays", generatedAt: nowLabel(), granularity: gState.gran, period: gPeriodActive() ? gState.period : "all" } : null,
      regData: scopedRegData(),
      getCoverageSnapshot: coverageRecords.length ? function (names) { return coverageSnapshotHTML(coverageKeysFor(names)); } : null
    });
  }
  function renderRuleBar() {
    if (!rulebar) return;
    rulebar.innerHTML = "";
    // Standalone "new-error" detector. Flags any Error Type/Detail value present in the
    // data that is on NEITHER the Include nor the Exclude list, so a reviewer knows an
    // include/exclude decision is still pending. Only shows once a rules file is loaded.
    var edKey = colMap && colMap.errorDetail;
    if (!ruleInclude || !rawRecords.length || !edKey) { rulebar.style.display = "none"; return; }
    var counts = new Map(), orig = new Map();
    rawRecords.forEach(function (r) { var v = r[edKey], k = ruleKey(v); if (k === "") return; if (!ruleInclude.has(k) && !ruleExclude.has(k)) { counts.set(k, (counts.get(k) || 0) + 1); if (!orig.has(k)) orig.set(k, String(v).trim()); } });
    var newRules = Array.from(counts, function (e) { return { rule: orig.get(e[0]) || e[0], n: e[1] }; }).sort(function (a, b) { return b.n - a.n; });
    var newWarnings = newRules.reduce(function (s, x) { return s + x.n; }, 0);
    var hasNew = newRules.length > 0;
    rulebar.style.display = "";

    var alert = document.createElement("div"); alert.className = "rulebar-alert" + (hasNew ? "" : " clean");
    var icon = document.createElement("span"); icon.className = "ra-icon";
    icon.innerHTML = hasNew
      ? '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
      : '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>';
    alert.appendChild(icon);
    var t = document.createElement("span"); t.className = "ra-text";
    var sub = document.createElement("span"); sub.className = "ra-sub";
    if (hasNew) {
      t.innerHTML = "<b>" + fmt(newRules.length) + "</b> error type" + (newRules.length === 1 ? "" : "s") + " need an Include / Exclude decision";
      sub.textContent = fmt(newWarnings) + " warning" + (newWarnings === 1 ? "" : "s") + " not on your Include or Exclude list" + (ruleFileName ? " · rules: " + ruleFileName : "");
    } else {
      t.textContent = "All error types are categorized";
      sub.textContent = fmt(ruleIncludeRaw.length) + " included · " + fmt(ruleExcludeRaw.length) + " excluded" + (ruleFileName ? " · " + ruleFileName : "");
    }
    alert.appendChild(t); alert.appendChild(sub);
    rulebar.appendChild(alert);

    if (hasNew) {
      var det = document.createElement("details"); det.className = "rulebar-new";
      var sm = document.createElement("summary"); sm.textContent = "Review " + newRules.length + " new / uncategorized error type" + (newRules.length === 1 ? "" : "s"); det.appendChild(sm);
      var list = document.createElement("div"); list.className = "rulebar-newlist";
      newRules.slice(0, 100).forEach(function (x) { var d = document.createElement("div"); d.className = "rulebar-newitem"; d.innerHTML = "<span>" + escText(x.rule) + "</span><b>" + fmt(x.n) + "</b>"; list.appendChild(d); });
      det.appendChild(list);
      if (newRules.length > 100) { var more = document.createElement("div"); more.className = "rulebar-newmore"; more.textContent = "+ " + fmt(newRules.length - 100) + " more"; det.appendChild(more); }
      rulebar.appendChild(det);
    }
  }
  function labelFromHeader(h, fallback) {
    var n = norm(h);
    if (n.indexOf("department") >= 0 || n.indexOf("dept") >= 0) return "department";
    if (n.indexOf("checkin") >= 0) return "employee";
    if (n.indexOf("user") >= 0) return "employee";
    return fallback;
  }

  function clearAll() { rawRecords = []; headers = []; fileNames = []; deptKey = null; empKey = null;
    copayRecords = []; copayHeaders = []; copayFileNames = []; activeModule = "bypass";
    regEntries = []; regMonth = ""; regFileNames = [];
    coverageRecords = []; coverageHeaders = []; coverageFileNames = []; covState = { dept: "", emp: "" };
    gState = { dept: "", emp: "", period: "" }; copayPeriod = "all"; periodInitialized = false;
    ruleInclude = null; ruleExclude = null; ruleIncludeRaw = []; ruleExcludeRaw = []; ruleFileName = ""; applyRules = true;
    fileInput.value = ""; setStatus("", ""); config.style.display = "none"; render(); }

  /* ---------- wire ---------- */
  fileInput.addEventListener("change", function () { handleFiles(fileInput.files); });
  ["dragenter", "dragover"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
  ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); if (ev === "dragleave" && e.target !== drop) return; drop.classList.remove("over"); }); });
  drop.addEventListener("drop", function (e) { if (e.dataTransfer && e.dataTransfer.files) handleFiles(e.dataTransfer.files); });
  drop.addEventListener("click", function (e) { if (["INPUT", "BUTTON", "LABEL"].indexOf(e.target.tagName) < 0) fileInput.click(); });
  deptSel.addEventListener("change", render);
  empSel.addEventListener("change", render);
  clearBtn.addEventListener("click", clearAll);
  render();
})();
