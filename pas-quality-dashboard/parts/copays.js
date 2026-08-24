/* ================= RMS Copays — aggregation + dashboard (self-contained) =================
   window.RMSCopays.mount(rootEl, payload)
   payload = { records:[{<hdr>:val}], headers:[...], colMap:{checkinUser,copayDue,copayPaid,historyUser,date},
               title, generatedAt, groupKey:'checkin'|'history' }
   Reuses the .rms* / .card / .kpi / .rms-table styles already present in the page.        */
(function () {
  "use strict";

  /* ---------- core (mirrors scratchpad copay_core.js; unit-tested) ---------- */
  function norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function cleanName(s) { s = String(s == null ? "" : s).trim(); return s.replace(/\s*\[[^\]]*\]\s*$/, "").trim(); }
  // Display names: keep letters, spaces and the Last, First comma — drop IDs, brackets, digits, other punctuation.
  function cleanDisplayName(s) { s = String(s == null ? "" : s); if (s === "(blank)") return s;
    var c = s.replace(/[^A-Za-z, ]+/g, " ").replace(/\s+/g, " ").replace(/\s*,\s*/g, ", ").replace(/^[,\s]+|[,\s]+$/g, "").trim();
    return c === "" ? "(blank)" : c; }
  function money(v) {
    if (v == null || v === "") return 0;
    if (typeof v === "number") return v;
    var s = String(v).replace(/[^0-9.\-]/g, "");
    if (s === "" || s === "-" || s === ".") return 0;
    var n = parseFloat(s); return isNaN(n) ? 0 : n;
  }
  function detectCopayCols(headers) {
    var N = headers.map(norm);
    function find(pred) { for (var i = 0; i < headers.length; i++) if (pred(N[i], headers[i])) return headers[i]; return null; }
    var cm = {};
    cm.checkinUser = find(function (n) { return n.indexOf("checkin") >= 0 && n.indexOf("user") >= 0; }) || find(function (n) { return n.indexOf("checkin") >= 0; });
    cm.copayDue    = find(function (n) { return n.indexOf("copaydue") >= 0 || (n.indexOf("copay") >= 0 && n.indexOf("due") >= 0); });
    cm.copayPaid   = find(function (n) { return n.indexOf("copaypaid") >= 0 || (n.indexOf("copay") >= 0 && n.indexOf("paid") >= 0); });
    cm.historyUser = find(function (n) { return n.indexOf("copayhistory") >= 0 || (n.indexOf("copay") >= 0 && n.indexOf("history") >= 0); });
    cm.date        = find(function (n) { return n === "date" || n === "apptdate" || n === "appointmentdate" || (n.indexOf("appt") >= 0 && n.indexOf("date") >= 0); }) || find(function (n) { return n.indexOf("date") >= 0; });
    cm.dept        = find(function (n) { return n === "dept" || n.indexOf("department") >= 0 || (n.indexOf("login") >= 0 && n.indexOf("dept") >= 0); });
    return cm;
  }
  function looksLikeCopays(headers) { var cm = detectCopayCols(headers); return !!(cm.copayDue && cm.copayPaid); }

  function hasVal(v) { return v != null && String(v).trim() !== ""; }
  function aggregate(records, cm, opts) {
    opts = opts || {};
    var keyCol = (opts.groupKey === "history" && cm.historyUser) ? cm.historyUser : cm.checkinUser;
    var groups = new Map();
    function g(name) { if (!groups.has(name)) groups.set(name, { employee: name, countDue: 0, countPaid: 0, sumDue: 0, sumPaid: 0 }); return groups.get(name); }
    records.forEach(function (r) {
      // exclude MyChart auto-collections (identical to the Bypass page's exclusion)
      if (cm.historyUser && cleanName(r[cm.historyUser]).toUpperCase() === "MYCHART, GENERIC") return;
      var name = cleanDisplayName(cleanName(r[keyCol])); if (name === "") name = "(blank)";
      var row = g(name);
      if (hasVal(r[cm.copayDue]))  { row.countDue++;  row.sumDue  += money(r[cm.copayDue]); }
      if (hasVal(r[cm.copayPaid])) { row.countPaid++; row.sumPaid += money(r[cm.copayPaid]); }
    });
    var rows = Array.from(groups.values()).filter(function (r) { return r.countDue || r.countPaid; }).map(function (r) {
      return { employee: r.employee, countDue: r.countDue, countPaid: r.countPaid, sumDue: r.sumDue, sumPaid: r.sumPaid,
               pctCount: r.countDue ? r.countPaid / r.countDue : 0, pctAmount: r.sumDue ? r.sumPaid / r.sumDue : 0 }; });
    rows.sort(function (a, b) { return String(a.employee).localeCompare(String(b.employee)); });
    var tot = rows.reduce(function (t, r) { t.countDue += r.countDue; t.countPaid += r.countPaid; t.sumDue += r.sumDue; t.sumPaid += r.sumPaid; return t; },
                          { countDue: 0, countPaid: 0, sumDue: 0, sumPaid: 0 });
    tot.pctCount = tot.countDue ? tot.countPaid / tot.countDue : 0;
    tot.pctAmount = tot.sumDue ? tot.sumPaid / tot.sumDue : 0;
    return { rows: rows, total: tot };
  }

  /* ---------- date parsing for month grouping ---------- */
  function parseDate(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim();
    var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) { var y = +m[3]; if (y < 100) y += 2000; return { y: y, mo: +m[1] }; }
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return { y: +m[1], mo: +m[2] };
    var d = new Date(s); if (!isNaN(d)) return { y: d.getFullYear(), mo: d.getMonth() + 1 };
    return null;
  }
  var MON = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  /* ---------- DOM helpers ---------- */
  function el(tag, attrs, kids) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === "class") e.className = attrs[k];
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k === "text") e.textContent = attrs[k];
      else if (k.slice(0, 2) === "on" && typeof attrs[k] === "function") e.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] != null) e.setAttribute(k, attrs[k]);
    }
    if (kids != null) (Array.isArray(kids) ? kids : [kids]).forEach(function (c) { if (c == null) return; e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }
  var SVGNS = "http://www.w3.org/2000/svg";
  function svg(tag, attrs) { var e = document.createElementNS(SVGNS, tag); if (attrs) for (var k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]); return e; }
  function num(n) { return (n == null ? 0 : n).toLocaleString("en-US"); }
  function usd(n) { return "$" + (Math.round((n || 0) * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 }); }
  function pct(n) { return (Math.round((n || 0) * 1000) / 10).toLocaleString("en-US") + "%"; }
  function trunc(s, n) { s = String(s == null ? "" : s); return s.length > n ? s.slice(0, n - 1) + "…" : s; }

  function kpi(label, value, sub, cls) { return el("div", { class: "kpi" + (cls ? " " + cls : "") }, [ el("div", { class: "k-label" }, [label]), el("div", { class: "k-val", html: value }), el("div", { class: "k-sub", text: sub || "" }) ]); }
  function card(title, note, body, span) { return el("div", { class: "card" + (span ? " span2" : "") }, [ el("h3", { text: title }), note ? el("p", { class: "c-note", text: note }) : null, body ]); }

  /* horizontal bars (self-contained; matches viewer look via .chart/.bar/.barcat/.barlabel) */
  function hbar(items, opts) {
    opts = opts || {}; var wrap = el("div", { class: "chart" });
    if (!items.length) { wrap.appendChild(el("div", { class: "empty-note", text: "No data" })); return wrap; }
    var max = Math.max.apply(null, items.map(function (i) { return i.value; })) || 1;
    var rowH = 42, gap = 6, top = 2, W = 640, valW = 70, plotW = W - valW - 2, barH = 16, barY = 22;
    var H = top * 2 + items.length * rowH + (items.length - 1) * gap;
    var s = svg("svg", { viewBox: "0 0 " + W + " " + H, role: "img" });
    items.forEach(function (it, i) {
      var y = top + i * (rowH + gap), bw = Math.max(2, (it.value / max) * plotW);
      s.appendChild(svg("text", { x: 1, y: y + 13, "text-anchor": "start", class: "barcat" })).textContent = trunc(it.label, 90);
      s.appendChild(svg("rect", { x: 1, y: y + barY, width: plotW, height: barH, rx: 5, fill: "var(--surface-2)" }));
      s.appendChild(svg("rect", { x: 1, y: y + barY, width: bw, height: barH, rx: 5, class: "bar" }));
      s.appendChild(svg("text", { x: 1 + bw + 7, y: y + barY + 12, class: "barlabel" })).textContent = opts.fmt ? opts.fmt(it.value) : num(it.value);
    });
    wrap.appendChild(s); return wrap;
  }

  /* "Paid vs. due" bullet bars: per employee, an outer track = the "due" figure (scaled to the largest
     due) with an inner fill = the "paid" figure on the same scale. Works for dollars OR counts via
     opts.fmt/opts.joiner/opts.suffix. Ranked by due. */
  function paidVsDueBar(items, opts) {
    opts = opts || {}; var f = opts.fmt || usd, joiner = opts.joiner || " out of ", suffix = opts.suffix || "";
    var wrap = el("div", { class: "chart" });
    if (!items.length) { wrap.appendChild(el("div", { class: "empty-note", text: "No data" })); return wrap; }
    var maxDue = Math.max.apply(null, items.map(function (i) { return i.due; })) || 1;
    var rowH = 42, gap = 6, top = 2, W = 640, valW = 210, plotW = W - valW - 2, barH = 16, barY = 22;
    var H = top * 2 + items.length * rowH + (items.length - 1) * gap;
    var s = svg("svg", { viewBox: "0 0 " + W + " " + H, role: "img" });
    items.forEach(function (it, i) {
      var y = top + i * (rowH + gap);
      var dueW = Math.max(2, (it.due / maxDue) * plotW);
      var paidW = Math.max(0, (Math.min(it.paid, it.due) / maxDue) * plotW);   // clamp fill so it never overruns the due track
      s.appendChild(svg("text", { x: 1, y: y + 13, "text-anchor": "start", class: "barcat" })).textContent = trunc(it.label, 90);
      s.appendChild(svg("rect", { x: 1, y: y + barY, width: dueW, height: barH, rx: 5, fill: "var(--series-1b)" }));
      if (paidW > 0) s.appendChild(svg("rect", { x: 1, y: y + barY, width: paidW, height: barH, rx: 5, class: "bar" }));
      s.appendChild(svg("text", { x: 1 + dueW + 7, y: y + barY + 12, class: "barlabel" })).textContent = f(it.paid) + joiner + f(it.due) + suffix;
    });
    wrap.appendChild(s); return wrap;
  }

  function th(label, cls) { return el("th", { class: cls || "" }, label); }

  function copayTable(rows, total) {
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table" });
    t.appendChild(el("thead", {}, el("tr", {}, [
      th("Check-in User"), th("Count of Copay Due", "num"), th("Count of Copay Paid", "num"),
      th("% Collected (Opportunities)", "num"),
      th("Sum of Copay Due", "num"), th("Sum of Copay Paid", "num"),
      th("% Collected (Amount)", "num") ])));
    var tb = el("tbody");
    rows.forEach(function (r) {
      tb.appendChild(el("tr", {}, [
        el("td", { text: r.employee }),
        el("td", { class: "num" }, el("span", { class: "pill", text: num(r.countDue) })),
        el("td", { class: "num", text: num(r.countPaid) }),
        el("td", { class: "num", text: pct(r.pctCount) }),
        el("td", { class: "num", text: usd(r.sumDue) }),
        el("td", { class: "num", text: usd(r.sumPaid) }),
        el("td", { class: "num", text: pct(r.pctAmount) }) ]));
    });
    tb.appendChild(el("tr", { class: "rms-totalrow" }, [
      el("td", {}, el("b", { text: "Grand Total" })),
      el("td", { class: "num" }, el("b", { text: num(total.countDue) })),
      el("td", { class: "num" }, el("b", { text: num(total.countPaid) })),
      el("td", { class: "num" }, el("b", { text: pct(total.pctCount) })),
      el("td", { class: "num" }, el("b", { text: usd(total.sumDue) })),
      el("td", { class: "num" }, el("b", { text: usd(total.sumPaid) })),
      el("td", { class: "num" }, el("b", { text: pct(total.pctAmount) })) ]));
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }

  /* Check-in User × month grid of copay collected ($), ranked by total, with a per-month TOTALS row.
     mlist = [{key,label}], perAgg = { <monthKey>: aggregate() result }. */
  function copayMonthMatrix(mlist, perAgg) {
    var emp = {};
    mlist.forEach(function (mm) { perAgg[mm.key].rows.forEach(function (r) {
      if (!emp[r.employee]) emp[r.employee] = { name: r.employee, m: {}, total: 0 };
      emp[r.employee].m[mm.key] = r.sumPaid; emp[r.employee].total += r.sumPaid; }); });
    var rows = Object.keys(emp).map(function (k) { return emp[k]; }).sort(function (a, b) { return b.total - a.total || String(a.name).localeCompare(String(b.name)); });
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table" });
    var head = [ th("Check-in User") ]; mlist.forEach(function (mm) { head.push(th(mm.label, "num")); }); head.push(th("Total collected", "num"));
    t.appendChild(el("thead", {}, el("tr", {}, head)));
    var tb = el("tbody"), colTot = {}, grand = 0; mlist.forEach(function (mm) { colTot[mm.key] = 0; });
    rows.forEach(function (g) {
      var tds = [ el("td", { text: g.name }) ];
      mlist.forEach(function (mm) { var v = g.m[mm.key] || 0; colTot[mm.key] += v; tds.push(el("td", { class: "num", text: usd(v) })); });
      tds.push(el("td", { class: "num" }, el("b", { text: usd(g.total) }))); grand += g.total;
      tb.appendChild(el("tr", {}, tds));
    });
    if (rows.length) { var tr = [ el("td", {}, el("b", { text: "TOTALS" })) ];
      mlist.forEach(function (mm) { tr.push(el("td", { class: "num" }, el("b", { text: usd(colTot[mm.key]) }))); });
      tr.push(el("td", { class: "num" }, el("b", { text: usd(grand) })));
      tb.appendChild(el("tr", { class: "rms-totalrow", style: "border-top:2px solid var(--border-2);font-weight:600" }, tr)); }
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }

  /* ---------- mount ---------- */
  /* Period granularity — see the viewer module for the rationale. Kept local because this module is
     injected into exported files on its own. Week keys are the Monday as "YYYY-MM-DD". */
  function padc2(n) { return String(n).length < 2 ? "0" + n : String(n); }
  function copayWeekKey(y, mo, d) { var dt = new Date(y, mo - 1, d); dt.setDate(dt.getDate() - ((dt.getDay() + 6) % 7));
    return dt.getFullYear() + "-" + padc2(dt.getMonth() + 1) + "-" + padc2(dt.getDate()); }
  function copayIsWeekKey(k) { return /^\d{4}-\d{2}-\d{2}$/.test(String(k)); }

  function mount(root, payload) {
    root.classList.add("rms"); root.innerHTML = "";
    var recs = payload.records || [], cm = payload.colMap || detectCopayCols(payload.headers || (recs[0] ? Object.keys(recs[0]) : []));
    var title = payload.title || "Copays", gen = payload.generatedAt || "";
    var groupKey = payload.groupKey || "checkin";

    /* Period buckets from the date column — by month, or by week when the export was written that way. */
    var gran = payload.granularity === "week" ? "week" : "month";
    var pnoun = gran === "week" ? "week" : "month";
    var months = new Map();
    recs.forEach(function (r) { var d = parseDate(r[cm.date]);
      var key = d ? (gran === "week" ? copayWeekKey(d.y, d.mo, d.d) : d.y + "-" + String(d.mo).padStart(2, "0")) : "";
      if (!months.has(key)) months.set(key, []); months.get(key).push(r); });
    var monthKeys = Array.from(months.keys()).filter(function (k) { return k; }).sort();

    var externalFilters = !!payload.externalFilters;   // dept/emp handled by the global bar; keep only the month picker here
    var state = { period: payload.period || "all", dept: "", emp: "" };
    var wrap = el("div", { class: "rms-wrap" }); root.appendChild(wrap);
    var head = el("div", { class: "rms-head2" }), body = el("div", {});
    wrap.appendChild(head); wrap.appendChild(body);

    function recsFor(p) { return p === "all" ? recs : (months.get(p) || []); }
    function deptVal(r) { return String(r[cm.dept] == null ? "" : r[cm.dept]).replace(/\s+/g, " ").trim(); }
    function empVal(r) { return cleanDisplayName(cleanName(r[cm.checkinUser])); }
    function uniqDepts() { var s = {}; recs.forEach(function (r) { var dp = deptVal(r); if (dp) s[dp] = 1; }); return Object.keys(s).sort(); }
    function uniqEmps(dp) { var s = {}; recs.forEach(function (r) { if (dp && deptVal(r) !== dp) return; s[empVal(r)] = 1; }); return Object.keys(s).sort(); }
    function filteredRecs() { return recsFor(state.period).filter(function (r) { if (state.dept && deptVal(r) !== state.dept) return false; if (state.emp && empVal(r) !== state.emp) return false; return true; }); }
    function periodLabel(p) { if (p === "all") return "All loaded dates"; var s = p.split("-");
      return copayIsWeekKey(p) ? "Wk " + MON[+s[1]] + " " + (+s[2]) : MON[+s[1]] + " " + s[0]; }

    function renderHead() {
      head.innerHTML = "";
      head.appendChild(el("div", { class: "rms-titlebox" }, [ el("h1", { text: "Copays — " + (payload.periodLabelOverride || periodLabel(state.period)) }),
        el("div", { class: "rms-subtitle", text: title }) ]));
      var right = el("div", { class: "rms-headright" });
      // Copays are treated as ONE clinic (the bypass clinic) — the copay sheet's own department
      // column is not split out, since a single upload is all the same clinic.
      if (!externalFilters && state.dept) {
        var esel = el("select", { class: "rms-select" });
        esel.appendChild(el("option", { value: "", text: "▸ All employees" }));
        uniqEmps(state.dept).forEach(function (e2) { esel.appendChild(el("option", { value: e2, text: e2.length > 34 ? e2.slice(0, 33) + "…" : e2 })); });
        esel.value = state.emp;
        esel.addEventListener("change", function () { state.emp = esel.value; rerender(); });
        right.appendChild(el("div", { class: "rms-picker" }, [ el("label", { text: "Employee" }), esel ]));
      }
      if (monthKeys.length > 1 && !externalFilters && !payload.hideMonthPicker) {   // in-app the Month filter lives in the top bar, and in an export the shared Data picker drives every tab; keep this one for standalone copay files
        var sel = el("select", { class: "rms-select" });
        sel.appendChild(el("option", { value: "all", text: "▸ All loaded dates" }));
        monthKeys.forEach(function (k) { sel.appendChild(el("option", { value: k, text: periodLabel(k) + " (" + num((months.get(k) || []).length) + " rows)" })); });
        sel.value = state.period; sel.addEventListener("change", function () { state.period = sel.value; if (payload.onPeriod) payload.onPeriod(state.period); rerender(); });
        right.appendChild(el("div", { class: "rms-picker" }, [ el("label", { text: "Month" }), sel ]));
      }
      head.appendChild(right);
    }

    function renderBody() {
      body.innerHTML = "";
      var subset = filteredRecs();
      var A = aggregate(subset, cm, { groupKey: groupKey });
      var t = A.total;

      var k = el("div", { class: "rms-kpis" });
      k.appendChild(kpi("Count of Copay Due", num(t.countDue), "encounters with a copay due", "accent"));
      k.appendChild(kpi("Count of Copay Paid", num(t.countPaid), "encounters with a payment"));
      k.appendChild(kpi("Sum of Copay Due", usd(t.sumDue), ""));
      k.appendChild(kpi("Sum of Copay Paid", usd(t.sumPaid), ""));
      k.appendChild(kpi("% Collected", pct(t.pctAmount), usd(t.sumPaid) + " of " + usd(t.sumDue)));
      body.appendChild(k);
      /* Period over period — only while every loaded period is on screen ("All loaded dates"). Picking a
         single period up top makes this that period's page, and the trend sections come off it.
         Windowed to the most recent 13 at week granularity so the grid stays printable. */
      if (state.period === "all" && monthKeys.length >= 2) {
        var winK = monthKeys.slice(gran === "week" ? -13 : -24), dropped = monthKeys.length - winK.length;
        var moList = winK.map(function (kk) { return { key: kk, label: periodLabel(kk) }; });
        var perAgg = {}; moList.forEach(function (mm) {
          var recsM = (months.get(mm.key) || []).filter(function (r) { if (state.dept && deptVal(r) !== state.dept) return false; if (state.emp && empVal(r) !== state.emp) return false; return true; });
          perAgg[mm.key] = aggregate(recsM, cm, { groupKey: groupKey }); });
        body.appendChild(el("div", { class: "rms-section", text: gran === "week" ? "Week over week" : "Month over month" }));
        var chartItems = moList.map(function (mm) { return { label: mm.label, value: perAgg[mm.key].total.sumPaid }; });
        var f = moList[0], l = moList[moList.length - 1], fa = perAgg[f.key].total, la = perAgg[l.key].total;
        var cnote = f.label + ": " + usd(fa.sumPaid) + " (" + pct(fa.pctAmount) + " collected) → " + l.label + ": " + usd(la.sumPaid) + " (" + pct(la.pctAmount) + ")";
        if (dropped > 0) cnote += "  • latest " + winK.length + " " + pnoun + "s (" + dropped + " earlier not shown)";
        body.appendChild(el("div", { class: "rms-grid one" }, [ card("Sum of Copay Paid by " + pnoun + " — whole clinic", cnote, hbar(chartItems, { fmt: function (v) { return usd(v); }, unit: " collected" })) ]));
        body.appendChild(el("div", { class: "rms-section", text: "Copay collected by Check-in User per " + pnoun }));
        body.appendChild(copayMonthMatrix(moList, perAgg));
      }

      // A filtered employee with no copay rows for the SELECTED month still shows zeros (KPIs above) instead of a blank page.
      if (!subset.length) { body.appendChild(el("div", { class: "empty-note", text: "No copay rows for this " + pnoun + " / selection." })); return; }

      // Copay opportunities — collections BY COUNT (opportunities = # of copays due). Paid copays filled
      // inside each employee's copays due, same progress-bar look as the dollar view below it.
      body.appendChild(el("div", { class: "rms-section", text: "Copay opportunities collected" }));
      var pvdOpp = A.rows.slice().sort(function (a, b) { return b.countDue - a.countDue || b.countPaid - a.countPaid; }).slice(0, 15)
        .map(function (r) { return { label: r.employee, due: r.countDue, paid: r.countPaid }; });
      body.appendChild(el("div", { class: "rms-grid one" }, [
        card("Copay opportunities collected", "Copays collected out of copays due (opportunities) • ranked by opportunities • " + num(t.countPaid) + " of " + num(t.countDue) + " copays overall (" + pct(t.pctCount) + ")",
          paidVsDueBar(pvdOpp, { fmt: num, joiner: " of ", suffix: " copays" }))
      ]));

      body.appendChild(el("div", { class: "rms-section", text: "Collections by employee ($)" }));
      var pvd = A.rows.slice().sort(function (a, b) { return b.sumDue - a.sumDue || b.sumPaid - a.sumPaid; }).slice(0, 15)
        .map(function (r) { return { label: r.employee, due: r.sumDue, paid: r.sumPaid }; });
      body.appendChild(el("div", { class: "rms-grid one" }, [
        card("Copay collected vs. due", "Dollars paid versus amount due • ranked by copay due • " + usd(t.sumPaid) + " out of " + usd(t.sumDue) + " overall", paidVsDueBar(pvd))
      ]));

      body.appendChild(el("div", { class: "rms-section", text: "Copay detail by Check-in User" }));
      body.appendChild(copayTable(A.rows, t));
    }

    function rerender() { renderHead(); renderBody(); }
    rerender();
    wrap.appendChild(el("div", { class: "rms-foot" }, [ el("span", { text: "St. Luke's • Copays" }), el("span", { class: "spacer" }), el("span", { text: gen ? "Generated: " + gen : "" }) ]));
    // Handle for the export's shared month picker — "" / "all" means every loaded month.
    return { setPeriod: function (pd) {
      pd = pd || "all";
      if (pd === state.period) return;
      state.period = pd; rerender();   // a month with no copay rows renders zeros + its own empty note
    } };
  }

  window.RMSCopays = { mount: mount, aggregate: aggregate, detectCopayCols: detectCopayCols, looksLikeCopays: looksLikeCopays, cleanName: cleanName, money: money };
})();