
/* ================= RMS Bypassed Warnings — dashboard viewer =================
   window.RMSViewer.mount(rootEl, payload)
   payload = {
     level: 'overview'|'department'|'employee',
     dept?: <department value>, emp?: <employee value>,
     records: [ {<Original Header>: value, ...}, ... ],   // raw rows
     headers: [<Original Header>, ...],                    // column order for tables
     colMap: { date, time, encDate, patient, messageType, errorDetail, errorText,
               workflow, loginDept, checkinUser, user },   // role -> original header
     deptKey: <header used to group departments>,
     empKey:  <header used to group employees>,
     title, generatedAt
   }
   Self-contained: embedded verbatim in every exported file and injected into the live app.  */
(function () {
  "use strict";
  var WD = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var WDL = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  var BLANK = "(blank)";

  /* ---------- DOM helpers ---------- */
  function el(tag, attrs, kids) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === "class") e.className = attrs[k];
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k === "text") e.textContent = attrs[k];
      else if (k.slice(0,2) === "on" && typeof attrs[k] === "function") e.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] != null) e.setAttribute(k, attrs[k]);
    }
    if (kids != null) (Array.isArray(kids) ? kids : [kids]).forEach(function (c) {
      if (c == null) return; e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    });
    return e;
  }
  var SVGNS = "http://www.w3.org/2000/svg";
  function svg(tag, attrs) { var e = document.createElementNS(SVGNS, tag); if (attrs) for (var k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]); return e; }
  function fmt(n) { return (n == null ? 0 : n).toLocaleString("en-US"); }
  function esc(s) { return String(s == null ? "" : s); }
  function trunc(s, n) { s = esc(s); return s.length > n ? s.slice(0, n - 1) + "…" : s; }

  /* ---------- value / role access ---------- */
  function val(r, header) { if (!header) return ""; var v = r[header]; return v == null ? "" : String(v).trim(); }
  function keyOf(r, header) { var v = val(r, header); return v === "" ? BLANK : v; }
  // Month-over-month record sets can arrive pre-collapsed (one row per month × group, carrying its
  // count in __n) so an exported file can hold every month's trend without every month's rows.
  function wOf(r) { var n = +r.__n; return n > 0 ? n : 1; }

  /* ---------- date/time parsing ---------- */
  function parseDate(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim();
    var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) { var y = +m[3]; if (y < 100) y += 2000; return dobj(y, +m[1], +m[2]); }
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return dobj(+m[1], +m[2], +m[3]);
    var d = new Date(s); if (!isNaN(d)) return dobj(d.getFullYear(), d.getMonth()+1, d.getDate());
    return null;
  }
  function dobj(y, mo, d) { var dt = new Date(y, mo - 1, d);
    return { y: y, mo: mo, d: d, ts: dt.getTime(), wd: dt.getDay(),
      key: y + "-" + String(mo).padStart(2,"0") + "-" + String(d).padStart(2,"0"),
      label: String(mo).padStart(2,"0") + "/" + String(d).padStart(2,"0") }; }
  function parseHour(v) {
    if (v == null || v === "") return null;
    var s = String(v).trim();
    var m = s.match(/(\d{1,2}):(\d{2})\s*([AaPp])/); if (m) { var h = +m[1] % 12; if (/[Pp]/.test(m[3])) h += 12; return h; }
    m = s.match(/^(\d{1,2}):(\d{2})/); if (m) return +m[1];
    return null;
  }
  function hourLabel(h) { var ap = h < 12 ? "a" : "p"; var hh = h % 12; if (hh === 0) hh = 12; return hh + ap; }
  function monthName(m) { return ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m] || m; }

  /* ---------- period granularity (month | week) ----------
     A period key is DERIVED from each row's date, never read from a column, so granularity is purely
     a view concern: the historical archive and the executive rollup keep their YYYY-MM keys whatever
     is selected here. Week keys are the MONDAY of the week as "YYYY-MM-DD" — they sort correctly, they
     never collide with a "YYYY-MM" month key, and they sidestep the ISO week-year edge cases (week 53,
     a Jan 1 that belongs to the previous year) that a "2026-W01" key would drag in.
     Week granularity exists for the first month with a new clinic, when a month is too coarse to show
     anyone whether the week's coaching landed. */
  var GRAN = "month";
  function setGranularity(g) { GRAN = (g === "week") ? "week" : "month"; }
  function granularity() { return GRAN; }
  function isWeekKey(p) { return /^\d{4}-\d{2}-\d{2}$/.test(String(p)); }
  function pad2v(n) { return String(n).length < 2 ? "0" + n : String(n); }
  // Monday of the week containing y/mo/d, as a Date.
  function weekStartOf(y, mo, d) { var dt = new Date(y, mo - 1, d); dt.setDate(dt.getDate() - ((dt.getDay() + 6) % 7)); return dt; }
  function weekKeyOf(y, mo, d) { var w = weekStartOf(y, mo, d); return w.getFullYear() + "-" + pad2v(w.getMonth() + 1) + "-" + pad2v(w.getDate()); }
  // dd = a parseDate() result. `gran` overrides the active granularity (registrations force "month").
  function periodKeyOf(dd, gran) {
    if (!dd) return "";
    if ((gran || GRAN) === "week") return weekKeyOf(dd.y, dd.mo, dd.d);
    return dd.y + "-" + pad2v(dd.mo);
  }
  function periodNoun(plural) { return GRAN === "week" ? (plural ? "weeks" : "week") : (plural ? "months" : "month"); }
  // "Apr 6" / "Apr 2026" for a table heading; long form for a tooltip or a sentence.
  function periodKeyLabel(p, long) {
    p = String(p);
    if (isWeekKey(p)) { var a = p.split("-");
      return (long ? "Week of " : "Wk ") + monthName(+a[1]) + " " + (+a[2]) + (long ? ", " + a[0] : ""); }
    var s = p.split("-"); return monthName(+s[1]) + " " + s[0];
  }

  /* ---------- counter ---------- */
  function counter() { var m = new Map(); return {
    add: function (k, n) { if (k == null || k === "") k = BLANK; m.set(k, (m.get(k) || 0) + (n || 1)); }, map: m,
    top: function (n) { var a = Array.from(m, function (e) { return { label: e[0], value: e[1] }; });
      a.sort(function (x, y) { return y.value - x.value || String(x.label).localeCompare(String(y.label)); }); return n ? a.slice(0, n) : a; },
    size: function () { return m.size; } }; }

  /* ---------- summarize a record set against a colMap ---------- */
  function summarize(recs, cm) {
    var byType = counter(), byDetail = counter(), byWf = counter(), byText = counter(),
        byPatient = counter(), byLoginDept = counter(), byDay = counter(),
        byWd = new Array(7).fill(0), byHour = new Array(24).fill(0), hourKnown = false;
    var dates = [], patients = new Set();
    recs.forEach(function (r) {
      byType.add(val(r, cm.messageType)); byDetail.add(val(r, cm.errorDetail)); byWf.add(val(r, cm.workflow));
      byText.add(val(r, cm.errorText)); byLoginDept.add(val(r, cm.loginDept));
      var p = val(r, cm.patient); if (p) { patients.add(p); byPatient.add(p); }
      var pd = parseDate(val(r, cm.date)); if (pd) { byDay.add(pd.key); byWd[pd.wd]++; dates.push(pd); }
      var h = parseHour(val(r, cm.time)); if (h != null && h >= 0 && h < 24) { byHour[h]++; hourKnown = true; }
    });
    dates.sort(function (a, b) { return a.ts - b.ts; });
    var peakWd = -1, peakWdN = -1; byWd.forEach(function (n, i) { if (n > peakWdN) { peakWdN = n; peakWd = i; } });
    var peakHour = -1, peakHourN = -1; byHour.forEach(function (n, i) { if (n > peakHourN) { peakHourN = n; peakHour = i; } });
    return { recs: recs, total: recs.length, byType: byType, byDetail: byDetail, byWf: byWf, byText: byText,
      byPatient: byPatient, byLoginDept: byLoginDept, byDay: byDay, byWd: byWd, byHour: byHour, hourKnown: hourKnown,
      patients: patients.size, first: dates[0], last: dates[dates.length - 1], daysActive: byDay.size(),
      peakDay: byDay.top(1)[0], peakWd: peakWd, peakWdN: peakWdN, peakHour: peakHour, peakHourN: peakHourN };
  }

  /* ---------- tooltip ---------- */
  var tip;
  function ensureTip() { if (!tip) { tip = el("div", { class: "rms-tip" }); document.body.appendChild(tip); } }
  function showTip(html, x, y) { return; ensureTip(); tip.innerHTML = html; tip.style.opacity = "1";
    var w = tip.offsetWidth, h = tip.offsetHeight, left = x + 14, top = y - h - 10;
    if (left + w > window.innerWidth - 8) left = x - w - 14; if (top < 8) top = y + 16;
    tip.style.left = left + "px"; tip.style.top = top + "px"; }
  function hideTip() { if (tip) tip.style.opacity = "0"; }

  /* ---------- charts ---------- */
  function hbar(items, opts) {
    opts = opts || {};
    var wrap = el("div", { class: "chart" });
    if (!items.length) { wrap.appendChild(el("div", { class: "empty-note", text: "No data" })); return wrap; }
    var max = Math.max.apply(null, items.map(function (i) { return i.value; })) || 1;
    var rowH = 42, gap = 6, top = 2, W = 640, valW = 70, plotW = W - valW - 2, barH = 16, barY = 22;
    var H = top * 2 + items.length * rowH + (items.length - 1) * gap;
    var s = svg("svg", { viewBox: "0 0 " + W + " " + H, role: "img" });
    items.forEach(function (it, i) {
      var y = top + i * (rowH + gap), bw = Math.max(2, (it.value / max) * plotW);
      s.appendChild(svg("text", { x: 1, y: y + 13, "text-anchor": "start", class: "barcat" })).textContent = trunc(it.label, 90);
      s.appendChild(svg("rect", { x: 1, y: y + barY, width: plotW, height: barH, rx: 5, fill: "var(--surface-2)" }));
      s.appendChild(svg("rect", { x: 1, y: y + barY, width: bw, height: barH, rx: 5, class: "bar" + (opts.flagFn && opts.flagFn(it) ? " flag" : "") }));
      s.appendChild(svg("text", { x: 1 + bw + 7, y: y + barY + 12, class: "barlabel" })).textContent = opts.fmt ? opts.fmt(it.value) : fmt(it.value);
      var hit = svg("rect", { x: 0, y: y, width: W, height: rowH, class: "hit" + (opts.onClick ? " clickable" : "") });
      var lbl = it.label, v = it.value, extra = it.tip || "";
      hit.addEventListener("mousemove", function (e) { showTip("<b>" + esc(lbl) + "</b><br>" + (opts.fmt ? opts.fmt(v) : fmt(v)) + (opts.unit || " warnings") + (opts.onClick ? "<br><span class='t-k'>click to open</span>" : "") + (extra ? "<br><span class='t-k'>" + esc(extra) + "</span>" : ""), e.clientX, e.clientY); });
      hit.addEventListener("mouseleave", hideTip);
      if (opts.onClick) hit.addEventListener("click", function () { hideTip(); opts.onClick(lbl); });
      s.appendChild(hit);
    });
    wrap.appendChild(s); return wrap;
  }
  function vbar(items, opts) {
    opts = opts || {};
    var wrap = el("div", { class: "chart" });
    if (!items.length || !items.some(function (i) { return i.value; })) { wrap.appendChild(el("div", { class: "empty-note", text: opts.emptyMsg || "No data" })); return wrap; }
    var max = Math.max.apply(null, items.map(function (i) { return i.value; })) || 1;
    var W = 640, H = 190, padB = 26, padT = 16, padX = 6, n = items.length, slot = (W - padX * 2) / n, bw = Math.min(46, slot * 0.62);
    var s = svg("svg", { viewBox: "0 0 " + W + " " + H, role: "img" }), base = H - padB;
    s.appendChild(svg("line", { x1: 0, y1: base + .5, x2: W, y2: base + .5, stroke: "var(--border)" }));
    items.forEach(function (it, i) {
      var cx = padX + slot * i + slot / 2, bh = (it.value / max) * (base - padT), x = cx - bw / 2, y = base - bh;
      s.appendChild(svg("rect", { x: x, y: y, width: bw, height: Math.max(bh, it.value ? 2 : 0), rx: 4, class: "bar" + (opts.flagIdx === i ? " flag" : "") }));
      if (it.value && (bw > 20 || max < 30)) s.appendChild(svg("text", { x: cx, y: y - 5, "text-anchor": "middle", class: "barlabel" })).textContent = fmt(it.value);
      s.appendChild(svg("text", { x: cx, y: base + 16, "text-anchor": "middle", class: "barcat", "font-size": "11" })).textContent = it.label;
      var hit = svg("rect", { x: padX + slot * i, y: 0, width: slot, height: base, class: "hit" }), lbl = it.full || it.label, v = it.value;
      hit.addEventListener("mousemove", function (e) { showTip("<b>" + esc(lbl) + "</b><br>" + fmt(v) + (opts.unit || " warnings"), e.clientX, e.clientY); });
      hit.addEventListener("mouseleave", hideTip); s.appendChild(hit);
    });
    wrap.appendChild(s); return wrap;
  }
  function timeline(days) {
    var wrap = el("div", { class: "chart" });
    if (!days.length) { wrap.appendChild(el("div", { class: "empty-note", text: "No dated records" })); return wrap; }
    var W = 960, H = 210, padL = 34, padR = 12, padT = 14, padB = 26;
    var max = Math.max.apply(null, days.map(function (d) { return d.value; })) || 1, n = days.length;
    var plotW = W - padL - padR, plotH = H - padT - padB, base = H - padB;
    function X(i) { return padL + (n === 1 ? plotW / 2 : (i / (n - 1)) * plotW); }
    function Y(v) { return base - (v / max) * plotH; }
    var s = svg("svg", { viewBox: "0 0 " + W + " " + H, role: "img" });
    niceTicks(max, 3).forEach(function (t) { var y = Y(t), g = svg("g", { class: "grid" });
      g.appendChild(svg("line", { x1: padL, y1: y, x2: W - padR, y2: y }));
      g.appendChild(svg("text", { x: padL - 6, y: y + 3, "text-anchor": "end", fill: "var(--text-3)", "font-size": "10" })).textContent = fmt(t); s.appendChild(g); });
    var lPath = "", dPath = "";
    days.forEach(function (d, i) { var x = X(i), y = Y(d.value); lPath += (i ? "L" : "M") + x.toFixed(1) + " " + y.toFixed(1) + " "; dPath += (i ? "L" : "M") + x.toFixed(1) + " " + y.toFixed(1) + " "; });
    s.appendChild(svg("path", { d: dPath + "L" + X(n - 1).toFixed(1) + " " + base + " L" + X(0).toFixed(1) + " " + base + " Z", class: "area" }));
    s.appendChild(svg("path", { d: lPath, class: "line" }));
    var step = Math.max(1, Math.ceil(n / 7));
    for (var i = 0; i < n; i += step) s.appendChild(svg("text", { x: X(i), y: base + 16, "text-anchor": "middle", class: "barcat", "font-size": "10" })).textContent = days[i].label;
    if (n <= 40) days.forEach(function (d, i) { s.appendChild(svg("circle", { cx: X(i), cy: Y(d.value), r: 3, class: "dot" })); });
    var cross = svg("line", { class: "cross", y1: padT, y2: base, x1: -9, x2: -9, style: "opacity:0" }), focus = svg("circle", { r: 4.5, class: "dot", style: "opacity:0" });
    s.appendChild(cross); s.appendChild(focus);
    var overlay = svg("rect", { x: padL, y: 0, width: plotW, height: base, class: "hit" });
    overlay.addEventListener("mousemove", function (e) { var rect = s.getBoundingClientRect(), rel = (e.clientX - rect.left) / rect.width * W;
      var i = Math.max(0, Math.min(n - 1, Math.round((rel - padL) / (plotW || 1) * (n - 1)))), x = X(i), y = Y(days[i].value);
      cross.setAttribute("x1", x); cross.setAttribute("x2", x); cross.style.opacity = "1"; focus.setAttribute("cx", x); focus.setAttribute("cy", y); focus.style.opacity = "1";
      showTip("<b>" + esc(days[i].full) + "</b><br>" + fmt(days[i].value) + " warnings", e.clientX, e.clientY); });
    overlay.addEventListener("mouseleave", function () { cross.style.opacity = "0"; focus.style.opacity = "0"; hideTip(); });
    s.appendChild(overlay); wrap.appendChild(s); return wrap;
  }
  function niceTicks(max, count) { var step = Math.max(1, Math.ceil(max / count)), mag = Math.pow(10, Math.floor(Math.log10(step)));
    step = Math.ceil(step / mag) * mag; var out = []; for (var v = step; v <= max + step * 0.001; v += step) out.push(v); if (!out.length) out.push(max); return out; }

  /* ---------- building blocks ---------- */
  function kpi(label, value, sub, cls, icon) { return el("div", { class: "kpi" + (cls ? " " + cls : "") }, [
    el("div", { class: "k-label" }, [icon ? iconEl(icon) : null, label]), el("div", { class: "k-val", html: value }), el("div", { class: "k-sub", text: sub || "" }) ]); }
  function iconEl(name) { var p = { alert: "M12 3l9 16H3z M12 10v4 M12 17v.5", clock: "M12 7v5l3 2 M12 3a9 9 0 100 18 9 9 0 000-18z",
    user: "M12 12a4 4 0 100-8 4 4 0 000 8z M5 20a7 7 0 0114 0", building: "M4 21V5a1 1 0 011-1h8a1 1 0 011 1v16 M14 21h5a1 1 0 001-1V9a1 1 0 00-1-1h-5 M7 8h3 M7 12h3 M7 16h3" };
    var s = svg("svg", { width: 14, height: 14, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", "stroke-width": 1.8, "stroke-linecap": "round", "stroke-linejoin": "round" });
    s.appendChild(svg("path", { d: p[name] || p.alert })); return s; }
  function card(title, note, body, span) { return el("div", { class: "card" + (span ? " span2" : "") }, [ el("h3", { text: title }), note ? el("p", { class: "c-note", text: note }) : null, body ]); }
  function flagIcon(hot) { var s = svg("svg", { class: "ic", viewBox: "0 0 24 24", fill: "none", stroke: hot ? "var(--crit)" : "var(--text-3)", "stroke-width": 1.9, "stroke-linecap": "round", "stroke-linejoin": "round" });
    s.appendChild(svg("path", { d: "M12 3l9 16H3z" })); s.appendChild(svg("path", { d: "M12 10v4" })); s.appendChild(svg("path", { d: "M12 16.5v.5" })); return s; }
  function flagList(items, opts) { opts = opts || {}; var wrap = el("div", { class: "flaglist" });
    if (!items.length) { wrap.appendChild(el("div", { class: "empty-note", text: "No data" })); return wrap; }
    items.forEach(function (it) { var crit = it.value >= (opts.hotAt || 3);
      wrap.appendChild(el("div", { class: "flagrow" + (crit && !opts.cool ? " hot" : "") + (opts.cool ? " cool" : "") }, [
        flagIcon(crit && !opts.cool), el("div", { class: "txt" }, [ el("b", { text: trunc(it.label, opts.labelChars || 34) }), el("div", { class: "m", text: (opts.cool ? "overridden " : "") + (crit ? "⚠ repeated" : "") }) ]),
        el("div", { class: "n" + (opts.cool || !crit ? " cool" : ""), text: fmt(it.value) + "×" }) ])); });
    return wrap; }

  /* ---------- day / hour series ---------- */
  /* How the "over time" line is bucketed. At month granularity a long span still collapses to months
     (one point per day over two years is a smear); at week granularity anything past a fortnight is
     drawn as weeks, which is the whole reason for the setting. Exposed so the card note can say what
     the reader is actually looking at instead of always claiming "By date". */
  function daySeriesMode(S) {
    if (!S.first) return "day";
    var span = Math.round((new Date(S.last.y, S.last.mo - 1, S.last.d) - new Date(S.first.y, S.first.mo - 1, S.first.d)) / 86400000);
    if (GRAN === "week") return span > 13 ? "week" : "day";
    return span > 180 ? "month" : "day";
  }
  function buildDaySeries(S) {
    if (!S.first) return [];
    var cur = new Date(S.first.y, S.first.mo - 1, S.first.d), end = new Date(S.last.y, S.last.mo - 1, S.last.d);
    var mode = daySeriesMode(S);
    if (mode === "month") { var mmap = new Map(); S.byDay.map.forEach(function (v, k) { var mk = k.slice(0, 7); mmap.set(mk, (mmap.get(mk) || 0) + v); });
      return Array.from(mmap.keys()).sort().map(function (k) { var p = k.split("-"); return { label: p[1] + "/" + p[0].slice(2), full: monthName(+p[1]) + " " + p[0], value: mmap.get(k) }; }); }
    if (mode === "week") {
      // Sum each day into its Monday, then walk Monday to Monday so a week with no warnings is a
      // zero on the line rather than a missing point that closes the gap and hides it.
      var wmap = new Map();
      S.byDay.map.forEach(function (v, k) { var p = k.split("-"); var wk = weekKeyOf(+p[0], +p[1], +p[2]); wmap.set(wk, (wmap.get(wk) || 0) + v); });
      var w = weekStartOf(S.first.y, S.first.mo, S.first.d), wEnd = weekStartOf(S.last.y, S.last.mo, S.last.d);
      var wout = [], wguard = 0;
      while (w <= wEnd && wguard++ < 600) {
        var wk2 = w.getFullYear() + "-" + pad2v(w.getMonth() + 1) + "-" + pad2v(w.getDate());
        wout.push({ label: pad2v(w.getMonth() + 1) + "/" + pad2v(w.getDate()),
                    full: "Week of " + monthName(w.getMonth() + 1) + " " + w.getDate(), value: wmap.get(wk2) || 0 });
        w = new Date(w.getFullYear(), w.getMonth(), w.getDate() + 7);
      }
      return wout;
    }
    var out = [], guard = 0;
    while (cur <= end && guard++ < 4000) { var key = cur.getFullYear() + "-" + String(cur.getMonth()+1).padStart(2,"0") + "-" + String(cur.getDate()).padStart(2,"0");
      out.push({ label: String(cur.getMonth()+1).padStart(2,"0") + "/" + String(cur.getDate()).padStart(2,"0"), full: WDL[cur.getDay()] + " " + monthName(cur.getMonth()+1) + " " + cur.getDate(), value: S.byDay.map.get(key) || 0 }); cur.setDate(cur.getDate() + 1); }
    return out;
  }
  function buildHourSeries(S) { var lo = 24, hi = -1; S.byHour.forEach(function (n, i) { if (n) { if (i < lo) lo = i; if (i > hi) hi = i; } });
    if (hi < 0) return []; lo = Math.max(0, lo - 1); hi = Math.min(23, hi + 1); var out = []; for (var h = lo; h <= hi; h++) out.push({ label: hourLabel(h), full: hourLabel(h).replace("a"," AM").replace("p"," PM"), value: S.byHour[h] }); return out; }

  /* ---------- summary table (grouped) ---------- */
  function summaryTable(recs, cm, groupKey, colLabel, onClick) {
    var groups = new Map();
    recs.forEach(function (r) { var k = keyOf(r, groupKey); if (!groups.has(k)) groups.set(k, { n: 0, pts: new Set(), types: new Set(), top: counter(), last: null });
      var g = groups.get(k); g.n++; var p = val(r, cm.patient); if (p) g.pts.add(p); var mt = val(r, cm.messageType); if (mt) g.types.add(mt); g.top.add(mt);
      var d = parseDate(val(r, cm.date)); if (d && (!g.last || d.ts > g.last.ts)) g.last = d; });
    var rows = Array.from(groups, function (e) { return { name: e[0], n: e[1].n, pts: e[1].pts.size, types: e[1].types.size, top: (e[1].top.top(1)[0] || {}).label || "—", last: e[1].last }; });
    rows.sort(function (a, b) { return b.n - a.n; });
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table" });
    t.appendChild(el("thead", {}, el("tr", {}, [ th(colLabel), th("Warnings", "num"), th("Patients", "num"), th("Warning types", "num"), th("Top warning type"), th("Last activity") ])));
    var tb = el("tbody");
    rows.forEach(function (r) { tb.appendChild(el("tr", {}, [
      el("td", {}, onClick ? el("button", { class: "rms-linkbtn", text: r.name, onclick: function () { onClick(r.name); } }) : document.createTextNode(r.name)),
      el("td", { class: "num" }, el("span", { class: "pill", text: fmt(r.n) })), el("td", { class: "num", text: fmt(r.pts) }), el("td", { class: "num", text: fmt(r.types) }),
      el("td", { text: trunc(r.top, 30) }), el("td", { text: r.last ? r.last.label + "/" + String(r.last.y).slice(2) : "—" }) ])); });
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }
  function th(label, cls) { return el("th", { class: cls || "" }, label); }

  /* Expandable employee table — each row drills down to a count of every warning type */
  function employeeDrillTable(recs, cm, groupKey, colLabel, onOpen) {
    var groups = new Map();
    recs.forEach(function (r) { var k = keyOf(r, groupKey); if (!groups.has(k)) groups.set(k, { recs: [], pts: new Set(), types: counter(), detail: counter(), last: null });
      var g = groups.get(k); g.recs.push(r); var p = val(r, cm.patient); if (p) g.pts.add(p); g.types.add(val(r, cm.messageType)); g.detail.add(val(r, cm.errorDetail));
      var d = parseDate(val(r, cm.date)); if (d && (!g.last || d.ts > g.last.ts)) g.last = d; });
    var rows = Array.from(groups, function (e) { return { name: e[0], g: e[1] }; });
    rows.sort(function (a, b) { return b.g.recs.length - a.g.recs.length; });
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table drill" });
    t.appendChild(el("thead", {}, el("tr", {}, [ th(""), th(colLabel), th("Warnings", "num"), th("Patients", "num"), th("Warning types", "num"), th("Last activity"), th("") ])));
    var tb = el("tbody");
    rows.forEach(function (r) {
      var g = r.g, open = false;
      var chev = el("span", { class: "exp-chev", text: "▸" });
      var detailRow = el("tr", { class: "drill-detail" });
      var td = el("td", { colspan: 7 });
      var typeItems = g.types.top(20).filter(function (it) { return it.label !== "(blank)"; });
      var detItems = g.detail.top(20).filter(function (it) { return it.label !== "(blank)"; });
      td.appendChild(el("div", { class: "emp-detail" }, [
        el("div", { class: "emp-detail-grid" }, [
          el("div", {}, [ el("div", { class: "emp-detail-h", text: "Count by warning type" }), hbar(typeItems, { labelW: 190, labelChars: 30 }) ]),
          el("div", {}, [ el("div", { class: "emp-detail-h", text: "Count by reason / error detail" }), hbar(detItems, { labelW: 190, labelChars: 30 }) ])
        ])
      ]));
      detailRow.appendChild(td); detailRow.style.display = "none";
      var mainRow = el("tr", { class: "drill-main" }, [
        el("td", { class: "exp-cell" }, chev),
        el("td", {}, onOpen ? el("button", { class: "rms-linkbtn", text: r.name, onclick: function (ev) { ev.stopPropagation(); onOpen(r.name); } }) : document.createTextNode(r.name)),
        el("td", { class: "num" }, el("span", { class: "pill", text: fmt(g.recs.length) })),
        el("td", { class: "num", text: fmt(g.pts.size) }),
        el("td", { class: "num", text: fmt(g.types.size()) }),
        el("td", { text: g.last ? g.last.label + "/" + String(g.last.y).slice(2) : "—" }),
        el("td", { class: "exp-hint", text: "details" })
      ]);
      function toggle() { open = !open; detailRow.style.display = open ? "" : "none"; chev.textContent = open ? "▾" : "▸"; mainRow.classList.toggle("open", open); }
      mainRow.addEventListener("click", toggle);
      mainRow.style.cursor = "pointer";
      tb.appendChild(mainRow); tb.appendChild(detailRow);
    });
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }

  /* ---------- records table ---------- */
  function recordsTable(recs, headers, cm) {
    var PAGE = 40, shown = PAGE, q = "", sortH = null, sortDir = 1, filtered = recs.slice();
    var dateH = cm.date, longH = cm.errorText;
    var wrap = el("div", { class: "rms-tablewrap" });
    var search = el("input", { type: "search", placeholder: "Search records…" }), count = el("span", { class: "count" });
    var scroll = el("div", { class: "rms-scroll" }), table = el("table", { class: "rms-table" }), thead = el("thead"), tbody = el("tbody"), moreWrap = el("div", { class: "tb-more" });
    table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table);
    wrap.appendChild(el("div", { class: "rms-tbtools" }, [search, count])); wrap.appendChild(scroll); wrap.appendChild(moreWrap);
    function buildHead() { var tr = el("tr"); headers.forEach(function (h) { var arrow = sortH === h ? el("span", { class: "arrow", text: sortDir > 0 ? " ▲" : " ▼" }) : null;
      tr.appendChild(el("th", { onclick: function () { if (sortH === h) sortDir *= -1; else { sortH = h; sortDir = 1; } apply(); } }, [h, arrow])); }); thead.innerHTML = ""; thead.appendChild(tr); }
    function apply() { var ql = q.trim().toLowerCase();
      filtered = recs.filter(function (r) { if (!ql) return true; return headers.some(function (h) { return String(r[h] == null ? "" : r[h]).toLowerCase().indexOf(ql) >= 0; }); });
      if (sortH) filtered.sort(function (a, b) { var av = a[sortH] == null ? "" : a[sortH], bv = b[sortH] == null ? "" : b[sortH];
        if (sortH === dateH) { var da = parseDate(av), db = parseDate(bv); av = da ? da.ts : 0; bv = db ? db.ts : 0; } return (av > bv ? 1 : av < bv ? -1 : 0) * sortDir; });
      shown = PAGE; render(); }
    function render() { buildHead(); tbody.innerHTML = "";
      filtered.slice(0, shown).forEach(function (r) { var tr = el("tr"); headers.forEach(function (h) { var v = r[h] == null ? "" : String(r[h]);
        tr.appendChild(el("td", { title: h === longH ? v : null, text: trunc(v, h === longH ? 54 : 26) })); }); tbody.appendChild(tr); });
      count.textContent = fmt(filtered.length) + " record" + (filtered.length === 1 ? "" : "s") + (q ? " matched" : "");
      moreWrap.innerHTML = ""; if (shown < filtered.length) moreWrap.appendChild(el("button", { text: "Show " + Math.min(PAGE, filtered.length - shown) + " more (" + fmt(filtered.length - shown) + " hidden)", onclick: function () { shown += PAGE; render(); } })); }
    search.addEventListener("input", function () { q = search.value; apply(); }); apply(); return wrap;
  }

  /* ---------- registrations (Detailed-View join) ---------- */
  function toLastFirst(s) { s = String(s == null ? "" : s).trim(); if (!s) return "";
    if (s.indexOf(",") >= 0) return s.toUpperCase();
    var t = s.split(/\s+/); if (t.length < 2) return s.toUpperCase();
    return (t[t.length - 1] + ", " + t.slice(0, t.length - 1).join(" ")).toUpperCase(); }
  function empNameKey(s) { s = String(s == null ? "" : s).trim(); if (!s) return "";
    var last, first;
    if (s.indexOf(",") >= 0) { var p = s.split(","); last = p[0]; first = (p[1] || "").trim().split(/\s+/)[0] || ""; }
    else { var t = s.split(/\s+/); first = t[0] || ""; last = t.length > 1 ? t[t.length - 1] : ""; }
    function nm(x) { return String(x).toLowerCase().replace(/[^a-z0-9]/g, ""); }
    return nm(last) + "|" + nm(first); }
  function regPeriodLabel(regData, S) { var m = regData && regData.month ? String(regData.month).trim() : ""; if (!m) return "";
    if (/\b\d{4}\b/.test(m)) return m; var y = S && S.first ? S.first.y : (S && S.last ? S.last.y : null); return y ? m + " " + y : m; }
  // Per-employee registration join: warnings, registrations, and the warnings-per-registration
  // ratio for each person. Shared by the registrations table and the per-employee leaderboard chart
  // so both show identical numbers.
  function registrationRows(recs, cm, empKey, regData, single, periods) {
    var lookup = {};
    (regData.entries || []).forEach(function (e) {
      if (periods && !periods[e.__period]) return;   // only the month(s) whose warnings are on screen
      var k = empNameKey(e.staff); if (k === "|") return;
      if (!lookup[k]) lookup[k] = { name: toLastFirst(e.staff), count: 0 }; lookup[k].count += (+e.count || 0); });
    var map = {};
    if (!single) { for (var lk in lookup) map[lk] = { emp: lookup[lk].name, warnings: 0, regs: lookup[lk].count }; }
    recs.forEach(function (r) { var disp = keyOf(r, empKey), k = empNameKey(disp);
      if (!map[k]) map[k] = { emp: disp, warnings: 0, regs: lookup[k] ? lookup[k].count : 0 };
      map[k].warnings++; map[k].emp = disp; });
    var rows = []; for (var mk in map) { var m = map[mk]; rows.push({ emp: m.emp, warnings: m.warnings, regs: m.regs, ratio: m.regs > 0 ? m.warnings / m.regs : 0 }); }
    return rows;
  }
  function registrationsTable(recs, cm, empKey, regData, single, periods) {
    var rows = registrationRows(recs, cm, empKey, regData, single, periods);
    rows.sort(function (a, b) { return String(a.emp).localeCompare(String(b.emp)); });
    var totW = 0, totR = 0; rows.forEach(function (r) { totW += r.warnings; totR += r.regs; });
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table" });
    t.appendChild(el("thead", {}, el("tr", {}, [ th("Employee"), th("Warnings Bypassed", "num"), th("Registrations", "num"), th("Bypassed Warnings per Registration", "num") ])));
    var tb = el("tbody");
    rows.forEach(function (r) { tb.appendChild(el("tr", {}, [
      el("td", { text: r.emp }),
      el("td", { class: "num" }, el("span", { class: "pill", text: fmt(r.warnings) })),
      el("td", { class: "num", text: fmt(r.regs) }),
      el("td", { class: "num", text: r.ratio.toFixed(2) }) ])); });
    if (!single && rows.length) tb.appendChild(el("tr", { class: "rms-totalrow", style: "border-top:2px solid var(--border-2);font-weight:600" }, [
      el("td", {}, el("b", { text: "TOTALS" })),
      el("td", { class: "num" }, el("b", { text: fmt(totW) })),
      el("td", { class: "num" }, el("b", { text: fmt(totR) })),
      el("td", { class: "num" }, el("b", { text: (totR > 0 ? totW / totR : 0).toFixed(2) })) ]));
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }

  /* ---------- period-over-period (only meaningful when the data spans ≥2 periods, i.e. the
     "All … (combined)" view). Buckets the SAME records by their date column. ---------- */
  /* The periods present in a record set, at the active granularity (or `gran` when forced).
     `partial` marks a bucket the data does not fully cover — the first and last week of an import
     window are almost always part-weeks, and without the flag they read as a real collapse in volume
     rather than as a week that simply has not finished yet. */
  function periodsPresent(recs, cm, gran, win) {
    var c = new Map(), lo = null, hi = null;
    recs.forEach(function (r) {
      var d = parseDate(val(r, cm.date)); if (!d) return;
      if (lo == null || d.ts < lo) lo = d.ts;
      if (hi == null || d.ts > hi) hi = d.ts;
      var k = periodKeyOf(d, gran); c.set(k, (c.get(k) || 0) + wOf(r));
    });
    /* An export collapses each bucket to a single synthetic row, so the real first/last dates are gone
       by the time this runs and every period would look complete. `win` carries the true window from
       the app that wrote the file, which is exactly where the part-week marker matters most. */
    if (win && win.length === 2) {
      var wl = parseDate(win[0]), wh = parseDate(win[1]);
      if (wl) lo = wl.ts;
      if (wh) hi = wh.ts;
    }
    return Array.from(c.keys()).sort().map(function (k) {
      var p = k.split("-"), y = +p[0], mo = +p[1], value = c.get(k);
      if (isWeekKey(k)) {
        var d0 = new Date(y, mo - 1, +p[2]), d1 = new Date(y, mo - 1, +p[2] + 6);
        var partial = (lo != null && d0.getTime() < lo) || (hi != null && d1.getTime() > hi);
        return { key: k, y: y, mo: mo, short: monthName(mo) + " " + (+p[2]), head: monthName(mo) + " " + (+p[2]),
                 full: "Week of " + monthName(mo) + " " + (+p[2]) + ", " + y, value: value, partial: partial };
      }
      return { key: k, y: y, mo: mo, short: monthName(mo), head: monthName(mo) + " " + y,
               full: monthName(mo) + " " + y, value: value, partial: false };
    });
  }
  // Registrations are uploaded a month at a time, so anything that has to line up with them asks for
  // month buckets explicitly rather than inheriting the active granularity.
  function monthsPresent(recs, cm) { return periodsPresent(recs, cm, "month"); }
  // group (employee / department) × period grid of warning counts, ranked by total, with a TOTALS row.
  function periodMatrixTable(recs, cm, groupKey, colLabel, months, onClick) {
    var groups = new Map();
    recs.forEach(function (r) {
      var d = parseDate(val(r, cm.date)); if (!d) return;   // rows without a parseable date can't be placed in a period
      var mk = periodKeyOf(d), k = keyOf(r, groupKey);
      if (!groups.has(k)) groups.set(k, { name: k, m: {}, total: 0 });
      var g = groups.get(k), w = wOf(r); g.m[mk] = (g.m[mk] || 0) + w; g.total += w;
    });
    /* Rank and total on the periods actually shown. The grid is windowed at week granularity
       (52 columns a year will not fit on a page), so summing every period would print a Total that
       does not add up across the row the reader can see. */
    var rows = Array.from(groups.values()).map(function (g) {
      var shown = 0; months.forEach(function (mm) { shown += g.m[mm.key] || 0; });
      g.shown = shown; return g;
    }).filter(function (g) { return g.shown > 0; })
      .sort(function (a, b) { return b.shown - a.shown || String(a.name).localeCompare(String(b.name)); });
    var wrap = el("div", { class: "rms-tablewrap" }), scroll = el("div", { class: "rms-scroll" }), t = el("table", { class: "rms-table" });
    var head = [ th(colLabel) ];
    months.forEach(function (mm) { head.push(th(mm.head + (mm.partial ? " *" : ""), "num")); });
    head.push(th("Total", "num"));
    t.appendChild(el("thead", {}, el("tr", {}, head)));
    var tb = el("tbody"), colTot = {}, grand = 0; months.forEach(function (mm) { colTot[mm.key] = 0; });
    rows.forEach(function (g) {
      var tds = [ el("td", { text: String(g.name) }) ];
      months.forEach(function (mm) { var n = g.m[mm.key] || 0; colTot[mm.key] += n; tds.push(el("td", { class: "num", text: fmt(n) })); });
      tds.push(el("td", { class: "num" }, el("b", { text: fmt(g.shown) }))); grand += g.shown;
      var tr = el("tr", {}, tds);
      if (onClick) { tr.style.cursor = "pointer"; tr.addEventListener("click", function () { onClick(g.name); }); }
      tb.appendChild(tr);
    });
    if (rows.length) {
      var trT = [ el("td", {}, el("b", { text: "TOTALS" })) ];
      months.forEach(function (mm) { trT.push(el("td", { class: "num" }, el("b", { text: fmt(colTot[mm.key]) }))); });
      trT.push(el("td", { class: "num" }, el("b", { text: fmt(grand) })));
      tb.appendChild(el("tr", { class: "rms-totalrow", style: "border-top:2px solid var(--border-2);font-weight:600" }, trT));
    }
    t.appendChild(tb); scroll.appendChild(t); wrap.appendChild(scroll); return wrap;
  }

  /* ---------- registrations ↔ month alignment ----------
     Registration counts are uploaded one month at a time, so they only line up with the warnings on
     screen when EVERY month being shown has its own registration counts. A Mar-only registrations
     file divided into a Jan–Mar warning count produces a ratio that is wrong and a heading that
     names the wrong period — so in that case the per-registration table/chart is held back until a
     single month is selected. Entries from older archives carry no period; those are used as-is. */
  function periodText(p) { var s = String(p).split("-"); return monthName(+s[1]) + " " + s[0]; }
  function regScope(recs, cm, regData) {
    if (!(regData && regData.entries && regData.entries.length)) return null;
    var have = {};
    regData.entries.forEach(function (e) { if (e.__period && /^\d{4}-\d{2}$/.test(e.__period)) have[e.__period] = 1; });
    var haveKeys = Object.keys(have).sort();
    var shown = monthsPresent(recs, cm);
    // untagged registrations, or records with no usable dates → keep the pre-period behaviour
    if (!haveKeys.length || !shown.length) return { ok: true, periods: null, label: null };
    var covered = shown.filter(function (m) { return have[m.key]; });
    var missing = shown.filter(function (m) { return !have[m.key]; });
    if (missing.length || !covered.length) {
      return { ok: false, haveLabel: haveKeys.map(periodText).join(", "),
        missingLabel: (missing.length ? missing : shown).map(function (m) { return m.full; }).join(", ") };
    }
    var periods = {}; covered.forEach(function (m) { periods[m.key] = 1; });
    return { ok: true, periods: periods,
      label: covered.length === 1 ? covered[0].full : covered[0].full + " – " + covered[covered.length - 1].full };
  }

  /* ---------- scope renderer ---------- */
  function renderScope(container, recs, cm, headers, ctx) {
    container.innerHTML = "";
    if (!recs.length) {
      // e.g. an employee who has copays/coverage but no bypassed warnings — show zeros, not a blank page.
      var kz = el("div", { class: "rms-kpis" });
      kz.appendChild(kpi("Bypassed Warnings", "0", ctx.kind === "employee" ? "for this " + ctx.empLabel : "", "accent", "alert"));
      kz.appendChild(kpi("Patients affected", "0", "", null, "user"));
      kz.appendChild(kpi("Warning types", "0", ""));
      container.appendChild(kz);
      container.appendChild(el("div", { class: "empty-note", text: "No bypassed warnings for this selection." }));
      return;
    }
    var kind = ctx.kind, S = summarize(recs, cm);
    var range = S.first ? (S.first.label + (S.last && S.last.key !== S.first.key ? " – " + S.last.label : "")) : "—";

    /* KPIs */
    var k = el("div", { class: "rms-kpis" });
    k.appendChild(kpi("Bypassed Warnings", fmt(S.total), kind === "overview" ? "across all " + ctx.deptLabelPlural : (kind === "department" ? "in this " + ctx.deptLabel : "for this " + ctx.empLabel), "accent", "alert"));
    if (kind === "overview") k.appendChild(kpi(cap(ctx.deptLabelPlural), fmt(distinct(recs, ctx.deptKey)), "with ≥1 bypass", null, "building"));
    if (kind === "overview" || kind === "department") k.appendChild(kpi(cap(ctx.empLabelPlural), fmt(distinct(recs, ctx.empKey)), "with ≥1 bypass", null, "user"));
    k.appendChild(kpi("Patients affected", fmt(S.patients), S.total ? (Math.round(S.total / (S.patients || 1) * 10) / 10) + " / patient" : "", null, "user"));
    k.appendChild(kpi("Warning types", fmt(S.byType.size()), "top: " + trunc((S.byType.top(1)[0] || {}).label, 20)));
    container.appendChild(k);

    /* Period over period — the combined view only. A page scoped to one period is that period's page,
       so the whole section (clinic trend + every employee's period-by-period counts) is left off it;
       it returns on "All … (combined)". Built from ctx.monthRecords, which spans every loaded period,
       so the combined view still shows periods the record set is filtered away from. */
    var momRecs = (ctx.monthRecords && ctx.monthRecords.length) ? ctx.monthRecords : recs;
    /* Period over period. At week granularity this is the point of the whole view, so the window is
       kept to the most recent PERIOD_WINDOW buckets: a year of weeks is 52 bars and 52 grid columns,
       which is neither readable on screen nor printable. What gets dropped is said out loud rather
       than silently trimmed. */
    var PERIOD_WINDOW = GRAN === "week" ? 13 : 24;
    var moAll = periodsPresent(recs, cm).length >= 2 ? periodsPresent(momRecs, cm, null, ctx.monthWindow) : [];
    var moM = moAll.slice(-PERIOD_WINDOW), trimmed = moAll.length - moM.length;
    if (moM.length >= 2) {
      container.appendChild(el("div", { class: "rms-section", text: GRAN === "week" ? "Week over week" : "Month over month" }));
      var peakIdx = 0; moM.forEach(function (m, i) { if (m.value > moM[peakIdx].value) peakIdx = i; });
      var mitems = moM.map(function (m) { return { label: m.short + (m.partial ? "*" : ""), full: m.full + (m.partial ? " (part " + periodNoun() + ")" : ""), value: m.value }; });
      var firstM = moM[0], lastM = moM[moM.length - 1], delta = firstM.value ? Math.round((lastM.value - firstM.value) / firstM.value * 100) : 0;
      var trendNote = firstM.full + ": " + fmt(firstM.value) + " → " + lastM.full + ": " + fmt(lastM.value) +
        (firstM.value ? "  (" + (delta >= 0 ? "+" : "") + delta + "% overall)" : "");
      if (trimmed > 0) trendNote += "  • latest " + moM.length + " " + periodNoun(true) + " (" + trimmed + " earlier " + (trimmed === 1 ? periodNoun() : periodNoun(true)) + " not shown)";
      var anyPartial = moM.some(function (m) { return m.partial; });
      if (anyPartial) trendNote += "  • * part " + periodNoun() + " — not fully covered by the imported dates";
      container.appendChild(el("div", { class: "rms-grid one" }, [ card(
        "Bypassed Warnings by " + periodNoun() + (kind === "employee" ? "" : " — whole clinic"),
        trendNote, vbar(mitems, { flagIdx: peakIdx }) ) ]));
      if (kind === "overview" || kind === "department") {
        var gKey = kind === "overview" ? ctx.deptKey : ctx.empKey;
        var gLabel = kind === "overview" ? cap(ctx.deptLabel) : "Employee";
        var gDrill = kind === "overview" ? ctx.onDrillDept : ctx.onDrillEmp;
        container.appendChild(el("div", { class: "rms-section", text: "Bypassed Warnings by " + (kind === "overview" ? ctx.deptLabelPlural : "employee") + " per " + periodNoun() + (gDrill ? " — click a row to open" : "") }));
        container.appendChild(periodMatrixTable(momRecs, cm, gKey, gLabel, moM, gDrill));
      }
    }

    /* Bypassed warnings per registration — directly under the KPI tiles, above the patterns/graphs.
       Needs an imported Detailed-View registrations file, and only for the month(s) it covers.
       A registrations file names its month in the count column's header and carries one figure for the
       whole month, so there is no honest way to cut it into weeks. This section therefore stays MONTHLY
       even at week granularity (regScope/monthsPresent force month buckets), and says so — dividing a
       month's registrations across its weeks would invent a denominator nobody reported. */
    var rs = regScope(recs, cm, ctx.regData);
    var regPeriod = rs && rs.ok ? (rs.label || regPeriodLabel(ctx.regData, S)) : "";
    if ((kind === "department" || kind === "employee") && rs) {
      container.appendChild(el("div", { class: "rms-section", text: "Bypassed Warnings per registration" + (regPeriod ? " — " + regPeriod : "") }));
      if (rs.ok && GRAN === "week") container.appendChild(el("div", { class: "c-note", text: "Registrations are reported monthly, so these rates stay by month while the rest of the page is by week." }));
      if (rs.ok) container.appendChild(registrationsTable(recs, cm, ctx.empKey, ctx.regData, kind === "employee", rs.periods));
      else container.appendChild(el("div", { class: "empty-note", text: "Registration counts are loaded for " + rs.haveLabel +
        " only, so warnings per registration can't be shown for " + rs.missingLabel + ". Pick a single month with registrations to see the rates." }));
    }

    /* Leaderboard + drill */
    if (kind === "overview" || kind === "department") {
      var isOverview = kind === "overview";
      var drill = isOverview ? ctx.onDrillDept : ctx.onDrillEmp;
      // The employee-level leaderboard is per Employee. With a registrations file loaded, plot the
      // warnings-per-registration RATE (identical numbers to the "per registration" table above),
      // ranked worst-first — not a raw count. Falls back to raw counts when no registrations exist.
      var rrows = (!isOverview && rs && rs.ok)
        ? registrationRows(recs, cm, ctx.empKey, ctx.regData, false, rs.periods).filter(function (r) { return r.regs > 0; })
        : [];
      if (rrows.length) {
        var totW = 0, totR = 0; rrows.forEach(function (r) { totW += r.warnings; totR += r.regs; });
        var baseRatio = totR > 0 ? totW / totR : 0;
        var ritems = rrows.slice().sort(function (a, b) { return b.ratio - a.ratio; }).slice(0, 15).map(function (r) {
          return { label: r.emp, value: r.ratio, tip: fmt(r.warnings) + " warnings / " + fmt(r.regs) + " registrations" }; });
        container.appendChild(el("div", { class: "rms-section", text: "Bypassed Warnings per Employee (ranked)" + (regPeriod ? " — " + regPeriod : "") }));
        container.appendChild(el("div", { class: "rms-grid one" }, [ card(
          "Bypassed Warnings per Employee", "Warnings ÷ registrations • " + fmt(rrows.length) + " Employees" + (drill ? " • Click colored bar to expand for employee" : "") + " • Red = Above clinic average",
          hbar(ritems, { unit: " per registration", fmt: function (v) { return v.toFixed(2); }, flagFn: function (it) { return baseRatio > 0 && it.value > baseRatio; }, onClick: drill }) ) ]));
      } else {
        var groupKey = isOverview ? ctx.deptKey : ctx.empKey;
        var groupLabel = isOverview ? ctx.deptLabel : "Employee";
        var groupLabelPlural = isOverview ? ctx.deptLabelPlural : "Employees";
        var lbc = counter(); recs.forEach(function (r) { lbc.add(keyOf(r, groupKey)); });
        var lb = lbc.top(15), mean = S.total / (lbc.size() || 1);
        container.appendChild(el("div", { class: "rms-section", text: "Bypassed" }));
        container.appendChild(el("div", { class: "rms-grid one" }, [ card(
          "Bypassed Warnings by " + groupLabel, "Ranked • " + fmt(lbc.size()) + " " + groupLabelPlural + (drill ? " • Click colored bar to expand for " + groupLabel.toLowerCase() : "") + " • Red = Above clinic average",
          hbar(lb, { labelW: 200, labelChars: 32, flagFn: function (it) { return mean > 0 && it.value > mean; }, onClick: drill }) ) ]));
      }
    }

    /* Charts */
    container.appendChild(el("div", { class: "rms-section", text: "Trends" }));
    var tlMode = daySeriesMode(S), tlNote = (tlMode === "week" ? "By week" : tlMode === "month" ? "By month" : "By date") + " • " + range;
    container.appendChild(el("div", { class: "rms-grid one" }, [ card("Bypassed Warnings over time", tlNote, timeline(buildDaySeries(S))) ]));
    container.appendChild(el("div", { class: "rms-grid" }, [
      card("By warning (message) type", "What kind of alert is overridden", hbar(S.byType.top(8), { labelW: 200, labelChars: 32 })),
      card("By reason / error detail", "Sub-category of the bypassed warning", hbar(S.byDetail.top(8), { labelW: 200, labelChars: 32 }))
    ]));
    container.appendChild(el("div", { class: "rms-grid" }, [
      card("By workflow", "Where in the workflow warnings were bypassed", hbar(S.byWf.top(6), { labelW: 200, labelChars: 32 })),
      card("By login department", "Login department on the record", hbar(S.byLoginDept.top(6), { labelW: 200, labelChars: 32 }))
    ]));

    /* Needs attention */
    container.appendChild(el("div", { class: "rms-section", text: "Needs attention" }));
    container.appendChild(el("div", { class: "rms-grid" }, [
      card("Repeat bypasses by patient", "Same patient overridden repeatedly — possible safety review", flagList(S.byPatient.top(6), { hotAt: 3 })),
      card("Most-overridden warning", "Alerts bypassed most often", flagList(S.byText.top(6), { hotAt: 8, cool: true, labelChars: 60 }))
    ]));

    /* Summary tables / drill-down */
    if (kind === "overview") { container.appendChild(el("div", { class: "rms-section", text: cap(ctx.deptLabel) + " summary" }));
      container.appendChild(summaryTable(recs, cm, ctx.deptKey, cap(ctx.deptLabel), ctx.onDrillDept)); }
    else if (kind === "department") { container.appendChild(el("div", { class: "rms-section", text: cap(ctx.empLabel) + " summary — click a row for details" }));
      container.appendChild(employeeDrillTable(recs, cm, ctx.empKey, cap(ctx.empLabel), ctx.onDrillEmp)); }

    /* Records — employee level only */
    if (kind === "employee") {
      container.appendChild(el("div", { class: "rms-section", text: "All records (" + fmt(recs.length) + ")" }));
      container.appendChild(recordsTable(recs, headers, cm));
    }
  }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }
  function distinct(recs, key) { var s = new Set(); recs.forEach(function (r) { s.add(keyOf(r, key)); }); return s.size; }
  function uniqSorted(recs, key) { var s = new Set(); recs.forEach(function (r) { s.add(keyOf(r, key)); }); return Array.from(s).sort(function (a, b) { return String(a).localeCompare(String(b)); }); }

  /* ---------- exports (self-contained: also runs inside exported files) ---------- */
  function rangeLabel(recs, cm) {
    var lo = null, hi = null; recs.forEach(function (r) { var d = parseDate(val(r, cm.date)); if (d) { if (!lo || d.ts < lo.ts) lo = d; if (!hi || d.ts > hi.ts) hi = d; } });
    function f(d) { return String(d.mo).padStart(2, "0") + "-" + String(d.d).padStart(2, "0") + "-" + d.y; }
    if (!lo) return ""; return lo.key === hi.key ? f(lo) : f(lo) + " to " + f(hi);
  }
  function monthLabel(recs, cm) {
    var lo = null, hi = null; recs.forEach(function (r) { var dd = parseDate(val(r, cm.date)); if (dd) { if (!lo || dd.ts < lo.ts) lo = dd; if (!hi || dd.ts > hi.ts) hi = dd; } });
    if (!lo) return "";
    var M = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    if (lo.y === hi.y && lo.mo === hi.mo) return M[lo.mo - 1] + " " + lo.y;
    if (lo.y === hi.y) return M[lo.mo - 1] + "–" + M[hi.mo - 1] + " " + lo.y;
    return M[lo.mo - 1] + " " + lo.y + " – " + M[hi.mo - 1] + " " + hi.y;
  }
  function safeName(s) { return String(s).replace(/[\\/:*?"<>|\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 90) || "export"; }
  function fileName(name, recs, cm) { var rg = rangeLabel(recs, cm); return safeName(name) + (rg ? " (" + rg + ")" : "") + ".html"; }
  function escHtml(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function safeJSON(o) { return JSON.stringify(o).replace(/</g, "\\u003c").replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029"); }
  function getCopayJS() {
    var needle = "RMSCopays" + " = {";
    var ss = document.getElementsByTagName("script");
    for (var i = 0; i < ss.length; i++) { var t = ss[i].textContent || ""; if (t.indexOf(needle) >= 0) return t; }
    return "";
  }
  // Grab THIS viewer's own source from the page so every exported file (and re-exports made from
  // inside an exported file) runs the current logic instead of a pre-baked snapshot. Matched by the
  // viewer's unique opening comment so the __RMS_ASSETS__ bundle / payload scripts can't be picked.
  function getViewerJS() {
    var SIG = "/* ================= RMS Bypassed Warnings — dashboard viewer";
    var ss = document.getElementsByTagName("script");
    for (var i = 0; i < ss.length; i++) { var t = ss[i].textContent || ""; if (t.replace(/^\s+/, "").indexOf(SIG) === 0) return t; }
    return (window.__RMS_ASSETS__ && window.__RMS_ASSETS__.js) || "";
  }
  function buildHTML(payload, docTitle, copay, coverageSnapshot) {
    var A = window.__RMS_ASSETS__ || { css: "", js: "" };
    var VJS = getViewerJS() || A.js;   // embed the current viewer code (falls back to the pre-baked bundle)
    var TABCSS = "<style>html,body{margin:0}body{background:#f4f4f2}html[data-theme=dark] body{background:#131316}@media(prefers-color-scheme:dark){html:not([data-theme=light]) body{background:#131316}}.xbar{display:flex;align-items:flex-end;gap:8px;max-width:1180px;margin:0 auto;padding:14px 22px 0;flex-wrap:wrap}.xbar .xspacer{flex:1 1 auto;min-width:12px}.xdata{display:flex;align-items:center;gap:7px;margin-bottom:2px;font:640 11px -apple-system,system-ui,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:#86857e}.xdata select{font:600 12.5px -apple-system,system-ui,sans-serif;letter-spacing:0;text-transform:none;color:#16161a;background:#fff;border:1px solid #d3d2cc;border-radius:9px;padding:6px 10px;cursor:pointer}html[data-theme=dark] .xdata select{background:#1c1c20;color:#f4f4f6;border-color:#3a3a42}@media(prefers-color-scheme:dark){html:not([data-theme=light]) .xdata select{background:#1c1c20;color:#f4f4f6;border-color:#3a3a42}}.xtab{font:600 13.5px -apple-system,system-ui,sans-serif;padding:8px 16px;border-radius:10px 10px 0 0;border:1px solid #d3d2cc;border-bottom:none;background:#fff;color:#55544f;cursor:pointer}.xtab.active{color:#1c5cab;border-color:#2a78d6}html[data-theme=dark] .xtab{background:#1c1c20;color:#b7b6b0;border-color:#3a3a42}html[data-theme=dark] .xtab.active{color:#8fbef2;border-color:#3987e5}@media(prefers-color-scheme:dark){html:not([data-theme=light]) .xtab{background:#1c1c20;color:#b7b6b0;border-color:#3a3a42}html:not([data-theme=light]) .xtab.active{color:#8fbef2;border-color:#3987e5}}</style>";
    var HEAD = "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>" + escHtml(docTitle) + "</title>\n" + TABCSS + "\n</head>\n<body>\n";
    /* Coverage Accuracy is a pre-rendered snapshot, so an export carries one per month keyed by
       period ("" = all months combined) and the Data picker swaps them. A plain string still works. */
    var covMap = (coverageSnapshot && typeof coverageSnapshot === "object") ? coverageSnapshot : (coverageSnapshot ? { "": coverageSnapshot } : null);
    var covFirst = covMap ? (covMap[payload.period || ""] || covMap[""] || "") : "";
    var divs = "<div id=\"xtabs\" class=\"xbar\"></div>\n<div id=\"rms-bypass\"></div>\n";
    if (copay) divs += "<div id=\"rms-copays\" style=\"display:none\"></div>\n";
    if (covMap) divs += "<div id=\"rms-coverage\" class=\"rms\" style=\"display:none\">" + covFirst + "</div>\n";
    var data = "<script>window.__RMS_PAYLOAD__=" + safeJSON(payload) + ";window.__RMS_ASSETS__={css:" + safeJSON(A.css) + ",js:" + safeJSON(VJS) + "};";
    if (copay) data += "window.__RMS_COPAY_PAYLOAD__=" + safeJSON(copay) + ";window.__RMS_COPAY_JS__=" + safeJSON(getCopayJS()) + ";";
    if (covMap) data += "window.__RMS_COV_SNAPS__=" + safeJSON(covMap) + ";";
    data += "<\/script>\n";
    var run = "<script>(function(){var s=document.createElement('style');s.textContent=window.__RMS_ASSETS__.css;document.head.appendChild(s);var j=document.createElement('script');j.textContent=window.__RMS_ASSETS__.js;document.head.appendChild(j);";
    if (copay) run += "var jc=document.createElement('script');jc.textContent=window.__RMS_COPAY_JS__;document.head.appendChild(jc);";
    run += "var vB=RMSViewer.mount(document.getElementById('rms-bypass'),window.__RMS_PAYLOAD__);var vC=null;";
    if (copay) run += "vC=RMSCopays.mount(document.getElementById('rms-copays'),window.__RMS_COPAY_PAYLOAD__);";
    run += "var tabs=[['rms-bypass','Bypassed Warnings']";
    if (copay) run += ",['rms-copays','Copays']";
    if (covMap) run += ",['rms-coverage','Coverage Accuracy']";
    run += "];var bar=document.getElementById('xtabs');";
    run += "if(tabs.length>1)tabs.forEach(function(t,i){var b=document.createElement('button');b.className='xtab'+(i===0?' active':'');b.textContent=t[1];b.onclick=function(){tabs.forEach(function(tt){document.getElementById(tt[0]).style.display='none';});document.getElementById(t[0]).style.display='';for(var c=0;c<bar.children.length;c++)if(bar.children[c].className.indexOf('xtab')===0)bar.children[c].className='xtab';b.className='xtab active';};bar.appendChild(b);});";
    /* Shared Data picker: one month or every loaded month combined, applied to all three tabs at once. */
    run += "var PDS=(window.__RMS_PAYLOAD__.periods||[]);if(PDS.length>1){if(tabs.length>1){var sp=document.createElement('span');sp.className='xspacer';bar.appendChild(sp);}";
    run += "var lb=document.createElement('label');lb.className='xdata';lb.appendChild(document.createTextNode('Data'));";
    var pNoun = (payload.granularity === "week") ? "week" : "month";
    run += "var sl=document.createElement('select');sl.title='Show one " + pNoun + ", or every loaded " + pNoun + " combined.';";
    run += "var oc=document.createElement('option');oc.value='';oc.textContent='All " + pNoun + "s combined';sl.appendChild(oc);";
    run += "PDS.forEach(function(pd){var o=document.createElement('option');o.value=pd.key;o.textContent=pd.label;sl.appendChild(o);});";
    run += "sl.value=window.__RMS_PAYLOAD__.period||'';";
    run += "sl.onchange=function(){var p=sl.value;if(vB&&vB.setPeriod)vB.setPeriod(p);if(vC&&vC.setPeriod)vC.setPeriod(p||'all');";
    run += "var cv=document.getElementById('rms-coverage');if(cv&&window.__RMS_COV_SNAPS__)cv.innerHTML=window.__RMS_COV_SNAPS__[p]||window.__RMS_COV_SNAPS__['']||'';";
    run += "try{window.scrollTo({top:0,behavior:'smooth'});}catch(e){window.scrollTo(0,0);}};";
    run += "lb.appendChild(sl);bar.appendChild(lb);}";
    run += "if(!bar.children.length)bar.style.display='none';";
    run += "})();<\/script>\n</body>\n</html>\n";
    return HEAD + divs + data + run;
  }
  function download(html, name) { var blob = new Blob([html], { type: "text/html" }), a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }
  function downloadMany(files, done) { var i = 0; (function next() { if (i >= files.length) { if (done) done(); return; }
    download(files[i].html, files[i].name); i++; setTimeout(next, 350); })(); }

  /* ---------- mount ---------- */
  function mount(root, payload) {
    // Granularity rides on the payload: the live app passes the filter-bar setting, and an exported
    // file carries the value chosen when it was written (there is no reader-facing control by design).
    setGranularity(payload.granularity);
    root.classList.add("rms"); ensureTip(); root.innerHTML = "";
    var recs = payload.records || [], headers = payload.headers || (recs[0] ? Object.keys(recs[0]) : []);
    var cm = payload.colMap || {}, deptKey = payload.deptKey, empKey = payload.empKey;
    var rootLevel = payload.level || "overview", gen = payload.generatedAt, title = payload.title || "SLHS — Bypassed Warnings";
    var deptLabel = payload.deptLabel || "department", empLabel = payload.empLabel || "employee";
    var deptLabelPlural = deptLabel + "s", empLabelPlural = empLabel + "s";
    var readOnly = !!payload.readOnly, onReset = payload.onReset;
    var externalFilters = !!payload.externalFilters;   // dept/emp filtering handled by the global bar; hide internal pickers + drill
    var copayData = payload.copayData || null;
    var regData = payload.regData || null;
    var getCoverageSnapshot = payload.getCoverageSnapshot || null;   // () -> static HTML of the Coverage Accuracy tab
    /* Periods carried by this file. An export ships EVERY loaded period's records plus this list, so
       the reader can switch between one period and all of them combined from the picker in the tab bar.
       state.period "" = all combined; "YYYY-MM" (or a "YYYY-MM-DD" week start) = that period only.
       Which of the two it is was fixed at export time and cannot be changed from inside the file. */
    var periods = (payload.periods || []).filter(function (p) { return p && p.key; });
    var state = { dept: payload.dept || null, emp: payload.emp || null, period: payload.period || "" };

    var wrap = el("div", { class: "rms-wrap" }); root.appendChild(wrap);
    var head = el("div", { class: "rms-head2" }), body = el("div", {});
    wrap.appendChild(head); wrap.appendChild(body);

    // Everything on the page reads through recsP() — the records for the selected month (or all of
    // them). deptRecsAll/empRecsAll/curRecsAll skip that filter: re-exports carry every month so the
    // downloaded file keeps its own month picker.
    function periodOf(r) { return periodKeyOf(parseDate(val(r, cm.date))); }
    function recsP() { return state.period ? recs.filter(function (r) { return periodOf(r) === state.period; }) : recs; }
    function deptRecsAll(d) { return recs.filter(function (r) { return keyOf(r, deptKey) === d; }); }
    function empRecsAll(d, e) { return recs.filter(function (r) { return (d == null || keyOf(r, deptKey) === d) && keyOf(r, empKey) === e; }); }
    function deptRecs(d) { return recsP().filter(function (r) { return keyOf(r, deptKey) === d; }); }
    function empRecs(d, e) { return recsP().filter(function (r) { return (d == null || keyOf(r, deptKey) === d) && keyOf(r, empKey) === e; }); }
    function curKind() { return state.emp ? "employee" : state.dept ? "department" : "overview"; }
    function curRecs() { return state.emp ? empRecs(state.dept, state.emp) : state.dept ? deptRecs(state.dept) : recsP(); }
    function curRecsAll() { return state.emp ? empRecsAll(state.dept, state.emp) : state.dept ? deptRecsAll(state.dept) : recs; }
    // All-month companion set for the Month-over-month section (records here can span months the page
    // itself is scoped away from). Narrowed to the same department / employee as the current view, and
    // handed to exports so a downloaded file keeps its month-over-month section.
    var monthRecords = payload.monthRecords || null;
    function momFor(dept, emp) {
      if (!monthRecords || !monthRecords.length) return null;
      if (dept == null && emp == null) return monthRecords;
      return monthRecords.filter(function (r) {
        if (dept != null && keyOf(r, deptKey) !== dept) return false;
        if (emp != null && keyOf(r, empKey) !== emp) return false;
        return true; });
    }

    var ctxBase = { deptKey: deptKey, empKey: empKey, deptLabel: deptLabel, empLabel: empLabel, deptLabelPlural: deptLabelPlural, empLabelPlural: empLabelPlural, regData: regData,
      onDrillDept: (readOnly || externalFilters || rootLevel === "employee") ? null : function (d) { state.dept = d; state.emp = null; rerender(); scrollTop(); },
      onDrillEmp: (readOnly || externalFilters || rootLevel === "employee") ? null : function (e) { state.emp = e; rerender(); scrollTop(); } };

    function scrollTop() { try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (e) { window.scrollTo(0, 0); } }

    function renderHead() {
      head.innerHTML = "";
      var kind = curKind();
      var titleText = kind === "overview" ? "Bypassed Warnings — All " + cap(deptLabelPlural) : (kind === "department" ? state.dept : state.emp);

      /* read-only (leadership / handout): title + subtitle only — the entire controls section is removed */
      if (readOnly) {
        var ml = monthLabel(curRecs(), cm);
        head.appendChild(el("div", { class: "rms-titlebox" }, [
          el("h1", { text: titleText }),
          el("div", { class: "rms-subtitle", text: (title || "Bypassed Warnings") + (ml ? " • " + ml : "") })
        ]));
        head.appendChild(el("div", { class: "rms-headright", style: "justify-content:flex-end;margin-top:2px" }, [ themeToggle() ]));
        return;
      }

      // header title — breadcrumb (internal drill) OR subtitle (when the global filter bar drives the view)
      if (externalFilters) {
        var mlx = monthLabel(curRecs(), cm);
        head.appendChild(el("div", { class: "rms-titlebox" }, [
          el("h1", { text: titleText }),
          el("div", { class: "rms-subtitle", text: (title || "Bypassed Warnings") + (mlx ? " • " + mlx : "") })
        ]));
      } else {
        var crumbs = el("div", { class: "rms-crumbs" });
        function crumb(label, active, onClick) { return el(onClick ? "button" : "span", { class: "crumb" + (active ? " active" : ""), text: label, onclick: onClick || undefined }); }
        var canTop = rootLevel === "overview";
        crumbs.appendChild(crumb("All " + deptLabelPlural, kind === "overview", (canTop && kind !== "overview") ? function () { state.dept = null; state.emp = null; rerender(); scrollTop(); } : null));
        if (state.dept) { crumbs.appendChild(el("span", { class: "sep", text: "›" }));
          crumbs.appendChild(crumb(state.dept, kind === "department", (rootLevel !== "employee" && kind === "employee") ? function () { state.emp = null; rerender(); scrollTop(); } : null)); }
        if (state.emp) { crumbs.appendChild(el("span", { class: "sep", text: "›" })); crumbs.appendChild(crumb(state.emp, true, null)); }
        head.appendChild(el("div", { class: "rms-titlebox" }, [ el("h1", { text: titleText }), crumbs ]));
      }

      var right = el("div", { class: "rms-headright" });
      if (!externalFilters && rootLevel === "overview") { var dsel = el("select", { class: "rms-select" });
        dsel.appendChild(el("option", { value: "", text: "▸ All " + deptLabelPlural + " (overview)" }));
        uniqSorted(recsP(), deptKey).forEach(function (d) { dsel.appendChild(el("option", { value: d, text: trunc(d, 40) + " (" + fmt(deptRecs(d).length) + ")" })); });
        dsel.value = state.dept || ""; dsel.addEventListener("change", function () { state.dept = dsel.value || null; state.emp = null; rerender(); scrollTop(); });
        right.appendChild(el("div", { class: "rms-picker" }, [ el("label", { text: cap(deptLabel) }), dsel ])); }
      if (!externalFilters && state.dept && rootLevel !== "employee") { var esel = el("select", { class: "rms-select" });
        esel.appendChild(el("option", { value: "", text: "▸ All " + empLabelPlural + " in " + trunc(state.dept, 24) }));
        uniqSorted(deptRecs(state.dept), empKey).forEach(function (e2) { esel.appendChild(el("option", { value: e2, text: trunc(e2, 34) + " (" + fmt(empRecs(state.dept, e2).length) + ")" })); });
        esel.value = state.emp || ""; esel.addEventListener("change", function () { state.emp = esel.value || null; rerender(); scrollTop(); });
        right.appendChild(el("div", { class: "rms-picker" }, [ el("label", { text: cap(empLabel) }), esel ])); }
      if (!externalFilters) {
        right.appendChild(themeToggle());
        if (onReset) right.appendChild(el("button", { class: "rms-themebtn", text: "↺ New import", title: "Load different files", onclick: function () { onReset(); } }));
      }
      if (right.childNodes.length) head.appendChild(right);

      // export toolbar (contextual)
      if (externalFilters) {
        // Exports live in the frozen page header; hand them up to the shell.
        var specs = [{ label: "⬇ Export for leadership", on: exportLeadership },
                     { label: "⬇ Clinic export", on: exportClinic }];
        if (kind === "employee") specs.push({ label: "⬇ Export this " + empLabel, on: exportCurrent });
        if (payload.onExports) payload.onExports(specs);
      } else {
        var xbar = el("div", { class: "rms-exports" });
        if (kind === "overview") {
          xbar.appendChild(xbtn("⬇ Export for leadership", "primary", function () { exportLeadership(); }));
          xbar.appendChild(xbtn("⬇ Export by " + deptLabel, "primary", function () { exportEachDept(); }));
        } else if (kind === "department") {
          xbar.appendChild(xbtn("⬇ Export by " + empLabel, "primary", function () { exportEachEmp(state.dept); }));
        } else {
          xbar.appendChild(xbtn("⬇ Export this " + empLabel, "primary", function () { exportCurrent(); }));
        }
        xbar.appendChild(el("span", { class: "rms-xstatus", id: "rms-xstatus" }));
        head.appendChild(xbar);
      }
    }
    function xbtn(label, kind, on) { return el("button", { class: "rms-xbtn " + kind, text: label, onclick: on }); }
    function setXStatus(msg) { var e = document.getElementById("rms-xstatus"); if (e) e.textContent = msg || ""; }

    /* export builders */
    function payloadFor(level, dept, emp, subset, ro, cpay) { return { level: level, dept: dept, emp: emp, records: subset, headers: headers, colMap: cm, granularity: granularity(),
      deptKey: deptKey, empKey: empKey, deptLabel: deptLabel, empLabel: empLabel, title: title, generatedAt: gen, regData: regData, readOnly: !!ro,
      periods: periods, period: state.period,
      monthRecords: momFor(dept === "All Departments" ? null : dept, emp), monthWindow: payload.monthWindow || null,
      copayData: cpay || null }; }
    // Copay payload for an exported file. When `names` is given, the copay rows are filtered to just
    // those employees (matched by check-in user name) so a per-employee export's Copays tab only shows
    // that person. `names` null/omitted → all copay rows.
    function copayFor(names) {
      if (!(copayData && copayData.records && copayData.records.length)) return null;
      var recsC = copayData.records;
      if (names && names.length && window.RMSCopays) {
        var ccm = RMSCopays.detectCopayCols(copayData.headers || (recsC[0] ? Object.keys(recsC[0]) : []));
        var want = {}; names.forEach(function (nm) { var k = empNameKey(nm); if (k && k !== "|") want[k] = 1; });
        if (ccm.checkinUser && Object.keys(want).length) recsC = recsC.filter(function (r) { return want[empNameKey(r[ccm.checkinUser])]; });
      }
      // keep every month's copay rows (the copay module buckets them itself) and the month it opens on,
      // so a re-export keeps its own month-over-month section
      return { records: recsC, headers: copayData.headers, title: copayData.title || "Copays", generatedAt: copayData.generatedAt || gen, groupKey: "checkin",
        period: state.period || "all", hideMonthPicker: periods.length > 1 };
    }
    function copayPayload() { return copayFor(null); }
    function coverageSnap(names) { try { return getCoverageSnapshot ? getCoverageSnapshot(names) : null; } catch (e) { return null; } }
    function exportLeadership() {
      // read-only snapshot of the current overview — controls stripped; coverage = all employees on this data
      var subset = recs, name = "Leadership — All " + cap(deptLabelPlural), cp = copayFor(null);   // every month; the file's own picker scopes it
      download(buildHTML(payloadFor("overview", null, null, subset, true, cp), name, cp, coverageSnap(null)), fileName("Leadership - All " + cap(deptLabelPlural), subset, cm));
      setXStatus("Downloaded leadership summary.");
    }
    // Single interactive file for the whole clinic (this sheet). Opening it, you can export by employee,
    // and each of those employee files is scoped to that person — copay included.
    function exportClinic() {
      var subset = state.dept ? deptRecsAll(state.dept) : recs, name = state.dept || ("All " + cap(deptLabelPlural)), cp = copayFor(uniqSorted(subset, empKey));
      download(buildHTML(payloadFor("department", state.dept || "All Departments", null, subset, false, cp), name, cp, coverageSnap(uniqSorted(subset, empKey))), fileName(name, subset, cm));
      setXStatus("Downloaded clinic export — open it to export by " + empLabel + ".");
    }
    function exportCurrent() { var subset = curRecsAll(), cp = copayFor(state.emp ? [state.emp] : uniqSorted(subset, empKey));
      download(buildHTML(payloadFor("employee", state.dept, state.emp, subset, true, cp), state.emp, cp, coverageSnap(state.emp ? [state.emp] : uniqSorted(subset, empKey))), fileName(state.emp, subset, cm));
      setXStatus("Downloaded " + fileName(state.emp, subset, cm));
    }
    function exportEachDept() {
      var depts = uniqSorted(recs, deptKey), files = depts.map(function (d) { var subset = deptRecsAll(d), cp = copayFor(uniqSorted(subset, empKey));
        return { name: fileName(d, subset, cm), html: buildHTML(payloadFor("department", d, null, subset, false, cp), d, cp, coverageSnap(uniqSorted(subset, empKey))) }; });
      setXStatus("Downloading " + files.length + " " + deptLabelPlural + " … approve the browser prompt if asked.");
      downloadMany(files, function () { setXStatus("Exported " + files.length + " " + deptLabel + " files."); });
    }
    function exportEachEmp(dept) {
      var emps = uniqSorted(deptRecsAll(dept), empKey), files = emps.map(function (e2) { var subset = empRecsAll(dept, e2), cp = copayFor([e2]);
        return { name: fileName(e2, subset, cm), html: buildHTML(payloadFor("employee", dept, e2, subset, true, cp), e2, cp, coverageSnap([e2])) }; });
      setXStatus("Downloading " + files.length + " " + empLabelPlural + " … approve the browser prompt if asked.");
      downloadMany(files, function () { setXStatus("Exported " + files.length + " " + empLabel + " files."); });
    }

    function rerender() { renderHead(); renderScope(body, curRecs(), cm, headers, Object.assign({ kind: curKind(), monthRecords: momFor(state.dept, state.emp), monthWindow: payload.monthWindow || null }, ctxBase)); }
    rerender();

    wrap.appendChild(el("div", { class: "rms-foot" }, [ el("span", { text: "St. Luke's • Bypassed Warnings Dashboard" }), el("span", { class: "spacer" }), el("span", { text: gen ? "Generated: " + gen : "" }) ]));
    // Handle for the export's shared month picker (see buildHTML): one control drives all three tabs.
    return { setPeriod: function (p) {
      p = p || "";
      if (p === state.period) return;
      state.period = p;
      // drop a drill-down that the newly-picked month has no rows for — but never past this file's own
      // root scope (an employee handout stays on that employee, and shows zeros for an empty month)
      if (state.emp && rootLevel !== "employee" && !empRecs(state.dept, state.emp).length) state.emp = null;
      if (state.dept && rootLevel === "overview" && !deptRecs(state.dept).length) { state.dept = null; state.emp = null; }
      rerender();
    } };
  }

  /* ---------- theme toggle ---------- */
  function themeToggle() {
    var btn = el("button", { class: "rms-themebtn", title: "Toggle light / dark" });
    function dark() { var d = document.documentElement.getAttribute("data-theme"); return d === "dark" || (d == null && window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches); }
    function paint() { btn.innerHTML = dark() ? "☀︎ Light" : "☾ Dark"; }
    btn.addEventListener("click", function () { document.documentElement.setAttribute("data-theme", dark() ? "light" : "dark"); paint(); });
    paint(); return btn;
  }

  window.RMSViewer = { mount: mount, summarize: summarize, buildHTML: buildHTML,
                       setGranularity: setGranularity, granularity: granularity, periodKeyLabel: periodKeyLabel, isWeekKey: isWeekKey };
})();

