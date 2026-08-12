/* Shared PPR parsing + PPTX generation logic.
 * Works in Node (module.exports) and browser (window.PPRCore).
 * generate(XLSX, PptxGenJS, workbook) -> configured pptx instance (caller calls writeFile). */
(function (global) {
  "use strict";

  // ---- palette lifted directly from the original PPR template -----------
  var FONT = "Calibri";
  var TITLEBLUE = "05387A";   // slide title
  var NAVY = "002060";        // section / table headers
  var BODYBLUE = "002776";    // body text (description, metrics, overall)
  var VALUEBLUE = "1F4187";   // sponsor values
  var GREEN = "00B050";
  var YELLOW = "FFFF00";
  var RED = "C00000";
  var NSGREY = "CBCDD6";
  var ROWGREY = "E7E8EC";     // milestone data rows
  var BOXGREY = "F3F4F6";     // bottom section boxes
  var GRIDGREY = "D9D9D9";    // thin borders on light tables
  var GREY = GRIDGREY;
  var BORDER = { type: "solid", color: GRIDGREY, pt: 1 };
  var WHITE_GRID = { type: "solid", color: "FFFFFF", pt: 1.5 };

  var SC = {
    C:  { fill: NAVY,  color: "FFFFFF", label: "Complete" },
    G:  { fill: GREEN, color: "000000", label: "On Track" },
    Y:  { fill: YELLOW, color: "000000", label: "At Risk" },
    R:  { fill: RED,   color: "FFFFFF", label: "Significant Risk w/ inadequate mitigation" },
    NS: { fill: NSGREY, color: "000000", label: "Not Started" }
  };
  // template legend shows exactly these four
  var LEGEND_ORDER = ["C", "G", "Y", "R"];
  function statusMeta(code) {
    var k = String(code || "").toUpperCase().replace(/\s+/g, "");
    return SC[k] || { fill: null, color: "000000", label: k };
  }

  // ---- grid helpers -----------------------------------------------------
  function toGrid(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: true });
  }
  function cell(G, r, c) {
    if (!G[r]) return "";
    var v = G[r][c];
    return v == null ? "" : String(v).trim();
  }
  function rowLen(G, r) { return G[r] ? G[r].length : 0; }
  function findCell(G, pred) {
    for (var r = 0; r < G.length; r++) {
      var L = rowLen(G, r);
      for (var c = 0; c < L; c++) {
        if (pred(cell(G, r, c), r, c)) return { r: r, c: c };
      }
    }
    return null;
  }
  function eqi(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }
  function has(a, sub) { return String(a).toLowerCase().indexOf(sub.toLowerCase()) >= 0; }
  function findColInRow(G, r, pred) {
    var L = rowLen(G, r);
    for (var c = 0; c < L; c++) if (pred(cell(G, r, c))) return c;
    return -1;
  }
  var SECTION_RE = /^(Project Team|Project Description|Overall Status|Accomplishments|Next Steps|Issues|Status Legend|Description|Milestones)$/i;

  // ---- sheet type detection --------------------------------------------
  function detectType(G) {
    var isMile = !!findCell(G, function (v) { return eqi(v, "Milestones"); });
    var isOverall = !!findCell(G, function (v) { return eqi(v, "Overall Status"); });
    if (isMile && isOverall) return "ppr";
    if (findCell(G, function (v) { return eqi(v, "Attribute"); }) &&
        findCell(G, function (v) { return eqi(v, "Definition"); })) return "key";
    if (findCell(G, function (v) { return has(v, "Email Communication Instructions") || eqi(v, "Email Template"); })) return "email";
    return "generic";
  }

  // ---- PPR sheet parser -------------------------------------------------
  function parsePpr(G, sheetName) {
    var out = { title: "", week: "", milestones: [], team: [], description: "",
                overall: { code: "", text: "" }, metrics: { rows: [] },
                acc: [], next: [], issues: [] };

    // title: first non-empty cell top rows
    var tc = findCell(G, function (v, r) { return r < 3 && v !== ""; });
    out.title = tc ? cell(G, tc.r, tc.c) : sheetName;

    var wc = findCell(G, function (v) { return has(v, "Progress week ending"); });
    if (wc) {
      var wtxt = cell(G, wc.r, wc.c);
      var idx = wtxt.indexOf(":");
      out.week = idx >= 0 ? wtxt.slice(idx + 1).trim() : wtxt;
    }

    // milestones
    var mh = findCell(G, function (v) { return eqi(v, "Milestones"); });
    if (mh) {
      var hr = mh.r, nameCol = mh.c;
      var statusCol = findColInRow(G, hr, function (v) { return eqi(v, "Status"); });
      var pctCol = findColInRow(G, hr, function (v) { return has(v, "% Complete") || eqi(v, "% Complete"); });
      var endCol = findColInRow(G, hr, function (v) { return has(v, "End Date"); });
      for (var r = hr + 1; r < G.length; r++) {
        var nm = cell(G, r, nameCol);
        var st = statusCol >= 0 ? cell(G, r, statusCol) : "";
        var pc = pctCol >= 0 ? cell(G, r, pctCol) : "";
        var en = endCol >= 0 ? cell(G, r, endCol) : "";
        if (nm === "" && st === "" && pc === "" && en === "") break;
        if (SECTION_RE.test(nm)) break;
        out.milestones.push({ name: nm, status: st, pct: pc, end: en });
      }
    }

    // team / sponsors
    var pt = findCell(G, function (v) { return eqi(v, "Project Team"); });
    if (pt) {
      var pc0 = pt.c, winEnd = pc0 + 6;
      for (var rr = pt.r + 1; rr < G.length; rr++) {
        var tokens = [];
        for (var cc = pc0; cc <= winEnd; cc++) {
          var val = cell(G, rr, cc);
          if (val !== "") tokens.push(val);
        }
        if (tokens.length === 0) break;
        if (SECTION_RE.test(tokens[0]) && !/:$/.test(tokens[0])) break;
        for (var t = 0; t < tokens.length; t++) {
          if (/:$/.test(tokens[t]) && t + 1 < tokens.length && !/:$/.test(tokens[t + 1])) {
            out.team.push({ label: tokens[t].replace(/:$/, "").trim(), value: tokens[t + 1] });
            t++;
          }
        }
      }
    }

    // description (collect one or more rows until a blank row or the next section)
    var pd = findCell(G, function (v) { return eqi(v, "Project Description"); });
    if (pd) {
      var descLines = [], started = false;
      for (var dr = pd.r + 1; dr < G.length; dr++) {
        var dv = cell(G, dr, pd.c);
        if (dv === "") { if (started) break; else continue; }
        if (SECTION_RE.test(dv)) break;
        dv.split("\n").forEach(function (ln) { if (ln.trim() !== "") descLines.push(ln.trim()); });
        started = true;
      }
      out.description = descLines.join("\n");
    }

    // overall status
    var os = findCell(G, function (v) { return eqi(v, "Overall Status"); });
    if (os) {
      for (var or = os.r + 1; or < Math.min(G.length, os.r + 6); or++) {
        var code = cell(G, or, os.c);
        if (code !== "") {
          out.overall.code = code;
          var L = rowLen(G, or);
          for (var oc = os.c + 1; oc < L; oc++) {
            var ov = cell(G, or, oc);
            if (ov !== "") { out.overall.text = ov; break; }
          }
          break;
        }
      }
    }

    // metrics
    var tgt = findCell(G, function (v) { return eqi(v, "Target"); });
    if (tgt) {
      var mr = tgt.r;
      var descCol = findColInRow(G, mr, function (v) { return eqi(v, "Description"); });
      var targetCol = tgt.c;
      var realCol = findColInRow(G, mr, function (v) { return has(v, "Target Realized"); });
      var mpctCol = findColInRow(G, mr, function (v) { return has(v, "% of Target"); });
      if (descCol < 0) descCol = 0;
      for (var mrow = mr + 1; mrow < G.length; mrow++) {
        var d = cell(G, mrow, descCol);
        if (d === "") break;
        if (SECTION_RE.test(d)) break;
        out.metrics.rows.push({
          desc: d,
          target: targetCol >= 0 ? cell(G, mrow, targetCol) : "",
          realized: realCol >= 0 ? cell(G, mrow, realCol) : "",
          pct: mpctCol >= 0 ? cell(G, mrow, mpctCol) : ""
        });
      }
    }

    // accomplishments / next steps / issues
    var ac = findCell(G, function (v) { return eqi(v, "Accomplishments"); });
    if (ac) {
      var ar = ac.r;
      var col1 = ac.c;
      var col2 = findColInRow(G, ar, function (v) { return eqi(v, "Next Steps"); });
      var col3 = findColInRow(G, ar, function (v) { return eqi(v, "Issues"); });
      function collect(colArr, targetKey, colIdx) {
        // placeholder
      }
      var cols = [{ idx: col1, key: "acc" }, { idx: col2, key: "next" }, { idx: col3, key: "issues" }];
      for (var sr = ar + 1; sr < G.length; sr++) {
        var vals = cols.map(function (o) { return o.idx >= 0 ? cell(G, sr, o.idx) : ""; });
        if (vals.every(function (v) { return v === ""; })) break;
        for (var k = 0; k < cols.length; k++) {
          var raw = vals[k];
          if (raw === "") continue;
          raw.split("\n").forEach(function (line) {
            var item = line.replace(/^[\s••\-*]+/, "").trim();
            if (item !== "") out[cols[k].key].push(item);
          });
        }
      }
    }

    return out;
  }

  function parseKey(G) {
    var out = { title: "Status Report Key", rows: [] };
    var t = findCell(G, function (v, r) { return r < 2 && v !== ""; });
    if (t) out.title = cell(G, t.r, t.c);
    var hdr = findCell(G, function (v) { return eqi(v, "Attribute"); });
    if (!hdr) return out;
    var ar = hdr.r, attrCol = hdr.c;
    var defCol = findColInRow(G, ar, function (v) { return eqi(v, "Definition"); });
    if (defCol < 0) defCol = attrCol + 1;
    for (var r = ar + 1; r < G.length; r++) {
      var a = cell(G, r, attrCol), d = cell(G, r, defCol);
      if (a === "" && d === "") continue;
      out.rows.push({ attr: a, def: d });
    }
    return out;
  }

  function parseGeneric(G, sheetName) {
    var lines = [];
    var title = sheetName;
    var t = findCell(G, function (v, r) { return r < 2 && v !== ""; });
    if (t) title = cell(G, t.r, t.c);
    for (var r = 0; r < G.length; r++) {
      var L = rowLen(G, r), parts = [];
      for (var c = 0; c < L; c++) { var v = cell(G, r, c); if (v !== "") parts.push(v); }
      lines.push(parts.join("    "));
    }
    // drop the title line (first non-empty)
    return { title: title, lines: lines };
  }

  // ======================================================================
  // PPTX rendering
  // ======================================================================
  function addTitle(slide, title) {
    slide.addText(title || "", { x: 0.3, y: 0.16, w: 12.7, h: 0.5, fontFace: FONT, fontSize: 24,
      bold: true, color: TITLEBLUE, align: "left", valign: "middle", margin: 0 });
  }
  function bar(slide, x, y, w, text) {
    slide.addText(text, { x: x, y: y, w: w, h: 0.28, fill: { color: NAVY }, color: "FFFFFF",
      bold: true, fontSize: 11, align: "left", valign: "middle", fontFace: FONT, margin: 4 });
  }

  // cell builder that tolerates array text (bulleted runs)
  function cc(text, o) {
    o = o || {};
    var opt = { fontFace: FONT, valign: o.valign || "middle" };
    Object.keys(o).forEach(function (k) { if (k !== "fill" && k !== "color") opt[k] = o[k]; });
    if (o.fill) opt.fill = { color: o.fill };
    if (o.color) opt.color = o.color;
    var t = Array.isArray(text) ? text : (text == null ? "" : String(text));
    return { text: t, options: opt };
  }
  function tcell(text, opts) { return cc(text, opts); }

  function renderPpr(pptx, data, pageNum) {
    var s = pptx.addSlide();
    s.background = { color: "FFFFFF" };

    // ---- Title (top-left) + week ending (top-right, date italic) ----
    s.addText(data.title || "", { x: 0.12, y: 0.08, w: 6.6, h: 0.44, fontFace: FONT, fontSize: 26,
      bold: true, color: TITLEBLUE, align: "left", valign: "middle", margin: 0 });
    if (data.week) {
      s.addText([
        { text: "Progress week ending: ", options: { bold: true } },
        { text: data.week, options: { bold: true, italic: true } }
      ], { x: 6.7, y: 0.2, w: 6.5, h: 0.34, fontFace: FONT, fontSize: 14, color: NAVY,
        align: "right", valign: "middle", margin: 0 });
    }

    // ---- Sponsors (left, top) ----
    var teamRows = [];
    for (var i = 0; i < data.team.length; i += 2) {
      var a = data.team[i], b = data.team[i + 1];
      teamRows.push([
        cc(a ? a.label + ": " : "", { fill: a ? NAVY : null, color: "FFFFFF", bold: true, align: "left", fontSize: 11 }),
        cc(a ? a.value : "", { color: VALUEBLUE, bold: true, align: "left", fontSize: 11 }),
        cc(b ? b.label + ": " : "", { fill: b ? NAVY : null, color: "FFFFFF", bold: true, align: "left", fontSize: 11 }),
        cc(b ? b.value : "", { color: VALUEBLUE, bold: true, align: "left", fontSize: 11 })
      ]);
    }
    if (teamRows.length === 0) teamRows.push([cc(""), cc(""), cc(""), cc("")]);
    s.addTable(teamRows, { x: 0.12, y: 0.63, w: 6.5, colW: [1.4, 1.85, 1.4, 1.85], border: BORDER,
      fontFace: FONT, valign: "middle", autoPage: false, rowH: 0.22, margin: 2 });

    // ---- Combined left table: Description / Overall Status / Metrics ----
    var L = [];
    L.push([cc("Project Description:", { colspan: 5, fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 11 })]);
    L.push([cc(data.description || "", { colspan: 5, color: "000000", fontSize: 10, align: "left", valign: "top" })]);
    L.push([cc("Overall Status:", { colspan: 5, fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 11 })]);
    var om = statusMeta(data.overall.code);
    L.push([
      cc(data.overall.code, { fill: om.fill, color: om.color, bold: true, align: "center", fontSize: 11 }),
      cc(data.overall.text, { colspan: 4, color: BODYBLUE, fontSize: 11, align: "left" })
    ]);
    L.push([
      cc("Description:", { colspan: 2, fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 11 }),
      cc("Target:", { fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 11 }),
      cc("Target Realized to Date:", { fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 9.5 }),
      cc("% of Target Realized:", { fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 9.5 })
    ]);
    data.metrics.rows.forEach(function (m) {
      L.push([
        cc(m.desc, { colspan: 2, color: BODYBLUE, fontSize: 10, align: "left" }),
        cc(m.target, { color: BODYBLUE, fontSize: 10, align: "left" }),
        cc(m.realized, { color: BODYBLUE, fontSize: 10, align: "left" }),
        cc(m.pct, { color: BODYBLUE, fontSize: 10, align: "left" })
      ]);
    });
    s.addTable(L, { x: 0.12, y: 1.19, w: 6.55, colW: [0.39, 1.92, 1.28, 1.47, 1.49], border: BORDER,
      fontFace: FONT, valign: "middle", autoPage: false, rowH: 0.2, margin: 3 });

    // ---- Milestones (right) ----
    var M = [[
      cc("Milestones", { fill: NAVY, color: "FFFFFF", bold: true, align: "left", fontSize: 11 }),
      cc("Status", { fill: NAVY, color: "FFFFFF", bold: true, align: "center", fontSize: 11 }),
      cc("% Complete", { fill: NAVY, color: "FFFFFF", bold: true, align: "center", fontSize: 11 }),
      cc("End Date", { fill: NAVY, color: "FFFFFF", bold: true, align: "center", fontSize: 11 })
    ]];
    data.milestones.forEach(function (m) {
      var sm = statusMeta(m.status);
      M.push([
        cc(m.name, { fill: ROWGREY, color: "000000", align: "left", fontSize: 10 }),
        cc(m.status, { fill: sm.fill, color: sm.color, bold: true, align: "center", fontSize: 11 }),
        cc(m.pct, { fill: ROWGREY, color: "000000", align: "center", fontSize: 11 }),
        cc(m.end, { fill: ROWGREY, color: "000000", align: "center", fontSize: 11 })
      ]);
    });
    s.addTable(M, { x: 6.82, y: 0.59, w: 6.28, colW: [3.65, 0.63, 1.03, 0.97], border: WHITE_GRID,
      fontFace: FONT, valign: "middle", autoPage: false, rowH: 0.26, margin: 3 });

    // ---- Bottom: section headers + 3-column content ----
    var bx = [0.36, 4.63, 8.9], bw = 4.27;
    var heads = ["Key Accomplishments", "Next Steps", "Key Issues / Risks"];
    heads.forEach(function (h, i) {
      s.addText(h, { x: bx[i], y: 4.62, w: bw, h: 0.3, fontFace: FONT, fontSize: 13, bold: true,
        color: NAVY, align: "center", valign: "middle", margin: 0 });
    });
    function bulletCell(arr) {
      if (!arr || arr.length === 0)
        return cc("None", { fill: BOXGREY, color: BODYBLUE, fontSize: 10, valign: "top", align: "left", margin: 6 });
      var runs = arr.map(function (it) {
        return { text: it, options: { bullet: { characterCode: "2022" }, color: BODYBLUE, fontSize: 10,
          breakLine: true, paraSpaceAfter: 3, align: "left" } };
      });
      return cc(runs, { fill: BOXGREY, valign: "top", align: "left", margin: 6 });
    }
    s.addTable([[bulletCell(data.acc), bulletCell(data.next), bulletCell(data.issues)]],
      { x: 0.36, y: 5.0, w: 12.81, colW: [4.27, 4.27, 4.27], border: BORDER, fontFace: FONT,
        fontSize: 10, valign: "top", autoPage: false, rowH: 2.05 });

    // ---- Legend (four swatches) + slide number ----
    var leg = LEGEND_ORDER.map(function (code) { var m = statusMeta(code); return { label: m.label, color: m.fill }; });
    var lx = [4.6, 5.78, 7.02, 8.16];
    var lw = [1.0, 1.05, 0.75, 4.6];
    leg.forEach(function (item, i) {
      s.addShape(pptx.ShapeType.rect, { x: lx[i], y: 7.31, w: 0.17, h: 0.17, fill: { color: item.color },
        line: { color: "999999", width: 0.5 } });
      s.addText(item.label, { x: lx[i] + 0.22, y: 7.22, w: lw[i], h: 0.34, fontFace: FONT, fontSize: 9,
        bold: true, color: "000000", align: "left", valign: "middle", margin: 0 });
    });
    if (pageNum) s.addText(String(pageNum), { x: 0.2, y: 7.18, w: 0.5, h: 0.25, fontFace: FONT,
      fontSize: 12, color: "333333", align: "left", valign: "middle", margin: 0 });
  }

  function renderKey(pptx, data) {
    var slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addTitle(slide, data.title);
    var rows = [[
      tcell("Attribute", { fill: NAVY, color: "FFFFFF", bold: true, align: "center" }),
      tcell("Definition", { fill: NAVY, color: "FFFFFF", bold: true, align: "left" })
    ]];
    var keyColor = { "Red (R)": "R", "Yellow (Y)": "Y", "Green (G)": "G",
                     "Not Started (NS)": "NS", "Complete (C)": "C" };
    data.rows.forEach(function (r) {
      var code = keyColor[r.attr];
      var m = code ? statusMeta(code) : { fill: null, color: "000000" };
      rows.push([
        tcell(r.attr, { fill: m.fill, color: m.color, bold: true, align: "center", valign: "top", fontSize: 9 }),
        tcell(r.def, { align: "left", valign: "top", fontSize: 8.5 })
      ]);
    });
    slide.addTable(rows, { x: 0.3, y: 0.75, w: 12.73, colW: [2.1, 10.63], border: BORDER,
      fontFace: FONT, fontSize: 8.5, valign: "top", autoPage: false, margin: 4 });
  }

  function renderEmail(pptx, data) {
    var slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addTitle(slide, data.title);
    var splitIdx = data.lines.findIndex(function (l) { return eqi(l, "Email Template"); });
    var leftLines = [], rightLines = [];
    data.lines.forEach(function (l, i) {
      if (i === 0) return; // title already shown
      if (splitIdx >= 0 && i >= splitIdx) { if (i > splitIdx) rightLines.push(l); }
      else leftLines.push(l);
    });
    // left: instructions
    bar(slide, 0.3, 0.9, 6.1, "Instructions");
    var leftPara = leftLines.filter(function (l) { return l !== ""; }).map(function (l) {
      var bold = /^(Purpose:|Steps:)/i.test(l);
      return { text: l, options: { fontSize: 11, bold: bold, breakLine: true, paraSpaceAfter: 6, align: "left" } };
    });
    slide.addText(leftPara, { x: 0.3, y: 1.2, w: 6.1, h: 5.9, fontFace: FONT, valign: "top",
      color: "000000", margin: 4, line: { color: GREY, pt: 0.5 } });
    // right: email template
    bar(slide, 6.7, 0.9, 6.3, "Email Template");
    var rightPara = rightLines.map(function (l) {
      return { text: l === "" ? " " : l, options: { fontSize: 11, breakLine: true, align: "left" } };
    });
    slide.addText(rightPara, { x: 6.7, y: 1.2, w: 6.3, h: 5.9, fontFace: FONT, valign: "top",
      color: "000000", fill: { color: "F2F2F2" }, margin: 6, line: { color: GREY, pt: 0.5 } });
  }

  function renderGeneric(pptx, data) {
    var slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addTitle(slide, data.title);
    var para = data.lines.slice(1).map(function (l) {
      return { text: l === "" ? " " : l, options: { fontSize: 11, breakLine: true, align: "left" } };
    });
    slide.addText(para, { x: 0.3, y: 0.9, w: 12.7, h: 6.3, fontFace: FONT, valign: "top", color: "000000", margin: 4 });
  }

  // ======================================================================
  function generate(XLSX, PptxGenJS, workbook) {
    var pptx = new PptxGenJS();
    pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
    pptx.layout = "WIDE";
    pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

    workbook.SheetNames.forEach(function (name, idx) {
      var ws = workbook.Sheets[name];
      var G = toGrid(XLSX, ws);
      var type = detectType(G);
      if (type === "ppr") renderPpr(pptx, parsePpr(G, name), idx + 1);
      else if (type === "key") renderKey(pptx, parseKey(G));
      else if (type === "email") renderEmail(pptx, parseGeneric(G, name));
      else renderGeneric(pptx, parseGeneric(G, name));
    });
    return pptx;
  }

  function summarize(XLSX, workbook) {
    return workbook.SheetNames.map(function (name) {
      var G = toGrid(XLSX, workbook.Sheets[name]);
      var type = detectType(G);
      var detail = "";
      if (type === "ppr") {
        var p = parsePpr(G, name);
        detail = p.milestones.length + " milestones";
      } else if (type === "key") {
        detail = parseKey(G).rows.length + " definitions";
      } else if (type === "email") {
        detail = "instructions + template";
      } else {
        detail = "text content";
      }
      var labels = { ppr: "Progress Report", key: "Status Key", email: "Email Guide", generic: "Content" };
      return { name: name, type: type, typeLabel: labels[type] || type, detail: detail };
    });
  }

  var api = { generate: generate, summarize: summarize };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  global.PPRCore = api;
})(typeof window !== "undefined" ? window : globalThis);
