/* Split PAS_Quality_Dashboard.html into editable parts, and reassemble them.
   The dashboard is one generated 1.2MB file whose script blocks are far too large to edit in place,
   and whose viewer JS is embedded TWICE (live source + a JSON-encoded copy in __RMS_ASSETS__).
   Splitting lets each block be edited as normal source; assemble.js re-encodes the asset copy so the
   two can never drift.

   Layout (1-indexed lines of the source file):
     474-1312   viewer JS      -> parts/viewer.js
     1314-1624  copay module   -> parts/copays.js
     1628-3076  app shell      -> parts/shell.js
     473        __RMS_ASSETS__ -> regenerated from parts/viewer.js + the CSS captured in parts/assets.css
   Everything else is copied through untouched.                                                    */
const fs = require("fs"), path = require("path");
const ROOT = path.join(__dirname, "..");
/* TEMPLATE is the pristine generated file and is never written to; SRC is the build output.
   Assembling always splices into TEMPLATE, because BLOCKS are TEMPLATE line numbers — reading the
   previous build back in would splice against boundaries the last edit had already shifted. */
const TEMPLATE = path.join(ROOT, "template.html");
const SRC = path.join(ROOT, "PAS_Quality_Dashboard.html");
const PARTS = path.join(ROOT, "parts");

const BLOCKS = [
  { name: "viewer.js", start: 474, end: 1312, strip: "<script>", tail: "</script>" },
  { name: "copays.js", start: 1314, end: 1624, strip: null, tail: null },
  { name: "shell.js", start: 1628, end: 3076, strip: null, tail: null },
];

function readLines() { return fs.readFileSync(TEMPLATE, "utf8").split("\n"); }

function split() {
  const L = readLines();
  fs.mkdirSync(PARTS, { recursive: true });
  for (const b of BLOCKS) {
    let text = L.slice(b.start - 1, b.end).join("\n");
    if (b.strip && text.startsWith(b.strip)) text = text.slice(b.strip.length);
    if (b.tail && text.endsWith(b.tail)) text = text.slice(0, -b.tail.length);
    fs.writeFileSync(path.join(PARTS, b.name), text);
    console.log("wrote", b.name, text.length, "chars");
  }
  // Pull the CSS out of the asset line so assemble.js can rebuild that line verbatim.
  const assetLine = L[472];
  const obj = eval("(" + assetLine.replace(/^<script>window\.__RMS_ASSETS__ = /, "").replace(/;<\/script>$/, "") + ")");
  fs.writeFileSync(path.join(PARTS, "assets.css"), obj.css);
  console.log("wrote assets.css", obj.css.length, "chars");
  // Sanity: the asset copy of the viewer must match the live copy we just extracted.
  const live = fs.readFileSync(path.join(PARTS, "viewer.js"), "utf8");
  console.log("asset js === live js:", obj.js.trim() === live.trim());
}

function assemble(outPath) {
  const L = readLines();
  const get = (n) => fs.readFileSync(path.join(PARTS, n), "utf8");
  const viewer = get("viewer.js"), css = get("assets.css");
  const out = [];
  let i = 0;
  while (i < L.length) {
    const n = i + 1;
    const b = BLOCKS.find((x) => x.start === n);
    if (n === 473) {
      out.push("<script>window.__RMS_ASSETS__ = { css: " + JSON.stringify(css) + ", js: " + JSON.stringify(viewer) + " };</script>");
      i += 1;
    } else if (b) {
      out.push((b.strip || "") + get(b.name) + (b.tail || ""));
      i = b.end;
    } else { out.push(L[i]); i += 1; }
  }
  fs.writeFileSync(outPath || SRC, out.join("\n"));
  console.log("assembled ->", outPath || SRC);
}

const cmd = process.argv[2];
if (cmd === "split") split();
else if (cmd === "assemble") assemble(process.argv[3]);
else { console.error("usage: node split.js split|assemble [out]"); process.exit(1); }
