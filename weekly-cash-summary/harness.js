const fs = require('fs'), vm = require('vm');

function el() {
  const target = function(){ return el(); };
  return new Proxy(target, {
    get(t, k) {
      if (k === 'classList') return { toggle(){}, add(){}, remove(){}, contains(){return false;} };
      if (k === 'style') return {};
      if (k === 'options') return [];
      if (k === 'value' || k === 'textContent' || k === 'innerHTML' || k === 'srcdoc' || k === 'title') return '';
      if (k === 'dataset') return {};
      if (k === 'files') return [];
      if (k === Symbol.iterator) return function*(){};
      if (k === 'length') return 0;
      return el();
    },
    set() { return true; },
    apply() { return el(); },
    has() { return true; },
  });
}
const doc = {
  getElementById: () => el(), querySelector: () => el(), querySelectorAll: () => [],
  createElement: () => el(), addEventListener(){}, body: el(), documentElement: el(),
};
const ctx = {
  console, TextEncoder, TextDecoder, Blob, URL, Date, Math, JSON, Number, String, Object, Array, Map, Set,
  Promise, Function, RegExp, Error, isNaN, parseInt, parseFloat, setTimeout, clearTimeout, DecompressionStream,
  Uint8Array, Uint32Array, DataView, ArrayBuffer, Proxy, Symbol, Intl, Response, ReadableStream, CompressionStream,
  document: doc,
  localStorage: { getItem: () => null, setItem(){}, removeItem(){} },
  indexedDB: { open: () => ({}) },
  navigator: { userAgent: 'node' },
  matchMedia: () => ({ matches: false, addEventListener(){}, addListener(){} }),
  alert(){}, confirm(){ return true; }, addEventListener(){}, removeEventListener(){},
  requestAnimationFrame: fn => setTimeout(fn, 0),
  FileReader: function(){},
};
ctx.window = ctx; ctx.globalThis = ctx; ctx.self = ctx;
vm.createContext(ctx);
const src = fs.readFileSync('app.js', 'utf8');
vm.runInContext(src + '\n;globalThis.__X = { state, parseRolling, computeRollingDash, writeRollingXlsx, ledgerNew, ledgerUpsertWeek, ledgerClone, ledgerWindow, ledgerSorted, ledgerFromLegacy, ledgerFromSheet, readSheetGrids, writeSheetsXlsx, rollingSheetXml, colRef2, dateKey, alnumKey, rollingCategoryOf, ROLLING_KEYS, BUCKET_META, bucketCategory, bucketName, rollWindowWeeks, buildWeekBuckets, rollingSpecialItems, ledgerRows, saveToRolling, wcsRolling, includedRows, computeWCS, rollingBucket, saveRollingEdits, specialItemsFor, weekRangeForMW, adjBlocked, planLedgerAdjustments, applyLedgerAdjustments };', ctx, { filename: 'app.js' });
module.exports = ctx.__X;
