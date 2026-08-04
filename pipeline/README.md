# Customer Service Report Builder — in-browser data pipeline

`../Customer_Service_Report_Builder.html` is a self-contained tool. Drop in **two files**:

1. the **12-month master workbook** (the file that contains the `Data Tables` sheet), and
2. this month's **raw-data export** (`.xlsx`),

and it will:

- calculate the new month, add it to the `Data Tables` grid,
- rebuild the on-screen dashboard (and the PPTX / HTML exports), and
- let you **download an updated 12-month workbook** with the new month filled in.

You can still drop just the master on its own to view the existing dashboard.

## How it works

The original workbook did all its number-crunching with ~15 Excel VBA macros that roll the raw
exports up into the `Data Tables` sheet, plus worksheet formulas that add a `Month` column and
compute derived fields. This pipeline ports that logic to JavaScript so it runs in the browser.

- **`stage_a.js`** — *ingestion*. Maps each raw export into the master's raw-sheet layout,
  adds the month, and computes the derived columns (IVR authentication status, CMS handle-time
  conversions, hospital reason merge, message column mapping, EPIC target-days lookup).
- **`pipeline.js`** — *aggregation + recalc*. Ports the 7 aggregation macros (EPIC WQ/Productivity,
  IVR, CMS phone, Flow Out, POS Hospital, MyChart & Email, Collections) that write the new month's
  column, then evaluates the `Data Tables` formula cells (SUM / ratios / IFERROR / header refs) for
  that column, because the browser can't recalculate Excel formulas.
- **`build_html.py`** — injects the two scripts into the original HTML and wires up the two-file
  drop flow, the status panel, and the **Download 12-Month Data** button.

Rebuild the HTML after changing the scripts:

```bash
python3 pipeline/build_html.py
```

## Validation

The port is verified against the master's own data. `pipeline/validate.js` re-derives a month's
`Data Tables` column from the raw sheets and compares it, cell by cell, to the value the Excel
macros produced.

```bash
cd pipeline && npm install xlsx && TGT=2026-06 node validate.js
```

- **June (the most recent fully-present month): 0 mismatches across all 9 tables.** Re-deriving
  June is the exact operation performed when adding a new month, so this proves correctness.
- IVR and CMS derived-column formulas reproduce the cached master values exactly (0 / 23,994 and
  0 / 198).
- The formula recalculator reproduces all 127 charted formula cells for June exactly (the 55
  `INDEX/MATCH` appendix formulas are left for Excel to recompute on open — no dashboard chart
  uses them).
- End-to-end in a real browser, the downloaded July workbook's driver totals match the validated
  pipeline (WQ 20,048 · Hospital 2,804 · CMS calls 12,218 · IVR $189,979.49).

## Two things that need your input each month

1. **Collections data is not in the standard raw export.** It comes from two tabs inside the
   master (`Monthly Collections Data` = team goal/actual, `Rep Collections Data` = per-rep). The
   monthly raw file's `Rep Collection Data` tab is usually empty. To include collections for the
   new month, put the per-rep rows (Date, Employee Name, Total Collections, Transaction Count,
   Name) into that tab and the monthly goal/actual into `Monthly Collections Data`, then re-drop.
   Otherwise the collection charts stay blank for the new month.
2. **The CMS per-rep appendix detail is not auto-computed** — its source time format is ambiguous
   and no dashboard chart uses it.

## Notes / limitations

- The downloaded workbook keeps every sheet and updates the `Data Tables` grid, but cell
  formatting, macros (VBA), and the in-file `Data Graphs` charts are **not** preserved by the
  in-browser writer. The on-screen dashboard and the PPTX/HTML exports replace those.
- The download grows as raw history accumulates (~14 MB with July). If that becomes unwieldy the
  raw sheets can be trimmed from the output.
