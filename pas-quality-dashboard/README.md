# PAS Quality Dashboard — weekly grouping

Adds a **Group by: Month / Week** control to the dashboard, so a clinic in its first month of
onboarding can be shown week-over-week numbers while everyone else keeps the monthly view.

## How it works

### Week start day

Clinics do not all run Monday-to-Sunday, so **Week starts** (Mon–Fri, shown only while grouping by
week) sets the weekday a reporting week begins on. A clinic whose week starts Wednesday gets Wed–Tue
buckets; left on Monday, every total would blend two of that clinic's real weeks. The setting is
baked into exports alongside the granularity, so a clinic's file cuts weeks exactly as the dashboard
did. Changing it re-cuts every bucket, so the selected period is dropped and re-defaulted.

The whole Monday assumption was three copies of `(getDay() + 6) % 7`; they are now
`(getDay() - weekStart + 7) % 7`. Card notes state the definition in words ("weeks run Wed–Tue") and
bar tooltips spell out the range ("Wed Jun 3 – Tue Jun 9") so a wrong setting is visible immediately.

A period key is *derived* from each row's date — it is never read from a column in the source
export. Month keys are `YYYY-MM`; week keys are the **Monday** of that week as `YYYY-MM-DD`
(sortable, never confusable with a month key, and free of the ISO week-year edge cases that a
`2026-W01` key would bring: week 53, a Jan 1 belonging to the previous year).

Because the key is derived, granularity is purely a view concern:

| | Month | Week |
|---|---|---|
| On-screen buckets, charts, grids | by month | by week |
| Exported clinic/employee HTML | by month | by week, **locked** — no reader control |
| Historical archive (`Period` column) | `YYYY-MM` | `YYYY-MM` — unchanged |
| Executive rollup (`.pasq.json`) | `version: 1`, month-keyed | `version: 1`, month-keyed — unchanged |

The archive and the rollup deliberately never see a week key. The Executive View merges rollups by
`clinic -> month` and de-duplicates a clinic by replacing its months, so a mix of month- and
week-keyed files in one folder would double-count. Nothing downstream needs updating.

### Notable behaviour

- **Windowed trends.** The period-over-period chart and the employee grid show the most recent 13
  weeks (24 months). What was dropped is stated in the card note rather than silently trimmed, and
  the Total column sums only the periods actually rendered.
- **Part-periods marked.** The first and last week of an import window are usually incomplete and
  would otherwise read as a collapse in volume. They carry a `*` and a footnote. Exported files get
  the true date window (`monthWindow`) so the markers survive the export's collapsed rows.
- **Registrations stay monthly.** A registrations file names its month in the count column's header
  and reports one figure for the whole month, so there is no honest way to cut it into weeks. That
  section stays by month at week granularity and says so — rather than disappearing, which is what
  the unmodified `regScope()` would have done.

## Executive export: Save as dialog

"Export for executive view (data)" opens a real **Save as** dialog (`showSaveFilePicker`) so the
rollup can be named and written straight into the shared executive folder, rather than landing in
Downloads under a generated name. The generated name is still offered as the default.

The API is Chrome/Edge only and needs the click's user activation, so `exportExecutiveData` builds
its JSON synchronously and awaits `saveBlobAs` immediately — anything asynchronous in between spends
the activation and the dialog is refused. Firefox, Safari, and policy-blocked Chrome fall back to the
ordinary download and the status line says so. Cancelling writes nothing and says nothing was saved.

Note: the historical archive still auto-downloads after every import (pre-existing behaviour), and
the clinic/employee HTML exports still use the plain download. `saveBlobAs` is generic if those
should get the dialog too.

## Editing this file

`PAS_Quality_Dashboard.html` is one generated 1.2 MB file. Its script blocks are too large to edit
in place, and the viewer JS is embedded **twice** — as live source and as a JSON-encoded copy inside
`__RMS_ASSETS__` that exported files fall back to. So:

```sh
node tools/split.js split      # template.html -> parts/*.js  (only needed after a new upstream drop)
$EDITOR parts/viewer.js        # or parts/copays.js, parts/shell.js
node tools/split.js assemble   # parts/*.js -> PAS_Quality_Dashboard.html, re-encoding the asset copy
```

`template.html` is the pristine upstream file and is never written to. `assemble` always splices
into it, because the block boundaries in `tools/split.js` are *template* line numbers — assembling
against a previous build would splice at boundaries the last edit had already shifted.

## Verified

Driven through Chromium against a generated `.xlsx`, importing via the app's own file input:

- Week/month bucketing, Monday boundaries, and the 2025→2026 year rollover (Sun Jan 4 belongs to the
  week of Mon Dec 29).
- Grid totals reconcile with the row count in both modes (150 = 67+61+22 monthly = 9+15+…+6 weekly).
- Exported clinic file carries `"granularity":"week"`, renders weekly, and contains exactly one
  `<select>` (the period picker) — no granularity control for the reader.
- Archive and rollup stay month-keyed while the screen is in week mode.
- Archive round-trip: re-importing the same archive twice does not double-count.
- Copays tab across both granularities: two weeks of copay rows total 30 / $750, and a single week
  is exactly half (15 / $375), on screen and inside an exported file.
- No console or page errors in the app or in an exported file.
- Executive export save dialog, with the picker stubbed, across all five outcomes: success (correct
  rollup JSON written, chosen filename reported), cancel (nothing written), unsupported browser and
  policy-blocked (both fall back to the download), and a write failure (reported, not swallowed).
  User activation confirmed active at the moment the picker is invoked.
