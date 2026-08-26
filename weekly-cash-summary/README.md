# Weekly Cash Summary — Rolling Ave ledger

`Weekly_Cash_Summary_ledger.html` replaces the Rolling Ave workbook's 52-column
grid with a flat **ledger**: one sheet, one row per `(week, payor bucket)` and
one row per `(week, special item)`.

## Why

The grid had to be edited in place on every save — week columns shifted left,
the dropped week archived into Sheet1, TOTAL formulas regenerated, payer rows
re-found by label prefix — and the workbook was rewritten from a bare template
each time. That cost extra sheets, all formatting, and every formula's cached
value; shifted formulas ended up pointing one column off; and a label that
matched no bucket silently dumped that payor's week into ALL OTHER PAYORS.

With a ledger there are no anchors to re-find and nothing to shift. Saving a
week deletes that week's rows and re-inserts them. The 12-month window is a
query over the ledger, so **dropping the oldest week from view no longer
touches stored history** — nothing is ever archived or discarded.

## Format

| Column | Meaning |
| --- | --- |
| Week Ending | the week's Thursday, a real Excel date (for pivots/sorting) |
| Week Range  | `mm/dd/yy-mm/dd/yy` — **authoritative**, everything keys off this |
| Type        | `Payor` or `Item` |
| Bucket      | canonical payor key — **authoritative** for the maths |
| Name        | display label (finance's own wording, carried over on migration) |
| Category    | dashboard grouping — re-derived on read |
| Amount      | the figure |
| Updated     | when the row was last written |

`Week Ending` and `Category` are written for people and are recomputed on read,
so hand-editing either can't skew the numbers.

## Migrating

Open the app, link the folder, press **Save**. The first save detects the old
grid, writes `<name> (pre-ledger backup <stamp>).xlsx` beside it, then converts.
The confirm dialog names anything that needed a judgement call.

What migration does with the old workbook:

- **Payor rows** map to canonical buckets by label prefix. A row matching no
  key keeps its own bucket rather than being folded into ALL OTHER PAYORS, and
  a row whose prefix is already taken (`CIGNA EDGE TRANS…` vs `CIGNA …`) keeps
  its own identity too — the dashboard averages **per row**, so merging two
  rows would move the category average.
- **Sheet1** is the archive of weeks earlier saves exiled out of the grid. Its
  *payor* rows are recovered (they sit above TOTAL at a fixed layout and tie to
  the Rolling Ave sheet across the time boundary), which can add years of
  history. Its rows *below* TOTAL are deliberately ignored: that list is
  append-only and has grown since the archive was written, so row N there is
  not the same item it is on the Rolling Ave sheet. Reading it by row number
  invents settlements that never happened.
- **The cross-foot row** is dropped, detected by behaviour — a row that keeps
  reproducing the week's payor total — rather than by position.
- **Labels** are trimmed to the 16-character payer-name column, but only when
  trimming leaves the row in the same category.

Migration is checked against the frozen pre-change implementation: same
categories, same weekly series, same averages.

## Adjustments

An adjustment is deducted from its payor's week and *also* recorded as a
labeled `Item` row. Items are never summed into any total — the money is
removed once and shown once.

A week the loaded report covers is rebuilt wholesale from that report's rows,
so re-saving is naturally idempotent. A week **outside** the report is adjusted
in place in the ledger instead: what moves the payor figure is the delta against
the item row already stored. Re-saving changes nothing, editing the amount
applies only the difference, and setting it to zero backs the adjustment out and
removes the item row. That week must already exist in the ledger.

The editor tints rows that need attention:

| Tint | Meaning |
| --- | --- |
| red    | not written — no amount, no valid week, or a week in neither the report nor the ledger |
| amber  | written, but no description, so it saves as `Adjustment` and a second undescribed row on the same week would collide with it |
| none (with a note) | an earlier week, applied straight to the ledger |

Amounts normalise when the field is committed, never while typing: `00` becomes
empty, `007.500` becomes `7.5`. Blocked rows are named in the Save dialog rather
than being dropped silently.

## The leader report

`buildLeaderReport()` exports the tabbed HTML report. Tabs run **months newest
first**, with the fiscal-year-to-date summary **last** — and that YTD tab is the
one the report opens on.

The YTD view covers the **fiscal year to date** (the year starts 1 October;
`FY_START_MONTH`), not a trailing twelve months. Its table reads newest month on
the left, October on the right, then a `YTD` column and the payor mix. The
month-over-month shading still compares each month with the one that really
precedes it — the columns are reversed only at render, the arithmetic stays
chronological.

The heading is the span plus `YTD`, so a fresh fiscal year reads `October YTD`
and fills out to `October – August YTD`. It draws on the whole ledger rather
than the 12-month window, because a fiscal year runs up to 53 lockbox weeks and
a windowed read would quietly drop October late in the year.

## Tests

    ./run-tests.sh

`test7` needs jsdom (`npm i jsdom`) and is skipped if it isn't installed.

- `test1` — migration off a synthetic legacy grid reproduces the old
  dashboard numbers exactly (the pre-change `computeRollingDash` is frozen
  inside the test as the baseline).
- `test2` — window/round-trip: ledger keeps all weeks, window trims to 52,
  xlsx round-trips, blank-week averaging matches the workbook rule.
- `test4` — end-to-end `saveToRolling` against a mock filesystem: bucketing,
  adjustment netting, backup, no history loss, re-save clears stale items.
- `test5` — `saveRollingEdits`, and cross-foot-row detection by behaviour.
- `test6` — earlier-week adjustments: they apply, re-saving doesn't
  double-subtract, an amount edit applies only the delta, zero backs it out,
  and blocked rows create no phantom weeks.
- `test7` — the adjustments editor in a real DOM (jsdom): row tinting, amount
  normalisation on commit but not mid-typing, and in-place status repaint.
- `test8` — loads the migrated production ledger into the full app and renders
  the rolling tab (`node migrate.js` produces the file it needs).
- `test9` — the exported leader report off the real ledger: fiscal-year window,
  tab order with YTD last and open, newest-first columns, the `YTD` column
  totalling its row, shading still pointing at the true prior month, and the
  October-only labelling.
