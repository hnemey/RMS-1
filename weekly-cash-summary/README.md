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
Weeks previously exiled to Sheet1 are recovered into the ledger. The confirm
dialog names any payer row whose label matched no standard bucket.

## Tests

    ./run-tests.sh

- `test1` — migration off a synthetic legacy grid reproduces the old
  dashboard numbers exactly (the pre-change `computeRollingDash` is frozen
  inside the test as the baseline).
- `test2` — window/round-trip: ledger keeps all weeks, window trims to 52,
  xlsx round-trips, blank-week averaging matches the workbook rule.
- `test4` — end-to-end `saveToRolling` against a mock filesystem: bucketing,
  adjustment netting, backup, no history loss, re-save clears stale items.
- `test5` — `saveRollingEdits`, and cross-foot-row detection by behaviour.
