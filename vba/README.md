# GAAP tab — actual vs. forecast formatter

`modGAAPActualForecast.bas` is an Excel VBA module that does the monthly
re-formatting of the **GAAP** tab in the FY forecast workbook, so a reader can
tell at a glance which columns are closed months and which are still forecast.

## What one run does

Given the last closed month (it asks), the macro:

| | |
|---|---|
| Banner (row 3) | Re-splits and re-merges `FY26 Actual` / `FY26 Forecast` at the right column, and labels each with the months it covers — e.g. *FY26 Actual (Oct - Jul)*, *FY26 Forecast (Aug - Sep)* |
| Month header (row 4) | Dark blue for closed months, dark amber for forecast months, grey for a column that mixes both |
| Statement body | Pale blue / pale amber / grey shading down the whole statement, in the same three states |
| Mixed columns | A quarter column is grey when only part of the quarter has closed — after a July close, **Q4** and **Total** are grey, because each is part actual and part forecast |
| Divider | Clears last month's divider and draws a heavy vertical rule on the current actual/forecast boundary, from the banner to the bottom of the statement |
| Legend | A three-cell colour key in row 2, above the block |

Nothing else is touched: no formulas, values, fonts, number formats, row
heights or column widths.

## Macros

| Macro | Does |
|---|---|
| `FormatGAAP_ActualVsForecast` | Formats the **GAAP** tab. Asks for the last closed month. |
| `FormatAllTabs_ActualVsForecast` | Same, across every tab in `TARGET_SHEETS` (**GAAP** and **Combined**), asking once so the tabs cannot drift apart |
| `ResetGAAP_Formatting` | Puts the GAAP tab back exactly as it was before the macro first ran |
| `ResetAllTabs_Formatting` | Same, for every tab in `TARGET_SHEETS` |

## Install

1. Open the forecast workbook and press `Alt`+`F11` for the VBA editor.
2. **File → Import File…**, pick `modGAAPActualForecast.bas`, and close the editor.
3. Save the workbook as **Excel Macro-Enabled Workbook (.xlsm)** — a plain
   `.xlsx` cannot store macros.
4. Run it: `Alt`+`F8` → `FormatGAAP_ActualVsForecast` → **Run**.

Optional, for one-click use: **Insert → Shapes**, draw a button on the GAAP
tab, right-click → **Assign Macro** → `FormatGAAP_ActualVsForecast`.

If you would rather keep the module in `PERSONAL.XLSB` and run it against
whatever workbook is open, set `USE_ACTIVE_WORKBOOK = True` at the top of the
module.

## Monthly use

Run the macro and answer the prompt with the month that just closed — `Jul`,
`July`, or the calendar month number `7` all work. The prompt is pre-filled
with last month and shows what the sheet is currently marked through, so a
double-check is one glance. Answer `0` if nothing has closed yet.

## Undo

Excel's Ctrl+Z does not undo macros, so the module keeps its own safety net.
The first time a tab is formatted, its original shading and vertical borders
are copied to a very-hidden sheet named `_fmtbak_<tab>`. Every run restores
that baseline before painting, so runs never stack up, and the `Reset...`
macros put the tab back to exactly how it looked before the macro was ever
used. Leave those `_fmtbak_` sheets in the workbook.

## Settings

All at the top of the module:

| Setting | Default | Meaning |
|---|---|---|
| `TARGET_SHEETS` | `"GAAP,Combined"` | Tabs the "AllTabs" macros handle |
| `MAIN_SHEET` | `"GAAP"` | Tab the plain macros handle |
| `FY_LABEL` | `"FY26"` | Fiscal year to find and label — change to `FY27` next October |
| `SHOW_LEGEND` | `True` | Write the colour key in row 2 |
| `KEEP_HIGHLIGHTS` | `True` | Cells already highlighted before the macro ran keep their colour |
| `BANNER_SHOWS_RANGE` | `True` | Put the month range in the banner text |
| `QUARTER_LINES` | `True` | Medium rule after each quarter column in the header rows |
| `USE_ACTIVE_WORKBOOK` | `False` | `True` if the module lives in `PERSONAL.XLSB` |

The colours are set in `InitPalette` as plain `RGB(...)` values — edit them
there if you want a different palette.

## How it finds things

The macro reads the sheet rather than relying on fixed cell addresses: it
looks for the `FY26` banner in row 3 above an `Oct` in row 4, walks right to
the `Total` column, and reads each column's `Oct…Sep` / `Q1…Q4` / `Total`
label to work out which fiscal months it covers. The bottom of the statement
is found by content, stopping above the numbered footnotes. So inserting a
line item, or the block moving a column or two, does not break it. What it
does assume, matching the current workbook:

- the banner is in row **3** and month labels in row **4** (`BANNER_ROW` / `MONTH_ROW`)
- the fiscal year runs **Oct → Sep**
- each FY block ends with a column labelled `Total`

If a tab does not match, the macro says so and leaves that tab alone.

## Note on the Combined tab

At the time this was written, **GAAP** was marked actual through Jun but
**Combined** was still marked actual through Feb — the two tabs had drifted.
`FormatAllTabs_ActualVsForecast` asks once and applies the same answer to
both, which keeps that from happening again.

## Not included

The macro is formatting only. It does not change the YTD formulas in column
`AM` (currently `=+W6+AA6+AE6`, i.e. whole quarters Q1–Q3). When a close lands
mid-quarter, that column still needs its own update — say the word and it can
be added.
