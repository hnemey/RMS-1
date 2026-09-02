# GAAP tab — move the Actual / Forecast split

`modGAAPActualForecast.bas` is an Excel VBA module that does the one job you
repeat every month on the **GAAP** tab: moving the `FY26 Actual` /
`FY26 Forecast` split in row 3 as another month closes.

It works in three steps:

1. **Un-merge** the banner row across the FY26 block.
2. **Merge it again correctly** — Actual over the closed months, Forecast over
   the rest — reusing the labels already on the sheet, word for word, and the
   banner formatting from the banner cells already there.
3. **Put white rules back on everything** in the two header rows: around the
   outside, between every column, and between the two rows.

Step 3 is why the header always comes out consistent — every edge is painted
fresh each run, so a rule left behind by last month's split cannot survive,
and nothing is left without its white outline. Inside a merged banner Excel
draws nothing, so the banner row reads as the two blocks with a single rule
between them.

The rule's style is read off the sheet rather than hardcoded — thick white on
GAAP, medium white on Combined — and kept as a theme colour, so it still
follows the workbook theme. Set `RULE_WEIGHT` if you want to force one.

Nothing outside the two header rows is touched: no colours, no month names, no
shading, no legend, and no formulas, values, fonts, number formats or column
widths.

## Macros

| Macro | Does |
|---|---|
| `UpdateGAAP_Split` | GAAP tab. Asks which month closed, pre-filled with last month — press Enter to accept |
| `UpdateGAAP_Split_Auto` | GAAP tab, no prompt at all: assumes last calendar month closed. For a button |
| `UpdateAllTabs_Split` | GAAP **and** Combined, asking once so the two cannot drift apart |
| `UpdateAllTabs_Split_Auto` | Same, no prompt |

## Install

1. Open the workbook, press `Alt`+`F11` for the VBA editor.
2. **File → Import File…**, pick `modGAAPActualForecast.bas`, close the editor.
3. Save as **Excel Macro-Enabled Workbook (.xlsm)** — a plain `.xlsx` cannot
   store macros.
4. Run it: `Alt`+`F8` → pick the macro → **Run**.

For one-click use: **Insert → Shapes**, draw a button on the GAAP tab,
right-click → **Assign Macro** → `UpdateGAAP_Split_Auto`.

If you would rather keep the module in `PERSONAL.XLSB` and run it against
whatever workbook is open, set `USE_ACTIVE_WORKBOOK = True` at the top.

## Monthly use

Run it and answer with the month that just closed — `Jul`, `July` or the
calendar month number `7` all work, and `0` means nothing has closed yet. The
box is pre-filled with last month and shows what the sheet is currently marked
through, so most months it is just Enter.

Excel has no undo for macros, but there is nothing to undo here: if you answer
with the wrong month, run it again with the right one.

Note the pre-filled month is **last calendar month**, which is what the `_Auto`
macros use with no prompt. Running in September therefore assumes August
closed, putting Forecast at `Sep | Q4 | Total`. Type a different month if you
are working a month behind.

## Settings

At the top of the module:

| Setting | Default | Meaning |
|---|---|---|
| `MAIN_SHEET` | `"GAAP"` | Tab the plain macros use |
| `TARGET_SHEETS` | `"GAAP,Combined"` | Tabs the "AllTabs" macros use |
| `FY_LABEL` | `"FY26"` | Fiscal year to find — change to `FY27` next October |
| `USE_ACTIVE_WORKBOOK` | `False` | `True` if the module lives in `PERSONAL.XLSB` |

## How it finds things

It reads the sheet rather than relying on fixed cell addresses: it looks for
the `FY26` banner in row 3 above an `Oct` in row 4, walks right to the `Total`
column, and reads each column's `Oct…Sep` / `Q1…Q4` / `Total` label to work
out which fiscal months it covers. The forecast banner starts at the first
column that is not entirely closed — so after a June close it starts at `Jul`
(with `Q3` still inside the actual banner, as it is now), and after a July
close it starts at `Aug`, leaving `Q4` and `Total` on the forecast side.

What it assumes, matching the workbook as it stands:

- the banner is in row **3**, month labels in row **4**
- the fiscal year runs **Oct → Sep**
- each FY block ends with a column labelled `Total`

If a tab does not match, the macro says so and leaves that tab alone.

## Note on the Combined tab

The **Combined** tab has the same layout and was marked actual through **Feb**
while GAAP said **Jun** — the two had drifted. `UpdateAllTabs_Split` asks once
and applies the same answer to both.

That tab also carried two stray white lines in the month header row, left
behind by splits that sat there in earlier months. Since step 3 repaints every
edge, they are tidied up on the first run.

## Not included

This is the banner only. The YTD column `AM` (`=+W6+AA6+AE6`, whole quarters
Q1–Q3) is a formula that still needs its own update when a close lands
mid-quarter — say the word if you want that automated too.
