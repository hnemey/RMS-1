# PAS Quality — executive rollout via a synced folder

Because the source data has to be pulled **per clinic**, the executive export is no longer a
self-contained HTML page. Each clinic now exports a small **data file**; every clinic's file lands in
one shared folder; the leader opens a separate viewer and points it at that folder.

## The two files

| File | Who uses it | What changed |
|---|---|---|
| `PAS_Quality_Dashboard.html` | Each clinic / analyst | The old **⬇ Export for leadership** button is now **⬇ Export for executive view (data)** and writes a `.pasq.json` rollup. Everything else — Clinic export, employee export, the Excel history archive, all three tabs — is untouched. |
| `PAS_Executive_View.html` | The leader | New. Links a folder, merges every rollup it finds, and renders the cross-clinic executive view. |

## The workflow

**Each clinic, once a month**

1. Open `PAS_Quality_Dashboard.html`, drop in that clinic's Epic exports as usual.
2. Click **⬇ Export for executive view (data)** — saves e.g. `PAS North Valley Family Medicine 2026-05 to 2026-06.pasq.json` (typically 5–70 KB).
3. Drop that file into the shared folder (OneDrive / SharePoint / network drive).

If a clinic filter is set on screen, the file covers just that clinic. With no filter it covers every
clinic in the loaded data — so one analyst pulling several clinics can ship them in a single file.

**The leader, once**

1. Open `PAS_Executive_View.html`, click **📁 Link Folder**, pick the shared folder.
2. Thereafter press **🔄 Update** to pull in whatever clinics have added. The browser remembers the
   folder between sessions — no re-linking.

## What the leader sees

- **Overview** — headline KPIs, month trend, warnings per 100 registrations, top warning types, clinic scorecard.
- **Clinics** — clinic × month matrices for warnings, warnings/100 registrations, and copay collection rate, each with a latest-vs-prior delta.
- **Employees** — every staff member across every reporting clinic in one sortable, searchable table.
- **Copays** — collection rate, dollars left on the table, by clinic and by month.
- **Coverage** — verification status mix and what needs work, by clinic.
- **Data sources** — which clinics reported, which months, when each file was exported, and a
  callout naming any clinic that has not yet reported the newest month.

Month and clinic filters at the top apply to every tab.

## The data format

`.pasq.json` — every number bucketed under **clinic → month**:

```
clinics["North Valley Family Medicine"].months["2026-05"] = {
  bypass:        { n, patients, types{}, details{}, byWeekday[7], byHour{}, byDay{}, employees{} },
  copays:        { countDue, countPaid, sumDue, sumPaid, employees{} },
  coverage:      { n, statuses{}, employees{} },
  registrations: { n, employees{} }
}
```

That bucketing is what makes the folder safe to drop files into repeatedly: for each
(clinic, month) pair the viewer keeps the block from the **most recently exported** file and ignores
the rest. A clinic re-exporting May replaces its own May rather than adding a second copy, and a
corrected re-export supersedes the original. The Data sources tab reports whenever this happens.

## Privacy

Rollups carry **counts only**. Patient identifiers are used to size a distinct-patient count and then
discarded — they never reach the file. Staff names are present, since the executive view drills to
employee level. Verified by an automated check that fails if any patient identifier appears in an
exported rollup.

One caveat on the numbers: "Patients affected" is a **sum of per-clinic-month distinct counts**, so
a patient seen in two months (or at two clinics) counts twice. The KPI is prefixed `≈` whenever more
than one bucket is being summed. Every other figure is exact and additive.

## Browser

**Link Folder** uses the File System Access API — **Chrome or Edge**. In Safari or Firefox the page
still works: drag the files (or the whole folder) onto the drop zone instead; you just re-drop each
visit rather than pressing Update.

Both files are fully self-contained and run locally — no server, no install, no network calls.

## Troubleshooting

| Symptom | Cause |
|---|---|
| "No .json rollup files in that folder" | Clinics saved the HTML export, not the data export. They need the **Export for executive view (data)** button. |
| A clinic is missing from the view | It has not dropped a file yet. The Data sources tab lists who is behind and on which month. |
| "N file(s) ignored" | Non-rollup `.json` files in the folder. Harmless — the Data sources tab names each one and why. |
| Update says the folder cannot be read | The folder was moved or renamed. Click **Link Folder** again. |
| Numbers look doubled | They should not be — check the Data sources tab, which reports any superseded clinic-month blocks. |

## Verification

Tested end to end in headless Chromium against generated Epic-shaped workbooks with known totals:
34/34 checks passing — per-clinic-month bypass, patient, copay (count and dollars), coverage and
registration figures all match the fixtures; MyChart auto-collections and blank-CSN rows excluded
exactly as the on-screen dashboard excludes them; no patient identifiers in the output; re-adding a
file leaves totals unchanged; a newer corrected file supersedes the old one. Scale check: 25 clinics
× 12 months (68,861 warnings, 1.7 MB) loads and renders in under a second, ~250 ms per tab.
