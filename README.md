# CRESTA Aggregate Tracker

Excel front end over an MS Access database for tracking catastrophe aggregates
by cedant, country and CRESTA zone. It is laid out the same way as the Universe
platform (screens → services → repositories → database, numbered migrations,
optimistic locking, audit log), using only Excel and Access.

## Download

**[Aggregate_Tracker_Kit.zip](https://github.com/iwadaya/cresta/raw/main/Aggregate_Tracker_Kit.zip)**: everything below in one zip. After downloading, right-click it → Properties → **Unblock**, then extract.

## What's in the repo

| Path | What it is |
|---|---|
| `Aggregate_Tracker.xlsx` | Clean workbook: Universe cedants, treaty types, classes, countries and currencies pre-loaded; no aggregates |
| `Aggregate_Tracker_DEMO.xlsx` | Same workbook with fictional "Demo -" cedants and aggregate rows, to see the reports filled in |
| `vba/*.bas` | The seven VBA modules (the data layer) |
| `Build-Workbook.cmd` / `.ps1` | One-click build: imports the modules into the `.xlsx` and saves `Aggregate_Tracker.xlsm` with buttons |
| `tools/build_workbook.py` | Generates both `.xlsx` files (`python tools/build_workbook.py .`, needs `openpyxl`) |

The built `.xlsm` and the `.accdb` database are git-ignored. The database holds
your live data and stays on your machine or a network drive.

## Setup (Windows)

1. Double-click `Build-Workbook.cmd`. It asks before temporarily enabling
   Excel's *Trust access to the VBA project object model* and restores the setting afterwards.
   Prefer not to? Open the `.xlsx`, press Alt+F11, use File → Import File on each `.bas`, save as `.xlsm`,
   then Alt+F8 → `InstallButtons`.
2. Open `Aggregate_Tracker.xlsm`, enable macros, click **1. Setup / create database**.
   This creates `AggregateTracker.accdb` next to the workbook, unless a path is set in `Start!C7`.
3. Requirements: Access, or the free *Microsoft Access Database Engine* with the same
   32/64-bit as Excel. Use a local or mapped-drive path, not a OneDrive web link.

## Daily use

1. **Pull everything from Access**.
2. New cedant? Click **Add cedant**, or type rows on *Cedants* and **Push lists + cedants**.
3. Enter aggregates on *Aggregates* (blue columns): 100% amounts per peril plus our **Written line %** and
   **Signed line %** for that row. To remove a saved row, type `X` in *Delete (X)*.
4. **Push aggregates**. The push saves everything or nothing. Fix any rows marked `ERROR` in the Status column.
5. Read the reports, all in USD with the UW-year filter on *Dashboard*:
   *100% Share*, *Written Share*, *Signed Share* (each with a View by Cedant / Country switch),
   *Factor Applied* (written and signed share x the factor in F4, default 0.35%), and *Matrix*
   (cedant x country on any basis).

## Written and signed lines

Each aggregate row carries our written line and signed line (0-100%). Written share = 100% x written
line; signed share = 100% x signed line. A push is refused if the signed line is above the written
line. A blank line means "not known yet" and gives a share of 0; the Dashboard counts these rows.

**Upgrading an existing database:** build the new `.xlsm`, keep it next to (or point Start!C7 at) the
existing `AggregateTracker.accdb` and click **Setup** once. Migration `006_written_signed_lines` adds
the two columns and the share views; your data is kept and pulled into the new workbook.

## Data model

`cedant_aggregate` has one row per cedant × country × CRESTA zone × UW year × treaty
type × class × currency (unique key), with separate EQ / WS / Flood / SRCC / Others
columns in original currency. Perils are never summed together; the headline figure is
the peak single-peril aggregate.

| Universe | Here |
|---|---|
| React screens | Excel sheets + `AggMacros.bas` |
| `services/`, `validation/` | `AggService.bas` |
| `modules/*/repositories` | `AggRepo.bas` |
| `db/pool.js`, `withTransaction.js` | `AggDb.bas` (late-bound ADO, parameterised commands) |
| `db/migrations` + `runMigrations.js` | `AggMigrate.bas` + `schema_migrations` |
| `startup/ensureReferenceData.js` | `AggSeed.bas` (insert-if-missing on every Setup) |
| `contract_cresta_data` | `cedant_aggregate` |
| `ref_exchange_rate.rate_to_usd` | `ref_exchange_rate` + `vw_fx_latest` |
| `db/optimisticLock.js` | `row_version` |
| `audit_log` | `audit_log` |

Saved queries in Access: `vw_aggregate_detail`, `vw_aggregate_usd`, `vw_agg_by_cedant`,
`vw_agg_by_country`, `vw_agg_by_cedant_country`, `vw_agg_missing_fx`, and for our share
`vw_aggregate_share_usd`, `vw_agg_share_by_cedant`, `vw_agg_share_by_country`,
`vw_agg_share_by_cedant_country` (100%, written and signed USD per peril).

To change the schema, add a new numbered migration in `AggMigrate.bas` and bump
`MIG_COUNT`. Never edit an applied migration.

## FX rates

Rates are USD per 1 unit. Only USD and the fixed pegs (SAR, AED, BHD, QAR, JOD, OMR) are
pre-filled. Enter the floating currencies on *Rates*. Until you do, rows in those
currencies are left out of the USD totals and counted on the Dashboard.
