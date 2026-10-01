# som-sell-in-report-automation (draft)

The SOM Sell In chain (monthly sales workbooks → `Primary_Sales.parquet`) on
GitHub Actions, reading its inputs from GCS instead of the SOM laptop.

It is a **draft running beside the local chain**, not a replacement yet: it
writes only under `gs://bucket_som/sales_parquet/raw/primary/sell_in/`, never
over the production `sales_parquet/Primary_Sales.parquet`, and runs on manual
trigger only.

## What it replaces from the local chain

Local `Automation\Sales-MonthlyReport\run_daily.ps1` runs 5 steps:

| # | Local step | Here |
|---|---|---|
| 0 | `preflight_sources.ps1` (force-run stale Odoo/Anchanto) | not needed - those run on GitHub themselves |
| 1 | `sales_monthly_report.py` → `<year> <MM> <Mon>.xlsx` | `sales_monthly_report.py` (same code, GCS inputs) |
| 2 | `primary_sales_csv_and_parquet.py` → `combined_sales.parquet` | `build_primary_parquet.py` |
| 3 | `refresh_sell_in.ps1` (Sell In Offline MT/GT, Excel COM) | **stays local** - needs Excel |
| 4 | `validate_chain.py` | `compare_with_production.py` |

## Inputs (GCS)

| File | GCS | Kept fresh by |
|---|---|---|
| Master Data Sales.xlsx | `sales_parquet/raw/master data/` | uploaded by hand |
| Odoo Report.xlsx | `sales_parquet/raw/primary/odoo/` | som-odoo-report-automation |
| Anchanto.parquet | `sales_parquet/raw/primary/anchanto/` | som-anchanto-report-automation |
| 0. 2025 Accurate.xlsx | `sales_parquet/raw/primary/accurate/` | **copied once - not synced yet** |
| E-Stock 2025.xlsx | `sales_parquet/raw/primary/e-stock/` | static since Apr 2025 |

### Anchanto from the parquet

Locally, Sell In reads `Anchanto Report <Y> Q<n>.xlsx`, which
`build_q3_anchanto_report.py` builds from the Anchanto CSVs. Here
`anchanto_from_parquet.py` applies that script's rules directly to
`Anchanto.parquet` (status filter, Marketplace → SalesType, Kode Pos / Channel /
Product lookups, SentOn = the parquet's SentOn (delivery, else dispatch, else scheduled date) else CreatedOn, Type of Item filter), so
the quarterly workbook is not needed.

`--anchanto-scope quarter` (default) reads the same source files the quarterly
workbook covers, so numbers match the local report. `all` reads every file,
which also picks up orders placed earlier and dispatched late.

## Outputs (GCS, under `sales_parquet/raw/primary/sell_in/`)

- `monthly_xlsx/<year>/<year> <MM> <Mon>.xlsx` - the rebuilt monthly workbooks
- `monthly_parquet/<year> <MM> <Mon>.parquet` - one per month
- `Primary_Sales.parquet` - all months stacked; same shape as production

Only the months a run rebuilds are recomputed; the rest come from
`monthly_parquet/`, seeded once from the laptop's `Sales csv` folder by
`seed_history.py`. The seeded history stacked together is value-for-value
identical to the local `combined_sales.parquet` (checked row by row).

## Setup

Secret `GCP_SA_KEY` - the same GCS service-account JSON the other SOM pipelines use.

One-time, on the SOM laptop:

```
python seed_history.py
```

## Run

Actions → *Sell In Pipeline (draft)* → Run workflow. `months` defaults to
`prev,current` (Jakarta dates), same as the local daily run.

Locally:

```
python download_inputs.py
python download_history.py
python sales_monthly_report.py --month prev --month current --no-backup
python build_primary_parquet.py month work/output/*/*.xlsx
python build_primary_parquet.py combine work/monthly_parquet work/Primary_Sales.parquet
python compare_with_production.py work/monthly_parquet/"2026 09 Sep.parquet"
```
