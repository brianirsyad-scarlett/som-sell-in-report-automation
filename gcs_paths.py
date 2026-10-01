"""Every GCS location this pipeline reads or writes, in one place."""

BUCKET = "bucket_som"

# Inputs: GCS object -> file name under work/inputs (the names sales_monthly_report expects).
INPUTS = {
    "sales_parquet/raw/master data/Master Data Sales.xlsx": "Master Data Sales.xlsx",
    "sales_parquet/raw/primary/odoo/Odoo Report.xlsx": "Odoo Report.xlsx",
    "sales_parquet/raw/primary/accurate/0. 2025 Accurate.xlsx": "0. 2025 Accurate.xlsx",
    "sales_parquet/raw/primary/e-stock/E-Stock 2025.xlsx": "E-Stock 2025.xlsx",
    # Written by som-anchanto-report-automation; has SentOn (delivery, else dispatch, else scheduled date).
    "sales_parquet/raw/primary/anchanto/Anchanto.parquet": "Anchanto.parquet",
}

# Outputs. All under sell_in/ - NOT sales_parquet/Primary_Sales.parquet, which the
# local pipeline still owns (the mirror daemon uploads it from combined_sales.parquet).
OUT_ROOT = "sales_parquet/raw/primary/sell_in"
OUT_WORKBOOKS = f"{OUT_ROOT}/monthly_xlsx"          # <year>/<year> <MM> <Mon>.xlsx
OUT_MONTHLY_PARQUET = f"{OUT_ROOT}/monthly_parquet"  # <year> <MM> <Mon>.parquet
OUT_COMBINED = f"{OUT_ROOT}/Primary_Sales.parquet"

# The production file, used only by compare_with_production.py.
PRODUCTION_COMBINED = "sales_parquet/Primary_Sales.parquet"
