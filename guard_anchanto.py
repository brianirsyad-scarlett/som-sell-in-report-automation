"""
Guard: fail the run if the Anchanto e-commerce units in the freshly combined Primary_Sales
do not match what Anchanto.parquet says they should be under the counting rule.

Runs between "Combine -> Primary_Sales.parquet" and "Upload Sell In outputs to GCS", so a bad build is
never uploaded or promoted to production (the later steps are skipped when this one fails, and the
"Email on failure" step sends the mail).

Why it exists (2026-10-08): the pipeline once published an October with every in-transit "dispatched" order
counted (547K e-commerce units instead of 398K) because stale code ran. Nothing checked the result.

The expected figure is recomputed here straight from the parquet with plain polars, deliberately NOT through
anchanto_from_parquet.py, so a bug or a stale copy of that module cannot hide itself:

    counted  = Order Status DELIVERED
               or (DISPATCHED and sale date <= today - STALE_DISPATCH_DAYS)      # Anchanto never flips some orders
    sale date = SentOn (delivery, else dispatch, else scheduled), else CreatedOn
    e-commerce = TikTok Shop / Shopee Mall / Shopee Star / Lazada / Zalora / Bukalapak / Tokopedia
    Type of Item = ACTIVITY or REGULER  (what Primary_Sales keeps)

and compared per SentOn month with the Anchanto e-commerce rows of Primary_Sales. Primary_Sales also keeps
return (-R) orders and internal rows that anchanto_report_v2 drops, so a small gap is normal (measured
0.03-0.1%); the tolerance below leaves room for that and nothing more.

    python guard_anchanto.py --parquet work/inputs/Anchanto.parquet --primary work/Primary_Sales.parquet
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

import polars as pl

STALE_DISPATCH_DAYS = 14          # keep equal to anchanto_from_parquet.STALE_DISPATCH_DAYS (checked below)
FIRST_MONTH = "2025-01"           # Primary_Sales history starts here
TOLERANCE = 0.005                 # 0.5 % of the expected units
MIN_ABS = 500                     # ... and at least this many units, so a tiny month cannot trip it
ECOMMERCE = ["TIKTOK SHOP", "SHOPEE MALL", "SHOPEE STAR", "LAZADA", "ZALORA", "BUKALAPAK",
             "TOKOPEDIA", "TOKOPEDIA & SHOP"]


def expected_units(parquet: Path, today: dt.date) -> pl.DataFrame:
    cutoff = today - dt.timedelta(days=STALE_DISPATCH_DAYS)
    lf = pl.scan_parquet(parquet).select(
        pl.col("Marketplace").str.to_uppercase(),
        pl.col("Order Status").str.to_uppercase().alias("status"),
        pl.coalesce(pl.col("SentOn"), pl.col("CreatedOn")).dt.date().alias("sale"),
        pl.col("Type of Item").str.to_uppercase().alias("type"),
        pl.col("Ordered Quantity").alias("q"),
    )
    salestype = pl.col("Marketplace").str.split("_").list.last().str.replace("NEXUS", "ENDORSE")
    lf = lf.filter(
        ((pl.col("status") == "DELIVERED") | ((pl.col("status") == "DISPATCHED") & (pl.col("sale") <= cutoff)))
        & salestype.is_in(ECOMMERCE)
        & pl.col("type").is_in(["ACTIVITY", "REGULER"])
    )
    return (lf.group_by(pl.col("sale").dt.strftime("%Y-%m").alias("ym"))
              .agg(pl.col("q").sum().alias("expected"))
              .collect())


def actual_units(primary: Path) -> pl.DataFrame:
    return (pl.scan_parquet(primary)
              .filter((pl.col("ReportSource") == "ANCHANTO") & (pl.col("Channel") == "E-COMMERCE"))
              .group_by(pl.col("SentOn").dt.strftime("%Y-%m").alias("ym"))
              .agg(pl.col("Quantity").sum().alias("actual"))
              .collect())


def check_constant() -> list[str]:
    """The rule constant in the pipeline module must equal the one used here."""
    try:
        import anchanto_from_parquet as afp
    except Exception as exc:                                  # pragma: no cover
        return [f"cannot import anchanto_from_parquet: {exc}"]
    if getattr(afp, "STALE_DISPATCH_DAYS", None) != STALE_DISPATCH_DAYS:
        return [f"STALE_DISPATCH_DAYS differs: pipeline={getattr(afp, 'STALE_DISPATCH_DAYS', None)} guard={STALE_DISPATCH_DAYS}"]
    return []


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--parquet", type=Path, required=True, help="Anchanto.parquet the build read")
    ap.add_argument("--primary", type=Path, required=True, help="the combined Primary_Sales.parquet to check")
    ap.add_argument("--today", type=dt.date.fromisoformat, default=None, help="override today (testing)")
    ap.add_argument("--first-month", default=FIRST_MONTH)
    args = ap.parse_args(argv)

    today = args.today or dt.datetime.now().date()
    problems = check_constant()

    exp, act = expected_units(args.parquet, today), actual_units(args.primary)
    j = (exp.join(act, on="ym", how="full", coalesce=True)
            .with_columns(pl.col("expected").fill_null(0), pl.col("actual").fill_null(0))
            .filter(pl.col("ym") >= args.first_month)
            .sort("ym")
            .with_columns((pl.col("actual") - pl.col("expected")).alias("diff"))
            .with_columns((pl.col("diff") / pl.col("expected").clip(lower_bound=1) * 100).alias("pct")))

    print(f"Anchanto e-commerce units, Primary_Sales vs {args.parquet.name} (today={today}, "
          f"dispatched counted after {STALE_DISPATCH_DAYS} days, tolerance {TOLERANCE:.1%} / {MIN_ABS} units)")
    print(f"{'month':8s} {'expected':>12s} {'primary':>12s} {'diff':>10s} {'%':>7s}")
    for r in j.iter_rows(named=True):
        bad = abs(r["diff"]) > MIN_ABS and abs(r["diff"]) > TOLERANCE * max(r["expected"], 1)
        if bad:
            problems.append(f"{r['ym']}: Primary_Sales {r['actual']:,} vs expected {r['expected']:,} ({r['pct']:+.2f}%)")
        print(f"{r['ym']:8s} {r['expected']:>12,} {r['actual']:>12,} {r['diff']:>+10,} {r['pct']:>+6.2f}% {'  <-- MISMATCH' if bad else ''}")

    if problems:
        print("\nGUARD FAILED - Primary_Sales is NOT being uploaded or promoted:")
        for p in problems:
            print("  -", p)
        print("Most likely cause: stale code or the wrong counting rule (see the module docstring).")
        return 1
    print("\nGUARD OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
