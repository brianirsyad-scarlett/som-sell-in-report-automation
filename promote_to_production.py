"""
Copy the GitHub-built parquets over the production names in sales_parquet/.

    python promote_to_production.py            # dry run: show what would change
    python promote_to_production.py --apply    # back up production, then copy

Every copy is server-side (no download). Before overwriting, each existing
production file is copied to sales_parquet/backup/<YYYYMMDD-HHMM WIB>/, so a
promotion can be undone by copying those back.

Note: the laptop's mirror daemon still uploads these production names from the
local chain. Until it stops, its next upload replaces what this wrote.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from zoneinfo import ZoneInfo

import pyarrow.parquet as pq
from google.cloud import storage

import gcs_paths

WIB = ZoneInfo("Asia/Jakarta")

# GitHub-built source -> production name
PROMOTIONS = {
    gcs_paths.OUT_COMBINED: "sales_parquet/Primary_Sales.parquet",
    "sales_parquet/raw/primary/anchanto/Anchanto.parquet": "sales_parquet/Anchanto.parquet",
    "sales_parquet/raw/primary/pcc/PCC_Sales.parquet": "sales_parquet/PCC_Sales.parquet",
    "sales_parquet/raw/primary/pcc/PCC_Order_Number.parquet": "sales_parquet/PCC_Order_Number.parquet",
}
BACKUP_ROOT = "sales_parquet/backup"


def describe(blob) -> str:
    return f"{blob.size / 1e6:,.1f} MB, updated {blob.updated.astimezone(WIB):%Y-%m-%d %H:%M} WIB"


def schema(blob) -> tuple[list[str], int]:
    """Column names and row count, read from the parquet footer only."""
    with blob.open("rb") as fh:
        meta = pq.ParquetFile(fh).metadata
        return meta.schema.to_arrow_schema().names, meta.num_rows


def copy(bucket, src, dst_key: str) -> None:
    """Server-side copy; rewrite() handles objects too large for a single copy call."""
    dst = bucket.blob(dst_key)
    token, _, _ = dst.rewrite(src)
    while token is not None:
        token, _, _ = dst.rewrite(src, token=token)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args(argv)

    bucket = storage.Client().bucket(gcs_paths.BUCKET)
    stamp = dt.datetime.now(WIB).strftime("%Y%m%d-%H%M")
    plan = []
    for src_key, dst_key in PROMOTIONS.items():
        src = bucket.get_blob(src_key)
        if src is None:
            print(f"MISSING source gs://{gcs_paths.BUCKET}/{src_key}")
            return 1
        dst = bucket.get_blob(dst_key)
        src_cols, src_rows = schema(src)
        print(f"\n{dst_key}")
        print(f"  new  <- {src_key}  ({describe(src)}, {src_rows:,} rows)")
        if dst is None:
            print("  old     (none)")
        else:
            dst_cols, dst_rows = schema(dst)
            print(f"  old     ({describe(dst)}, {dst_rows:,} rows)")
            added = [c for c in src_cols if c not in dst_cols]
            dropped = [c for c in dst_cols if c not in src_cols]
            if added or dropped:
                print(f"  columns: +{added}  -{dropped}")
            else:
                print("  columns: same")
        plan.append((src, dst, dst_key))

    if not args.apply:
        print("\ndry run - nothing copied (pass --apply)")
        return 0

    print()
    for src, dst, dst_key in plan:
        if dst is not None:
            backup_key = f"{BACKUP_ROOT}/{stamp}/{dst_key.rsplit('/', 1)[-1]}"
            copy(bucket, dst, backup_key)
            print(f"backed up {dst_key} -> {backup_key}")
        copy(bucket, src, dst_key)
        print(f"copied    {src.name} -> {dst_key}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
