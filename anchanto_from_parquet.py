"""
The Anchanto rows of one month, read from Anchanto.parquet instead of the local
"Anchanto Report <Y> Q<n>.xlsx" workbooks.

The quarterly workbook is built by Data\\Report\\Sales\\Anchanto Report\\
build_q3_anchanto_report.py. That script's rules are repeated here unchanged, so
the rows handed to Sell In match what it would have read from the workbook:

  1. Order Status in {delivered, dispatched}
  2. SalesType = Marketplace suffix after "_", Nexus -> Endorse, upper-cased
  3. Kode Pos (Postcode -> Province/City), Channel (SalesType) and Product
     (ItemName) looked up from Master Data Sales.xlsx
  4. SentOn = the parquet's SentOn (Delivery Date, else Dispatch Date, else Scheduled
     Date), else CreatedOn. NOT the local report's Dispatch Date - on purpose.
  5. Type of Item blank / ACTIVITY / REGULER only

Only the source files the quarterly workbook would have covered are read: the
three months of the month's quarter plus the "(3)" file of the month before the
quarter (scope="quarter", the default). scope="all" reads every file instead,
which also catches orders placed earlier than that and dispatched late - more
complete, but no longer comparable with the local report.

Anchanto.parquet stores Marketplace / Order Status / Item Name upper-cased, so
steps 1 and 2 compare upper-case values. The raw values are lower-case status
codes and fixed marketplace names, so the result is the same.
"""

from __future__ import annotations

import calendar
import datetime as dt
import logging
import re
from pathlib import Path

import pandas as pd
import polars as pl
import pyarrow as pa
import pyarrow.compute as pc
import pyarrow.parquet as pq

log = logging.getLogger("anchanto_from_parquet")

MONTH_ABBR = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

PARQUET_COLUMNS = [
    "Source", "Marketplace", "CreatedOn", "SentOn", "Order Number",
    "Item Name", "Order Status", "Shipping Postcode", "Ordered Quantity", "Unit Price",
]

# Same order as build_q3_anchanto_report.FINAL_COLUMNS - the quarterly workbook's header.
FINAL_COLUMNS = [
    "CreatedOn", "SentOn", "SalesType", "NoPesanan", "ItemName", "Quantity", "Unit Price",
    "Province", "City", "Channel Group", "Channel", "Sub Channel",
    "Brand", "Category", "Sub Category", "Variant", "Product Name", "Type of Item",
]


# --------------------------------------------------------------------------- #
# Which source files
# --------------------------------------------------------------------------- #

def quarter_source_tags(year: int, month: int) -> list[str]:
    """build_q3_anchanto_report's _month_tags + _spillover_tag, for the quarter of year-month."""
    quarter = (month - 1) // 3 + 1
    months = list(range((quarter - 1) * 3 + 1, (quarter - 1) * 3 + 4))
    tags = [f"{year % 100:02d} {m:02d} {MONTH_ABBR[m]}" for m in months]
    prev_month, prev_year = months[0] - 1, year
    if prev_month == 0:
        prev_month, prev_year = 12, year - 1
    tags.append(f"{prev_year % 100:02d} {prev_month:02d} {MONTH_ABBR[prev_month]} (3)")
    return tags


def _row_group_sources(pf: pq.ParquetFile) -> list[str | None]:
    """The Source value of each row group, from column statistics alone.

    BigQuery_Anchanto.py writes one source file per write_table call, so every
    row group holds a single Source. None means the statistics cannot prove that
    and the group has to be read to find out.
    """
    idx = pf.schema_arrow.get_field_index("Source")
    out = []
    for i in range(pf.metadata.num_row_groups):
        stats = pf.metadata.row_group(i).column(idx).statistics
        if stats is not None and stats.has_min_max and stats.min == stats.max:
            out.append(stats.min)
        else:
            out.append(None)
    return out


def read_parquet_rows(parquet_path: Path, year: int, month: int, scope: str) -> pd.DataFrame:
    pf = pq.ParquetFile(str(parquet_path))
    if scope == "quarter":
        pattern = re.compile("|".join(re.escape(t) for t in quarter_source_tags(year, month)),
                             re.IGNORECASE)

        def wanted(src: str) -> bool:
            # build_q3's filter on file names: a tag match, "Report" in the name, no "$".
            return bool(pattern.search(src)) and "REPORT" in src.upper() and "$" not in src

        sources = _row_group_sources(pf)
        groups = [i for i, s in enumerate(sources) if s is None or wanted(s)]
        kept = sorted({s for s in sources if s is not None and wanted(s)})
        log.info("  Anchanto sources (quarter scope): %s", kept)
    elif scope == "all":
        groups = list(range(pf.metadata.num_row_groups))
        wanted = None
    else:
        raise ValueError(f"unknown scope {scope!r}")

    if not groups:
        return pd.DataFrame(columns=PARQUET_COLUMNS)
    table = pf.read_row_groups(groups, columns=PARQUET_COLUMNS)

    if wanted is not None:
        # Row groups whose statistics were inconclusive were read whole - filter them now.
        srcs = table.column("Source").to_pylist()
        table = table.filter(pa.array([wanted(s or "") for s in srcs]))

    # Cheap pre-filter on the sale date before going to pandas: the quarterly
    # sale date is the parquet's SentOn, else CreatedOn. The exact month filter
    # happens again in the caller, on the normalised date.
    start = pa.scalar(dt.datetime(year, month, 1), pa.timestamp("us"))
    last = calendar.monthrange(year, month)[1]
    end = pa.scalar(dt.datetime(year, month, last) + dt.timedelta(days=1), pa.timestamp("us"))
    sent = pc.coalesce(table.column("SentOn"), table.column("CreatedOn"))
    table = table.filter(pc.and_(pc.greater_equal(sent, start), pc.less(sent, end)))
    return table.to_pandas()


# --------------------------------------------------------------------------- #
# Master data - verbatim from build_q3_anchanto_report.load_masters
# --------------------------------------------------------------------------- #

def load_masters(master_excel: Path):
    kode_pos = pd.read_excel(master_excel, sheet_name="Kode Pos", dtype=str)
    kode_pos = kode_pos.rename(columns={"Area Level 1": "Province", "Area Level 2": "City"})
    kode_pos["Kode Pos"] = kode_pos["Kode Pos"].astype(str).str.strip()
    kode_pos = kode_pos[kode_pos["Kode Pos"].notna() & (kode_pos["Kode Pos"] != "") & (kode_pos["Kode Pos"] != "nan")]
    kode_pos = kode_pos.drop_duplicates(subset=["Kode Pos"], keep="first")
    kode_pos = kode_pos[["Kode Pos", "Province", "City"]]

    channel = pd.read_excel(master_excel, sheet_name="Channel")
    channel = channel[["Channel Group", "SalesType", "Channel", "Sub Channel"]].copy()
    channel["SalesType"] = channel["SalesType"].astype(str).str.strip()

    product = pd.read_excel(master_excel, sheet_name="Product", skiprows=2, header=0)
    product = product.drop(columns=["No"], errors="ignore")
    keep_cols = ["ItemName", "Brand", "Category", "Sub Category", "Variant", "Product Name", "Type of Item"]
    product = product[keep_cols].copy()
    for c in keep_cols:
        # astype(str) turns a blank cell into "NAN" here, which the Type of Item
        # filter below then drops. Kept as-is: it is what the local report does.
        product[c] = product[c].astype(str).str.upper()

    log.info("  Anchanto masters: Kode_Pos=%d, Channel=%d, Product=%d",
             len(kode_pos), len(channel), len(product))
    return kode_pos, channel, product


# --------------------------------------------------------------------------- #
# Transform - build_q3_anchanto_report.transform on the parquet's columns
# --------------------------------------------------------------------------- #

def transform(df: pd.DataFrame, kode_pos, channel, product) -> pd.DataFrame:
    quantity = pd.to_numeric(df["Ordered Quantity"], errors="coerce").fillna(0).astype("int64")
    unit_price = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0).astype("int64")

    order_date = pd.to_datetime(df["CreatedOn"], errors="coerce").dt.normalize()
    sent_date = pd.to_datetime(df["SentOn"], errors="coerce").dt.normalize()

    status = df["Order Status"].astype(str).str.strip()
    mask = status.isin(["DELIVERED", "DISPATCHED"])

    marketplace_suffix = df["Marketplace"].astype(str).str.split("_", n=1, expand=True)
    marketplace_suffix = (marketplace_suffix[1] if 1 in marketplace_suffix.columns
                          else pd.Series([None] * len(df), index=df.index)).fillna("")
    marketplace_suffix = marketplace_suffix.str.replace("NEXUS", "ENDORSE", regex=False)

    out = pd.DataFrame({
        "CreatedOn": order_date,
        "SalesType": marketplace_suffix.str.upper(),
        "NoPesanan": df["Order Number"],
        "ItemName": df["Item Name"].astype(str).str.upper(),
        "Quantity": quantity,
        "Unit Price": unit_price,
        "PostCode": df["Shipping Postcode"],
        "SentOn": sent_date,
    })
    out = out[mask].reset_index(drop=True)

    out["PostCode"] = out["PostCode"].fillna("-")
    out.loc[out["PostCode"].astype(str).str.strip() == "", "PostCode"] = "-"

    out = out.merge(kode_pos, how="left", left_on="PostCode", right_on="Kode Pos")
    out = out.drop(columns=["Kode Pos", "PostCode"])
    out["City"] = out["City"].fillna("-")
    out["Province"] = out["Province"].fillna("-")

    out = out.merge(channel, how="left", on="SalesType")
    out = out.merge(product, how="left", on="ItemName")

    out.loc[out["SentOn"].isna(), "SentOn"] = out.loc[out["SentOn"].isna(), "CreatedOn"]

    out = out[FINAL_COLUMNS]

    type_mask = out["Type of Item"].isna() | out["Type of Item"].isin(["ACTIVITY", "REGULER"])
    return out[type_mask].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Entry point used by sales_monthly_report
# --------------------------------------------------------------------------- #

def anchanto_month(parquet_path: Path, master_excel: Path, year: int, month: int,
                   scope: str = "quarter", masters=None) -> pl.DataFrame:
    """The quarterly workbook's rows for one month, as sales_monthly_report reads them.

    Returned the way fastexcel hands a workbook sheet back: dates as Date,
    blanks as null, one row per workbook row.
    """
    masters = masters or load_masters(master_excel)
    raw = read_parquet_rows(parquet_path, year, month, scope)
    log.info("  Anchanto parquet rows for %d-%02d (pre-transform): %d", year, month, len(raw))
    out = transform(raw, *masters)
    # The workbook stores dates without a time; sales_monthly_report then types them as Date.
    out["CreatedOn"] = out["CreatedOn"].dt.date
    out["SentOn"] = out["SentOn"].dt.date
    return pl.from_pandas(out)
