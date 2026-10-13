"""Tiny GCS copy helper for the workflow.

    python gcs_copy.py get "sales_parquet/x.csv" work/x.csv
    python gcs_copy.py get "sales_order_number/" work/orders/      # every object under a prefix
    python gcs_copy.py put work/x.parquet "sales_parquet/raw/primary/pcc/x.parquet"
"""

import sys
import time
from pathlib import Path

from google.api_core.exceptions import NotFound
from google.cloud import storage

import gcs_paths


def main(argv) -> int:
    cmd, a, b = argv
    bucket = storage.Client().bucket(gcs_paths.BUCKET)
    if cmd == "get":
        if a.endswith("/"):
            dest = Path(b)
            dest.mkdir(parents=True, exist_ok=True)
            # A prefix can be re-exported while we copy it (sales_order_number/ is rewritten by the
            # BigQuery SOM_Daily export): a shard listed a moment ago is then gone (404). Start over.
            for attempt in range(1, 6):
                n = 0
                try:
                    for blob in bucket.list_blobs(prefix=a):
                        if blob.name.endswith("/"):
                            continue
                        blob.download_to_filename(str(dest / blob.name.rsplit("/", 1)[-1]), timeout=900)
                        n += 1
                    break
                except NotFound as exc:
                    print(f"  attempt {attempt}: {exc.message[:120]} - the prefix changed while copying, retrying")
                    if attempt == 5:
                        raise
                    time.sleep(45)
                    for old in dest.iterdir():
                        old.unlink()
            print(f"downloaded {n} object(s) from gs://{gcs_paths.BUCKET}/{a}")
        else:
            Path(b).parent.mkdir(parents=True, exist_ok=True)
            bucket.blob(a).download_to_filename(b, timeout=900)
            print(f"downloaded gs://{gcs_paths.BUCKET}/{a}")
    elif cmd == "put":
        bucket.blob(b).upload_from_filename(a, timeout=900)
        print(f"uploaded {a} -> gs://{gcs_paths.BUCKET}/{b}")
    else:
        raise SystemExit(f"unknown command {cmd!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
