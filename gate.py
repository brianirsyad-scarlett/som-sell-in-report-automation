"""
Decide whether the Sell In -> PCC chain should run now.

The upstream pipelines (Odoo hourly, Anchanto every 6 hours, Accurate daily) run in
their own repos and GitHub starts scheduled runs late, so a fixed start time cannot
guarantee they are done. Instead this workflow polls every 20 minutes and builds only when ALL of these hold:

  * neither the Odoo nor the Anchanto pipeline is queued or running (so the build
    always sees both of their finished outputs, never one half-way through), and
  * every input is recent enough: Odoo <= 3 h, Anchanto <= 8 h, Accurate <= 30 h, and
  * at least one input is newer than the last Primary_Sales build.

Writes run=true|false and reason=... to $GITHUB_OUTPUT. With --final (the 07:00 slot),
an input that is still too old is an error, so it gets emailed.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import urllib.request
from zoneinfo import ZoneInfo

from google.cloud import storage

import gcs_paths

WIB = ZoneInfo("Asia/Jakarta")
UPSTREAM = {
    "Odoo": "sales_parquet/raw/primary/odoo/Odoo Report.xlsx",
    "Anchanto": "sales_parquet/raw/primary/anchanto/Anchanto.parquet",
    "Accurate": "sales_parquet/raw/primary/accurate/0. 2025 Accurate.xlsx",
}
# How old an input may be before the build waits for it.
MAX_AGE = {
    "Odoo": dt.timedelta(hours=3),
    "Anchanto": dt.timedelta(hours=8),
    "Accurate": dt.timedelta(hours=30),
}
# Odoo changes every hour, Anchanto every ~6 h. A build downloads ~1 GB, so an Odoo-only
# change rebuilds at most this often; a new Anchanto output rebuilds as soon as it lands.
MIN_REBUILD = dt.timedelta(hours=3)
# Pipelines whose runs in other repos must be finished before a build starts.
RUNNING_CHECKS = {
    "Odoo": ("brianirsyad-scarlett/som-odoo-report-automation", "odoo-export.yml"),
    "Anchanto": ("brianirsyad-scarlett/som-anchanto-report-automation", "anchanto-pipeline.yml"),
}


def updated(bucket, key):
    b = bucket.get_blob(key)
    return b.updated.astimezone(WIB) if b else None


def active_runs(repo: str, workflow: str) -> int | None:
    """Queued + in-progress runs of a workflow (both repos are public, so the token
    this job already has can read them). None if GitHub could not be asked."""
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    total = 0
    for status in ("queued", "in_progress"):
        url = f"https://api.github.com/repos/{repo}/actions/workflows/{workflow}/runs?status={status}&per_page=1"
        req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json",
                                                   **({"Authorization": f"Bearer {token}"} if token else {})})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                total += json.load(resp).get("total_count", 0)
        except Exception as exc:  # noqa: BLE001 - never block a build on a status lookup
            print(f"  could not read {repo} runs ({status}): {exc}")
            return None
    return total


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--final", action="store_true")
    args = ap.parse_args(argv)

    bucket = storage.Client().bucket(gcs_paths.BUCKET)
    now = dt.datetime.now(WIB)
    times = {name: updated(bucket, key) for name, key in UPSTREAM.items()}
    out_time = updated(bucket, gcs_paths.OUT_COMBINED)
    for name, t in times.items():
        print(f"{name:9s} {t:%Y-%m-%d %H:%M} WIB ({(now - t).total_seconds() / 3600:.1f} h ago)" if t
              else f"{name:9s} MISSING")
    print(f"Sell In   {out_time:%Y-%m-%d %H:%M} WIB" if out_time else "Sell In   (never built)")

    stale = [n for n, t in times.items() if t is None or now - t > MAX_AGE[n]]
    running = []
    for name, (repo, wf) in RUNNING_CHECKS.items():
        n = active_runs(repo, wf)
        print(f"{name:9s} pipeline runs queued/in progress: {'?' if n is None else n}")
        if n:
            running.append(name)
    newest = max(t for t in times.values() if t) if any(times.values()) else None

    if args.force:
        run, reason = True, "forced"
    elif running:
        run, reason = False, f"waiting for the {', '.join(running)} pipeline to finish"
    elif stale:
        run, reason = False, f"waiting for a fresh {', '.join(stale)}"
    elif out_time and newest and out_time >= newest:
        run, reason = False, "already built after the latest Odoo, Anchanto and Accurate"
    elif out_time and times["Anchanto"] and out_time >= times["Anchanto"] and now - out_time < MIN_REBUILD:
        run, reason = False, f"built {(now - out_time).total_seconds() / 3600:.1f} h ago; only Odoo/Accurate changed since (rebuild at most every 3 h)"
    else:
        run, reason = True, "inputs fresh and newer than the last build"
    print(f"run={run} ({reason})")

    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as fh:
            fh.write(f"run={'true' if run else 'false'}\nreason={reason}\n")
    if args.final and not run and stale and not running:
        print(f"Last slot of the morning and still {reason}.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
