"""
Seed the metric history store from previously extracted report text.

Reporter writes the raw text of every PDF it reads to "Insights Output/Debug Text"
as "<dealership> - <Mon YYYY> - sales.txt" / "- service.txt". That folder is the only
record of months whose source PDFs are no longer on this machine, so it is what the
history store can be seeded from.

Re-running is safe: a dealership-month already in the store is overwritten with the
same figures, never duplicated.

    python tools/backfill_metric_history.py
    python tools/backfill_metric_history.py --debug-text "<path>" --dry-run
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import dealership_insights_local as core  # noqa: E402
from reporting import metric_history  # noqa: E402

FILENAME = re.compile(r"^(?P<dealership>.+?) - (?P<month>[A-Za-z]{3,9} \d{4}) - (?P<kind>sales|service)\.txt$")


def collect(debug_dir: Path) -> dict[tuple[str, str], dict[str, Path]]:
    pairs: dict[tuple[str, str], dict[str, Path]] = defaultdict(dict)
    for path in sorted(debug_dir.glob("*.txt")):
        match = FILENAME.match(path.name)
        if not match:
            continue
        pairs[(match.group("dealership"), match.group("month"))][match.group("kind")] = path
    return pairs


def main() -> int:
    parser = argparse.ArgumentParser(description="Seed metric history from extracted report text.")
    parser.add_argument("--debug-text", default=str(core.DEBUG_TEXT_DIR), help="Folder holding the extracted text.")
    parser.add_argument("--store", default=None, help="History store path (defaults to data/metric_history.json).")
    parser.add_argument("--dry-run", action="store_true", help="Report what would be recorded without writing.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.WARNING, format="%(message)s")
    debug_dir = Path(args.debug_text)
    if not debug_dir.is_dir():
        print(f"No extracted text at {debug_dir}", file=sys.stderr)
        return 1

    store_path = Path(args.store) if args.store else None
    store = metric_history.load_store(store_path)
    pairs = collect(debug_dir)

    recorded = skipped = 0
    months: set[str] = set()
    for (dealership, month_label), files in sorted(pairs.items()):
        if "sales" not in files or "service" not in files:
            skipped += 1
            continue
        try:
            sales_text = files["sales"].read_text(encoding="utf-8", errors="replace")
            service_text = files["service"].read_text(encoding="utf-8", errors="replace")
            metrics = core.parse_report_metrics(sales_text, service_text)
        except Exception as exc:  # noqa: BLE001
            print(f"  skipped {dealership} {month_label}: {exc}", file=sys.stderr)
            skipped += 1
            continue
        if not metrics:
            skipped += 1
            continue
        metric_history.record_month(store, dealership, month_label, metrics)
        recorded += 1
        months.add(month_label)

    print(f"{recorded} dealership-months recorded, {skipped} skipped")
    print(f"months: {', '.join(sorted(months, key=lambda m: metric_history.month_key(m) or (0, 0)))}")
    print(f"dealerships: {len(list(metric_history.dealership_names(store)))}")

    if args.dry_run:
        print("dry run - nothing written")
        return 0
    target = store_path or metric_history.store_path()
    if not metric_history.save_store(store, store_path):
        return 1
    print(f"written to {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
