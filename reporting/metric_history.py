"""
Persistent metric history.

The trend machinery in dealership_insights_local reads prior months by looking for
*last month's PDFs* in the same dealership folder. Operators do not keep old PDFs
there, so in the July 2026 run 344 of 490 metric instances had no history at all and
almost every insight could only be an observation, never a comparison.

This module keeps the parsed figures instead. Every run records what it read, so the
tool accumulates its own baseline: month on month, the same month last year, the
dealership's own rolling average, and whether a move is the sharpest in N months.

The store is deliberately outside the output folder, which is cleared per month.
"""

from __future__ import annotations

import json
import logging
import os
import re
import tempfile
from pathlib import Path
from typing import Any, Iterable

STORE_VERSION = 1
DEFAULT_STORE_PATH = Path(__file__).resolve().parent.parent / "data" / "metric_history.json"

_MONTH_NUMBERS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def store_path() -> Path:
    override = os.getenv("REPORTER_METRIC_HISTORY", "").strip()
    return Path(override) if override else DEFAULT_STORE_PATH


def month_key(month_label: str) -> tuple[int, int] | None:
    """('Jul 2026') -> (2026, 7). Returns None for anything unparseable."""
    match = re.search(r"([A-Za-z]{3,9})\s+(\d{4})", str(month_label or ""))
    if not match:
        return None
    month = _MONTH_NUMBERS.get(match.group(1)[:3].lower())
    return (int(match.group(2)), month) if month else None


def load_store(path: Path | None = None) -> dict[str, Any]:
    target = path or store_path()
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {"version": STORE_VERSION, "dealerships": {}}
    except (OSError, ValueError) as exc:
        logging.warning("Metric history unreadable at %s (%s); starting a new store.", target, exc)
        return {"version": STORE_VERSION, "dealerships": {}}
    if not isinstance(raw, dict) or not isinstance(raw.get("dealerships"), dict):
        return {"version": STORE_VERSION, "dealerships": {}}
    return raw


def save_store(store: dict[str, Any], path: Path | None = None) -> bool:
    """Write atomically. History is a convenience, never load-bearing: failures are logged."""
    target = path or store_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=str(target.parent), prefix=".metric_history-", suffix=".tmp", delete=False
        ) as handle:
            json.dump(store, handle, indent=1, sort_keys=True)
            temporary = Path(handle.name)
        temporary.replace(target)
        return True
    except OSError as exc:
        logging.warning("Could not write metric history to %s: %s", target, exc)
        return False


def record_month(
    store: dict[str, Any],
    dealership_name: str,
    month_label: str,
    metrics: dict[str, dict[str, Any]],
) -> None:
    """Record one dealership-month. Re-running a month overwrites that month only."""
    if not dealership_name or month_key(month_label) is None:
        return
    snapshot: dict[str, Any] = {}
    for metric_name, metric in (metrics or {}).items():
        value = metric.get("current_value", metric.get("current"))
        if not isinstance(value, (int, float)):
            continue
        snapshot[metric_name] = {
            "v": float(value),
            "t": str(metric.get("metric_type", "count")),
            "r": bool(metric.get("is_reliable", False)),
            "s": bool(metric.get("is_suspicious", False)),
        }
    if not snapshot:
        return
    store.setdefault("dealerships", {}).setdefault(dealership_name, {})[str(month_label)] = snapshot


def recorded_months(store: dict[str, Any], dealership_name: str) -> list[str]:
    months = list((store.get("dealerships", {}).get(dealership_name) or {}).keys())
    dated = [(month_key(label), label) for label in months]
    return [label for key, label in sorted((d for d in dated if d[0]), key=lambda item: item[0])]


def metric_series(
    store: dict[str, Any],
    dealership_name: str,
    metric_name: str,
    *,
    before_month: str | None = None,
    reliable_only: bool = True,
) -> list[tuple[str, float]]:
    """Oldest-first (month_label, value) for one metric, optionally strictly before a month."""
    months = store.get("dealerships", {}).get(dealership_name) or {}
    cutoff = month_key(before_month) if before_month else None
    series: list[tuple[tuple[int, int], str, float]] = []
    for label, snapshot in months.items():
        key = month_key(label)
        if key is None or (cutoff is not None and key >= cutoff):
            continue
        entry = (snapshot or {}).get(metric_name)
        if not isinstance(entry, dict) or not isinstance(entry.get("v"), (int, float)):
            continue
        if reliable_only and (not entry.get("r", False) or entry.get("s", False)):
            continue
        series.append((key, label, float(entry["v"])))
    series.sort(key=lambda item: item[0])
    return [(label, value) for _, label, value in series]


def value_at(
    store: dict[str, Any],
    dealership_name: str,
    metric_name: str,
    month_label: str,
    *,
    reliable_only: bool = True,
) -> float | None:
    entry = ((store.get("dealerships", {}).get(dealership_name) or {}).get(month_label) or {}).get(metric_name)
    if not isinstance(entry, dict) or not isinstance(entry.get("v"), (int, float)):
        return None
    if reliable_only and (not entry.get("r", False) or entry.get("s", False)):
        return None
    return float(entry["v"])


def same_month_last_year_label(month_label: str) -> str | None:
    key = month_key(month_label)
    if key is None:
        return None
    year, month = key
    for name, number in _MONTH_NUMBERS.items():
        if number == month:
            return f"{name.capitalize()} {year - 1}"
    return None


def dealership_names(store: dict[str, Any]) -> Iterable[str]:
    return (store.get("dealerships") or {}).keys()
