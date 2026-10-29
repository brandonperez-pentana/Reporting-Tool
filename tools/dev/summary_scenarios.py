"""Sweep synthetic months and flag any summary opener that disagrees with the metrics.

Kept in the repo because the bug it found -- an opening line claiming the opposite of
the month's own numbers -- is invisible on any single dealership and only shows up when
sales and service move in different directions. Run it after touching lane selection,
sentiment, or anything that decides how a month is characterised.

    python tools/dev/summary_scenarios.py
"""
import itertools
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
import dealership_insights_local as core  # noqa: E402

DEBUG = ROOT / "Insights Output" / "Debug Text"
BASE_SALES = (DEBUG / "Demo Site 1 - Jul 2026 - sales.txt").read_text(encoding="utf-8")
BASE_SERVICE = (DEBUG / "Demo Site 1 - Jul 2026 - service.txt").read_text(encoding="utf-8")


def sales(leads, deals, conv, appts, drives, a2s, overdue, email, same_day, response):
    t = BASE_SALES
    t = t.replace("TOTAL 331 91 176 114 53% 125%",
                  f"TOTAL {leads[0]} {leads[1]} {deals[0]} {deals[1]} {conv[0]}% {conv[1]}%")
    t = t.replace("Jun-26 0 Jun-26 50:00 min Jun-26 4% Jun-26 138%",
                  f"Jun-26 0 Jun-26 50:00 min Jun-26 {appts[0]}% Jun-26 {drives[0]}%")
    t = t.replace("Jul-26 0 Jul-26 16:35 hrs Jul-26 9% Jul-26 47%",
                  f"Jul-26 0 Jul-26 {response} Jul-26 {appts[1]}% Jul-26 {drives[1]}%")
    t = t.replace("Jun-26 8% Jun-26 136 Jun-26 75% Jun-26 68%",
                  f"Jun-26 {a2s[0]}% Jun-26 {overdue[0]} Jun-26 {email[0]}% Jun-26 {same_day[0]}%")
    t = t.replace("Jul-26 7% Jul-26 146 Jul-26 69% Jul-26 18%",
                  f"Jul-26 {a2s[1]}% Jul-26 {overdue[1]} Jul-26 {email[1]}% Jul-26 {same_day[1]}%")
    return t


def service(rem, enrolled, tasks, bookings, roi):
    t = BASE_SERVICE
    t = t.replace("Jun-26 70 Jun-26 70 Jun-26 0", f"Jun-26 {rem[0]} Jun-26 {enrolled[0]} Jun-26 {tasks[0]}")
    t = t.replace("Jul-26 300 Jul-26 300 Jul-26 0", f"Jul-26 {rem[1]} Jul-26 {enrolled[1]} Jul-26 {tasks[1]}")
    t = t.replace("Jun-26 7 $10,059", f"Jun-26 {bookings[0]} ${roi[0]:,}")
    t = t.replace("Jul-26 65 $44,183", f"Jul-26 {bookings[1]} ${roi[1]:,}")
    return t


SALES_CASES = {
    "sales up":            ((300, 345), (150, 190), (50, 62), (4, 6), (100, 122), (8, 10), (140, 120), (70, 79), (60, 70), "40:00 min"),
    "sales up, tasks up":  ((300, 345), (150, 190), (50, 62), (4, 6), (100, 122), (8, 10), (140, 168), (70, 66), (60, 70), "58:00 min"),
    "deals+conv up only":  ((300, 298), (150, 192), (50, 63), (4, 4), (100, 99), (8, 8), (140, 141), (70, 70), (60, 60), "50:00 min"),
    "sales flat":          ((300, 301), (150, 151), (50, 50), (4, 4), (100, 100), (8, 8), (140, 140), (70, 70), (60, 60), "50:00 min"),
    "sales down":          ((331, 210), (176, 120), (53, 41), (9, 5), (138, 95), (8, 6), (120, 165), (80, 66), (70, 52), "72:00 min"),
}
SERVICE_CASES = {
    "svc up":   ((70, 300), (70, 300), (10, 6), (7, 65), (10059, 44183)),
    "svc flat": ((100, 101), (100, 101), (10, 10), (20, 20), (12000, 12100)),
    "svc down": ((300, 90), (300, 90), (6, 22), (65, 8), (44183, 9000)),
}

NEGATIVE_WORDS = ("step down", "softer", "softening", "softness", "pressure", "weaker", "lighter", "heavier", "workload")
POSITIVE_WORDS = ("improvement", "right way", "stronger", "firmer", "faster", "quicker")


def balance(metrics):
    up = down = 0
    for name, metric in metrics.items():
        if core.metric_usable_for_reporting(metric):
            if core.metric_is_adverse(metrics, name):
                down += 1
            else:
                up += 1
    return up, down


def main() -> int:
    rows, mismatches = [], []
    for (sname, sargs), (vname, vargs) in itertools.product(SALES_CASES.items(), SERVICE_CASES.items()):
        metrics = core.parse_report_metrics(sales(*sargs), service(*vargs))
        up, down = balance(metrics)
        out = core.evaluate_rule_based_insights(metrics, "Jul 2026", f"seed|{sname}|{vname}")
        summary, sentiment = out[1], out[6]
        opener = summary.split(", with")[0] if ", with" in summary else summary[:80]
        low = opener.lower()
        tone = "neg" if any(w in low for w in NEGATIVE_WORDS) else ("pos" if any(w in low for w in POSITIVE_WORDS) else "neutral")
        truth = "pos" if up > down * 1.5 else ("neg" if down > up * 1.5 else "neutral")
        bad = (tone == "neg" and truth == "pos") or (tone == "pos" and truth == "neg")
        rows.append((f"{sname} / {vname}", f"{up}/{down}", truth, tone, sentiment, opener, bad))
        if bad:
            mismatches.append((f"{sname} / {vname}", up, down, sentiment, summary))

    print(f"{'scenario':<28} {'up/dn':>6} {'truth':>7} {'opener':>7}  {'sentiment':<22} text")
    print("-" * 140)
    for label, ud, truth, tone, sent, opener, bad in rows:
        print(f"{'*' if bad else ' '}{label:<27} {ud:>6} {truth:>7} {tone:>7}  {sent:<22} {opener}")
    print(f"\n{len(mismatches)} contradiction(s)")
    for label, up, down, sent, summary in mismatches:
        print(f"\n>>> {label}: {up} up / {down} down, sentiment={sent}\n    {summary}")
    return 1 if mismatches else 0


if __name__ == "__main__":
    raise SystemExit(main())
