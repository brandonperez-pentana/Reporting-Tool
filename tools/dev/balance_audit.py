"""Does the finished output reflect the whole month, or only one side of it?

A month is very often mixed inside a single department -- leads down while campaign
returns climb. When every one of the four insights, the summary and the SMS point the
same way, the owner is given a false perspective of their own month.

This drives synthesised months through the full pipeline and reports the polarity mix
of what actually reaches the client, against the polarity mix of the metrics.

    python tools/dev/balance_audit.py
"""
import itertools
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
import dealership_insights_local as core  # noqa: E402
from tools.dev.summary_scenarios import SALES_CASES, SERVICE_CASES, sales, service  # noqa: E402

MATERIAL_BANDS = {"meaningful", "strong"}


def metric_split(metrics):
    """Usable metrics that moved materially, split by direction."""
    up, down = [], []
    for name, metric in metrics.items():
        if not core.metric_usable_for_reporting(metric):
            continue
        if str(metric.get("significance_band", "")) not in MATERIAL_BANDS:
            continue
        (down if core.metric_is_adverse(metrics, name) else up).append(name)
    return up, down


def line_mix(lines):
    counts = {"positive": 0, "negative": 0, "mixed": 0, "neutral": 0}
    for line in lines:
        counts[core._insight_line_polarity(line)] += 1
    return counts


def main() -> int:
    one_sided = 0
    for (sname, sargs), (vname, vargs) in itertools.product(SALES_CASES.items(), SERVICE_CASES.items()):
        metrics = core.parse_report_metrics(sales(*sargs), service(*vargs))
        up, down = metric_split(metrics)
        if not up or not down:
            continue  # genuinely one-directional month; nothing to balance
        out = core.evaluate_rule_based_insights(metrics, "Jul 2026", f"seed|{sname}|{vname}")
        bullets, summary, sms = out[0], out[1], out[2]
        mix = line_mix(bullets)
        covers_both = (mix["positive"] + mix["mixed"]) > 0 and (mix["negative"] + mix["mixed"]) > 0
        sms_low = sms.lower()
        sms_both = " but " in sms_low or " while " in sms_low
        flag = "" if covers_both else "  <-- ONE-SIDED"
        if not covers_both:
            one_sided += 1
        print(f"\n=== {sname} / {vname}{flag}")
        print(f"    metrics: {len(up)} material improvements, {len(down)} material declines")
        print(f"    insight polarity: +{mix['positive']} -{mix['negative']} ~{mix['mixed']} .{mix['neutral']}"
              f"   sms shows both sides: {sms_both}")
        for line in bullets:
            print(f"      [{core._insight_line_polarity(line):8}] {line}")
        print(f"    SUMMARY: {summary}")
        print(f"    SMS:     {sms.splitlines()[-1] if sms else ''}")
    print(f"\n{one_sided} mixed month(s) presented one-sidedly")
    return 1 if one_sided else 0


if __name__ == "__main__":
    raise SystemExit(main())
