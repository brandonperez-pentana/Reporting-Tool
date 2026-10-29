"""Print the month model for a demo dealership, so it can be read rather than trusted.

The model is the layer everything above it will depend on, and it produces no sentences
of its own -- which makes it invisible unless something prints it. This is that
something.

    python tools/dev/month_model_view.py "Demo Site 1"
    python tools/dev/month_model_view.py            # every dealership in the debug set
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import dealership_insights_local as core  # noqa: E402
from reporting import month_model as mm  # noqa: E402
from reporting import report_breakdowns as rb  # noqa: E402

DEBUG = ROOT / "Insights Output" / "Debug Text"


def movement(m) -> str:
    if m is None:
        return "not reported"
    if not m.usable:
        return f"{m.previous:,.0f} -> {m.current:,.0f}  (held back: {m.band})"
    change = f"{m.percent_change:+.1f}%" if m.percent_change is not None else "from zero"
    return f"{m.previous:,.0f} -> {m.current:,.0f}  {change}  {'firmer' if m.improving else 'softer'}  [{m.band}]"


def show(name: str) -> None:
    sales_file = DEBUG / f"{name} - Jul 2026 - sales.txt"
    service_file = DEBUG / f"{name} - Jul 2026 - service.txt"
    if not sales_file.is_file():
        print(f"no debug text for {name}")
        return
    sales_text = sales_file.read_text(encoding="utf-8")
    service_text = service_file.read_text(encoding="utf-8") if service_file.is_file() else ""
    metrics = core.parse_report_metrics(sales_text, service_text)
    model = mm.build_month_model(
        dealership=name, month_label="Jul 2026",
        metrics=metrics, breakdowns=rb.parse_breakdowns(sales_text, service_text),
    )

    counts = model.breadth()
    print(f"\n{'=' * 78}")
    print(f"{name}  ·  {model.month_label}  ·  verdict: {model.verdict().upper()}"
          f"   ({counts['firmed']} firmer, {counts['softened']} softer)")
    print("=" * 78)

    for funnel in (model.sales, model.service):
        print(f"\n{funnel.name.upper()} FUNNEL")
        for station in funnel.stations:
            print(f"  {station.label:<26}{movement(station.movement)}")
            if station.attribution:
                a = station.attribution
                share = f" — {a.share_of_total}% of the stated total" if a.share_of_total is not None else ""
                tail = f", {a.unnamed} not named" if a.unnamed else ""
                print(f"  {'':<26}driven by {a.driver} ({a.value:+,.0f}{share}{tail})")
            if station.note:
                print(f"  {'':<26}note: {station.note}")
        leak = funnel.largest_leak()
        print(f"  {'first leak':<26}{leak.label if leak else 'none'}")

    print("\nHYGIENE")
    for station in model.hygiene:
        print(f"  {station.label:<26}{movement(station.movement)}")

    print("\nBOTTOM LINE")
    for key, m in model.bottom_line().items():
        print(f"  {key:<26}{movement(m)}")

    print("\nDIAGNOSTIC PAIR (does the floor work the leads it gets?)")
    for station in model.diagnostic_pair():
        print(f"  {station.label:<26}{movement(station.movement)}")

    if model.lead_sources:
        print("\nLEAD SOURCES")
        for row in model.lead_sources:
            print(f"  {row.source:<26}leads {row.leads_previous:>5} -> {row.leads_current:<5}"
                  f"  deals {row.deals_previous:>5} -> {row.deals_current:<5}"
                  f"  conv {row.conversion_previous:>4.0f}% -> {row.conversion_current:.0f}%")

    for label, block in (("SALES", model.sales_campaigns), ("SERVICE", model.service_campaigns)):
        if block is None:
            continue
        print(f"\n{label} CAMPAIGNS", end="")
        if not block.runs_campaigns:
            print("   none run this month")
            continue
        covered = "covers the stated total" if block.accounts_for_whole_total else f"{block.unnamed_count} not named"
        print(f"   ({covered})")
        for campaign in block.named:
            flag = "  LOSING MONEY" if campaign.roi < 0 else ""
            print(f"  {campaign.name:<34}{campaign.count:>4}   ${campaign.roi:>10,}{flag}")

    if model.missing():
        print(f"\nnot reported this month: {', '.join(model.missing())}")


def main() -> int:
    if len(sys.argv) > 1:
        show(" ".join(sys.argv[1:]))
        return 0
    names = sorted({p.name.split(" - ")[0] for p in DEBUG.glob("Demo Site * - sales.txt")})
    for name in names:
        show(name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
