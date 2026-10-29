"""Both engines on the same month, side by side, so the choice is made by reading.

The funnel engine is off by default and switching it on changes what every dealership
reads. That decision should not be taken from a description of the design; it should be
taken from the actual copy, on actual reports. This runs a month through both paths and
prints them next to each other.

    python tools/dev/engine_compare.py                     # every dealership with debug text
    python tools/dev/engine_compare.py "Demo Site 1"       # one of them
    python tools/dev/engine_compare.py --month "Jun 2026"  # a different month

Neither path is asked to write anything and the operator's settings file is not touched:
the engine is forced per run, in memory. Exits non-zero if the funnel engine could not
produce a month for a dealership the rule engine handled, since that is the one outcome
that would silently cost coverage after switching.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import dealership_insights_local as core  # noqa: E402
import settings_manager  # noqa: E402

DEBUG = ROOT / "Insights Output" / "Debug Text"
WIDTH = 96


def force_engine(engine: str):
    """Swap load_settings for the length of one run. The file on disk is never written."""
    original = settings_manager.load_settings

    def loader(*args, **kwargs):
        data = dict(original(*args, **kwargs))
        data["insight_engine"] = engine
        return data

    settings_manager.load_settings = loader
    return original


def copy_for(engine: str, name: str, month: str, sales: str, service: str):
    original = settings_manager.load_settings
    force_engine(engine)
    try:
        result = core.generate_insights_from_metrics(
            core.parse_report_metrics(sales, service), "success", name, month,
            sales_text=sales, service_text=service,
        )
    finally:
        settings_manager.load_settings = original
    bullets, summary, sms, segments = result[0], result[1], result[2], result[3]
    return bullets, summary, sms, ("funnel_engine" in result[7])


def wrap(text: str, indent: str = "    ") -> list[str]:
    words, lines, line = str(text).split(), [], indent
    for word in words:
        if len(line) + len(word) + 1 > WIDTH and line.strip():
            lines.append(line.rstrip())
            line = indent
        line += word + " "
    if line.strip():
        lines.append(line.rstrip())
    return lines


def show_engine(title: str, bullets, summary, sms) -> None:
    print(f"\n  {title}")
    print(f"  {'-' * (WIDTH - 2)}")
    for bullet in bullets:
        for i, line in enumerate(wrap(bullet, "      ")):
            print(("    -" + line[5:]) if i == 0 else line)
    print("\n    SUMMARY")
    for line in wrap(summary, "      "):
        print(line)
    print(f"\n    SMS ({len(sms)} of {core.SMS_CHAR_LIMIT} characters)")
    for line in sms.split("\n"):
        print(f"      {line}")


NOTHING_TO_COMPARE = object()


def available() -> dict[str, list[str]]:
    """Every month with debug text on disk, and the dealerships inside each.

    Debug text is written by a run, so this is the honest account of what can be
    compared -- and the thing to print when a name or a month turns up nothing.
    """
    found: dict[str, set[str]] = {}
    for path in DEBUG.glob("* - * - sales.txt"):
        stem = path.name[: -len(" - sales.txt")]
        name, separator, month = stem.rpartition(" - ")
        if separator and name:
            found.setdefault(month, set()).add(name)
    return {month: sorted(names) for month, names in found.items()}


def newest_month(months: list[str]) -> str:
    """The most recently written month, by the debug text's own timestamps.

    Better than a hard-coded default. This is run after a month, and the month just
    run is the one worth reading; a fixed default silently compares nothing the
    moment the calendar moves on.
    """
    def written(month: str) -> float:
        return max((p.stat().st_mtime for p in DEBUG.glob(f"* - {month} - sales.txt")), default=0.0)

    return max(months, key=written)


def compare(name: str, month: str):
    sales_file = DEBUG / f"{name} - {month} - sales.txt"
    service_file = DEBUG / f"{name} - {month} - service.txt"
    if not sales_file.is_file():
        print(f"\nno debug text for {name} {month}")
        return NOTHING_TO_COMPARE
    sales = sales_file.read_text(encoding="utf-8")
    service = service_file.read_text(encoding="utf-8") if service_file.is_file() else ""

    rules = copy_for("rules", name, month, sales, service)
    funnel = copy_for("funnel", name, month, sales, service)

    print(f"\n{'=' * WIDTH}\n{name}  ·  {month}\n{'=' * WIDTH}")
    show_engine("RULES (today)", *rules[:3])
    show_engine("FUNNEL (proposed)" + ("" if funnel[3] else "   [FELL BACK TO RULES]"), *funnel[:3])
    return funnel[3]


def main() -> int:
    args = list(sys.argv[1:])
    month = ""
    if "--month" in args:
        index = args.index("--month")
        if index + 1 >= len(args):
            print('--month needs a value, e.g. --month "Jun 2026"')
            return 2
        month = args[index + 1]
        del args[index:index + 2]

    # Everything below refuses to claim a comparison it did not make. The first
    # version defaulted to a hard-coded month, found nothing, and still reported
    # "both engines produced a month" -- which is the one thing a verification tool
    # must never do.
    catalogue = available()
    if not catalogue:
        print(f"No debug text in {DEBUG}")
        print("Run a month in Reporter first; this reads what that run wrote.")
        return 1

    if not month:
        month = newest_month(list(catalogue))
        print(f'month: {month}   (latest on disk; override with --month "Jun 2026")')

    known = catalogue.get(month)
    if known is None:
        print(f"No debug text for {month}. Months on disk: {', '.join(sorted(catalogue))}")
        return 1

    if args:
        wanted = " ".join(args)
        if wanted not in known:
            print(f'No debug text for "{wanted}" in {month}.')
            print(f"Dealerships in {month}: {', '.join(known)}")
            return 1
        names = [wanted]
    else:
        names = known

    outcomes = {name: compare(name, month) for name in names}
    compared = [n for n, o in outcomes.items() if o is not NOTHING_TO_COMPARE]
    fell_back = [n for n, o in outcomes.items() if o is False]

    print(f"\n{'=' * WIDTH}")
    if not compared:
        print("nothing was compared")
        return 1
    if fell_back:
        print(f"{len(fell_back)} of {len(compared)} fell back to the rule engine: {', '.join(fell_back)}")
        return 1
    print(f"{len(compared)} dealership(s), both engines produced a month")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
