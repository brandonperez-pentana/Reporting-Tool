#!/usr/bin/env python3
"""Execute a monthly Reporter run for an arbitrary month.

This replaces the old per-month scripts (run_may_2026.py, run_june_2026.py,
run_jul_2026.py, rerun_apr_2026.py), which were hand-copied and had quietly
drifted from each other: later copies gained integrity checks (SMS length,
semantics, sales/service PDF counts) that earlier copies never had. Because
every month now runs through this one script, every month gets identical
checks by construction.

Usage:
    python run_month.py --month "Aug 2026"
    python run_month.py --month "Aug 2026" --output-dir "Insights Output STAGING"

The old run_<month>_2026.py filenames still work as thin wrappers that call
main() with their month hard-coded, for anyone/anything that still invokes
them directly.
"""

from __future__ import annotations

import argparse
import json
import sys
import traceback
from calendar import month_abbr
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import report_engine
import dealership_insights_local as core
from settings_manager import load_settings


def yn(value: bool) -> str:
    return "YES" if value else "NO"


def _previous_month_label(month_label: str) -> str:
    """Best-effort 'prior calendar month' label for the comparison_month field."""
    month_num, year_num, _ = core.parse_month_label(month_label)
    prev_month = month_num - 1 or 12
    prev_year = year_num - 1 if month_num == 1 else year_num
    return f"{month_abbr[prev_month]} {prev_year}"


def _month_slug(month_label: str) -> str:
    """'Jul 2026' -> 'jul_2026', used for log filenames."""
    return month_label.strip().lower().replace(" ", "_")


def main(month: str | None = None, output_dir: str | None = None, all_dealerships: bool = False) -> int:
    settings = load_settings()
    accounts = settings["accounts_folder"]
    # An empty setting used to resolve to Path("") -> the current directory, so
    # every subfolder of the app itself (logs, tests, templates) was treated as
    # a dealership and the run "succeeded" with placeholder reports. Fail loudly.
    if not accounts or not Path(accounts).expanduser().is_dir():
        print(
            f"Accounts folder not usable: {accounts or '(not set)'}\n"
            "Set it in the app under Settings, or set REPORTER_ACCOUNTS_DIR."
        )
        return 1
    output = str((Path(output_dir) if output_dir else Path(__file__).parent / "Insights Output").resolve())
    resolved_month = month or settings.get("month") or datetime.now().strftime("%b %Y")
    month_num, year_num, month_label = core.parse_month_label(resolved_month)
    comparison_month = _previous_month_label(month_label)
    template = settings.get("template_path")
    contacts = settings.get("recipient_mapping_path")
    portfolio = [] if all_dealerships else [str(n) for n in settings.get("portfolio") or [] if str(n).strip()]
    pdf_folder = str(settings.get("pdf_source_folder") or "").strip()
    if pdf_folder and not Path(pdf_folder).expanduser().is_dir():
        print(f"PDF source folder not usable: {pdf_folder}")
        return 1

    print(f"Starting {month_label} reporting run...")
    print(f"Accounts: {accounts}")
    print(f"PDF source: {pdf_folder or accounts}")
    print(f"Output: {output}")
    print(f"Dealerships: {f'{len(portfolio)} in My dealerships' if portfolio else 'all'}")

    try:
        session = report_engine.run_reporting(
            accounts, output, month_label, template, contacts, portfolio=portfolio or None,
            pdf_folder=pdf_folder or None,
        )
    except Exception as exc:
        print("PIPELINE ERROR:", exc)
        traceback.print_exc()
        return 1

    print(f"Pipeline complete: {len(session.dealerships)} dealerships")

    email_errors: list[dict[str, str]] = []
    missing_placeholders: list[dict[str, object]] = []
    render_issues: list[dict[str, object]] = []
    anomaly_accounts: list[str] = []
    blocked_accounts: list[dict[str, object]] = []
    skipped: list[dict[str, object]] = []
    reports_generated = 0
    sales_reports = 0
    service_reports = 0
    email_drafts_generated = 0

    for name in sorted(session.dealerships.keys(), key=str.lower):
        state = session.dealerships[name]
        status = str(state.result.report_status).lower()
        tags = state.result.review_reason_tags or []
        is_unreviewable = state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder"

        if state.result.sales_pdf is not None:
            sales_reports += 1
        if state.result.service_pdf is not None:
            service_reports += 1

        if status == "missing file" or state.result.sales_pdf is None or state.result.service_pdf is None:
            skipped.append(
                {
                    "dealership": name,
                    "status": state.result.report_status,
                    "notes": list(state.result.notes or []),
                }
            )
        else:
            reports_generated += 1

        if "anomaly_detected" in tags:
            anomaly_accounts.append(name)

        try:
            preview = report_engine.build_email_preview(name)
        except Exception as exc:
            email_errors.append({"dealership": name, "error": str(exc)})
            preview = {}

        # A "successful" preview built from placeholder/fallback data is not a
        # real email draft: don't count it, and surface it separately so an
        # operator can see exactly which accounts need their source PDFs
        # before anything can be sent for them.
        if preview.get("blocked") or is_unreviewable:
            blocked_accounts.append(
                {
                    "dealership": name,
                    "reason": preview.get("block_reason", "no_source_data"),
                }
            )
        elif name not in [e["dealership"] for e in email_errors]:
            email_drafts_generated += 1

        details = report_engine.get_dealership_report_text(name)
        missing = details.get("missing_placeholders") or []
        if missing:
            missing_placeholders.append({"dealership": name, "missing": list(missing)})
        warnings = (state.email_preview or {}).get("render_audit_warnings") or []
        if warnings:
            render_issues.append({"dealership": name, "warnings": list(warnings)})

    rows = report_engine.get_dealership_results()
    review_required = [r["dealership"] for r in rows if str(r.get("review_required", "")).lower() == "yes"]
    review_blocked = [r["dealership"] for r in rows if str(r.get("review_severity", "")).lower() == "blocked"]
    review_flagged = [r["dealership"] for r in rows if str(r.get("review_severity", "")).lower() == "flagged"]
    manual_review = bool(review_required or anomaly_accounts or missing_placeholders or email_errors)

    integrity: list[dict[str, str]] = []
    for name, state in session.dealerships.items():
        sms = state.result.sms_version or ""
        if state.result.report_status == "success":
            if not sms:
                integrity.append({"dealership": name, "issue": "SMS missing"})
            elif len(sms) > core.SMS_CHAR_LIMIT:
                integrity.append({"dealership": name, "issue": "SMS over 160 characters", "sms": sms[:120]})
            for issue in core.validate_sms_semantics(sms, state.result.bullet_insights or []):
                integrity.append({"dealership": name, "issue": issue, "sms": sms[:120]})
        if state.result.report_status == "success" and not state.result.bullet_insights:
            integrity.append({"dealership": name, "issue": "No priority insights on success"})

    summary = {
        "reporting_month": month_label,
        "reporting_period": f"1 {month_label.split()[0]} {year_num} report period",
        "comparison_month": comparison_month,
        "accounts_processed": len(session.dealerships),
        "reports_generated": reports_generated,
        "sales_reports_generated": sales_reports,
        "service_reports_generated": service_reports,
        "email_drafts_generated": email_drafts_generated,
        "accounts_skipped": len(skipped),
        "skipped_details": skipped,
        "anomaly_detected_reviewed": len(anomaly_accounts),
        "anomaly_accounts": anomaly_accounts,
        "manual_review_required": manual_review,
        "review_required_count": len(review_required),
        "review_required": review_required,
        "review_blocked_count": len(review_blocked),
        "review_blocked": review_blocked,
        "review_flagged_count": len(review_flagged),
        "review_flagged": review_flagged,
        "blocked_accounts_no_source_data": blocked_accounts,
        "unresolved_placeholders": bool(missing_placeholders),
        "missing_placeholder_details": missing_placeholders,
        "email_render_issues": bool(render_issues or email_errors),
        "email_errors": email_errors,
        "render_issues": render_issues,
        "data_integrity_issues": bool(integrity),
        "integrity_details": integrity,
        "output_folder": output,
    }

    log_path = Path(__file__).parent / "logs" / f"{_month_slug(month_label)}_run_summary.json"
    log_path.parent.mkdir(exist_ok=True)
    log_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"{month_label} reporting run complete: YES")
    print(f"Reporting month: {month_label}")
    print(f"Comparison month: {comparison_month}")
    print(f"Sales reports found: {sales_reports}")
    print(f"Service reports found: {service_reports}")
    print(f"Email drafts generated: {email_drafts_generated}")
    print(f"Accounts needing source data before anything can be sent: {len(blocked_accounts)}")
    if blocked_accounts:
        for entry in blocked_accounts:
            print(f"  - {entry['dealership']}")
    print(f"Review required: {len(review_required)} (blocked={len(review_blocked)}, flagged={len(review_flagged)})")
    print(f"Manual review required: {yn(manual_review)}")
    print(f"Data integrity issues: {yn(bool(integrity))}")
    print(f"Summary written to: {log_path}")

    return 0


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--month",
        required=False,
        help='Reporting month, e.g. "Aug 2026" (defaults to user_settings.json\'s saved month).',
    )
    parser.add_argument(
        "--output-dir",
        required=False,
        help="Override the output folder (defaults to 'Insights Output'). Use a staging folder for test runs.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        dest="all_dealerships",
        help="Ignore My dealerships and run every workbook in the accounts folder.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    sys.exit(main(month=args.month, output_dir=args.output_dir, all_dealerships=args.all_dealerships))
