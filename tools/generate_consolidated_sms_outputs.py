#!/usr/bin/env python3
"""Consolidate final Priority Insights SMS outputs across dealerships and months."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document

SMS_CHAR_LIMIT = 160
MONTH_LABEL_RE = re.compile(r"^[A-Za-z]{3} \d{4}$")
FILE_SUFFIX_RE = re.compile(r"^(.+) - ([A-Za-z]{3} \d{4})\.(txt|docx)$", re.IGNORECASE)
EXCLUDED_SUBDIR_NAMES = {"debug text", "archive", "archived"}
DEFAULT_OUTPUT_BASENAME = "Consolidated_SMS_Outputs"
DEFAULT_OUTPUT_MD = ROOT / f"{DEFAULT_OUTPUT_BASENAME}.md"
DEFAULT_OUTPUT_DOCX = ROOT / f"{DEFAULT_OUTPUT_BASENAME}.docx"
DEFAULT_OUTPUT_TXT = ROOT / f"{DEFAULT_OUTPUT_BASENAME}.txt"
PRODUCTION_OUTPUT_DIR = ROOT / "Insights Output"
TXT_SEPARATOR = "-" * 60


@dataclass
class SmsCandidate:
    dealership: str
    month_label: str
    source_path: Path
    sms_text: str
    modified_at: float
    folder_priority: int
    parse_error: str = ""


@dataclass
class SmsRow:
    dealership: str
    month_label: str
    sms_text: str
    char_count: int
    source_file: str
    over_limit: bool
    validation_note: str = ""


@dataclass
class ValidationReport:
    consolidation_complete: bool = False
    months_found: list[str] = field(default_factory=list)
    rows_written: int = 0
    missing_sms_outputs: list[str] = field(default_factory=list)
    sms_over_160: list[str] = field(default_factory=list)
    duplicate_outputs: list[str] = field(default_factory=list)
    source_files_modified: list[str] = field(default_factory=list)
    production_reports_modified: bool = False
    output_md_path: str = ""
    output_docx_path: str = ""
    output_txt_path: str = ""
    monthly_files_generated: list[str] = field(default_factory=list)
    sms_text_preservation_failures: list[str] = field(default_factory=list)
    character_count_failures: list[str] = field(default_factory=list)
    spacing_separator_validation: str = "PASS"
    parse_failures: list[str] = field(default_factory=list)
    excluded_unknown_accounts: list[str] = field(default_factory=list)

    @property
    def final_verdict(self) -> str:
        if not self.consolidation_complete:
            return "FAIL"
        if (
            self.missing_sms_outputs
            or self.sms_over_160
            or self.parse_failures
            or self.production_reports_modified
            or self.duplicate_outputs
            or self.sms_text_preservation_failures
            or self.character_count_failures
            or self.spacing_separator_validation == "FAIL"
        ):
            return "PASS_WITH_WARNINGS"
        return "PASS"


def discover_output_roots(root: Path = ROOT) -> list[tuple[int, Path]]:
    roots: list[tuple[int, Path]] = []
    for path in sorted(root.iterdir()):
        if not path.is_dir():
            continue
        name = path.name
        if name == PRODUCTION_OUTPUT_DIR.name:
            roots.append((0, path))
        elif name.startswith("Insights Output"):
            if "TEST" in name.upper():
                continue
            roots.append((1, path))
    if PRODUCTION_OUTPUT_DIR.is_dir() and not any(p == PRODUCTION_OUTPUT_DIR for _, p in roots):
        roots.insert(0, (0, PRODUCTION_OUTPUT_DIR))
    return sorted(roots, key=lambda item: (item[0], item[1].name.lower()))


def load_known_dealership_names(accounts_dir: Path | None) -> set[str] | None:
    """Whitelist of real dealership names, derived from the Accounts folder.

    Excluding folders whose *directory* name merely doesn't contain "TEST"
    (the old rule in discover_output_roots) is not enough: test/demo report
    files placed directly inside the production "Insights Output" folder
    slip straight through that check, because it only filters at the
    directory level. Matching every candidate's inferred dealership name
    against the accounts Reporter actually knows about (top-level Accounts
    subfolders, plus anything archived) closes that gap regardless of where
    a stray file ends up.

    Returns None (meaning "don't filter") if accounts_dir is missing or
    empty, so this stays safe to call in environments/tests without a real
    Accounts folder available.
    """
    if not accounts_dir:
        return None
    base = Path(accounts_dir)
    if not base.is_dir():
        return None

    names: set[str] = set()
    for child in base.iterdir():
        if not child.is_dir():
            continue
        if child.name.strip().lower() in {"archieve", "archive", "archived"}:
            for grandchild in child.iterdir():
                if grandchild.is_dir():
                    names.add(grandchild.name.strip().lower())
            continue
        names.add(child.name.strip().lower())
    return names or None


def _dealership_is_known(dealership: str, known_dealerships: set[str] | None) -> bool:
    if known_dealerships is None:
        return True
    normalised = dealership.strip().lower()
    if normalised in known_dealerships:
        return True
    # Grouped/brand-suffixed report names (e.g. "Demo Site 1 Ford")
    # should still match their parent account folder ("Demo Site 1").
    return any(
        normalised.startswith(known) or known.startswith(normalised)
        for known in known_dealerships
    )


def infer_month_from_text(text: str) -> str:
    match = re.search(r"^Month:\s*(.+)$", text, re.MULTILINE)
    if not match:
        return ""
    label = match.group(1).strip()
    return label if MONTH_LABEL_RE.match(label) else ""


def infer_month(path: Path, text: str | None = None) -> str:
    match = FILE_SUFFIX_RE.match(path.name)
    if match:
        return match.group(2)
    if text:
        return infer_month_from_text(text)
    return ""


def infer_dealership_from_text(text: str) -> str:
    match = re.search(r"^Dealership:\s*(.+)$", text, re.MULTILINE)
    return match.group(1).strip() if match else ""


def infer_dealership(path: Path, text: str | None = None) -> str:
    match = FILE_SUFFIX_RE.match(path.name)
    if match:
        return match.group(1).strip()
    if text:
        return infer_dealership_from_text(text)
    return path.stem


def extract_sms_from_text(text: str) -> tuple[str, str]:
    if "SMS Version:" not in text:
        return "", "SMS Version section not found"
    block = text.split("SMS Version:", 1)[1]
    for stop in ("\n\nSource Files:", "\n\nExtracted Metrics:", "\n\nRules Triggered:"):
        if stop in block:
            block = block.split(stop, 1)[0]
            break
    sms = block.strip()
    if not sms:
        return "", "SMS Version section is empty"
    return sms, ""


def extract_sms_from_docx(path: Path) -> tuple[str, str]:
    try:
        document = Document(path)
    except Exception as exc:
        return "", f"Unable to read DOCX: {exc}"

    capture = False
    lines: list[str] = []
    for paragraph in document.paragraphs:
        style_name = (paragraph.style.name if paragraph.style else "") or ""
        text = paragraph.text.strip()
        if style_name.startswith("Heading") and text.lower() == "sms version":
            capture = True
            continue
        if capture:
            if style_name.startswith("Heading") and text:
                break
            if text:
                lines.append(text)
    sms = "\n".join(lines).strip()
    if not sms:
        return "", "SMS Version section not found in DOCX"
    return sms, ""


def extract_sms_text(path: Path) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix == ".txt":
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            return "", f"Unable to read TXT: {exc}"
        return extract_sms_from_text(text)
    if suffix == ".docx":
        return extract_sms_from_docx(path)
    return "", f"Unsupported file type: {suffix}"


def discover_sms_outputs(
    root: Path = ROOT,
    known_dealerships: set[str] | None = None,
    report: ValidationReport | None = None,
) -> list[SmsCandidate]:
    candidates: list[SmsCandidate] = []
    for folder_priority, output_root in discover_output_roots(root):
        for path in sorted(output_root.iterdir()):
            if not path.is_file():
                continue
            if path.suffix.lower() not in {".txt", ".docx"}:
                continue
            if not FILE_SUFFIX_RE.match(path.name):
                continue

            text: str | None = None
            if path.suffix.lower() == ".txt":
                try:
                    text = path.read_text(encoding="utf-8")
                except OSError:
                    text = None

            month_label = infer_month(path, text)
            dealership = infer_dealership(path, text)
            sms_text, parse_error = extract_sms_text(path)
            if not month_label:
                parse_error = parse_error or "Unable to infer month"
            if not dealership:
                parse_error = parse_error or "Unable to infer dealership"

            if dealership and not _dealership_is_known(dealership, known_dealerships):
                if report is not None:
                    report.excluded_unknown_accounts.append(f"{dealership} / {month_label}: {path}")
                continue

            candidates.append(
                SmsCandidate(
                    dealership=dealership,
                    month_label=month_label,
                    source_path=path.resolve(),
                    sms_text=sms_text,
                    modified_at=path.stat().st_mtime,
                    folder_priority=folder_priority,
                    parse_error=parse_error,
                )
            )
    return candidates


def choose_best_candidate(group: list[SmsCandidate]) -> tuple[SmsCandidate, list[Path]]:
    def sort_key(item: SmsCandidate) -> tuple[int, int, float, str]:
        prefer_txt = 0 if item.source_path.suffix.lower() == ".txt" else 1
        return (item.folder_priority, prefer_txt, -item.modified_at, item.source_path.name.lower())

    ordered = sorted(group, key=sort_key)
    best = ordered[0]
    duplicates: list[Path] = []
    for item in ordered[1:]:
        same_folder = item.source_path.parent == best.source_path.parent
        alternate_format = (
            same_folder
            and item.source_path.stem == best.source_path.stem
            and item.source_path.suffix.lower() != best.source_path.suffix.lower()
        )
        if alternate_format:
            continue
        duplicates.append(item.source_path)
    return best, duplicates


def month_sort_key(month_label: str) -> datetime:
    try:
        return datetime.strptime(month_label, "%b %Y")
    except ValueError:
        return datetime.min


def month_heading(month_label: str) -> str:
    try:
        parsed = datetime.strptime(month_label, "%b %Y")
        return parsed.strftime("%B %Y")
    except ValueError:
        return month_label


def safe_month_filename(month_label: str) -> str:
    return month_heading(month_label).replace(" ", "_")


def group_rows_by_month(rows: list[SmsRow]) -> dict[str, list[SmsRow]]:
    grouped: dict[str, list[SmsRow]] = {}
    for row in rows:
        grouped.setdefault(row.month_label, []).append(row)
    for month_label in grouped:
        grouped[month_label].sort(key=lambda row: row.dealership.lower())
    return grouped


def validate_sms_rows(rows: Iterable[SmsRow], report: ValidationReport) -> None:
    report.sms_over_160 = [
        f"{row.dealership} / {row.month_label} ({row.char_count} chars)"
        for row in rows
        if row.over_limit
    ]


def build_consolidated_rows(candidates: list[SmsCandidate], report: ValidationReport) -> list[SmsRow]:
    grouped: dict[tuple[str, str], list[SmsCandidate]] = {}
    for candidate in candidates:
        if not candidate.month_label or not candidate.dealership:
            report.parse_failures.append(f"{candidate.source_path}: metadata inference failed")
            continue
        grouped.setdefault((candidate.dealership, candidate.month_label), []).append(candidate)

    rows: list[SmsRow] = []
    for (dealership, month_label), group in sorted(grouped.items(), key=lambda item: item[0]):
        best, duplicate_paths = choose_best_candidate(group)
        validation_note = ""
        if duplicate_paths:
            dup_list = ", ".join(str(path) for path in duplicate_paths)
            validation_note = f"Duplicate outputs ignored (newer/preferred source kept): {dup_list}"
            report.duplicate_outputs.append(f"{dealership} / {month_label}: {dup_list}")

        if best.parse_error and not best.sms_text:
            report.missing_sms_outputs.append(f"{dealership} / {month_label}: {best.parse_error}")
            report.parse_failures.append(f"{best.source_path}: {best.parse_error}")
            continue

        if not best.sms_text.strip():
            report.missing_sms_outputs.append(f"{dealership} / {month_label}: empty SMS text")
            continue

        char_count = len(best.sms_text)
        rows.append(
            SmsRow(
                dealership=dealership,
                month_label=month_label,
                sms_text=best.sms_text,
                char_count=char_count,
                source_file=str(best.source_path),
                over_limit=char_count > SMS_CHAR_LIMIT,
                validation_note=validation_note,
            )
        )

    rows.sort(key=lambda row: (month_sort_key(row.month_label), row.dealership.lower()))
    report.months_found = sorted({row.month_label for row in rows}, key=month_sort_key)
    report.rows_written = len(rows)
    validate_sms_rows(rows, report)
    return rows


def format_sms_block(row: SmsRow, *, block_format: str = "md") -> list[str]:
    if block_format == "txt":
        lines = [
            row.dealership,
            "",
            row.sms_text,
            "",
            f"Character Count: {row.char_count}",
            f"Source File: {row.source_file}",
            "",
            TXT_SEPARATOR,
            "",
        ]
        return lines

    dealership_heading = f"### {row.dealership}"
    lines = [
        dealership_heading,
        "",
        row.sms_text,
        "",
        f"Character Count: {row.char_count}",
        f"Source File: {row.source_file}",
        "",
        "---",
        "",
    ]
    return lines


def write_markdown_output(
    rows_by_month: dict[str, list[SmsRow]],
    output_path: Path,
    *,
    title: str = "Consolidated SMS Outputs",
    include_validation: bool = False,
    report: ValidationReport | None = None,
) -> None:
    lines: list[str] = [f"# {title}", ""]
    for month_label in sorted(rows_by_month, key=month_sort_key):
        lines.append(f"## {month_heading(month_label)}")
        lines.append("")
        for row in rows_by_month[month_label]:
            lines.extend(format_sms_block(row, block_format="md"))

    if include_validation and report is not None:
        lines.extend(_validation_markdown_section(report))
        if report.duplicate_outputs:
            lines.append("### Duplicate Output Notes")
            lines.append("")
            for note in report.duplicate_outputs:
                lines.append(f"- {note}")
            lines.append("")

    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def write_text_output(
    rows_by_month: dict[str, list[SmsRow]],
    output_path: Path,
    *,
    title: str = "Consolidated SMS Outputs",
) -> None:
    lines: list[str] = [title, "=" * len(title), ""]
    for month_label in sorted(rows_by_month, key=month_sort_key):
        month_title = month_heading(month_label)
        lines.append(month_title)
        lines.append("-" * len(month_title))
        lines.append("")
        for row in rows_by_month[month_label]:
            lines.extend(format_sms_block(row, block_format="txt"))
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def write_docx_output(
    rows_by_month: dict[str, list[SmsRow]],
    output_path: Path,
    *,
    title: str = "Consolidated SMS Outputs",
    include_validation: bool = False,
    report: ValidationReport | None = None,
) -> None:
    document = Document()
    document.add_heading(title, level=0)

    for month_label in sorted(rows_by_month, key=month_sort_key):
        document.add_heading(month_heading(month_label), level=1)
        for row in rows_by_month[month_label]:
            document.add_heading(row.dealership, level=2)
            for sms_line in row.sms_text.splitlines():
                document.add_paragraph(sms_line)
            document.add_paragraph(f"Character Count: {row.char_count}")
            document.add_paragraph(f"Source File: {row.source_file}")
            document.add_paragraph("")
            document.add_paragraph("---")
            document.add_paragraph("")

    if include_validation and report is not None:
        document.add_heading("Validation Summary", level=1)
        for line in _validation_markdown_section(report):
            if line.startswith("|") and "---" not in line and "Field" not in line:
                document.add_paragraph(line.strip("| ").replace(" | ", ": ", 1))

    document.save(output_path)


def monthly_output_paths(output_root: Path, month_label: str) -> dict[str, Path]:
    month_token = safe_month_filename(month_label)
    basename = f"{DEFAULT_OUTPUT_BASENAME}_{month_token}"
    return {
        "md": output_root / f"{basename}.md",
        "docx": output_root / f"{basename}.docx",
        "txt": output_root / f"{basename}.txt",
    }


def write_monthly_outputs(
    rows_by_month: dict[str, list[SmsRow]],
    output_root: Path,
    report: ValidationReport,
) -> None:
    report.monthly_files_generated = []
    for month_label in sorted(rows_by_month, key=month_sort_key):
        month_rows = {month_label: rows_by_month[month_label]}
        paths = monthly_output_paths(output_root, month_label)
        title = f"Consolidated SMS Outputs - {month_heading(month_label)}"
        write_markdown_output(month_rows, paths["md"], title=title)
        write_text_output(month_rows, paths["txt"], title=title)
        write_docx_output(month_rows, paths["docx"], title=title)
        for path in paths.values():
            report.monthly_files_generated.append(str(path.resolve()))


def validate_written_exports(
    rows: list[SmsRow],
    rows_by_month: dict[str, list[SmsRow]],
    output_md: Path,
    output_txt: Path,
    report: ValidationReport,
) -> None:
    md_content = output_md.read_text(encoding="utf-8")
    txt_content = output_txt.read_text(encoding="utf-8")

    for row in rows:
        label = f"{row.dealership} / {row.month_label}"
        if row.sms_text not in md_content:
            report.sms_text_preservation_failures.append(f"{label}: SMS text missing from full Markdown export")
        if row.sms_text not in txt_content:
            report.sms_text_preservation_failures.append(f"{label}: SMS text missing from full TXT export")
        if f"### {row.dealership}" not in md_content:
            report.sms_text_preservation_failures.append(f"{label}: dealership heading missing from Markdown export")
        if row.dealership not in txt_content:
            report.sms_text_preservation_failures.append(f"{label}: dealership missing from TXT export")
        if f"Character Count: {row.char_count}" not in md_content:
            report.character_count_failures.append(f"{label}: character count missing from Markdown export")
        if f"Character Count: {row.char_count}" not in txt_content:
            report.character_count_failures.append(f"{label}: character count missing from TXT export")
        if row.source_file not in md_content:
            report.sms_text_preservation_failures.append(f"{label}: source file missing from Markdown export")
        if row.source_file not in txt_content:
            report.sms_text_preservation_failures.append(f"{label}: source file missing from TXT export")

    spacing_issues: list[str] = []
    separator_count_md = md_content.count("\n---\n")
    separator_count_txt = txt_content.count(f"\n{TXT_SEPARATOR}\n")
    if separator_count_md < len(rows):
        spacing_issues.append(
            f"Full Markdown export has insufficient separators: expected at least {len(rows)}, found {separator_count_md}"
        )
    if separator_count_txt < len(rows):
        spacing_issues.append(
            f"Full TXT export has insufficient separators: expected at least {len(rows)}, found {separator_count_txt}"
        )

    for month_label, month_rows in rows_by_month.items():
        paths = monthly_output_paths(output_md.parent, month_label)
        for kind, path in paths.items():
            if not path.exists():
                spacing_issues.append(f"Missing monthly {kind.upper()} file for {month_heading(month_label)}")
        if not paths["md"].exists():
            continue

        month_md = paths["md"].read_text(encoding="utf-8")
        month_txt = paths["txt"].read_text(encoding="utf-8")
        for other_month in rows_by_month:
            if other_month == month_label:
                continue
            other_heading = month_heading(other_month)
            if f"## {other_heading}" in month_md:
                spacing_issues.append(
                    f"Monthly Markdown for {month_heading(month_label)} incorrectly includes {other_heading}"
                )

        for row in month_rows:
            label = f"{row.dealership} / {row.month_label}"
            if row.sms_text not in month_md:
                spacing_issues.append(f"{label}: SMS text missing from monthly Markdown export")
            if row.sms_text not in month_txt:
                spacing_issues.append(f"{label}: SMS text missing from monthly TXT export")
            if f"### {row.dealership}" not in month_md:
                spacing_issues.append(f"{label}: dealership heading missing from monthly Markdown export")

    if spacing_issues:
        report.spacing_separator_validation = "FAIL"
        report.sms_text_preservation_failures.extend(spacing_issues)
    else:
        report.spacing_separator_validation = "PASS"


def _validation_markdown_section(report: ValidationReport) -> list[str]:
    return [
        "## Validation Summary",
        "",
        "| Field | Result |",
        "| ----- | ------ |",
        f"| Consolidation complete | {'YES' if report.consolidation_complete else 'NO'} |",
        f"| Months found | {', '.join(month_heading(m) for m in report.months_found) or 'none'} |",
        f"| Dealership SMS rows written | {report.rows_written} |",
        f"| Missing SMS outputs | {len(report.missing_sms_outputs)} |",
        f"| SMS over 160 chars | {len(report.sms_over_160)} |",
        f"| Duplicate dealership/month outputs | {len(report.duplicate_outputs)} |",
        f"| SMS text preservation failures | {len(report.sms_text_preservation_failures)} |",
        f"| Character count failures | {len(report.character_count_failures)} |",
        f"| Spacing/separator validation | {report.spacing_separator_validation} |",
        f"| Source files modified | {', '.join(report.source_files_modified) or 'none'} |",
        f"| Production reports modified | {'YES' if report.production_reports_modified else 'NO'} |",
        f"| Emails sent | NO |",
        f"| SMS sent | NO |",
        f"| Output document path | {report.output_md_path} |",
        f"| Final verdict | {report.final_verdict} |",
        "",
    ]


def capture_production_mtimes(root: Path = ROOT) -> dict[Path, float]:
    if not PRODUCTION_OUTPUT_DIR.is_dir():
        return {}
    mtimes: dict[Path, float] = {}
    for path in PRODUCTION_OUTPUT_DIR.rglob("*"):
        if path.is_file():
            mtimes[path.resolve()] = path.stat().st_mtime
    return mtimes


def detect_production_modifications(before: dict[Path, float]) -> tuple[bool, list[str]]:
    modified: list[str] = []
    if not PRODUCTION_OUTPUT_DIR.is_dir():
        return False, modified
    for path in PRODUCTION_OUTPUT_DIR.rglob("*"):
        if not path.is_file():
            continue
        resolved = path.resolve()
        previous = before.get(resolved)
        current = path.stat().st_mtime
        if previous is None or current != previous:
            modified.append(str(resolved))
    return bool(modified), modified


def print_console_summary(report: ValidationReport) -> None:
    print("")
    print("| Field | Result |")
    print("| ------------------------------------ | -------------------------------- |")
    print(f"| Consolidated export patch complete | {'YES' if report.consolidation_complete else 'NO'} |")
    modified_display = ", ".join(report.source_files_modified) if report.source_files_modified else "tools/generate_consolidated_sms_outputs.py"
    print(f"| Source files modified | {modified_display} |")
    print(f"| Dealership reports modified | {'YES' if report.production_reports_modified else 'NO'} |")
    print(f"| Full consolidated MD generated | {'YES' if report.output_md_path else 'NO'} + {report.output_md_path or 'n/a'} |")
    print(f"| Full consolidated DOCX generated | {'YES' if report.output_docx_path else 'NO'} + {report.output_docx_path or 'n/a'} |")
    print(f"| Full consolidated TXT generated | {'YES' if report.output_txt_path else 'NO'} + {report.output_txt_path or 'n/a'} |")
    monthly_display = ", ".join(Path(path).name for path in report.monthly_files_generated) if report.monthly_files_generated else "none"
    print(f"| Monthly consolidated files generated | {len(report.monthly_files_generated)} ({monthly_display}) |")
    months_display = ", ".join(month_heading(m) for m in report.months_found) or "none"
    print(f"| Months processed | {months_display} |")
    print(f"| Dealership SMS rows written | {report.rows_written} |")
    missing_display = "; ".join(report.missing_sms_outputs) if report.missing_sms_outputs else "none"
    print(f"| Missing SMS outputs | {missing_display} |")
    dup_display = "; ".join(report.duplicate_outputs) if report.duplicate_outputs else "none"
    print(f"| Duplicate dealership/month outputs | {dup_display} |")
    preserve_display = "; ".join(report.sms_text_preservation_failures) if report.sms_text_preservation_failures else "none"
    print(f"| SMS text preservation failures | {preserve_display} |")
    count_display = "; ".join(report.character_count_failures) if report.character_count_failures else "none"
    print(f"| Character count failures | {count_display} |")
    print(f"| Spacing/separator validation | {report.spacing_separator_validation} |")
    excluded_display = "; ".join(report.excluded_unknown_accounts) if report.excluded_unknown_accounts else "none"
    print(f"| Excluded unknown-account files | {excluded_display} |")
    print("| Emails sent | NO |")
    print("| SMS sent | NO |")
    print(f"| Final verdict | {report.final_verdict} |")


def _default_accounts_dir() -> Path | None:
    try:
        from settings_manager import load_settings

        accounts_folder = (load_settings() or {}).get("accounts_folder")
        return Path(accounts_folder) if accounts_folder else None
    except Exception:
        return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate consolidated dealership SMS outputs.")
    parser.add_argument("--root", type=Path, default=ROOT, help="Repository root path")
    parser.add_argument("--output-md", type=Path, default=DEFAULT_OUTPUT_MD, help="Markdown output path")
    parser.add_argument("--output-docx", type=Path, default=DEFAULT_OUTPUT_DOCX, help="DOCX output path")
    parser.add_argument("--output-txt", type=Path, default=DEFAULT_OUTPUT_TXT, help="Plain-text output path")
    parser.add_argument(
        "--accounts-dir",
        type=Path,
        default=None,
        help=(
            "Accounts folder used to whitelist real dealership names (defaults to "
            "user_settings.json's accounts_folder). Files whose inferred dealership "
            "isn't a known account are excluded instead of being consolidated — this "
            "is what stops stray test/demo report files from reaching a client-facing "
            "export. Pass an empty/non-existent path to disable filtering."
        ),
    )
    args = parser.parse_args(argv)

    report = ValidationReport()
    production_before = capture_production_mtimes(args.root)

    try:
        accounts_dir = args.accounts_dir if args.accounts_dir is not None else _default_accounts_dir()
        known_dealerships = load_known_dealership_names(accounts_dir)
        if known_dealerships is None:
            print(
                "WARNING: no Accounts folder found — consolidating without a known-dealership "
                "whitelist. Output may include stray/test files.",
                file=sys.stderr,
            )
        candidates = discover_sms_outputs(args.root, known_dealerships=known_dealerships, report=report)
        rows = build_consolidated_rows(candidates, report)
        rows_by_month = group_rows_by_month(rows)

        report.consolidation_complete = True
        report.output_md_path = str(args.output_md.resolve())
        report.output_docx_path = str(args.output_docx.resolve())
        report.output_txt_path = str(args.output_txt.resolve())

        write_markdown_output(
            rows_by_month,
            args.output_md,
            title="Consolidated SMS Outputs",
            include_validation=True,
            report=report,
        )
        write_text_output(rows_by_month, args.output_txt, title="Consolidated SMS Outputs")
        write_docx_output(
            rows_by_month,
            args.output_docx,
            title="Consolidated SMS Outputs",
            include_validation=True,
            report=report,
        )
        write_monthly_outputs(rows_by_month, args.root, report)
        validate_written_exports(rows, rows_by_month, args.output_md, args.output_txt, report)
    except Exception as exc:
        print(f"Consolidation failed: {exc}", file=sys.stderr)
        print_console_summary(report)
        return 1

    production_modified, modified_paths = detect_production_modifications(production_before)
    report.production_reports_modified = production_modified
    if production_modified:
        report.source_files_modified.extend(modified_paths)

    print_console_summary(report)
    return 0 if report.final_verdict != "FAIL" else 1


if __name__ == "__main__":
    raise SystemExit(main())
