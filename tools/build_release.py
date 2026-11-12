"""Build a clean, data-screened release zip from the committed Git tree.

Run after committing the release candidate:
    python tools/build_release.py

The working tree must be clean. The archive is built from HEAD (not the current
folder), then checked for paths and payloads that must never travel with a release.
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path, PurePosixPath
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
PREFIX = "Reporter-Web"
CONTACTS_HEADER = [
    "dealership", "contact_name", "email", "cc", "bcc", "active", "notes", "last_updated",
]
_ALLOWED_ROOT_TEXT = {"README.txt", "requirements-web.txt"}
_ALLOWED_DATA = {"dealership_aliases.json", "contacts.csv"}
_ALLOWED_DOCUMENT_DIRS = {"accounts", "templates", "tests"}
_FORBIDDEN_NAMES = {
    ".venv", "venv", "logs", "insights output", "performance intelligence",
    "reporting_audit", "user_settings.json", "shared_settings.json",
    "python-path.txt", "metric_history.json",
}
_FORBIDDEN_SUFFIXES = {".eml", ".msg", ".oft", ".pst", ".zip"}
_DOCUMENT_SUFFIXES = {".docx", ".xlsx", ".pdf"}


class ReleaseSafetyError(RuntimeError):
    pass


def _relative_parts(name: str) -> tuple[str, ...]:
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts:
        raise ReleaseSafetyError(f"Unsafe archive path: {name}")
    parts = path.parts
    if parts and parts[0] == PREFIX:
        parts = parts[1:]
    return parts


def validate_archive(archive: zipfile.ZipFile) -> list[str]:
    """Return all release-policy violations found in a candidate archive."""
    errors: list[str] = []
    file_names = {info.filename for info in archive.infolist() if not info.is_dir()}

    for name in sorted(file_names):
        try:
            parts = _relative_parts(name)
        except ReleaseSafetyError as exc:
            errors.append(str(exc))
            continue
        if not parts:
            continue

        lowered = tuple(part.casefold() for part in parts)
        path = PurePosixPath(*parts)
        suffix = path.suffix.casefold()
        basename = path.name.casefold()

        if any(part in _FORBIDDEN_NAMES for part in lowered):
            errors.append(f"forbidden path: {name}")
        if (basename.startswith(".env") and basename != ".env.example") or basename.endswith(".local.json"):
            errors.append(f"machine-specific or secret file: {name}")
        if any(part.startswith("backup-") for part in lowered):
            errors.append(f"settings backup: {name}")
        if suffix in _FORBIDDEN_SUFFIXES:
            errors.append(f"private mail or nested archive: {name}")

        if len(parts) == 1:
            if suffix in {".txt", ".docx", ".csv", ".xlsx", ".json", ".pdf"} and path.name not in _ALLOWED_ROOT_TEXT:
                errors.append(f"unexpected root-level data file: {name}")

        if lowered[0] == "data" and len(parts) > 1 and parts[1].casefold() not in _ALLOWED_DATA:
            errors.append(f"unapproved data file: {name}")
        if suffix == ".csv" and not (
            lowered[:2] == ("data", "contacts.csv") or lowered[0] == "tests"
        ):
            errors.append(f"unexpected CSV data: {name}")
        if suffix in _DOCUMENT_SUFFIXES and lowered[0] not in _ALLOWED_DOCUMENT_DIRS:
            errors.append(f"unexpected report/document: {name}")
        if lowered[0] == "accounts" and len(parts) > 1:
            account_name = PurePosixPath(parts[1]).stem
            account_match = re.fullmatch(r"demo site (?:[1-9]|10)", account_name, re.IGNORECASE)
            if not account_match:
                errors.append(f"non-demo account data: {name}")
            elif len(parts) == 2 and suffix != ".xlsx":
                errors.append(f"unexpected top-level demo account file: {name}")
            elif len(parts) == 3:
                expected_pdf = re.fullmatch(
                    rf"{re.escape(account_name)} - (?:sales|service) report - july 26\.pdf",
                    path.name,
                    re.IGNORECASE,
                )
                if suffix != ".pdf" or not expected_pdf:
                    errors.append(f"unexpected demo account report: {name}")
            elif len(parts) > 3:
                errors.append(f"unexpected nested demo account file: {name}")

    contacts_name = f"{PREFIX}/data/contacts.csv"
    if contacts_name in file_names:
        try:
            content = archive.read(contacts_name).decode("utf-8-sig")
            rows = list(csv.reader(io.StringIO(content)))
            if not rows or rows[0] != CONTACTS_HEADER or len(rows) != 1:
                errors.append("data/contacts.csv must contain only the standard header row")
        except (UnicodeDecodeError, csv.Error, OSError):
            errors.append("data/contacts.csv is not a valid UTF-8 header-only starter file")

    for required in ("reporter_web.py", "README.md", "requirements-web.txt", "VERSION"):
        if f"{PREFIX}/{required}" not in file_names:
            errors.append(f"required release file is missing: {required}")
    return errors


def build_release(*, force: bool = False) -> Path:
    status = subprocess.run(
        ["git", "status", "--porcelain", "--untracked-files=normal"],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()
    if status:
        raise ReleaseSafetyError("Refusing to release: commit changes and remove untracked files first.")

    version = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
    if not re.fullmatch(r"[0-9]{4}\.[0-9]{2}\.[0-9]{2}(?:[-+][A-Za-z0-9.-]+)?", version):
        raise ReleaseSafetyError(f"Unexpected VERSION value: {version!r}")

    output_dir = ROOT / "dist"
    output_dir.mkdir(exist_ok=True)
    output = output_dir / f"Reporter-Web-{version}.zip"
    if output.exists() and not force:
        raise ReleaseSafetyError(f"{output.name} already exists; pass --force to replace it.")

    try:
        subprocess.run(
            ["git", "archive", "--format=zip", f"--prefix={PREFIX}/", "--output", str(output), "HEAD"],
            cwd=ROOT,
            check=True,
        )
        with zipfile.ZipFile(output) as archive:
            errors = validate_archive(archive)
        if errors:
            raise ReleaseSafetyError("Release contents failed safety checks:\n- " + "\n- ".join(errors))
    except Exception:
        output.unlink(missing_ok=True)
        raise
    return output


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true", help="Replace an existing zip for this version.")
    args = parser.parse_args(argv)
    try:
        output = build_release(force=args.force)
    except (OSError, subprocess.CalledProcessError, ReleaseSafetyError) as exc:
        print(f"Release not built: {exc}", file=sys.stderr)
        return 1
    print(f"Release ready: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
