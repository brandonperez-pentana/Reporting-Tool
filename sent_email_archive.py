"""PDF copies of the emails Reporter builds.

Two layouts, because two callers want different things:

  * "month_folder" — <destination>/<Month Label>/<Dealership>.pdf. What the desktop
    app has always written, after a confirmed send.
  * "beside_source" — <folder>/<Dealership> - Email - <Month Label>.pdf, written
    straight into whichever folder that dealership's source reports came from, with
    no subfolder. The month moves into the filename because without a month folder
    there is nothing else keeping January's copy from overwriting February's.

Rendering strategy, in order of preference:
  1. Microsoft Word via COM (present wherever classic Outlook is) — highest
     fidelity, keeps the branded table layout and images.
  2. An .html file alongside, if Word is unavailable — never silently produce
     nothing, since the archive is the audit trail.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import settings_manager

_LOG = logging.getLogger("sent_email_archive")

# Word's WdSaveFormat for PDF.
WD_FORMAT_PDF = 17


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', " ", str(value or "")).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "email"


def archive_folder_for_month(base_folder: str | Path, month_label: str) -> Path:
    return Path(base_folder).expanduser() / safe_filename(month_label)


def _open_word() -> Any | None:
    """Start a private Word instance for rendering.

    DispatchEx, not Dispatch: Word is a multi-use COM server, so Dispatch
    attaches to the operator's own running Word. This code then hides it,
    suppresses its alerts and calls Quit() on it — closing their open documents
    without asking. A dedicated instance cannot touch their session.
    """
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        _LOG.info("pywin32 unavailable; cannot render PDF via Word.")
        return None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        return word
    except Exception as exc:  # noqa: BLE001 - Word may not be installed
        _LOG.warning("Could not start Word for PDF rendering: %s", exc)
        return None


def _close_word(word: Any | None) -> None:
    if word is None:
        return
    try:
        word.Quit()
    except Exception:  # noqa: BLE001
        _LOG.warning("Word did not quit cleanly after archiving.", exc_info=True)


def _html_to_pdf_via_word(html_path: Path, pdf_path: Path, word: Any | None = None) -> bool:
    """Render one HTML file to PDF. Reuses `word` when given, so a batch pays
    a single Word start-up instead of one per email."""
    own_instance = word is None
    if own_instance:
        word = _open_word()
    if word is None:
        return False

    document = None
    try:
        document = word.Documents.Open(str(html_path), ConfirmConversions=False, ReadOnly=True)
        document.SaveAs(str(pdf_path), FileFormat=WD_FORMAT_PDF)
        return True
    except Exception as exc:  # noqa: BLE001 - Word may be missing or busy
        _LOG.warning("Word could not render %s to PDF: %s", html_path.name, exc)
        return False
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:  # noqa: BLE001
            pass
        if own_instance:
            _close_word(word)


def _html_to_pdf_via_browser(html_path: Path, pdf_path: Path) -> bool:
    """The server's renderer: Microsoft does not support automating Word from a service."""
    try:
        import performance_report
    except Exception:  # noqa: BLE001
        return False
    return performance_report.render_html_file(html_path, pdf_path)


def archive_sent_email(
    *,
    dealership_name: str,
    month_label: str,
    html_body: str,
    destination_folder: str | Path,
    word: Any | None = None,
    layout: str = "month_folder",
) -> dict[str, Any]:
    """Write a PDF (or HTML fallback) copy of one email.

    Never raises. When this runs after a send the email has already gone, and when it
    runs on its own a failed copy must not read as a failed email — either way the
    caller gets a row describing what happened rather than an exception.
    """
    result: dict[str, Any] = {"dealership": dealership_name, "path": "", "format": "", "error": ""}
    try:
        if layout == "beside_source":
            folder = Path(destination_folder).expanduser()
            stem = safe_filename(f"{dealership_name} - Email - {month_label}")
        else:
            folder = archive_folder_for_month(destination_folder, month_label)
            stem = safe_filename(dealership_name)
        folder.mkdir(parents=True, exist_ok=True)

        html_path = folder / f"{stem}.html"
        html_path.write_text(html_body or "", encoding="utf-8")

        pdf_path = folder / f"{stem}.pdf"
        if settings_manager.server_mode():
            rendered = _html_to_pdf_via_browser(html_path, pdf_path)
        else:
            rendered = _html_to_pdf_via_word(html_path, pdf_path, word)
        if rendered and pdf_path.exists():
            # Rendered; the intermediate HTML is no longer needed.
            try:
                html_path.unlink()
            except OSError:
                pass
            result["path"] = str(pdf_path)
            result["format"] = "pdf"
        else:
            result["path"] = str(html_path)
            result["format"] = "html"
            result["error"] = (
                "The PDF renderer is unavailable — saved as HTML instead of PDF."
                if settings_manager.server_mode()
                else "Word unavailable — saved as HTML instead of PDF."
            )
        return result
    except Exception as exc:  # noqa: BLE001
        _LOG.exception("Unable to archive sent email for %s", dealership_name)
        result["error"] = str(exc)
        return result


def archive_sent_emails(
    *,
    sent_results: list[dict[str, Any]],
    month_label: str,
    destination_folder: str | Path = "",
    layout: str = "month_folder",
) -> list[dict[str, Any]]:
    """Archive every email in the batch that did not fail.

    An item may carry its own "destination_folder", which wins over the shared one.
    That is how "beside the source reports" works: every dealership has a different
    answer, and only the caller that built the session knows what it is.
    """
    archived: list[dict[str, Any]] = []
    pending = [item for item in sent_results if item.get("status") != "failed"]
    if not pending:
        return archived

    # One Word instance for the whole batch. Starting and quitting Word per
    # email cost 1-3 seconds each, so a 27-email run spent up to a minute in
    # start-up alone.
    word = None if settings_manager.server_mode() else _open_word()
    try:
        for item in pending:
            archived.append(
                archive_sent_email(
                    dealership_name=str(item.get("dealership", "")),
                    month_label=month_label,
                    html_body=str(item.get("archive_html", "") or item.get("html_body", "")),
                    destination_folder=str(item.get("destination_folder") or destination_folder),
                    word=word,
                    layout=layout,
                )
            )
    finally:
        _close_word(word)
        if settings_manager.server_mode():
            # The browser belongs to this worker thread; left open it could not be closed later.
            try:
                import performance_report

                performance_report.shutdown()
            except Exception:  # noqa: BLE001
                _LOG.debug("Renderer shutdown failed", exc_info=True)
    return archived
