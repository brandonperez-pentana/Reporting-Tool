"""JSON-facing wrapper around report_engine, shared by every web route.

The desktop app (reporter_ui.py) and this web app are two front ends over the same
engine. Nothing here re-implements reporting, email assembly or sending: every call
below delegates to report_engine, exactly as the Tk UI does, so the two stay in step
by construction. Keeping that boundary is the whole point of this module -- if you
find yourself parsing a PDF or formatting an insight here, it belongs in the engine.

Two things this layer adds that a desktop UI gets for free:

* Serialisation. report_engine keeps one module-level session, and a browser can fire
  overlapping requests in a way a single-threaded Tk loop never does. Every engine call
  goes through one re-entrant lock.
* Jobs. Reporting runs and send batches take minutes. They run on a worker thread and
  the browser polls for progress, which is the web equivalent of the Tk worker queue.
"""

from __future__ import annotations

import contextvars
import hashlib
import logging
import os
import sys
import threading
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import ai_config  # noqa: E402
import dealership_insights_local as reporting_core  # noqa: E402
import email_builder  # noqa: E402
import excel_reader  # noqa: E402
import report_engine  # noqa: E402
import settings_manager  # noqa: E402

LOGGER = logging.getLogger("reporter.web")


class _PerUserLock:
    """Serialises engine calls per signed-in user, since each user has their own session.

    The desktop (no user) gets one lock, exactly as before.
    """

    def __init__(self) -> None:
        self._locks: dict[str, threading.RLock] = {}
        self._guard = threading.Lock()

    def _lock(self) -> threading.RLock:
        user = settings_manager.current_user() or ""
        with self._guard:
            return self._locks.setdefault(user, threading.RLock())

    def __enter__(self) -> "_PerUserLock":
        self._lock().acquire()
        return self

    def __exit__(self, *exc: Any) -> None:
        self._lock().release()


# report_engine holds a session per user in module state. A browser can send two
# requests at once, so calls are serialised per user.
_ENGINE_LOCK = _PerUserLock()
# The reporting pipeline and the shared browser are process-wide: one at a time, for everyone.
_RUN_LOCK = threading.Lock()

_JOBS: dict[str, dict[str, Any]] = {}
_JOBS_LOCK = threading.Lock()
MAX_REMEMBERED_JOBS = 40


class ApiError(Exception):
    """A failure with a message meant for the operator, not a stack trace."""

    def __init__(self, message: str, status: int = 400) -> None:
        super().__init__(message)
        self.status = status
        self.message = message


def friendly_error(exc: BaseException) -> str:
    """Plain-language rendering of the errors this app actually raises.

    Mirrors reporter_ui.friendly_error_message so both front ends say the same thing.
    """
    text = str(exc).strip() or exc.__class__.__name__
    lowered = text.lower()
    if "invalid class string" in lowered or "-2147221005" in text:
        return (
            "Windows could not resolve 'Outlook.Application', so Reporter cannot drive "
            "Outlook. Classic Outlook being installed and open is not enough on its own: "
            "the usual causes are a bitness mismatch between Python and Office, running "
            "Reporter from an elevated prompt when Office is registered per user, or a "
            "registration that needs an Office Quick Repair. Reporting and the SMS "
            "document are unaffected. Run tools/diagnose_outlook.py to see which it is."
        )
    if "pywin32 is required" in lowered:
        return (
            "Outlook automation needs the pywin32 package on this machine. "
            "Install it with: pip install pywin32"
        )
    if isinstance(exc, PermissionError):
        return f"{text}. The file may be open in another program."
    if isinstance(exc, FileNotFoundError):
        return f"{text}. Check the paths in Settings."
    return text


# --- jobs ----------------------------------------------------------------------

def _job_create(kind: str) -> str:
    job_id = uuid.uuid4().hex[:12]
    with _JOBS_LOCK:
        _JOBS[job_id] = {
            "id": job_id,
            "kind": kind,
            "owner": settings_manager.current_user() or "",
            "state": "running",
            "message": "Starting…",
            "phase": "",
            "processed": None,
            "total": None,
            "result": None,
            "error": "",
        }
        if len(_JOBS) > MAX_REMEMBERED_JOBS:
            for stale in sorted(_JOBS, key=lambda key: _JOBS[key]["state"] != "running")[:10]:
                if _JOBS[stale]["state"] != "running":
                    _JOBS.pop(stale, None)
    return job_id


def _job_update(job_id: str, **fields: Any) -> None:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job:
            job.update(fields)


def job_status(job_id: str) -> dict[str, Any]:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if not job or job.get("owner", "") != (settings_manager.current_user() or ""):
            raise ApiError("That job is no longer available. Refresh and try again.", 404)
        return dict(job)


def _hold_run_lock(job_id: str) -> None:
    if not _RUN_LOCK.acquire(blocking=False):
        _job_update(job_id, message="Waiting for another user's run to finish…")
        _RUN_LOCK.acquire()


def _start_job(kind: str, work: Callable[[str], Any]) -> dict[str, Any]:
    job_id = _job_create(kind)
    # A new thread starts with an empty context; carry the signed-in user across.
    context = contextvars.copy_context()

    def runner() -> None:
        try:
            result = work(job_id)
        except Exception as exc:  # noqa: BLE001
            LOGGER.exception("%s job failed", kind)
            _job_update(job_id, state="failed", error=friendly_error(exc), message="Failed.")
            return
        _job_update(job_id, state="done", result=result, message="Finished.")

    threading.Thread(target=context.run, args=(runner,), name=f"reporter-web-{kind}", daemon=True).start()
    return {"job_id": job_id}


# --- settings ------------------------------------------------------------------

SETTING_KEYS = (
    "accounts_folder",
    "pdf_source_folder",
    "output_folder",
    "recipient_mapping_path",
    "month",
    "sent_email_pdf_folder",
    "test_send_address",
    "eml_export_folder",
    "default_bcc",
    "insight_engine",
    "ai_insights_enabled",
    "performance_report",
)

# What a run cannot proceed without. The first-run screen asks for exactly these and
# nothing else, so a new operator is not made to fill in a form before seeing the app.
REQUIRED_SETTING_KEYS = ("accounts_folder", "output_folder", "recipient_mapping_path")


def get_settings() -> dict[str, Any]:
    settings = settings_manager.load_settings()
    payload = {key: str(settings.get(key, "") or "") for key in SETTING_KEYS}
    payload["portfolio"] = portfolio_from(settings)
    payload["settings_path"] = str(settings_manager.get_settings_path())
    payload["default_contacts_path"] = str(ensure_contacts_file(ROOT / "data" / "contacts.csv"))
    payload["missing_required"] = [
        key for key in REQUIRED_SETTING_KEYS if not str(settings.get(key, "") or "").strip()
    ]
    # A copy that has never been configured, as opposed to one mid-setup: the settings
    # file itself does not exist yet. This is what the guided first run keys off, so a
    # returning operator who cleared one field is not sent back through setup.
    payload["first_run"] = not settings_manager.get_settings_path().exists()
    payload["server_mode"] = settings_manager.server_mode()
    payload["user"] = settings_manager.current_user() or ""
    payload["is_admin"] = settings_manager.is_admin()
    payload["shared_keys"] = [key for key in SETTING_KEYS if key not in settings_manager.USER_SETTING_KEYS]
    if payload["server_mode"]:
        # A new user inherits the administrator's setup; only a gap in it needs the setup screen.
        payload["first_run"] = bool(payload["missing_required"]) and payload["is_admin"]
    # Connection status for the optional AI layer. Read-only and always masked:
    # the key lives in a .env file and is never sent to the browser.
    payload["ai"] = ai_status()
    payload["performance_engine"] = performance_engine_description()
    return payload


def save_settings(payload: dict[str, Any]) -> dict[str, Any]:
    changes: dict[str, Any] = {
        key: str(payload.get(key) or "").strip() for key in SETTING_KEYS if key in payload
    }
    if "portfolio" in payload:
        changes["portfolio"] = clean_portfolio(payload.get("portfolio"))

    if not settings_manager.server_mode():
        settings_manager.save_settings(changes)
        return get_settings()

    current = settings_manager.load_settings()
    personal = {key: value for key, value in changes.items() if key in settings_manager.USER_SETTING_KEYS}
    # The settings form posts every field; only a value that actually moved counts as a change.
    shared = {
        key: value for key, value in changes.items()
        if key not in settings_manager.USER_SETTING_KEYS and value != str(current.get(key, "") or "")
    }
    if shared and not settings_manager.is_admin():
        raise ApiError(f"Only an administrator can change: {', '.join(sorted(shared))}.", 403)
    if shared:
        settings_manager.save_settings(shared, shared=True)
    settings_manager.save_settings(personal)
    return get_settings()


def _require_admin_on_server(action: str) -> None:
    if settings_manager.server_mode() and not settings_manager.is_admin():
        raise ApiError(f"Only an administrator can {action} on the server.", 403)


def clean_portfolio(values: Any) -> list[str]:
    """Bare workbook filenames, de-duplicated case-insensitively."""
    if not isinstance(values, list):
        return []
    names: list[str] = []
    seen: set[str] = set()
    for value in values:
        name = Path(str(value or "")).name.strip()
        if name and name.casefold() not in seen:
            seen.add(name.casefold())
            names.append(name)
    return names


def portfolio_from(settings: dict[str, Any]) -> list[str]:
    return clean_portfolio(settings.get("portfolio"))


def portfolio_options() -> dict[str, Any]:
    """The workbooks in the accounts folder, for picking "My dealerships".

    Listed by filename only: opening 150 workbooks on OneDrive to read their titles is
    the very cost this list exists to avoid.
    """
    settings = settings_manager.load_settings()
    accounts = str(settings.get("accounts_folder") or "").strip()
    if not accounts:
        raise ApiError("Set the Accounts folder in Settings first.")
    if not Path(accounts).expanduser().is_dir():
        raise ApiError(f"The accounts folder does not exist: {accounts}")
    names = [path.name for path in excel_reader.workbook_paths(accounts, recursive=False)]
    present = {name.casefold() for name in names}
    selected = portfolio_from(settings)
    return {
        "accounts_folder": accounts,
        "workbooks": names,
        "selected": selected,
        "missing": [name for name in selected if name.casefold() not in present],
    }


def performance_engine_description() -> str:
    """How the Performance Intelligence report can be rendered here, for Settings."""
    try:
        import performance_report

        return performance_report.describe()
    except Exception as exc:  # noqa: BLE001
        LOGGER.info("Performance renderer unavailable: %s", exc)
        return "unavailable"


def ai_status() -> dict[str, Any]:
    """Masked view of the AI configuration, safe for the browser and for logs."""
    try:
        return ai_config.load().status()
    except Exception as exc:  # noqa: BLE001 - config problems must not break Settings
        LOGGER.warning("AI config could not be read: %s", exc)
        return {"configured": False, "active": False, "reason": friendly_error(exc)}


def ai_test() -> dict[str, Any]:
    """One round trip to the configured provider, for a Test connection button."""
    try:
        import ai_client

        return ai_client.check_connection()
    except Exception as exc:  # noqa: BLE001
        LOGGER.exception("AI connection test failed")
        return {"ok": False, "message": friendly_error(exc), "status": ai_status()}


def path_check(payload: dict[str, Any]) -> dict[str, Any]:
    """Tell the operator whether a typed path exists, since there is no file dialog."""
    _require_admin_on_server("check paths")
    raw = str(payload.get("path") or "").strip()
    kind = str(payload.get("kind") or "any")
    if not raw:
        return {"path": "", "exists": False, "kind": kind, "note": "Not set."}
    resolved = Path(raw).expanduser()
    exists = resolved.exists()
    is_dir = resolved.is_dir()
    note = "Found." if exists else "No file or folder at that path."
    if exists and kind == "folder" and not is_dir:
        note = "That is a file, not a folder."
    if exists and kind == "file" and is_dir:
        note = "That is a folder, not a file."
    return {"path": str(resolved), "exists": exists, "is_dir": is_dir, "kind": kind, "note": note}


BROWSE_MAX_ENTRIES = 400


def _browse_places() -> list[dict[str, str]]:
    """Somewhere sensible to start, since typing a full path is the thing we are avoiding."""
    home = Path.home()
    candidates = [
        ("This app", ROOT),
        ("Home", home),
        ("Desktop", home / "Desktop"),
        ("Downloads", home / "Downloads"),
        ("Documents", home / "Documents"),
    ]
    for child in sorted(home.glob("OneDrive*")):
        if child.is_dir():
            candidates.append((child.name, child))
    places = []
    seen: set[str] = set()
    for label, path in candidates:
        try:
            if not path.is_dir():
                continue
        except OSError:
            continue
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        places.append({"label": label, "path": key})
    return places


def _browse_drives() -> list[str]:
    if os.name != "nt":
        return []
    drives = []
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        root = f"{letter}:\\"
        if os.path.exists(root):
            drives.append(root)
    return drives


def browse(payload: dict[str, Any]) -> dict[str, Any]:
    """List one directory on this machine so Settings can point at real paths.

    A browser cannot open a native file dialog on a server's behalf, and typing a full
    Windows path from memory is exactly the step this app should not ask for. The server
    runs on the machine that holds the files, so it lists them itself.

    Read-only, and no wider than Settings already implies: the operator can already send
    the app at any folder on this machine by typing it.
    """
    _require_admin_on_server("browse folders")
    raw = str(payload.get("path") or "").strip()
    kind = str(payload.get("kind") or "folder")
    extensions = tuple(
        item.strip().lower()
        for item in str(payload.get("extensions") or "").split(",")
        if item.strip()
    )

    places = _browse_places()
    if not raw:
        raw = places[0]["path"] if places else str(ROOT)

    current = Path(raw).expanduser()
    if not current.is_dir():
        # A file path or a typo: start from the nearest real folder rather than erroring.
        current = current.parent if current.parent.is_dir() else Path(places[0]["path"] if places else ROOT)

    entries: list[dict[str, Any]] = []
    truncated = False
    try:
        with os.scandir(current) as scanner:
            for item in scanner:
                if item.name.startswith("."):
                    continue
                try:
                    is_dir = item.is_dir()
                except OSError:
                    continue
                if not is_dir:
                    if kind == "folder":
                        continue
                    if extensions and not item.name.lower().endswith(extensions):
                        continue
                size = 0
                if not is_dir:
                    try:
                        size = item.stat().st_size
                    except OSError:
                        size = 0
                entries.append({"name": item.name, "path": str(Path(item.path)), "is_dir": is_dir, "size": size})
                if len(entries) >= BROWSE_MAX_ENTRIES:
                    truncated = True
                    break
    except PermissionError as exc:
        raise ApiError(f"No permission to list {current}", 403) from exc
    except OSError as exc:
        raise ApiError(f"Could not list {current}: {friendly_error(exc)}") from exc

    entries.sort(key=lambda item: (not item["is_dir"], item["name"].lower()))
    parent = str(current.parent) if current.parent != current else ""
    return {
        "path": str(current),
        "parent": parent,
        "entries": entries,
        "truncated": truncated,
        "places": places,
        "drives": _browse_drives(),
        "separator": os.sep,
    }


def month_options() -> dict[str, Any]:
    settings = settings_manager.load_settings()
    accounts = str(settings.get("accounts_folder") or "").strip()
    detected: list[str] = []
    if accounts:
        try:
            detected = report_engine.detect_available_report_months(
                accounts, include=portfolio_from(settings) or None
            )
        except Exception:  # noqa: BLE001
            detected = []
    # Default to the newest month the accounts actually hold. Falling back to last
    # calendar month was how a fresh run landed on a month with no figures in it and
    # reported every dealership as blocked.
    chosen = str(settings.get("month") or "").strip()
    if not chosen:
        chosen = detected[0] if detected else report_engine.default_month_label()
    return {
        "detected": detected,
        "all": report_engine.available_month_options(),
        "default": chosen,
    }


# --- reporting run --------------------------------------------------------------

def start_reporting(payload: dict[str, Any]) -> dict[str, Any]:
    settings = settings_manager.load_settings()
    if settings_manager.server_mode():
        # Folders come from settings only: a request must not point the server elsewhere.
        payload = {key: payload.get(key) for key in ("month", "all_dealerships")}
    accounts = str(payload.get("accounts_folder") or settings.get("accounts_folder") or "").strip()
    pdf_folder = str(payload.get("pdf_source_folder") or settings.get("pdf_source_folder") or "").strip()
    output = str(payload.get("output_folder") or settings.get("output_folder") or "").strip()
    contacts = str(payload.get("recipient_mapping_path") or settings.get("recipient_mapping_path") or "").strip()
    month = str(payload.get("month") or settings.get("month") or "").strip()
    portfolio = [] if payload.get("all_dealerships") else portfolio_from(settings)

    if not accounts:
        raise ApiError("Set the Accounts folder in Settings before running reporting.")
    if not Path(accounts).expanduser().is_dir():
        raise ApiError(f"The accounts folder does not exist: {accounts}")
    if pdf_folder and not Path(pdf_folder).expanduser().is_dir():
        raise ApiError(f"The PDF source folder does not exist: {pdf_folder}")
    if not output:
        raise ApiError("Set the Output folder in Settings before running reporting.")
    if not month:
        raise ApiError("Choose a month to report on.")

    def work(job_id: str) -> dict[str, Any]:
        def on_progress(update: report_engine.ProgressUpdate) -> None:
            _job_update(
                job_id,
                phase=update.phase,
                message=update.message,
                processed=update.processed_count,
                total=update.total_count,
            )

        _hold_run_lock(job_id)
        try:
            with _ENGINE_LOCK:
                report_engine.run_reporting(
                    accounts_folder=accounts,
                    output_folder=output,
                    month=month,
                    template_path=None,
                    recipient_mapping_path=contacts or None,
                    progress_callback=on_progress,
                    portfolio=portfolio or None,
                    pdf_folder=pdf_folder or None,
                )
                return session_state()
        finally:
            _RUN_LOCK.release()

    return _start_job("reporting", work)


def session_state() -> dict[str, Any]:
    with _ENGINE_LOCK:
        if not report_engine.has_active_session():
            return {"has_session": False, "summary": None, "results": [], "queue": []}
        return {
            "has_session": True,
            "summary": report_engine.get_session_summary(),
            "results": report_engine.get_dealership_results(),
            "queue": report_engine.get_send_queue(),
        }


# --- one dealership -------------------------------------------------------------

def dealership_report(name: str) -> dict[str, Any]:
    with _ENGINE_LOCK:
        _require_dealership(name)
        details = report_engine.get_dealership_report_text(name)
        approvable, reason = report_engine.can_approve(name)
        details["can_approve"] = approvable
        details["approve_blocked_reason"] = reason
        details["source_documents"] = report_engine.get_dealership_source_documents(name)
        return details


def save_dealership_text(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    fields = {
        key: payload[key]
        for key in (
            "executive_summary", "priority_insights", "summary", "sms_version",
            "to_recipients", "cc_recipients", "bcc_recipients",
        )
        if key in payload
    }
    with _ENGINE_LOCK:
        _require_dealership(name)
        report_engine.update_dealership_text(name, **fields)
        return dealership_report(name)


def save_working_email(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    fields = {
        key: payload[key]
        for key in ("subject", "to_recipients", "cc_recipients", "bcc_recipients", "body")
        if key in payload
    }
    with _ENGINE_LOCK:
        _require_dealership(name)
        report_engine.update_working_email(name, **fields)
        return dealership_report(name)


def set_approved(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    approved = bool(payload.get("approved", True))
    with _ENGINE_LOCK:
        _require_dealership(name)
        if approved:
            ok, reason = report_engine.can_approve(name)
            if not ok:
                raise ApiError(reason or "This dealership cannot be approved yet.")
        report_engine.mark_approved(name, approved)
        return dealership_report(name)


def set_skipped(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    with _ENGINE_LOCK:
        _require_dealership(name)
        report_engine.mark_skipped(name, bool(payload.get("skipped", True)))
        return dealership_report(name)


def set_ready(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    with _ENGINE_LOCK:
        _require_dealership(name)
        report_engine.set_ready_to_send(name, bool(payload.get("ready", True)))
        return {"queue": report_engine.get_send_queue()}


def start_email_build(name: str, payload: dict[str, Any]) -> dict[str, Any]:
    """Assemble the branded email. Slow enough to belong on a worker thread."""
    manual_to = payload.get("to")
    manual_cc = payload.get("cc")
    manual_bcc = payload.get("bcc")

    def work(job_id: str) -> dict[str, Any]:
        _job_update(job_id, message=f"Building email for {name}…")
        with _ENGINE_LOCK:
            _require_dealership(name)
            report_engine.build_email_preview(
                name, manual_to=manual_to, manual_cc=manual_cc, manual_bcc=manual_bcc
            )
            return dealership_report(name)

    return _start_job("email", work)


def _require_dealership(name: str) -> None:
    if not report_engine.has_active_session():
        raise ApiError("No reporting run is loaded. Run reporting first.", 409)
    if not report_engine.has_dealership(name):
        raise ApiError(f"'{name}' is not part of the current reporting run.", 404)


# --- drafts and copies ----------------------------------------------------------
#
# This front end cannot send. Every route here stops at something a person then acts
# on themselves — a draft sitting in Outlook, an .eml on disk, a PDF copy — so no
# click in a browser tab can put mail in front of a dealership. report_engine still
# has send_emails(send=True) because the desktop app uses it; nothing in webapp/ ever
# passes True.

def start_drafts(payload: dict[str, Any]) -> dict[str, Any]:
    """Save the ticked emails as Outlook drafts. Never sends."""
    names = [str(value) for value in (payload.get("dealerships") or []) if str(value).strip()]
    if not names:
        raise ApiError("Tick at least one dealership first.")
    if payload.get("send") or payload.get("test_address"):
        raise ApiError(
            "Reporter cannot send email. It saves drafts in Outlook, writes .eml files and "
            "saves PDF copies; sending is done from Outlook by a person."
        )
    if settings_manager.server_mode():
        raise ApiError(
            "Outlook drafts are not available when Reporter runs on the server — they would "
            "land in the server's mailbox, not yours. Use Download .eml instead."
        )

    def work(job_id: str) -> dict[str, Any]:
        _job_update(job_id, message=f"Saving drafts for {len(names)} email(s)…", total=len(names), processed=0)
        with _ENGINE_LOCK:
            results = report_engine.send_emails(names, send=False)
        # send_emails catches per dealership and stores str(exc), so a COM code reached
        # the operator unexplained. Translate each one the same way a raised error is.
        for row in results:
            if row.get("error"):
                row["error"] = friendly_error(Exception(str(row["error"])))
        succeeded = [row for row in results if row.get("status") != "failed"]
        return {
            "results": results,
            "succeeded": len(succeeded),
            "failed": len(results) - len(succeeded),
            "queue": session_state()["queue"],
        }

    return _start_job("drafts", work)


def start_pdf_copies(payload: dict[str, Any]) -> dict[str, Any]:
    """Write a PDF copy of each ticked email, beside that dealership's source reports."""
    names = [str(value) for value in (payload.get("dealerships") or []) if str(value).strip()]
    if not names:
        raise ApiError("Tick at least one dealership first.")
    settings = settings_manager.load_settings()
    if settings_manager.server_mode():
        payload = {key: value for key, value in payload.items() if key != "folder"}
    override = str(payload.get("folder") or settings.get("sent_email_pdf_folder") or "").strip()

    def work(job_id: str) -> dict[str, Any]:
        _job_update(job_id, message=f"Rendering {len(names)} PDF copy(ies)…", total=len(names), processed=0)
        # The renderer is shared with reporting runs, so this waits its turn like one.
        _hold_run_lock(job_id)
        try:
            with _ENGINE_LOCK:
                results = report_engine.archive_email_copies(names, override)
        finally:
            _RUN_LOCK.release()
        for row in results:
            if row.get("error"):
                row["error"] = friendly_error(Exception(str(row["error"])))
        written = [row for row in results if row.get("path")]
        folders = sorted({str(Path(row["path"]).parent) for row in written})
        note = ""
        if written and not override:
            note = (
                f"Saved beside the source reports in {len(folders)} folder(s)."
                if len(folders) > 1
                else f"Saved to {folders[0]}."
            )
        elif written:
            note = f"Saved to {override}."
        if any(row.get("format") == "html" for row in written):
            note += " PDF rendering was unavailable, so some copies were saved as HTML rather than PDF."
        return {
            "results": results,
            "succeeded": len(written),
            "failed": len(results) - len(written),
            "archived_note": note,
            "queue": session_state()["queue"],
        }

    return _start_job("pdf", work)


def start_export(payload: dict[str, Any]) -> dict[str, Any]:
    """Write the ticked emails to disk as .eml files for manual import into Outlook."""
    names = [str(value) for value in (payload.get("dealerships") or []) if str(value).strip()]
    if not names:
        raise ApiError("Tick at least one dealership first.")

    settings = settings_manager.load_settings()
    if settings_manager.server_mode():
        # A folder on the server is no use to the person at the browser; the job's
        # files are handed back as a download instead (see export_download).
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        folder = str(settings_manager.user_folder() / "eml" / f"{stamp}-{uuid.uuid4().hex[:6]}")
    else:
        folder = str(payload.get("folder") or settings.get("eml_export_folder") or "").strip()
    if not folder:
        raise ApiError(
            "Choose where the .eml files should go first — Settings, 'Save .eml files to'."
        )

    def work(job_id: str) -> dict[str, Any]:
        _job_update(job_id, message=f"Writing {len(names)} .eml file(s)…", total=len(names), processed=0)
        with _ENGINE_LOCK:
            results = report_engine.export_emails_as_eml(names, folder)
        for row in results:
            if row.get("error"):
                row["error"] = friendly_error(Exception(str(row["error"])))
        succeeded = [row for row in results if row.get("status") == "exported"]
        return {
            "results": results,
            "succeeded": len(succeeded),
            "failed": len(results) - len(succeeded),
            "folder": folder,
            "download": f"/api/jobs/{job_id}/download" if settings_manager.server_mode() and succeeded else "",
            "queue": session_state()["queue"],
        }

    return _start_job("export", work)


def export_download(job_id: str) -> tuple[bytes, str, str]:
    """The files a finished .eml export wrote: one .eml as-is, several as a zip."""
    import io
    import zipfile

    job = job_status(job_id)
    result = job.get("result") or {}
    if job.get("kind") != "export" or job.get("state") != "done" or not result.get("download"):
        raise ApiError("There is nothing to download for that job.", 404)
    files = sorted(Path(result["folder"]).glob("*.eml"))
    if not files:
        raise ApiError("The exported files are no longer on the server. Export again.", 404)
    if len(files) == 1:
        return files[0].read_bytes(), "message/rfc822", files[0].name
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, arcname=path.name)
    return buffer.getvalue(), "application/zip", "Reporter emails.zip"


def outlook_status() -> dict[str, Any]:
    """Can this machine drive Outlook at all?

    Checked before a batch rather than discovered after failing every email in it. Only
    creates the Application object; it opens nothing and sends nothing.
    """
    if settings_manager.server_mode():
        return {
            "available": False,
            "reason": "Outlook drafts are off on the server. Use Download .eml and open the files in your own Outlook.",
            "checked": True,
            "server_mode": True,
        }
    if os.name != "nt":
        return {"available": False, "reason": "Outlook automation only works on Windows.", "checked": True}
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return {
            "available": False,
            "reason": friendly_error(Exception("pywin32 is required for Outlook draft creation.")),
            "checked": True,
        }
    try:
        import pythoncom  # type: ignore[import-not-found]

        pythoncom.CoInitialize()
        try:
            win32com.client.Dispatch("Outlook.Application")
        finally:
            pythoncom.CoUninitialize()
    except Exception as exc:  # noqa: BLE001
        return {"available": False, "reason": friendly_error(exc), "checked": True}
    return {"available": True, "reason": "", "checked": True}


# --- SMS stage ------------------------------------------------------------------

def sms_document() -> dict[str, Any]:
    """The consolidated SMS document, identical in shape to the desktop stage."""
    with _ENGINE_LOCK:
        rows = report_engine.get_send_queue() if report_engine.has_active_session() else []
    limit = reporting_core.SMS_CHAR_LIMIT
    blocks: list[str] = []
    entries: list[dict[str, Any]] = []
    over = 0
    for row in rows:
        name = str(row["dealership"])
        sms = str(row.get("sms_version", "")).strip()
        length = len(sms)
        if length > limit:
            over += 1
        flag = f"  [OVER by {length - limit}]" if length > limit else ""
        blocks.append(f"{name}  ({length}/{limit}){flag}\n{'-' * 60}\n{sms or '(no SMS recorded)'}")
        entries.append({"dealership": name, "sms": sms, "length": length, "over": length > limit})
    return {
        "document": "\n\n".join(blocks),
        "entries": entries,
        "dealership_count": len(rows),
        "over_limit": over,
        "limit": limit,
    }


def export_sms() -> dict[str, Any]:
    with _ENGINE_LOCK:
        if not report_engine.has_active_session():
            raise ApiError("No reporting run is loaded.", 409)
        return report_engine.export_sms_breakdowns_for_session()


def save_sms_document(payload: dict[str, Any]) -> dict[str, Any]:
    """Write the SMS document beside the month's other outputs."""
    text = str(payload.get("text") or "")
    if not text.strip():
        raise ApiError("There is nothing to save yet.")
    with _ENGINE_LOCK:
        if not report_engine.has_active_session():
            raise ApiError("No reporting run is loaded.", 409)
        session = report_engine.get_current_session()
        folder = Path(session.output_folder)
        slug = report_engine.month_label_slug(session.month_label)
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / f"SMS Messages - {slug}.txt"
    destination.write_text(text, encoding="utf-8")
    return {"path": str(destination)}


# --- contacts -------------------------------------------------------------------

# One contacts file is shared by every user on a server.
_CONTACTS_LOCK = threading.Lock()


def contacts(path: str | None = None) -> dict[str, Any]:
    if settings_manager.server_mode():
        path = None
    resolved = _contacts_path(path)
    # GETs are reachable cross-site, so only the configured file is ever created on read.
    if not path:
        resolved = ensure_contacts_file(resolved)
    try:
        with _CONTACTS_LOCK:
            rows = email_builder.load_contact_rows(resolved)
            revision = _contacts_revision(resolved)
    except Exception as exc:  # noqa: BLE001
        raise ApiError(f"Could not read the contacts file: {friendly_error(exc)}") from exc
    return {
        "path": str(resolved),
        "fieldnames": list(email_builder.CONTACT_CSV_FIELDNAMES),
        "rows": rows,
        "revision": revision,
        "dealership_names": dealership_names(),
    }


def save_contacts(payload: dict[str, Any]) -> dict[str, Any]:
    resolved = _contacts_path(None if settings_manager.server_mode() else payload.get("path"))
    rows = payload.get("rows")
    if not isinstance(rows, list):
        raise ApiError("Contacts must be a list of rows.")

    problems: list[str] = []
    cleaned: list[dict[str, str]] = []
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            raise ApiError(f"Contact row {index} must be an object.")
        clean = {key: str(row.get(key, "") or "").strip() for key in email_builder.CONTACT_CSV_FIELDNAMES}
        cleaned.append(clean)
        for error in email_builder.validate_contact_row_for_save(clean):
            problems.append(f"Row {index} ({clean.get('dealership', '')}): {error}")
    if problems:
        raise ApiError(" | ".join(problems[:12]))

    try:
        with _CONTACTS_LOCK:
            if "revision" in payload and payload["revision"] != _contacts_revision(resolved):
                raise ApiError(
                    "The contacts file changed since you opened it. Reopen Contacts and merge your edits before saving.",
                    409,
                )
            email_builder.save_contact_rows(resolved, cleaned)
    except OSError as exc:
        raise ApiError(f"Contacts were not saved to {resolved}: {friendly_error(exc)}") from exc

    # Remember a selected file, rather than reverting to the old/default contacts.csv
    # on the next open or run. On a server the configured shared path remains authoritative.
    configured = str(settings_manager.load_settings().get("recipient_mapping_path") or "").strip()
    if not settings_manager.server_mode() and (
        not configured or Path(configured).expanduser().resolve() != resolved
    ):
        try:
            settings_manager.save_settings({"recipient_mapping_path": str(resolved)})
        except OSError as exc:
            raise ApiError(
                f"Contacts were saved to {resolved}, but the selected path could not be remembered: {friendly_error(exc)}"
            ) from exc

    refreshed = False
    with _ENGINE_LOCK:
        if report_engine.has_active_session():
            try:
                report_engine.reload_session_recipient_mapping(resolved)
                refreshed = True
            except Exception:  # noqa: BLE001
                refreshed = False
    report_engine.append_workflow_evidence_record(
        "contacts_file_saved", {"path": str(resolved), "rows": len(cleaned), "source": "web"}
    )
    # The saved file is returned, not just a count. The panel re-renders from this,
    # so what the operator sees after saving is what is actually on disk — an entry
    # that did not make it cannot look like it did.
    saved = contacts(str(resolved))
    return {
        "path": str(resolved),
        "rows": len(cleaned),
        "dealerships": sorted({row["dealership"] for row in cleaned if row.get("dealership")}),
        "session_refreshed": refreshed,
        "note": "" if refreshed else "Rebuild the email preview to refresh greeting and recipients.",
        "contacts": saved,
    }


def dealership_names() -> list[str]:
    """The dealership names as the accounts folder spells them.

    Contacts are matched to reports by this string, so a contact typed as "Demo
    Site 2" against a folder called "Demo Site 2 Eastside" silently produces an
    email with no recipient. Offering the real names removes the class of mistake
    rather than validating it after the fact.

    Folders are the source of truth because contacts are usually filled in before the
    month has been run; the loaded session is merged in as well so a name that has been
    parsed is never missing from the list.
    """
    names: set[str] = set()
    folder = str(settings_manager.load_settings().get("accounts_folder") or "").strip()
    if folder:
        root = Path(folder).expanduser()
        if root.is_dir():
            for child in root.iterdir():
                if child.is_dir() and not child.name.startswith("."):
                    names.add(child.name)
    try:
        with _ENGINE_LOCK:
            if report_engine.has_active_session():
                for row in report_engine.get_dealership_results():
                    name = str(row.get("dealership") or "").strip()
                    if name:
                        names.add(name)
    except Exception:  # noqa: BLE001 - a picker is never worth failing the page for
        logging.getLogger("webapp").debug("Could not read session dealerships", exc_info=True)
    return sorted(names, key=str.casefold)


def ensure_contacts_file(path: str | Path) -> Path:
    """Create an empty contacts.csv at `path` if nothing is there yet.

    The package used to ship a header-only contacts.csv, which made reinstalling over a
    copy already in use a way to silently destroy every contact the operator had
    entered. Nothing is shipped now; the file is written the first time the app looks
    for it, so a fresh install still finds one and an existing install keeps its own.
    """
    resolved = Path(path).expanduser()
    if not resolved.exists():
        try:
            email_builder.save_contact_rows(resolved, [])
        except OSError:
            # A read-only or unreachable location is the operator's to fix. The panel
            # already reports a contacts file it cannot read, and guessing is worse.
            pass
    return resolved


def _contacts_revision(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else ""


def _contacts_path(path: str | None) -> Path:
    raw = str(path or "").strip()
    if not raw:
        raw = str(settings_manager.load_settings().get("recipient_mapping_path") or "").strip()
    if not raw:
        raw = str(ROOT / "data" / "contacts.csv")
    return Path(raw).expanduser().resolve()
