from __future__ import annotations

"""Reporter engine and session API.

This module encapsulates the boundary between the core reporting engine,
email template rendering, and desktop session state.

Responsibilities:
- Run the reporting pipeline and build a session from core output
- Expose dealership result rows for UI queue rendering
- Maintain editable session state for email previews and manual edits
- Provide reusable preview, draft, validation, and approval APIs
- Keep core reporting and settings separate from UI behavior
"""

import csv
import json
import logging
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable
import shutil

import dealership_insights_local as reporting_core
import email_assembler
import email_builder
import email_export
import sent_email_archive
import settings_manager
from settings_manager import load_settings


EXECUTIVE_SUMMARY_FIELD_MAP = {
    "overall month rating": "overall_month_rating",
    "primary strength": "primary_strength",
    "primary risk": "primary_risk",
    "recommended focus": "recommended_focus",
}


@dataclass
class DealershipSessionState:
    result: reporting_core.DealershipResult
    executive_summary_text: str
    priority_insights_text: str
    summary_text: str
    sms_text: str
    to_recipients: str = ""
    cc_recipients: str = ""
    bcc_recipients: str = ""
    contact_name: str = ""
    contact_greeting_name: str = ""
    contact_matched: bool = False
    contact_row_active: bool = False
    contact_ready: bool = False
    approved: bool = False
    skipped: bool = False
    email_preview: dict[str, Any] | None = None
    working_email: WorkingEmailState | None = None
    draft_created: bool = False
    email_sent: bool = False
    # Written to disk as .eml for manual import, which is a delivery route of its
    # own — tracked separately so "sent" never covers work the operator still has
    # to carry into Outlook themselves.
    eml_exported: bool = False
    # Operator has ticked this email as ready to send on the send queue.
    ready_to_send: bool = False


@dataclass
class WorkingEmailState:
    to: str
    cc: str
    subject: str
    body: str
    body_text: str
    html_body: str
    preview_text: str
    attachments: list[Path]
    template_type: str
    template_path: str
    template_encoding: str
    placeholders_found: list[str]
    replaced_placeholders: list[str]
    missing_placeholders: list[str]
    contact: dict[str, Any]
    report_reference: dict[str, Any]
    email_context: dict[str, str]
    bcc: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
            "subject": self.subject,
            "body": self.body,
            "body_text": self.body_text,
            "html_body": self.html_body,
            "preview_text": self.preview_text,
            "attachments": [str(path) for path in self.attachments],
            "template_type": self.template_type,
            "template_path": self.template_path,
            "template_encoding": self.template_encoding,
            "placeholders_found": list(self.placeholders_found),
            "replaced_placeholders": list(self.replaced_placeholders),
            "missing_placeholders": list(self.missing_placeholders),
            "contact": dict(self.contact),
            "report_reference": dict(self.report_reference),
            "email_context": dict(self.email_context),
        }


@dataclass
class ReportingSession:
    accounts_folder: Path
    output_folder: Path
    template_path: Path | None
    recipient_mapping_path: Path | None
    month: int
    year: int
    month_label: str
    artifacts: reporting_core.ReportingRunArtifacts
    recipient_mapping: dict[str, dict[str, str]]
    dealerships: dict[str, DealershipSessionState]
    email_template_assets_folder: Path | None = None
    email_assembly_mode: str = "assembler"


@dataclass
class ProgressUpdate:
    phase: str
    message: str
    processed_count: int | None = None
    total_count: int | None = None
    dealership_name: str | None = None


CURRENT_SESSION: ReportingSession | None = None
# On a server each signed-in user has their own run; CURRENT_SESSION is the desktop's.
_USER_SESSIONS: dict[str, ReportingSession] = {}


def _active_session() -> ReportingSession | None:
    user = settings_manager.current_user()
    return _USER_SESSIONS.get(user) if user else CURRENT_SESSION


def _set_active_session(session: ReportingSession | None) -> None:
    global CURRENT_SESSION
    user = settings_manager.current_user()
    if not user:
        CURRENT_SESSION = session
    elif session is None:
        _USER_SESSIONS.pop(user, None)
    else:
        _USER_SESSIONS[user] = session

# These diagnostic files are rewritten in full on every append, so an uncapped
# history makes each write progressively slower (one had already grown past
# 700KB in normal use). Keep only the most recent records.
MAX_DIAGNOSTIC_RECORDS = 500
EVIDENCE_JSON_PATH = Path(__file__).resolve().parent / "logs" / "email_workflow_upgrade_evidence.json"
CONTENT_MAPPING_EVIDENCE_PATH = Path(__file__).resolve().parent / "logs" / "email_build_content_mapping_evidence.json"
PIPELINE_SOURCE_OF_TRUTH_PATH = Path(__file__).resolve().parent / "logs" / "email_pipeline_source_of_truth.json"
CONTACT_EMAIL_WIRING_PATH = Path(__file__).resolve().parent / "logs" / "contact_email_wiring_validation.json"
OUTLOOK_DRAFT_RENDER_TRACE_PATH = Path(__file__).resolve().parent / "logs" / "outlook_draft_render_trace.json"


def month_label_slug(month_label: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", month_label.strip().lower())
    return slug.strip("_") or "month"


def append_outlook_draft_render_trace(record: dict[str, Any]) -> None:
    payload = {"timestamp": datetime.utcnow().isoformat() + "Z"}
    payload.update(record)
    _append_diagnostic(OUTLOOK_DRAFT_RENDER_TRACE_PATH, payload)


# Two users' requests can append at once; read-modify-write would drop one record.
_DIAGNOSTICS_LOCK = threading.Lock()


def _append_diagnostic(path: Path, payload: dict[str, Any]) -> None:
    try:
        with _DIAGNOSTICS_LOCK:
            path.parent.mkdir(parents=True, exist_ok=True)
            existing: list[Any] = []
            if path.exists():
                try:
                    loaded = json.loads(path.read_text(encoding="utf-8"))
                    if isinstance(loaded, list):
                        existing = loaded
                except (json.JSONDecodeError, OSError):
                    existing = []
            existing.append(payload)
            existing = existing[-MAX_DIAGNOSTIC_RECORDS:]
            path.write_text(json.dumps(existing, indent=2), encoding="utf-8")
    except Exception:  # noqa: BLE001 - diagnostics must never fail the operation they describe
        pass


def append_contact_email_wiring_validation(record: dict[str, Any]) -> None:
    payload = {"timestamp": datetime.utcnow().isoformat() + "Z"}
    payload.update(record)
    _append_diagnostic(CONTACT_EMAIL_WIRING_PATH, payload)


def reload_session_recipient_mapping(mapping_path: str | Path | None = None) -> None:
    """Reload contacts from disk into the active session and clear stale email previews."""
    session = _ensure_session()
    path = mapping_path if mapping_path is not None else session.recipient_mapping_path
    if path is None:
        return
    resolved_path = Path(path).expanduser().resolve()
    if not resolved_path.exists():
        return
    session.recipient_mapping = email_builder.load_recipient_mappings(resolved_path)
    session.recipient_mapping_path = resolved_path
    for st in session.dealerships.values():
        r = email_builder.resolve_contact_for_dealership(st.result.dealership_name, session.recipient_mapping)
        # Only rebuild the dealerships whose contact details actually moved.
        # This used to clear every preview and working email unconditionally,
        # so editing one contact row silently threw away every manual subject
        # and body edit made across the whole run.
        changed = (
            st.to_recipients != r.to
            or st.cc_recipients != r.cc
            or st.bcc_recipients != _bcc_with_defaults(r.bcc)
            or st.contact_greeting_name != r.greeting_name
            or st.contact_ready != r.ready
        )
        st.to_recipients = r.to
        st.cc_recipients = r.cc
        st.bcc_recipients = _bcc_with_defaults(r.bcc)
        st.contact_name = r.contact_name
        st.contact_greeting_name = r.greeting_name
        st.contact_matched = r.matched
        st.contact_row_active = r.row_active
        st.contact_ready = r.ready
        if changed:
            st.email_preview = None
            st.working_email = None


def append_pipeline_source_of_truth(record: dict[str, Any]) -> None:
    payload = {"timestamp": datetime.utcnow().isoformat() + "Z"}
    payload.update(record)
    _append_diagnostic(PIPELINE_SOURCE_OF_TRUTH_PATH, payload)


def append_content_mapping_evidence(record: dict[str, Any]) -> None:
    payload = {"timestamp": datetime.utcnow().isoformat() + "Z"}
    payload.update(record)
    _append_diagnostic(CONTENT_MAPPING_EVIDENCE_PATH, payload)


def append_workflow_evidence_record(event_type: str, details: dict[str, Any]) -> None:
    record: dict[str, Any] = {"timestamp": datetime.utcnow().isoformat() + "Z", "event": event_type}
    record.update(details)
    _append_diagnostic(EVIDENCE_JSON_PATH, record)


def _record_email_preview_evidence(
    dealership_name: str,
    preview: dict[str, Any],
    *,
    contact_ready: bool,
    email_context: dict[str, str],
) -> None:
    append_workflow_evidence_record(
        "email_preview_built",
        {
            "dealership": dealership_name,
            "template_path": preview.get("template_path", ""),
            "template_type": preview.get("template_type", ""),
            "placeholders_found": preview.get("placeholders_found", []),
            "replaced_placeholders": preview.get("replaced_placeholders", []),
            "missing_placeholders": preview.get("missing_placeholders", []),
            "subject": preview.get("subject", ""),
            "to": preview.get("to", ""),
            "cc": preview.get("cc", ""),
            "contact_ready": contact_ready,
            "template_variables_used": sorted(email_context.keys()),
            "rendered_plain_body_length": len(str(preview.get("preview_text", ""))),
            "html_body_length": len(str(preview.get("html_body", ""))),
            "attachment_count": len(preview.get("attachments", []) or []),
        },
    )


class SessionNotInitializedError(RuntimeError):
    pass


class DealershipNotInSessionError(KeyError):
    """Raised when a dealership name is not part of the current reporting run.

    Dealership names legitimately change between runs: a folder holding several
    entities (e.g. "Demo Site 1") is reported per-entity ("Demo Site 1 North",
    "Demo Site 1 South") when its source PDFs are found, but collapses to the
    folder name when they are not. A name held over from a previous run must
    therefore fail with an explanation rather than a bare KeyError.
    """

    def __init__(self, dealership_name: str) -> None:
        self.dealership_name = dealership_name
        super().__init__(
            f"'{dealership_name}' is not part of the current reporting run. "
            "Choose a dealership from the queue, or run reporting again."
        )

    def __str__(self) -> str:  # KeyError's default __str__ would re-quote the message
        return self.args[0] if self.args else ""


def has_dealership(dealership_name: str) -> bool:
    """True when the name belongs to the active session."""
    session = _active_session()
    return session is not None and dealership_name in session.dealerships


def _dealership_state(session: ReportingSession, dealership_name: str) -> DealershipSessionState:
    try:
        return session.dealerships[dealership_name]
    except KeyError:
        raise DealershipNotInSessionError(dealership_name) from None



def default_month_label() -> str:
    _month, _year, month_label = reporting_core.previous_month()
    return month_label


_MONTH_TOKEN_RE = re.compile(
    rf"({reporting_core.MONTH_NAME_PATTERN})\.?[\s_\-]+((?:19|20)\d{{2}})",
    re.IGNORECASE,
)


def detect_available_report_months(
    accounts_folder: str | Path | None,
    limit: int = 36,
    include: Iterable[str] | None = None,
) -> list[str]:
    """Scan the accounts folder for 'Month YYYY' tokens in PDF filenames.

    Used to tell the operator which months actually have source reports when a
    run produced nothing but placeholders — almost always because the wrong
    month is selected. Only called on that failure path, so the walk cost is
    acceptable; it is still bounded in case the folder is very large.
    """
    if not accounts_folder:
        return []
    root = Path(accounts_folder).expanduser()
    if not root.is_dir():
        return []
    # Workbooks say which month they hold, so ask them before scanning filenames.
    try:
        import excel_reader

        months = excel_reader.available_months(root, include=include)
        if months:
            return months[:limit]
    except Exception:  # noqa: BLE001 - fall through to the filename scan
        LOGGER.info("Workbook month detection unavailable; scanning filenames instead.")

    found: set[tuple[int, int]] = set()
    scanned = 0
    try:
        for pdf in root.rglob("*.pdf"):
            scanned += 1
            if scanned > 8000:
                break
            for name, year in _MONTH_TOKEN_RE.findall(pdf.stem):
                month_number = reporting_core.month_from_name(name)
                if month_number:
                    found.add((int(year), month_number))
    except OSError:
        return []
    ordered = sorted(found, reverse=True)[:limit]
    return [datetime(year=year, month=month, day=1).strftime("%b %Y") for year, month in ordered]



def available_month_options(month_count: int = 18) -> list[str]:
    options: list[str] = []
    reference = datetime.now().replace(day=1)
    for offset in range(-2, month_count):
        month, year, month_label = reporting_core.shift_month(reference.month, reference.year, -offset)
        options.append(datetime(year=year, month=month, day=1).strftime("%b %Y"))
    deduped: list[str] = []
    for option in options:
        if option not in deduped:
            deduped.append(option)
    return deduped



def build_executive_summary_text(result: reporting_core.DealershipResult) -> str:
    return "\n".join(
        [
            f"Overall Month Rating: {result.overall_month_rating}",
            f"Primary Strength: {result.primary_strength}",
            f"Primary Risk: {result.primary_risk}",
        ]
    )



def _apply_executive_summary_text(result: reporting_core.DealershipResult, executive_summary_text: str) -> None:
    values = {
        "overall_month_rating": result.overall_month_rating,
        "primary_strength": result.primary_strength,
        "primary_risk": result.primary_risk,
        "recommended_focus": result.recommended_focus,
    }
    for line in executive_summary_text.splitlines():
        if ":" not in line:
            continue
        label, value = line.split(":", 1)
        attribute_name = EXECUTIVE_SUMMARY_FIELD_MAP.get(label.strip().lower())
        if attribute_name:
            values[attribute_name] = value.strip()

    result.overall_month_rating = values["overall_month_rating"]
    result.primary_strength = values["primary_strength"]
    result.primary_risk = values["primary_risk"]
    result.recommended_focus = values["recommended_focus"]



def _split_priority_insights(priority_insights_text: str) -> list[str]:
    lines = [line.rstrip() for line in priority_insights_text.splitlines() if line.strip()]
    return [reporting_core.sanitise_priority_insight_text(line) for line in lines]



def _bcc_with_defaults(row_bcc: str) -> str:
    """A contact's own BCC plus the standing BCC list from settings.

    Merged in one helper and applied at every point that sets state.bcc_recipients,
    because reload_session_recipient_mapping compares the resolved value against the
    stored one to decide what to rebuild — apply the defaults in only some of those
    places and every dealership looks changed on every contacts save.
    """
    default = str(load_settings().get("default_bcc") or "").strip()
    if not default:
        return row_bcc
    return email_builder.join_addresses(
        email_builder.split_addresses(row_bcc) + email_builder.split_addresses(default)
    )


def _sender_identity() -> tuple[str, str]:
    """Who signs the email, from this operator's settings.

    Both blank is the shipped default and means no sign-off block at all, so a
    new install never signs mail with whoever built the package.
    """
    settings = load_settings()
    return (
        str(settings.get("sender_name") or "").strip(),
        str(settings.get("sender_title") or "").strip(),
    )


def _attachments_for_result(result: reporting_core.DealershipResult) -> list[Path]:
    attachments: list[Path] = []
    for pdf_path in [result.sales_pdf, result.service_pdf]:
        if pdf_path is None:
            continue
        resolved = Path(pdf_path)
        if resolved.exists():
            attachments.append(resolved)
    return attachments


def _report_reference_payload(state: DealershipSessionState) -> dict[str, Any]:
    attachments = _attachments_for_result(state.result)
    return {
        "executive_summary": state.executive_summary_text,
        "priority_insights": state.priority_insights_text,
        "summary": state.summary_text,
        "sms_version": state.sms_text,
        "sentiment": state.result.sentiment,
        "review_reason_tags": list(state.result.review_reason_tags or []),
        "source_pdf_paths": [str(path) for path in attachments],
    }


def _contact_payload(state: DealershipSessionState) -> dict[str, Any]:
    return {
        "dealership": state.result.dealership_name,
        "contact_name": state.contact_name,
        "email": state.to_recipients,
        "cc": state.cc_recipients,
        "bcc": state.bcc_recipients,
        "matched": state.contact_matched,
        "ready": state.contact_ready,
    }


def _ensure_session() -> ReportingSession:
    session = _active_session()
    if session is None:
        raise SessionNotInitializedError("Run reporting first so the desktop session has dealership results to work with.")
    return session


def has_active_session() -> bool:
    """Return True when a reporting session is loaded."""
    return _active_session() is not None


def get_current_session() -> ReportingSession:
    """Return the current active reporting session."""
    return _ensure_session()


def get_dealership_names() -> list[str]:
    """Return a sorted list of dealership names in the current session."""
    session = _ensure_session()
    return sorted(session.dealerships.keys(), key=str.lower)


def clear_session() -> None:
    """Clear the current active reporting session."""
    _set_active_session(None)


def _build_working_email_state(
    preview: dict[str, Any],
    *,
    contact: dict[str, Any],
    report_reference: dict[str, Any],
    email_context: dict[str, str],
) -> WorkingEmailState:
    attachments = [Path(path) for path in preview.get("attachments", [])]
    body_text = str(
        preview.get("plain_text_body")
        or preview.get("body_text")
        or preview.get("body")
        or "",
    )
    preview_text = str(preview.get("preview_text") or body_text)
    return WorkingEmailState(
        to=str(preview.get("to", "")),
        cc=str(preview.get("cc", "")),
        bcc=str(preview.get("bcc", "")),
        subject=str(preview.get("subject", "")),
        body=preview_text,
        body_text=body_text,
        html_body=str(preview.get("html_body", "")),
        preview_text=preview_text,
        attachments=attachments,
        template_type=str(preview.get("template_type", "")),
        template_path=str(preview.get("template_path", "")),
        template_encoding=str(preview.get("template_encoding", "")),
        placeholders_found=list(preview.get("placeholders_found", [])),
        replaced_placeholders=list(preview.get("replaced_placeholders", [])),
        missing_placeholders=list(preview.get("missing_placeholders", [])),
        contact=dict(contact),
        report_reference=dict(report_reference),
        email_context=dict(email_context),
    )



def _save_updated_outputs(session: ReportingSession, state: DealershipSessionState) -> None:
    reporting_core.save_txt_output(session.output_folder, state.result)
    reporting_core.save_docx_output(session.output_folder, state.result)



def _build_state_for_result(
    result: reporting_core.DealershipResult,
    recipient_mapping: dict[str, dict[str, str]],
) -> DealershipSessionState:
    resolved = email_builder.resolve_contact_for_dealership(result.dealership_name, recipient_mapping)
    return DealershipSessionState(
        result=result,
        executive_summary_text=build_executive_summary_text(result),
        priority_insights_text="\n".join(result.bullet_insights),
        summary_text=result.summary,
        sms_text=result.sms_version,
        to_recipients=resolved.to,
        cc_recipients=resolved.cc,
        bcc_recipients=_bcc_with_defaults(resolved.bcc),
        contact_name=resolved.contact_name,
        contact_greeting_name=resolved.greeting_name,
        contact_matched=resolved.matched,
        contact_row_active=resolved.row_active,
        contact_ready=resolved.ready,
    )



def run_reporting(
    accounts_folder: str | Path,
    output_folder: str | Path,
    month: str,
    template_path: str | Path | None,
    recipient_mapping_path: str | Path | None = None,
    progress_callback: Callable[[ProgressUpdate], None] | None = None,
    portfolio: Iterable[str] | None = None,
    pdf_folder: str | Path | None = None,
) -> ReportingSession:
    month_number, year_number, month_label = reporting_core.parse_month_label(month)
    accounts_path = Path(accounts_folder).expanduser().resolve()
    pdf_path = Path(pdf_folder).expanduser().resolve() if pdf_folder else None
    output_path = Path(output_folder).expanduser().resolve()
    # Keep accepting template_path for older callers, but never use a user-supplied
    # email template. The branded Jinja templates and assets ship with the app.
    template = None
    email_assets_folder = (Path(__file__).resolve().parent / "email_assets").resolve()
    assembly_mode = "assembler"
    logging.getLogger("report_engine").info("Using the bundled email templates and assets.")

    recipient_mapping: dict[str, dict[str, str]] = {}
    recipient_mapping_resolved: Path | None = None
    if recipient_mapping_path:
        recipient_mapping = email_builder.load_recipient_mappings(recipient_mapping_path)
        recipient_mapping_resolved = Path(recipient_mapping_path).expanduser().resolve()
    else:
        data_contacts = Path(__file__).resolve().parent / "data" / "contacts.csv"
        root_contacts = Path(__file__).resolve().parent / "contacts.csv"
        if data_contacts.exists():
            recipient_mapping = email_builder.load_recipient_mappings(data_contacts)
            recipient_mapping_resolved = data_contacts
        elif root_contacts.exists():
            recipient_mapping = email_builder.load_recipient_mappings(root_contacts)
            recipient_mapping_resolved = root_contacts

    def _handle_progress(progress_payload: dict[str, Any]) -> None:
        if progress_callback is None:
            return
        progress_callback(
            ProgressUpdate(
                phase=str(progress_payload.get("phase", "working")),
                message=str(progress_payload.get("message", "Working...")),
                processed_count=progress_payload.get("processed_count"),
                total_count=progress_payload.get("total_count"),
                dealership_name=progress_payload.get("dealership_name"),
            )
        )

    artifacts = reporting_core.run_reporting_pipeline(
        accounts_path,
        output_path,
        month_number,
        year_number,
        month_label=month_label,
        progress_callback=_handle_progress,
        include_workbooks=portfolio,
        pdf_dir=pdf_path,
    )

    dealerships = {
        result.dealership_name: _build_state_for_result(result, recipient_mapping)
        for result in artifacts.results
    }

    session = ReportingSession(
        accounts_folder=accounts_path,
        output_folder=output_path,
        template_path=template,
        recipient_mapping_path=recipient_mapping_resolved,
        month=month_number,
        year=year_number,
        month_label=month_label,
        artifacts=artifacts,
        recipient_mapping=recipient_mapping,
        dealerships=dealerships,
        email_template_assets_folder=email_assets_folder,
        email_assembly_mode=assembly_mode,
    )
    _set_active_session(session)
    matched = sum(1 for s in dealerships.values() if s.contact_matched)
    missing = sum(1 for s in dealerships.values() if not s.contact_ready)
    append_workflow_evidence_record(
        "reporting_session_started",
        {
            "month": month_label,
            "dealerships_processed": sorted(dealerships.keys()),
            "dealership_count": len(dealerships),
            "contacts_matched": matched,
            "contacts_missing_or_not_ready": missing,
            "contacts_file": str(recipient_mapping_resolved) if recipient_mapping_resolved else "",
        },
    )
    return session



def get_dealership_results() -> list[dict[str, Any]]:
    session = _ensure_session()
    rows: list[dict[str, Any]] = []
    for dealership_name, state in sorted(session.dealerships.items(), key=lambda item: item[0].lower()):
        attachments = _attachments_for_result(state.result)
        rows.append(
            {
                "dealership": dealership_name,
                "month": session.month_label,
                "report_status": state.result.report_status,
                "review_required": state.result.review_required,
                "review_severity": state.result.review_severity,
                "contact_ready": "yes" if state.contact_ready else "no",
                "contact_queue": "yes" if (state.contact_matched and state.contact_row_active) else "no",
                "email_addr_ok": "yes" if state.contact_ready else "no",
                "email_ready": "yes" if state.email_preview else "no",
                "attachments_found": len(attachments),
                "approved": state.approved,
                "skipped": state.skipped,
                "sms_version": state.sms_text,
                # Authoritative "can this be approved/sent" flag, taken from the
                # same fields the send-blocking paths use rather than the
                # derived review_severity, which may be unset.
                "approvable": not (
                    state.result.parsing_status == "fallback"
                    or state.result.insights_source == "placeholder"
                ),
            }
        )
    return rows



def _rendered_email_plain_and_html(state: DealershipSessionState) -> tuple[str, str]:
    """Bodies from email_preview snapshot only (never report_reference or session insight fields)."""
    preview = state.email_preview or {}
    plain = str(
        preview.get("plain_text_body")
        or preview.get("body_text")
        or preview.get("preview_text")
        or preview.get("body")
        or "",
    )
    html_out = str(preview.get("html_body") or "")
    if plain.strip() or html_out.strip():
        return plain, html_out
    if state.working_email is not None:
        we = state.working_email
        plain = str(we.body_text or we.preview_text or we.body or "")
        if not html_out.strip():
            html_out = str(we.html_body or "")
    return plain, html_out


def _sync_email_preview_from_working(state: DealershipSessionState) -> None:
    """Keep email_preview aligned with working_email after explicit user edits (To/CC/subject/body)."""
    if state.email_preview is None or state.working_email is None:
        return
    we = state.working_email
    snap = state.email_preview
    snap["to"] = we.to
    snap["cc"] = we.cc
    snap["subject"] = we.subject
    snap["plain_text_body"] = we.preview_text
    snap["body"] = we.preview_text
    snap["body_text"] = we.body_text
    snap["preview_text"] = we.preview_text
    snap["html_body"] = we.html_body
    snap["attachments"] = [str(path) for path in we.attachments]


def get_dealership_source_documents(dealership_name: str) -> dict[str, str]:
    """Return the sales/service source PDF paths for a dealership, individually
    (as opposed to get_dealership_report_text's combined attachments list), so
    the UI can label and open each one separately. Empty string when the PDF
    was never found or no longer exists on disk."""
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)
    result = state.result

    def _existing_path(candidate: Any) -> str:
        if not candidate:
            return ""
        resolved = Path(candidate)
        return str(resolved) if resolved.exists() else ""

    return {
        # The report Reporter builds from the workbook comes first: it is the document
        # the operator reviews. The spreadsheet's own exports stay available beside it
        # for anyone who wants to check a figure against its source.
        "performance_pdf": _existing_path(getattr(result, "performance_pdf", None)),
        "performance_note": str(getattr(result, "performance_note", "") or ""),
        "sales_pdf": _existing_path(result.sales_pdf),
        "service_pdf": _existing_path(result.service_pdf),
    }


def get_dealership_report_text(dealership_name: str) -> dict[str, Any]:
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)
    attachments = _attachments_for_result(state.result)
    preview = state.email_preview or {}
    working_source = state.working_email.as_dict() if state.working_email else preview

    if session.email_assembly_mode != "legacy_msg":
        plain_text_body = str(
            preview.get("plain_text_body")
            or preview.get("preview_text")
            or preview.get("body_text")
            or "",
        )
        html_body_rendered = str(preview.get("html_body") or "")
    else:
        plain_text_body, html_body_rendered = _rendered_email_plain_and_html(state)

    if state.working_email is not None:
        final_email = dict(state.working_email.as_dict())
        final_email["plain_text_body"] = plain_text_body
        final_email["body_text"] = plain_text_body
        final_email["preview_text"] = plain_text_body
        final_email["html_body"] = html_body_rendered
    else:
        final_email = {
            "to": state.to_recipients,
            "cc": state.cc_recipients,
            "bcc": state.bcc_recipients,
            "subject": str(preview.get("subject", "")),
            "plain_text_body": plain_text_body,
            "body": plain_text_body,
            "body_text": plain_text_body,
            "html_body": html_body_rendered,
            "preview_text": plain_text_body,
            "attachments": [str(path) for path in attachments],
            "template_type": str(preview.get("template_type", "")),
            "template_path": str(preview.get("template_path", "")),
            "template_encoding": str(preview.get("template_encoding", "")),
            "placeholders_found": list(preview.get("placeholders_found", [])),
            "replaced_placeholders": list(preview.get("replaced_placeholders", [])),
            "missing_placeholders": list(preview.get("missing_placeholders", [])),
            "contact": _contact_payload(state),
            "report_reference": _report_reference_payload(state),
            "email_context": dict(preview.get("email_context", {})),
        }

    return {
        "dealership": state.result.dealership_name,
        "month": state.result.month_label,
        "executive_summary": state.executive_summary_text,
        "priority_insights": state.priority_insights_text,
        "summary": state.summary_text,
        "sms_version": state.sms_text,
        "sentiment": state.result.sentiment,
        "review_reason_tags": state.result.review_reason_tags or [],
        "review_severity": state.result.review_severity,
        "blocked": bool(preview.get("blocked")) or state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder",
        "source_pdf_paths": [str(path) for path in attachments],
        "to": state.to_recipients,
        "cc": state.cc_recipients,
        "bcc": state.bcc_recipients,
        "contact_ready": state.contact_ready,
        "contact_matched": state.contact_matched,
        "contact_row_active": state.contact_row_active,
        "subject": str(working_source.get("subject", "")),
        "plain_text_body": plain_text_body,
        "email_body": plain_text_body,
        "body_text": plain_text_body,
        "html_body": html_body_rendered,
        "attachments": [str(path) for path in working_source.get("attachments", attachments)],
        "template_type": str(working_source.get("template_type", preview.get("template_type", ""))),
        "template_path": str(working_source.get("template_path", preview.get("template_path", ""))),
        "template_encoding": str(working_source.get("template_encoding", preview.get("template_encoding", ""))),
        "placeholders_found": list(working_source.get("placeholders_found", preview.get("placeholders_found", []))),
        "replaced_placeholders": list(working_source.get("replaced_placeholders", preview.get("replaced_placeholders", []))),
        "missing_placeholders": list(working_source.get("missing_placeholders", preview.get("missing_placeholders", []))),
        "working_email_ready": state.working_email is not None,
        "approved": state.approved,
        "skipped": state.skipped,
        "final_email": final_email,
        "email_context": dict(final_email.get("email_context") or {}),
    }



def update_dealership_text(
    dealership_name: str,
    *,
    executive_summary: str | None = None,
    priority_insights: str | None = None,
    summary: str | None = None,
    sms_version: str | None = None,
    to_recipients: str | None = None,
    cc_recipients: str | None = None,
    bcc_recipients: str | None = None,
    approved: bool | None = None,
    skipped: bool | None = None,
) -> dict[str, Any]:
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)
    invalidate_working_email = False

    if executive_summary is not None:
        executive_summary_value = executive_summary.strip()
        if executive_summary_value != state.executive_summary_text:
            state.executive_summary_text = executive_summary_value
            _apply_executive_summary_text(state.result, state.executive_summary_text)
            invalidate_working_email = True
    if priority_insights is not None:
        priority_insights_value = priority_insights.strip()
        if priority_insights_value != state.priority_insights_text:
            state.priority_insights_text = priority_insights_value
            state.result.bullet_insights = _split_priority_insights(state.priority_insights_text)
            invalidate_working_email = True
    if summary is not None:
        summary_value = summary.strip()
        if summary_value != state.summary_text:
            state.summary_text = summary_value
            state.result.summary = state.summary_text
            invalidate_working_email = True
    if sms_version is not None:
        sms_value = sms_version.strip()
        if sms_value != state.sms_text:
            state.sms_text = sms_value
            state.result.sms_version = state.sms_text
            invalidate_working_email = True
    if to_recipients is not None:
        to_value = to_recipients.strip()
        if to_value != state.to_recipients:
            state.to_recipients = to_value
            if state.working_email:
                state.working_email.to = state.to_recipients
    if cc_recipients is not None:
        cc_value = cc_recipients.strip()
        if cc_value != state.cc_recipients:
            state.cc_recipients = cc_value
            if state.working_email:
                state.working_email.cc = state.cc_recipients
    if bcc_recipients is not None:
        bcc_value = bcc_recipients.strip()
        if bcc_value != state.bcc_recipients:
            state.bcc_recipients = bcc_value
            if state.working_email:
                state.working_email.bcc = state.bcc_recipients
    if approved is not None:
        state.approved = approved
    if skipped is not None:
        state.skipped = skipped

    if invalidate_working_email:
        state.email_preview = None
        state.working_email = None

    _save_updated_outputs(session, state)
    return get_dealership_report_text(dealership_name)


def update_working_email(
    dealership_name: str,
    *,
    subject: str | None = None,
    to_recipients: str | None = None,
    cc_recipients: str | None = None,
    bcc_recipients: str | None = None,
    body: str | None = None,
) -> dict[str, Any]:
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)
    if state.working_email is None:
        build_email_preview(dealership_name)

    if state.working_email is None:
        raise RuntimeError("Build an email preview before editing the working email.")

    # The UI flushes whatever is currently in its entry boxes back through here
    # whenever the operator clicks a different row. When a dealership was on
    # screen *before* its email finished building, those boxes are still empty,
    # and that empty flush used to wipe the freshly built email — leaving a
    # blank preview and no subject for the first dealership in the run. An
    # empty field can never overwrite content that already exists; clearing a
    # field is not an edit anyone performs deliberately.
    def _would_blank(new_value: str | None, existing: str) -> bool:
        return new_value is not None and not new_value.strip() and bool((existing or "").strip())

    _log = logging.getLogger("report_engine")
    if _would_blank(subject, state.working_email.subject):
        _log.info("Ignored empty subject flush for %s | built subject preserved", dealership_name)
        subject = None
    if _would_blank(body, state.working_email.body_text):
        _log.info("Ignored empty body flush for %s | built HTML preserved", dealership_name)
        body = None

    if subject is not None:
        state.working_email.subject = subject
    if to_recipients is not None:
        state.to_recipients = to_recipients.strip()
        state.working_email.to = state.to_recipients
        c = dict(state.working_email.contact)
        c["email"] = state.to_recipients
        state.working_email.contact = c
    if cc_recipients is not None:
        state.cc_recipients = cc_recipients.strip()
        state.working_email.cc = state.cc_recipients
    if bcc_recipients is not None:
        state.bcc_recipients = bcc_recipients.strip()
        state.working_email.bcc = state.bcc_recipients
        c = dict(state.working_email.contact)
        c["cc"] = state.cc_recipients
        state.working_email.contact = c
    if body is not None:
        preserved_body = body.replace("\r\n", "\n").replace("\r", "\n")
        # Only rebuild the HTML when the text actually changed. The UI flushes
        # the current body back through here whenever the operator merely clicks
        # a different dealership, and converting that unchanged text would
        # replace the branded HTML (banners, button, inline images) with a plain
        # conversion — silently destroying the email's layout on every click.
        if preserved_body != state.working_email.body_text:
            state.working_email.body = preserved_body
            state.working_email.body_text = preserved_body
            state.working_email.preview_text = preserved_body
            state.working_email.html_body = email_builder.convert_text_to_html(preserved_body)
            if state.email_preview is not None:
                state.email_preview["html_body_outlook"] = ""
                state.email_preview["assembler_inline_assets"] = []
            logging.getLogger("report_engine").info(
                "Body edited for %s | branded HTML replaced by plain-text conversion", dealership_name
            )

    _sync_email_preview_from_working(state)

    return get_dealership_report_text(dealership_name)



def _coalesce_priority_insights_raw(state: DealershipSessionState) -> Any:
    """Prefer session text; fall back to core bullet_insights; then report_reference keys."""
    text = (state.priority_insights_text or "").strip()
    if text:
        return state.priority_insights_text
    bullets = getattr(state.result, "bullet_insights", None) or []
    if bullets:
        return list(bullets)
    ref = _report_reference_payload(state)
    for key in ("priority_insights", "Priority Insights", "Priority insights"):
        v = ref.get(key)
        if v is not None and str(v).strip():
            return v
    return ""


def build_email_preview(
    dealership_name: str,
    manual_to: str | None = None,
    manual_cc: str | None = None,
    manual_bcc: str | None = None,
) -> dict[str, Any]:
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)

    if manual_to is not None or manual_cc is not None or manual_bcc is not None:
        resolved = email_builder.resolve_contact_for_dealership(
            dealership_name,
            session.recipient_mapping,
            manual_to=manual_to or "",
            manual_cc=manual_cc or "",
            manual_bcc=manual_bcc or "",
        )
    else:
        resolved = email_builder.resolve_contact_for_dealership(
            dealership_name,
            session.recipient_mapping,
        )
    state.to_recipients = resolved.to
    state.cc_recipients = resolved.cc
    state.bcc_recipients = _bcc_with_defaults(resolved.bcc)
    state.contact_name = resolved.contact_name
    state.contact_greeting_name = resolved.greeting_name
    state.contact_matched = resolved.matched
    state.contact_row_active = resolved.row_active
    state.contact_ready = resolved.ready

    sales_p = str(state.result.sales_pdf) if state.result.sales_pdf else ""
    service_p = str(state.result.service_pdf) if state.result.service_pdf else ""
    greeting = (resolved.greeting_name.strip() or state.result.dealership_name.strip())

    ref_for_mapping = _report_reference_payload(state)
    pi_raw = _coalesce_priority_insights_raw(state)
    summary_resolved = email_builder.resolve_summary_for_email(
        summary=state.summary_text or "",
        report_summary=getattr(state.result, "summary", "") or "",
        executive_summary=state.executive_summary_text or "",
        report_reference=ref_for_mapping,
    )
    pi_extracted = email_builder.extract_priority_insight_lines(pi_raw)
    logging.getLogger("report_engine").info(
        "Email content mapping | dealership=%r | pi_type=%s | pi_len=%s | pi_lines=%s | summary_len=%s | raw_priority_insights=%r",
        state.result.dealership_name,
        type(pi_raw).__name__,
        len(pi_extracted),
        pi_extracted,
        len(summary_resolved),
        pi_raw,
    )
    logging.getLogger("report_engine").info(
        "Email content mapping | dealership=%r | summary_preview=%r",
        state.result.dealership_name,
        summary_resolved[:500] if len(summary_resolved) > 500 else summary_resolved,
    )

    email_context = email_builder.build_email_context_object(
        dealership_name=state.result.dealership_name,
        greeting_name=greeting,
        month_label=state.result.month_label,
        priority_insights=pi_raw,
        summary=summary_resolved,
        sales_report_path=sales_p,
        service_report_path=service_p,
        email=state.to_recipients,
        cc=state.cc_recipients,
    )

    _log = logging.getLogger("report_engine")

    _log.info("Using bundled EmailAssembler")
    assembled = email_assembler.assemble_dealership_email(
        dealership_name=state.result.dealership_name,
        month_label=state.result.month_label,
        greeting_name=greeting,
        priority_insights=pi_raw,
        summary_resolved=summary_resolved,
        to_recipients=state.to_recipients,
        cc_recipients=state.cc_recipients,
        pdf_attachments=_attachments_for_result(state.result),
        report_reference=ref_for_mapping,
        assets_dir=session.email_template_assets_folder,
        sales_report_path=sales_p,
        service_report_path=service_p,
        sender_name=_sender_identity()[0],
        sender_title=_sender_identity()[1],
    )
    preview = email_assembler.assembled_email_to_preview_dict(
        assembled,
        to_recipients=state.to_recipients,
        cc_recipients=state.cc_recipients,
        bcc_recipients=state.bcc_recipients,
    )
    email_context = dict(assembled.email_context)
    preview["email_context"] = email_context
    preview["plain_text_body"] = assembled.plain_text_body
    preview["html_body"] = assembled.html_body
    preview["html_body_outlook"] = assembled.html_body_outlook
    preview["subject"] = assembled.subject
    preview["body"] = assembled.plain_text_body
    preview["body_text"] = assembled.plain_text_body
    preview["preview_text"] = assembled.plain_text_body
    preview["pipeline_used"] = "assembler"
    preview["ui_bound_to"] = "assembler"

    pad_chk = email_builder.INSIGHT_SLOT_PAD
    plain_chk = assembled.plain_text_body
    missing_asm: list[str] = []
    for slot_name in ("insight_1", "insight_2", "insight_3", "insight_4"):
        slot_val = str(email_context.get(slot_name, ""))
        if not slot_val.strip() or slot_val == pad_chk:
            missing_asm.append(slot_name)
        elif slot_val.strip() not in plain_chk:
            missing_asm.append(slot_name)
    if missing_asm:
        _log.error(
            "ASSEMBLER OUTPUT MISSING CONTENT | dealership=%r | missing=%s",
            state.result.dealership_name,
            missing_asm,
        )

    audit, rci, rcs = email_builder.validate_rendered_email_mapping(
        preview_text=str(preview.get("preview_text", "")),
        html_body=str(preview.get("html_body", "")),
        insight_1=str(email_context.get("insight_1", "")),
        insight_2=str(email_context.get("insight_2", "")),
        insight_3=str(email_context.get("insight_3", "")),
        insight_4=str(email_context.get("insight_4", "")),
        summary_resolved=email_builder.normalize_text_for_preview(summary_resolved),
        sms_version=state.sms_text,
        omit_sms_from_email=True,
        had_raw_insight_lines=len(pi_extracted) > 0,
    )
    preview["render_audit_warnings"] = list(audit)
    preview["render_contains_insights"] = rci
    preview["render_contains_summary"] = rcs
    state.email_preview = preview
    state.working_email = _build_working_email_state(
        preview,
        contact=_contact_payload(state),
        report_reference=_report_reference_payload(state),
        email_context=email_context,
    )
    _record_email_preview_evidence(
        state.result.dealership_name,
        preview,
        contact_ready=state.contact_ready,
        email_context=email_context,
    )
    append_workflow_evidence_record(
        "jinja_email_rendered",
        {
            "dealership": state.result.dealership_name,
            "plain_body_length": len(str(preview.get("preview_text", ""))),
            "html_body_length": len(str(preview.get("html_body", ""))),
            "contact_ready": state.contact_ready,
        },
    )

    i1, i2, i3, i4 = (
        email_context.get("insight_1", ""),
        email_context.get("insight_2", ""),
        email_context.get("insight_3", ""),
        email_context.get("insight_4", ""),
    )
    pad = email_builder.INSIGHT_SLOT_PAD
    append_content_mapping_evidence(
        {
            "dealership": state.result.dealership_name,
            "has_contact": state.contact_matched,
            "contact_name_used": greeting,
            "priority_insights_raw_type": type(pi_raw).__name__,
            "priority_insights_count": len(pi_extracted),
            "insight_1_present": bool(i1.strip() and i1 != pad),
            "insight_2_present": bool(i2.strip() and i2 != pad),
            "insight_3_present": bool(i3.strip() and i3 != pad),
            "insight_4_present": bool(i4.strip() and i4 != pad),
            "summary_present": bool(summary_resolved.strip()),
            "rendered_contains_insights": bool(preview.get("render_contains_insights")),
            "rendered_contains_summary": bool(preview.get("render_contains_summary")),
            "warnings": list(preview.get("render_audit_warnings") or []),
        }
    )

    plain_for_log = str(preview.get("plain_text_body") or preview.get("preview_text") or "")
    html_for_log = str(preview.get("html_body") or "")
    pi1_log = str(email_context.get("insight_1", ""))
    sum_log = str(email_context.get("summary", "")).strip()
    pad_log = email_builder.INSIGHT_SLOT_PAD
    needs_insight_1 = bool(pi1_log.strip() and pi1_log != pad_log)
    needs_summary_in_body = bool(sum_log and sum_log != "[Missing report summary]")
    append_pipeline_source_of_truth(
        {
            "dealership": state.result.dealership_name,
            "pipeline_used": str(preview.get("pipeline_used") or "assembler"),
            "plain_body_length": len(plain_for_log),
            "contains_insights": (not needs_insight_1) or (pi1_log.strip() in plain_for_log),
            "contains_summary": (not needs_summary_in_body) or (sum_log in plain_for_log),
            "html_length": len(html_for_log),
            "ui_bound_to": str(preview.get("ui_bound_to") or "assembler"),
        }
    )

    row_w = email_builder.lookup_contact_row(state.result.dealership_name, session.recipient_mapping)
    active_parsed = email_builder.parse_active_flag(str(row_w.get("active", "yes"))) if row_w else None
    if active_parsed is None and row_w:
        active_parsed = True
    contact_active_w = bool(row_w and active_parsed)
    cname_w = str(row_w.get("contact_name", "")).strip() if row_w else ""
    email_w = str(row_w.get("email") or row_w.get("to") or "").strip() if row_w else ""
    cc_w = str(row_w.get("cc", "")).strip() if row_w else ""
    name_used_ctx = str(email_context.get("name", "")).strip()
    dealer_nm = state.result.dealership_name.strip()
    wiring_warnings: list[str] = []
    if row_w and contact_active_w and email_w and not email_builder.is_probable_email(email_w):
        wiring_warnings.append("active contact row has email that does not look valid")
    append_contact_email_wiring_validation(
        {
            "dealership": state.result.dealership_name,
            "contact_found": bool(row_w),
            "contact_active": contact_active_w,
            "contact_name": cname_w,
            "email": email_w,
            "cc": cc_w,
            "name_used_in_context": name_used_ctx,
            "to_field": state.to_recipients,
            "cc_field": state.cc_recipients,
            "greeting_contains_contact_name": bool(cname_w) and bool(cname_w in plain_for_log),
            "fallback_used": name_used_ctx == dealer_nm,
            "warnings": wiring_warnings,
        }
    )

    if state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder":
        # Source PDFs were missing this month, so everything above is generic
        # placeholder text, not real dealership figures. Overwrite the
        # deliverable content so a "successful" preview can never be mistaken
        # for a real one and accidentally drafted/sent to a client.
        block_notice = (
            "NO SOURCE DATA FOR THIS MONTH — Sales/Service PDFs were missing, so nothing in this "
            "report reflects real dealership figures. This is a placeholder only. Do not send until "
            "the source reports are available and this month is regenerated."
        )
        preview["blocked"] = True
        preview["block_reason"] = "no_source_data"
        preview["subject"] = f"[DO NOT SEND — no source data] {preview.get('subject', '')}".strip()
        preview["body"] = block_notice
        preview["body_text"] = block_notice
        preview["preview_text"] = block_notice
        preview["plain_text_body"] = block_notice
        preview["html_body"] = f"<p style='color:#c00000;font-weight:bold;'>{block_notice}</p>"
        preview["html_body_outlook"] = preview["html_body"]
        state.email_preview = preview
    else:
        preview["blocked"] = False

    return preview



def _resolve_email_body(
    state: DealershipSessionState,
) -> tuple[str, list[tuple[str, Path]] | None, str, list[str]]:
    """Pick the HTML that goes on the wire and the inline images that belong with it.

    Shared by the Outlook draft and the .eml export so the two cannot drift. An email
    saved to disk for manual sending has to be the same email that would have been
    drafted, down to which branding images are embedded and which fallback was taken.

    Returns (html, inline assets or None, which source was used, warnings).
    """
    working_email = state.working_email
    preview_snap = state.email_preview or {}
    html_outlook = str(preview_snap.get("html_body_outlook") or "").strip()
    html_preview = str(preview_snap.get("html_body") or "").strip()
    we_html = str(getattr(working_email, "html_body", "") or "").strip()

    inline_pairs: list[tuple[str, Path]] = []
    for item in preview_snap.get("assembler_inline_assets") or []:
        if isinstance(item, dict) and item.get("cid") and item.get("path"):
            inline_pairs.append((str(item["cid"]), Path(str(item["path"]))))

    warnings: list[str] = []
    if html_outlook:
        # The only branch with working images: cid: references plus the files to match.
        return html_outlook, inline_pairs, "html_body_outlook", warnings
    if html_preview:
        if inline_pairs:
            warnings.append(
                "Inline CID assets dropped: html_body_outlook empty while html_body uses file:// preview URLs.",
            )
        return html_preview, None, "html_body", warnings
    if we_html and "<html" in we_html.lower():
        return we_html, None, "html_body", warnings
    warnings.append(
        "No HTML in email snapshot; using convert_text_to_html (branded layout may be lost).",
    )
    return (
        email_builder.convert_text_to_html(getattr(working_email, "body_text", "") or ""),
        None,
        "plain_text",
        warnings,
    )


def create_outlook_draft(
    dealership_name: str,
    manual_to: str | None = None,
    manual_cc: str | None = None,
    manual_bcc: str | None = None,
    *,
    send: bool = False,
    override_to: str | None = None,
) -> dict[str, Any]:
    """Create an Outlook draft, or send it outright when send=True.

    override_to redirects the message (and clears CC and BCC) for test sends, so a
    whole batch can be checked in a real inbox without any of it reaching a dealership.
    """
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)

    if state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder":
        verb = "send" if send else "create an Outlook draft for"
        raise RuntimeError(
            f"Refusing to {verb} {dealership_name}: source Sales/Service "
            "PDFs were missing this month, so the report contains only placeholder text, not real "
            "figures. Add the missing PDFs and regenerate this month before drafting an email."
        )

    if manual_to is not None or manual_cc is not None or manual_bcc is not None:
        update_dealership_text(
            dealership_name,
            to_recipients=(manual_to if manual_to is not None else state.to_recipients),
            cc_recipients=(manual_cc if manual_cc is not None else state.cc_recipients),
            bcc_recipients=(manual_bcc if manual_bcc is not None else state.bcc_recipients),
        )

    if state.working_email is None:
        build_email_preview(
            dealership_name, manual_to=manual_to, manual_cc=manual_cc, manual_bcc=manual_bcc
        )

    if state.working_email is None:
        raise RuntimeError("Unable to prepare the working email for Outlook draft creation.")

    working_email = state.working_email
    preview_snap = state.email_preview or {}
    final_html, inline_for_draft, draft_body_source, trace_warnings = _resolve_email_body(state)
    # Kept only for the render trace, which records what was available as well as
    # what was chosen — a zero here is how a lost branded body gets diagnosed later.
    html_outlook = str(preview_snap.get("html_body_outlook") or "").strip()
    html_preview = str(preview_snap.get("html_body") or "").strip()

    plain_len = len(str(working_email.body_text or ""))
    pdf_attachment_count = sum(
        1
        for p in working_email.attachments
        if Path(p).exists() and str(p).lower().endswith(".pdf")
    )
    html_for_metrics = email_builder.ensure_outlook_html_document(final_html)
    low_m = html_for_metrics.lower()
    append_outlook_draft_render_trace(
        {
            "dealership": dealership_name,
            "subject": str(working_email.subject or ""),
            "html_preview_length": len(html_preview),
            "html_outlook_length": len(html_outlook),
            "plain_text_length": plain_len,
            "draft_body_source": draft_body_source,
            "contains_html_tag": "<html" in low_m,
            "contains_table": "<table" in low_m,
            "contains_inline_styles": "style=" in low_m,
            "contains_cid_images": "cid:" in low_m,
            "inline_asset_count": len(inline_for_draft or []),
            "pdf_attachment_count": pdf_attachment_count,
            "warnings": list(trace_warnings),
        }
    )

    redirected = bool(str(override_to or "").strip())
    send_to = str(override_to).strip() if redirected else working_email.to
    send_cc = "" if redirected else working_email.cc
    # A test send goes to one inbox and nowhere else. A BCC surviving the redirect would
    # reach a real recipient invisibly, which is the worst way to discover it was there.
    send_bcc = "" if redirected else str(getattr(working_email, "bcc", "") or "")
    subject_line = working_email.subject
    if redirected:
        subject_line = f"[TEST — would go to {working_email.to or 'no recipient'}] {subject_line}"

    draft_details = email_builder.create_outlook_draft(
        to_recipients=send_to,
        bcc_recipients=send_bcc,
        cc_recipients=send_cc,
        subject=subject_line,
        body_text=working_email.body_text,
        html_body=final_html,
        attachments=list(working_email.attachments),
        inline_cid_attachments=inline_for_draft,
        send=send,
    )
    draft_details["dealership"] = dealership_name
    draft_details["test_mode"] = redirected
    draft_details["intended_to"] = working_email.to
    draft_details["html_body"] = final_html
    # final_html references its images as cid:, which only resolves inside a
    # MIME message. The archive writes a standalone file and hands it to Word,
    # so it needs the data-URI build or every archived PDF loses its branding.
    draft_details["archive_html"] = str(preview_snap.get("html_body") or "") or final_html
    state.draft_created = True
    if send and not redirected:
        state.email_sent = True
    append_workflow_evidence_record(
        "outlook_draft_created",
        {
            "dealership": dealership_name,
            "draft_body_source": draft_body_source,
            "html_body_length": len(final_html),
            "attachment_count": len(working_email.attachments),
            "to_set": bool(str(working_email.to or "").strip()),
        },
    )
    return draft_details


def validate_template(template_path: str | Path | None) -> dict[str, Any]:
    return email_builder.validate_template_file(template_path)



def can_approve(dealership_name: str) -> tuple[bool, str]:
    """Whether a dealership's report may be approved, and why not if it can't.

    Mirrors the send-blocking rule: a report built from placeholder text because
    its source PDFs were missing must never reach the send queue.
    """
    session = _ensure_session()
    state = _dealership_state(session, dealership_name)
    if state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder":
        return False, (
            f"{dealership_name} has no source data for {session.month_label} — its report is "
            "placeholder text, not real figures. Add the missing Sales/Service PDFs and re-run "
            "reporting before approving."
        )
    return True, ""


def set_ready_to_send(dealership_name: str, ready: bool) -> None:
    session = _ensure_session()
    _dealership_state(session, dealership_name).ready_to_send = bool(ready)


def get_send_queue() -> list[dict[str, Any]]:
    """Approved dealerships, in the order they appear in the queue."""
    session = _ensure_session()
    rows: list[dict[str, Any]] = []
    for name, state in sorted(session.dealerships.items(), key=lambda item: item[0].lower()):
        if not state.approved:
            continue
        rows.append(
            {
                "dealership": name,
                "to": state.to_recipients,
                "cc": state.cc_recipients,
                "contact_ready": state.contact_ready,
                "email_ready": state.email_preview is not None,
                "ready_to_send": state.ready_to_send,
                "sent": state.email_sent,
                # Carried here so the SMS page can build its document straight
                # from the queue instead of assembling a full ~30-key report
                # dict per dealership just to read one string.
                "sms_version": state.sms_text,
            }
        )
    return rows


def source_folder_for(dealership_name: str) -> Path | None:
    """The folder this dealership's source reports came from.

    Derived from the PDF the report was parsed out of, not from the accounts folder
    plus a guessed subfolder name: the folder is often not spelled the same as the
    dealership, and every operator arranges their accounts folder differently.
    """
    session = _ensure_session()
    state = session.dealerships.get(dealership_name)
    if state is None:
        return None
    for pdf in (state.result.sales_pdf, state.result.service_pdf):
        if pdf:
            parent = Path(pdf).expanduser().parent
            if parent.is_dir():
                return parent
    return None


def archive_email_copies(
    dealership_names: list[str],
    destination_folder: str | Path = "",
) -> list[dict[str, Any]]:
    """Write a PDF copy of each dealership's email.

    With no destination_folder each copy lands beside that dealership's own source
    reports, which is where the operator already looks for that month's paperwork.
    A destination_folder overrides that for anyone whose accounts folder is read-only
    or who wants one central archive; then the old month-subfolder layout applies.
    """
    session = _ensure_session()
    override = str(destination_folder or "").strip()
    items: list[dict[str, Any]] = []
    results: list[dict[str, Any]] = []

    for name in dealership_names:
        try:
            if name not in session.dealerships:
                raise DealershipNotInSessionError(name)
            state = _dealership_state(session, name)
            if state.working_email is None:
                build_email_preview(name)
            preview = state.email_preview or {}
            # The data-URI build, never the cid: one: Word renders a standalone file,
            # and cid: only resolves inside a MIME message — archiving it produced
            # PDFs with four broken images.
            html = str(preview.get("html_body") or "")
            if not html:
                raise RuntimeError("No email has been built for this dealership yet.")

            target = override
            if not target:
                folder = source_folder_for(name)
                if folder is None:
                    raise RuntimeError(
                        "Could not tell where this dealership's source reports live, so "
                        "there is nowhere to put the copy. Set a PDF copies folder in Settings."
                    )
                target = str(folder)
            items.append({"status": "ok", "dealership": name, "archive_html": html, "destination_folder": target})
        except Exception as exc:  # noqa: BLE001 - one bad copy must not stop the rest
            logging.getLogger("report_engine").exception("Could not prepare a PDF copy for %s", name)
            results.append({"dealership": name, "status": "failed", "path": "", "format": "", "error": str(exc)})

    if items:
        archived = sent_email_archive.archive_sent_emails(
            sent_results=items,
            month_label=session.month_label,
            destination_folder=override,
            layout="month_folder" if override else "beside_source",
        )
        for row in archived:
            row["status"] = "failed" if not row.get("path") else row.get("format", "pdf")
            results.append(row)

    append_workflow_evidence_record(
        "email_pdf_copies_written",
        {
            "requested": len(dealership_names),
            "written": sum(1 for r in results if r.get("path")),
            "failed": sum(1 for r in results if not r.get("path")),
            "override_folder": override,
        },
    )
    return results


def export_emails_as_eml(
    dealership_names: list[str],
    folder: str | Path,
) -> list[dict[str, Any]]:
    """Write each approved email to disk as .eml, for importing into Outlook by hand.

    The escape hatch for a machine where Outlook automation will not run: same body,
    same branding images, same PDF attachments, but produced entirely by the standard
    library so nothing about Outlook has to be working. Independent per dealership,
    like send_emails — one bad export never costs the rest of the batch.
    """
    session = _ensure_session()
    target = Path(folder).expanduser()
    results: list[dict[str, Any]] = []

    for name in dealership_names:
        try:
            if name not in session.dealerships:
                raise DealershipNotInSessionError(name)
            state = _dealership_state(session, name)

            # Same refusal as the Outlook path. A file on disk is easier to send by
            # accident than a draft, so placeholder figures must not reach one.
            if state.result.parsing_status == "fallback" or state.result.insights_source == "placeholder":
                raise RuntimeError(
                    f"Refusing to export {name}: source Sales/Service PDFs were missing this "
                    "month, so the report contains only placeholder text, not real figures. "
                    "Add the missing PDFs and regenerate this month before exporting."
                )

            if state.working_email is None:
                build_email_preview(name)
            working_email = state.working_email
            if working_email is None:
                raise RuntimeError("Unable to prepare the working email for export.")

            final_html, inline_assets, body_source, _warnings = _resolve_email_body(state)
            exported = email_export.write_eml(
                target,
                f"{name} - {state.result.month_label}",
                subject=str(working_email.subject or ""),
                to_recipients=str(working_email.to or ""),
                cc_recipients=str(working_email.cc or ""),
                bcc_recipients=str(getattr(working_email, "bcc", "") or ""),
                body_text=str(working_email.body_text or ""),
                html_body=email_builder.ensure_outlook_html_document(final_html),
                inline_assets=list(inline_assets or []),
                attachments=list(working_email.attachments),
            )
            state.eml_exported = True
            results.append(
                {
                    "dealership": name,
                    "status": "exported",
                    "to": str(working_email.to or ""),
                    "subject": str(working_email.subject or ""),
                    "path": str(exported.path),
                    "filename": exported.path.name,
                    "attachments": list(exported.attachment_names),
                    "inline_assets": exported.inline_asset_count,
                    "bytes": exported.bytes_written,
                    "body_source": body_source,
                    "error": "",
                }
            )
        except Exception as exc:  # noqa: BLE001 - one bad export must not stop the rest
            logging.getLogger("report_engine").exception("Export failed for %s", name)
            results.append(
                {
                    "dealership": name,
                    "status": "failed",
                    "to": "",
                    "subject": "",
                    "path": "",
                    "filename": "",
                    "attachments": [],
                    "inline_assets": 0,
                    "bytes": 0,
                    "body_source": "",
                    "error": str(exc),
                }
            )

    append_workflow_evidence_record(
        "emails_exported_as_eml",
        {
            "folder": str(target),
            "requested": len(dealership_names),
            "exported": sum(1 for r in results if r["status"] == "exported"),
            "failed": sum(1 for r in results if r["status"] == "failed"),
        },
    )
    return results


def _attachments_for_name(dealership_name: str) -> list[Path]:
    """The PDFs attached to one dealership's email, for reporting back to the operator."""
    session = _ensure_session()
    state = session.dealerships.get(dealership_name)
    return _attachments_for_result(state.result) if state else []


def send_emails(
    dealership_names: list[str],
    *,
    send: bool = True,
    test_address: str | None = None,
) -> list[dict[str, Any]]:
    """Send (or draft) each email independently.

    One failure never aborts the batch: every dealership gets its own result so
    the operator can see exactly what went out and what did not.
    """
    session = _ensure_session()
    results: list[dict[str, Any]] = []
    test_to = str(test_address or "").strip()
    for name in dealership_names:
        try:
            if name not in session.dealerships:
                raise DealershipNotInSessionError(name)
            details = create_outlook_draft(name, send=send, override_to=test_to or None)
            results.append(
                {
                    "dealership": name,
                    "status": "sent" if send else "drafted",
                    "to": details.get("to", ""),
                    "subject": details.get("subject", ""),
                    "html_body": details.get("html_body", ""),
                    "archive_html": details.get("archive_html", "") or details.get("html_body", ""),
                    "attachments": [Path(p).name for p in _attachments_for_name(name)],
                    "error": "",
                }
            )
        except Exception as exc:  # noqa: BLE001 - one bad email must not stop the rest
            logging.getLogger("report_engine").exception("Send failed for %s", name)
            results.append(
                {
                    "dealership": name,
                    "status": "failed",
                    "to": "",
                    "subject": "",
                    "html_body": "",
                    "archive_html": "",
                    "attachments": [],
                    "error": str(exc),
                }
            )
    append_workflow_evidence_record(
        "emails_sent" if send else "emails_drafted",
        {
            "month": session.month_label,
            "requested": len(dealership_names),
            "succeeded": sum(1 for r in results if r["status"] != "failed"),
            "failed": sum(1 for r in results if r["status"] == "failed"),
            "test_mode": bool(test_to),
        },
    )
    return results


def mark_approved(dealership_name: str, approved: bool = True) -> dict[str, Any]:
    return update_dealership_text(dealership_name, approved=approved, skipped=False if approved else None)



def mark_skipped(dealership_name: str, skipped: bool = True) -> dict[str, Any]:
    return update_dealership_text(dealership_name, skipped=skipped, approved=False if skipped else None)


def export_sms_breakdowns_for_session() -> dict[str, Any]:
    """Write consolidated SMS rows for all dealerships in the session to the output folder."""
    session = _ensure_session()
    slug = month_label_slug(session.month_label)
    out_path = session.output_folder / f"sms_breakdowns_{slug}.csv"
    generated_at = datetime.utcnow().isoformat() + "Z"
    fieldnames = [
        "dealership",
        "month",
        "contact_name",
        "email",
        "sms_text",
        "source_sales_pdf",
        "source_service_pdf",
        "generated_at",
    ]
    rows_written = 0
    with out_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for dealership_name, state in sorted(session.dealerships.items(), key=lambda item: item[0].lower()):
            writer.writerow(
                {
                    "dealership": dealership_name,
                    "month": session.month_label,
                    "contact_name": state.contact_name,
                    "email": state.to_recipients,
                    "sms_text": state.sms_text,
                    "source_sales_pdf": str(state.result.sales_pdf or ""),
                    "source_service_pdf": str(state.result.service_pdf or ""),
                    "generated_at": generated_at,
                }
            )
            rows_written += 1
    append_workflow_evidence_record(
        "sms_export",
        {"sms_export_path": str(out_path), "rows_written": rows_written, "month": session.month_label},
    )
    return {"path": str(out_path), "rows_written": rows_written, "month": session.month_label}


def get_session_summary() -> dict[str, Any]:
    session = _ensure_session()
    return {
        "accounts_folder": str(session.accounts_folder),
        "output_folder": str(session.output_folder),
        "template_path": str(session.template_path) if session.template_path else "",
        "recipient_mapping_path": str(session.recipient_mapping_path) if session.recipient_mapping_path else "",
        "month": session.month_label,
        "dealership_count": len(session.dealerships),
    }


def export_reports_for_month(export_path: str | Path) -> dict[str, Any]:
    """Export all generated dealership reports for the current month to the specified folder."""
    session = _ensure_session()
    export_path = Path(export_path)
    export_path.mkdir(parents=True, exist_ok=True)
    
    output_folder = session.output_folder
    month_label = session.month_label
    
    exported_files = []
    for file_path in output_folder.glob("*.txt"):
        if month_label in file_path.name:
            shutil.copy2(file_path, export_path / file_path.name)
            exported_files.append(file_path.name)
    
    return {
        "exported_count": len(exported_files),
        "export_path": str(export_path),
        "exported_files": exported_files,
    }
