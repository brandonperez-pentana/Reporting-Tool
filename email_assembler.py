"""Email assembly layer: branded Jinja templates, local assets, CID/outlook + preview URLs.

Does not send mail. SMS is never included in assembled bodies.
"""
from __future__ import annotations

import contextlib
import json
import logging
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from jinja2 import Environment, StrictUndefined, TemplateError

import email_builder

PACKAGE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = PACKAGE_DIR / "templates"
HTML_TEMPLATE_PATH = TEMPLATE_DIR / "dealer_snapshot_email.html.j2"
TEXT_TEMPLATE_PATH = TEMPLATE_DIR / "dealer_snapshot_email.txt.j2"
ASSET_MANIFEST_PATH = TEMPLATE_DIR / "email_asset_manifest.json"
ASSEMBLER_VALIDATION_PATH = PACKAGE_DIR / "logs" / "email_assembler_validation.json"

# These diagnostic files are rewritten in full on every append, so an uncapped
# history makes each write progressively slower (one had already grown past
# 700KB in normal use). Keep only the most recent records.
MAX_DIAGNOSTIC_RECORDS = 500

JINJA_SLOT_NAMES = (
    "name",
    "month",
    "insight_1",
    "insight_2",
    "insight_3",
    "insight_4",
    "summary",
    "dealership",
    "sender_name",
    "sender_title",
)
JINJA_VAR_RE = re.compile(r"\{\{-?\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*-?\}\}")


@dataclass
class AssembledEmail:
    subject: str
    plain_text_body: str
    html_body: str
    html_body_outlook: str
    attachments: list[Path]
    inline_assets: list[tuple[str, Path]]
    report_reference: dict[str, Any]
    email_context: dict[str, str]
    jinja_slots_found: list[str]
    legacy_fallback_used: bool
    template_used: str
    assembler_metadata: dict[str, Any] = field(default_factory=dict)


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _load_asset_manifest(assets_dir: Path | None) -> dict[str, str]:
    """Map logical asset key -> filename. Project manifest lists keys; user dir holds files."""
    base: dict[str, str] = {}
    if ASSET_MANIFEST_PATH.exists():
        try:
            raw = json.loads(ASSET_MANIFEST_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                base = {str(k): str(v).strip() for k, v in raw.items() if str(v).strip()}
        except (json.JSONDecodeError, OSError):
            pass
    if assets_dir and assets_dir.is_dir():
        local_manifest = assets_dir / "email_asset_manifest.json"
        if local_manifest.exists():
            try:
                raw = json.loads(local_manifest.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    for k, v in raw.items():
                        if str(v).strip():
                            base[str(k)] = str(v).strip()
            except (json.JSONDecodeError, OSError):
                pass
    return base


# Link targets used by the bundled email template.
HELP_PORTAL_URL = "https://www.pentanasolutions.com/customer-service-portal"
FOOTER_LOGO_URL = "https://www.pentanasolutions.com/"


def _base_asset_placeholders() -> dict[str, str]:
    return {
        "asset_banner_top": "",
        "asset_banner_hero": "",
        "asset_button_help": "",
        "asset_footer_logo": "",
        "help_portal_url": HELP_PORTAL_URL,
        "footer_logo_url": FOOTER_LOGO_URL,
    }


# The branded HTML template is a faithful rebuild of the .msg snapshot template
# and its four images appear in the same document order as the .msg's inline
# images. Verified against the July 2026 template: the .msg <img> tags are
# 712x118, 714x239, 310x66, 255x170, and the .j2 declares matching widths for
# button_help (310) and footer_logo (255).
MSG_ASSET_ORDER = ("banner_top", "banner_hero", "button_help", "footer_logo")

_IMG_CID_RE = re.compile(r"<img[^>]*\ssrc=[\"']cid:([^\"']+)[\"'][^>]*>", re.IGNORECASE)
_MIME_BY_SUFFIX = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
}


def msg_asset_cache_dir(msg_path: Path) -> Path:
    return PACKAGE_DIR / "logs" / "template_assets" / re.sub(r"[^A-Za-z0-9_.-]+", "_", msg_path.stem)


# Extracting a .msg re-reads and re-decodes every attachment. Assembling one
# email asks for the assets twice, and a monthly run assembles ~27 emails from
# the same unchanged template, so the result is cached against the file's
# identity (path + mtime + size) and recomputed only when the operator drops in
# a new template. The lock matters because approvals build emails on background
# threads that share the same cache folder.
_MSG_ASSET_CACHE: dict[tuple[str, int, int], list[Path]] = {}
_MSG_ASSET_LOCK = threading.Lock()


def extract_msg_inline_assets(msg_path: Path | None) -> list[Path]:
    """Extract a .msg template's inline (CID) images to a cache folder, in the
    order they appear in that template's HTML.

    The .msg the operator copies month on month already carries the branding
    images, so sourcing them from the template itself keeps the email correct
    without depending on a separate assets folder that may not exist on this
    machine.
    """
    if not msg_path:
        return []
    try:
        resolved = Path(msg_path).expanduser().resolve()
    except Exception:  # noqa: BLE001
        return []
    if not resolved.is_file() or resolved.suffix.lower() != ".msg":
        return []

    try:
        stat = resolved.stat()
        cache_key = (str(resolved), int(stat.st_mtime), int(stat.st_size))
    except OSError:
        cache_key = None
    if cache_key is not None:
        with _MSG_ASSET_LOCK:
            cached = _MSG_ASSET_CACHE.get(cache_key)
        if cached is not None and all(path.exists() for path in cached):
            return list(cached)

    try:
        import extract_msg  # type: ignore[import-not-found]
    except ImportError:
        logging.getLogger("email_assembler").warning(
            "extract-msg not installed; cannot read inline images from the .msg template."
        )
        return []

    try:
        with contextlib.closing(extract_msg.Message(str(resolved))) as message:
            return _extract_and_cache_msg_assets(message, resolved, cache_key)
    except Exception as exc:  # noqa: BLE001
        logging.getLogger("email_assembler").warning("Unable to read inline images from %s: %s", resolved, exc)
        return []


def _extract_and_cache_msg_assets(
    message: Any, resolved: Path, cache_key: tuple[str, int, int] | None
) -> list[Path]:
    try:
        html_raw = getattr(message, "htmlBody", None)
        html_text = html_raw.decode("utf-8", errors="replace") if isinstance(html_raw, bytes) else str(html_raw or "")

        # Index every attachment by both CID and filename so either reference
        # form in the HTML resolves.
        by_key: dict[str, tuple[str, bytes]] = {}
        for att in message.attachments:
            data = att.data
            if not isinstance(data, (bytes, bytearray)):
                continue
            filename = str(getattr(att, "longFilename", "") or getattr(att, "shortFilename", "") or "").strip()
            cid = str(getattr(att, "cid", "") or "").strip()
            for key in filter(None, {cid, filename}):
                by_key[key.lower()] = (filename or key, bytes(data))

        ordered: list[tuple[str, bytes]] = []
        seen: set[str] = set()
        for ref in _IMG_CID_RE.findall(html_text):
            key = ref.strip().lower()
            entry = by_key.get(key) or by_key.get(key.split("@", 1)[0])
            if entry is None or key in seen:
                continue
            seen.add(key)
            ordered.append(entry)
        if not ordered:
            return []

        cache_dir = msg_asset_cache_dir(resolved)
        # Two threads writing the same cache file can interleave and leave a
        # corrupt image in an email, so the write side is serialised.
        with _MSG_ASSET_LOCK:
            cache_dir.mkdir(parents=True, exist_ok=True)
            written: list[Path] = []
            for index, (filename, data) in enumerate(ordered):
                suffix = Path(filename).suffix.lower() or ".png"
                slot = MSG_ASSET_ORDER[index] if index < len(MSG_ASSET_ORDER) else f"asset_{index}"
                out_path = cache_dir / f"{index:02d}_{slot}{suffix}"
                try:
                    if not out_path.exists() or out_path.read_bytes() != data:
                        out_path.write_bytes(data)
                    written.append(out_path)
                except OSError as exc:
                    logging.getLogger("email_assembler").warning(
                        "Unable to cache template asset %s: %s", out_path, exc
                    )
            if cache_key is not None and written:
                _MSG_ASSET_CACHE[cache_key] = list(written)
        return written
    except Exception as exc:  # noqa: BLE001
        logging.getLogger("email_assembler").warning("Unable to read inline images from %s: %s", resolved, exc)
        return []


def _data_uri(path: Path) -> str:
    """Inline an image as a data: URI so the preview is fully self-contained."""
    import base64

    mime = _MIME_BY_SUFFIX.get(path.suffix.lower(), "application/octet-stream")
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def resolve_asset_urls(
    assets_dir: Path | None,
    *,
    for_outlook: bool,
    template_path: Path | None = None,
) -> tuple[dict[str, str], list[tuple[str, Path]]]:
    """Build Jinja context keys asset_banner_top, ... and list of (cid, path) for Outlook.

    Assets are sourced from the configured assets folder first; any slot that
    folder does not supply falls back to the images embedded in the .msg
    template itself. Preview URLs are data: URIs so the rendered preview is
    self-contained and always displays, rather than depending on the preview
    widget resolving file:// or cid: references.
    """
    manifest = _load_asset_manifest(assets_dir)
    ctx: dict[str, str] = {}
    inline: list[tuple[str, Path]] = []
    resolved_paths: dict[str, Path] = {}

    if assets_dir and assets_dir.is_dir():
        for key, fname in manifest.items():
            if not fname:
                continue
            path = (assets_dir / fname).resolve()
            if not path.is_file():
                logging.getLogger("email_assembler").warning("Asset file missing | key=%s path=%s", key, path)
                continue
            resolved_paths[key] = path

    missing = [key for key in MSG_ASSET_ORDER if key not in resolved_paths]
    if missing:
        msg_assets = extract_msg_inline_assets(template_path)
        for index, key in enumerate(MSG_ASSET_ORDER):
            if key in resolved_paths or index >= len(msg_assets):
                continue
            resolved_paths[key] = msg_assets[index]
        if msg_assets:
            logging.getLogger("email_assembler").info(
                "Using %s inline image(s) from the .msg template for slots: %s",
                len(msg_assets),
                ", ".join(k for k in missing if k in resolved_paths),
            )

    for key, path in resolved_paths.items():
        jinja_key = f"asset_{key}"
        if for_outlook:
            ctx[jinja_key] = f"cid:{key}"
            inline.append((key, path))
        else:
            try:
                ctx[jinja_key] = _data_uri(path)
            except OSError as exc:
                logging.getLogger("email_assembler").warning("Unable to inline asset %s: %s", path, exc)
                ctx[jinja_key] = path.as_uri()
    return ctx, inline


def collect_slots_from_templates(html_src: str, text_src: str) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for src in (html_src, text_src):
        for m in JINJA_VAR_RE.finditer(src):
            name = m.group(1).lower()
            if name not in seen:
                seen.add(name)
                ordered.append(name)
    return ordered


# One environment for the process: it holds no per-render state, and building
# it three times per email (preview HTML, Outlook HTML, plain text) is pure
# overhead.
_JINJA_ENV = Environment(undefined=StrictUndefined, autoescape=False)
# Insights, summary and greeting are plain text (operator-editable), so the HTML body escapes them.
_JINJA_HTML_ENV = Environment(undefined=StrictUndefined, autoescape=True)


def _render_jinja(template_str: str, context: dict[str, Any], *, html: bool = False) -> str:
    env = _JINJA_HTML_ENV if html else _JINJA_ENV
    try:
        return env.from_string(template_str).render(**context)
    except TemplateError as exc:
        raise email_builder.TemplateValidationError(
            "Unable to render email assembler template.",
            technical_message=str(exc),
        ) from exc


def _split_subject_and_body_plain(raw: str) -> tuple[str, str]:
    subject = "DealerSocket Insights"
    body = raw.strip()
    lines = raw.splitlines()
    if lines and lines[0].lower().startswith("subject:"):
        subject = lines[0].split(":", 1)[1].strip() or subject
        body = "\n".join(lines[1:]).lstrip("\n")
    return subject, body


def append_assembler_validation(record: dict[str, Any]) -> None:
    payload = {"timestamp": datetime.now(timezone.utc).isoformat()}
    payload.update(record)
    try:
        ASSEMBLER_VALIDATION_PATH.parent.mkdir(parents=True, exist_ok=True)
        existing: list[Any] = []
        if ASSEMBLER_VALIDATION_PATH.exists():
            try:
                loaded = json.loads(ASSEMBLER_VALIDATION_PATH.read_text(encoding="utf-8"))
                if isinstance(loaded, list):
                    existing = loaded
            except (json.JSONDecodeError, OSError):
                existing = []
        existing.append(payload)
        existing = existing[-MAX_DIAGNOSTIC_RECORDS:]
        ASSEMBLER_VALIDATION_PATH.write_text(json.dumps(existing, indent=2), encoding="utf-8")
    except Exception:
        pass


def build_email_context_contract(
    *,
    dealership_name: str,
    greeting_name: str,
    month_label: str,
    priority_insights: Any,
    summary: str,
    sales_report_path: str = "",
    service_report_path: str = "",
    sender_name: str = "",
    sender_title: str = "",
    email: str = "",
    cc: str = "",
) -> dict[str, str]:
    """Canonical lowercase context; same contract as email_builder.build_email_context_object."""
    return email_builder.build_email_context_object(
        dealership_name=dealership_name,
        greeting_name=greeting_name,
        month_label=month_label,
        priority_insights=priority_insights,
        summary=summary,
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name,
        sender_title=sender_title,
        email=email,
        cc=cc,
    )


def assemble_dealership_email(
    *,
    dealership_name: str,
    month_label: str,
    greeting_name: str,
    priority_insights: Any,
    summary_resolved: str,
    to_recipients: str,
    cc_recipients: str,
    pdf_attachments: list[Path],
    report_reference: dict[str, Any],
    assets_dir: Path | None,
    sales_report_path: str = "",
    service_report_path: str = "",
    template_path: Path | None = None,
    sender_name: str = "",
    sender_title: str = "",
) -> AssembledEmail:
    """Render branded HTML + plain text; build CID vs file:// HTML variants."""
    html_src = _read_text(HTML_TEMPLATE_PATH)
    text_src = _read_text(TEXT_TEMPLATE_PATH)
    slots = collect_slots_from_templates(html_src, text_src)

    core = build_email_context_contract(
        dealership_name=dealership_name,
        greeting_name=greeting_name,
        month_label=month_label,
        priority_insights=priority_insights,
        summary=summary_resolved,
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name or email_builder.DEFAULT_SENDER_NAME,
        sender_title=sender_title or email_builder.DEFAULT_SENDER_TITLE,
        email=to_recipients,
        cc=cc_recipients,
    )

    asset_ctx_preview, _inline_preview_unused = resolve_asset_urls(
        assets_dir, for_outlook=False, template_path=template_path
    )
    asset_ctx_outlook, inline_outlook = resolve_asset_urls(
        assets_dir, for_outlook=True, template_path=template_path
    )

    render_base = {**_base_asset_placeholders(), **core, **asset_ctx_preview}
    html_body = _render_jinja(html_src, render_base, html=True)

    render_outlook = {**_base_asset_placeholders(), **core, **asset_ctx_outlook}
    html_body_outlook = _render_jinja(html_src, render_outlook, html=True)

    plain_raw = _render_jinja(text_src, core)
    subject, plain_text_body = _split_subject_and_body_plain(plain_raw)

    sms_excluded = True
    legacy_fallback_used = False
    template_used = str(HTML_TEMPLATE_PATH)

    meta = {
        "dealership": dealership_name,
        "template_used": template_used,
        "jinja_slots_found": slots,
        "legacy_fallback_used": legacy_fallback_used,
        "inline_assets_found": [str(p) for _, p in inline_outlook],
        "inline_assets_attached": len(inline_outlook),
        "insight_1_present": bool(str(core.get("insight_1", "")).strip() and core.get("insight_1") != email_builder.INSIGHT_SLOT_PAD),
        "insight_2_present": bool(str(core.get("insight_2", "")).strip() and core.get("insight_2") != email_builder.INSIGHT_SLOT_PAD),
        "insight_3_present": bool(str(core.get("insight_3", "")).strip() and core.get("insight_3") != email_builder.INSIGHT_SLOT_PAD),
        "insight_4_present": bool(str(core.get("insight_4", "")).strip() and core.get("insight_4") != email_builder.INSIGHT_SLOT_PAD),
        "summary_present": bool(
            str(core.get("summary", "")).strip()
        ),
        "html_body_length": len(html_body),
        "plain_text_body_length": len(plain_text_body),
        "sms_excluded": sms_excluded,
        "attachment_count": len(pdf_attachments),
    }
    append_assembler_validation(meta)

    return AssembledEmail(
        subject=subject,
        plain_text_body=email_builder.normalize_text_for_preview(plain_text_body),
        html_body=html_body,
        html_body_outlook=html_body_outlook,
        attachments=list(pdf_attachments),
        inline_assets=inline_outlook,
        report_reference=dict(report_reference),
        email_context=dict(core),
        jinja_slots_found=slots,
        legacy_fallback_used=False,
        template_used=template_used,
        assembler_metadata=meta,
    )


def assembled_email_to_preview_dict(
    assembled: AssembledEmail,
    *,
    to_recipients: str,
    cc_recipients: str,
    bcc_recipients: str = "",
) -> dict[str, Any]:
    """Shape compatible with report_engine email_preview / email_builder.build_email_preview."""
    plain = assembled.plain_text_body
    return {
        "to": to_recipients,
        "cc": cc_recipients,
        "bcc": bcc_recipients,
        "subject": assembled.subject,
        "plain_text_body": plain,
        "body": plain,
        "body_text": plain,
        "preview_text": plain,
        "html_body": assembled.html_body,
        "html_body_outlook": assembled.html_body_outlook,
        "assembler_inline_assets": [{"cid": c, "path": str(p)} for c, p in assembled.inline_assets],
        "attachments": assembled.attachments,
        "template_type": "html",
        "template_path": assembled.template_used,
        "template_encoding": "assembler",
        "placeholders_found": [s.upper() for s in assembled.jinja_slots_found],
        "replaced_placeholders": [],
        "missing_placeholders": [],
        "render_audit_warnings": [],
        "render_contains_insights": assembled.assembler_metadata.get("insight_1_present", False),
        "render_contains_summary": assembled.assembler_metadata.get("summary_present", False),
        "email_context": dict(assembled.email_context),
        "legacy_fallback_used": False,
        "email_assembler": True,
    }
