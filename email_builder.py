from __future__ import annotations

import csv
import html
import json
import logging
import os
import re
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from jinja2 import Environment, StrictUndefined, TemplateError

SUPPORTED_TEMPLATE_EXTENSIONS = {".txt", ".html", ".htm", ".msg"}
SUPPORTED_TEMPLATE_ENCODINGS = ["utf-8", "utf-16", "cp1252", "latin-1"]

# Preferred Jinja variables for dealership workflow emails (lowercase names in templates).
PREFERRED_EMAIL_REQUIRED_VARS = ("name", "month", "insight_1", "insight_2", "insight_3", "insight_4")
PREFERRED_EMAIL_OPTIONAL_VARS = (
    "dealership",
    "summary",
    "sales_report_path",
    "service_report_path",
    "sender_name",
    "sender_title",
)
PRIORITY_INSIGHT_COUNT = 4
SMS_CHAR_LIMIT = 160
NO_ADDITIONAL_INSIGHT_MARKER = "[No additional insight]"
# Legacy alias; preview/debug padding only — never use in final-ready output.
INSIGHT_SLOT_PAD = NO_ADDITIONAL_INSIGHT_MARKER
MISSING_INSIGHT_MARKER = INSIGHT_SLOT_PAD

# Empty on purpose. A fresh install belongs to whoever set it up, so nothing is
# signed until they fill in Sender details, and the sign-off block is skipped
# entirely while these are blank rather than rendering "Best regards," over a gap.
DEFAULT_SENDER_NAME = ""
DEFAULT_SENDER_TITLE = ""

DEFAULT_SUBJECT_TEMPLATE = "DealerSocket {{ MONTH }} Insights - {{ DEALERSHIP_NAME }}"
DEFAULT_BODY_TEMPLATE = (
    "Hi {{ name }},\n\n"
    "Attached are your Sales and Service reports for {{ month }}. Below are the four headline moves.\n\n"
    "Key Insights:\n"
    "{{ insight_1 }}\n"
    "{{ insight_2 }}\n"
    "{{ insight_3 }}\n"
    "{{ insight_4 }}\n\n"
    "{% if summary %}{{ summary }}\n\n{% endif %}"
    "Reply if you want any of these unpacked against last month or the PDFs.\n\n"
    "{% if sender_name or sender_title %}Best regards,\n"
    "{{ sender_name }}\n"
    "{{ sender_title }}{% endif %}"
)
DEFAULT_HTML_BODY_TEMPLATE = (
    "<html><body>\n"
    "<p>Hi {{ name }},</p>\n"
    "<p>Attached are your Sales and Service reports for {{ month }}. Below are the four headline moves, then a short wrap-up in the same order of emphasis.</p>\n"
    "<p><strong>Key Insights:</strong></p>\n"
    "<ul>\n"
    "<li>{{ insight_1 }}</li>\n"
    "<li>{{ insight_2 }}</li>\n"
    "<li>{{ insight_3 }}</li>\n"
    "<li>{{ insight_4 }}</li>\n"
    "</ul>\n"
    "<p>{{ summary }}</p>\n"
    "<p>Reply if you want any of these unpacked against last month or the PDFs.</p>\n"
    "{% if sender_name or sender_title %}<p>Best regards,<br>\n"
    "{{ sender_name }}<br>\n"
    "{{ sender_title }}</p>\n{% endif %}"
    "</body></html>"
)
EXPLICIT_PLACEHOLDER_PATTERN = re.compile(r"{{\s*([A-Za-z0-9_]+)\s*}}")
LEGACY_PLACEHOLDER_PATTERN = re.compile(
    r"\{(dealership_name|dealership|name|month|month_label|executive_summary|priority_insights|summary|sms_version|sentiment|review_reason_tags|to|cc|subject|point_[1-4]|insight_[1-4]|sales_report_path|service_report_path|sender_name|sender_title)\}",
    flags=re.IGNORECASE,
)
SQUARE_BRACKET_PLACEHOLDER_PATTERN = re.compile(
    r"\[\s*(dealership_name|dealership|name|month|month_label|executive_summary|priority_insights|summary|sms_version|sentiment|review_reason_tags|to|cc|subject|point_[1-4]|insight_[1-4]|sales_report_path|service_report_path|sender_name|sender_title)\s*\]",
    flags=re.IGNORECASE,
)
HTML_SUBJECT_COMMENT_PATTERN = re.compile(r"<!--\s*SUBJECT\s*:(.*?)-->", flags=re.IGNORECASE | re.DOTALL)
HTML_SUBJECT_META_PATTERN = re.compile(
    r"<meta\s+name=[\"']subject[\"']\s+content=[\"'](.*?)[\"']\s*/?>",
    flags=re.IGNORECASE | re.DOTALL,
)
TEXT_SUBJECT_LINE_PATTERN = re.compile(r"^\s*subject\s*:\s*(.*)$", flags=re.IGNORECASE)
RECOGNISED_PLACEHOLDERS = {
    "DEALERSHIP_NAME",
    "DEALERSHIP",
    "NAME",
    "MONTH",
    "MONTH_LABEL",
    "EXECUTIVE_SUMMARY",
    "PRIORITY_INSIGHTS",
    "SUMMARY",
    "SMS_VERSION",
    "SENTIMENT",
    "REVIEW_REASON_TAGS",
    "TO",
    "CC",
    "SUBJECT",
    "POINT_1",
    "POINT_2",
    "POINT_3",
    "POINT_4",
    "INSIGHT_1",
    "INSIGHT_2",
    "INSIGHT_3",
    "INSIGHT_4",
    "SALES_REPORT_PATH",
    "SERVICE_REPORT_PATH",
    "SENDER_NAME",
    "SENDER_TITLE",
}

JINJA_VAR_PATTERN = re.compile(r"\{\{-?\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*-?\}\}")


class TemplateValidationError(RuntimeError):
    def __init__(self, friendly_message: str, technical_message: str | None = None) -> None:
        super().__init__(friendly_message)
        self.friendly_message = friendly_message
        self.technical_message = technical_message or friendly_message


class InsightValidationError(TemplateValidationError):
    """Raised when final-ready priority insights are missing or padded."""


@dataclass
class LoadedTemplate:
    template_path: Path | None
    template_type: str
    subject_template: str
    body_template: str
    raw_text: str
    encoding: str
    placeholders_found: list[str]


@dataclass
class RenderedEmail:
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
    render_audit_warnings: list[str] = field(default_factory=list)
    render_contains_insights: bool = False
    render_contains_summary: bool = False


@dataclass(frozen=True)
class ResolvedContact:
    """Local contacts.csv resolution for a dealership row."""

    to: str
    cc: str
    contact_name: str
    greeting_name: str
    matched: bool
    ready: bool
    row_active: bool = True
    bcc: str = ""


CONTACT_CSV_FIELDNAMES = [
    "dealership",
    "contact_name",
    "email",
    "cc",
    "bcc",
    # Kept in the file but no longer shown in the app. Existing values are preserved
    # rather than dropped, so a contacts.csv written by an older build still behaves
    # the way its author intended.
    "active",
    "notes",
    "last_updated",
]
EMAIL_SIMPLE_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
ADDRESS_SPLIT = re.compile(r"[;,]")


def split_addresses(value: str) -> list[str]:
    """One field, several addresses — contacts.csv accepts either separator."""
    return [part.strip() for part in ADDRESS_SPLIT.split(str(value or "")) if part.strip()]


def is_probable_email(value: str) -> bool:
    return bool(value.strip()) and bool(EMAIL_SIMPLE_PATTERN.match(value.strip()))


def is_probable_email_list(value: str) -> bool:
    """True when every address in a list looks like one, and there is at least one.

    A dealership with two contacts resolves to "sam@x; pat@x", which the single-address
    check rejected outright — the merge landed and the whole dealership then had no
    recipient. Every part must be valid rather than any: dropping the malformed half of
    a pair silently would send half the email and say nothing, and a typo the operator
    can see and fix is the better failure.
    """
    parts = split_addresses(value)
    return bool(parts) and all(is_probable_email(part) for part in parts)


def parse_active_flag(value: str) -> bool | None:
    low = value.strip().lower()
    if low in {"yes", "true", "1", "y"}:
        return True
    if low in {"no", "false", "0", "n"}:
        return False
    return None


def normalize_contact_row(row_lower: dict[str, str]) -> dict[str, str]:
    dealership = row_lower.get("dealership") or row_lower.get("dealership_name") or ""
    email = (
        row_lower.get("email")
        or row_lower.get("to")
        or row_lower.get("recipient")
        or row_lower.get("recipients")
        or ""
    ).strip()
    return {
        "dealership": str(dealership).strip(),
        "contact_name": (row_lower.get("contact_name") or "").strip(),
        "email": email,
        "to": email,
        "cc": (row_lower.get("cc") or "").strip(),
        "bcc": (row_lower.get("bcc") or "").strip(),
        "active": (row_lower.get("active") or "yes").strip() or "yes",
        "notes": (row_lower.get("notes") or "").strip(),
        "last_updated": (row_lower.get("last_updated") or "").strip(),
    }


def normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def normalize_placeholder_name(value: str) -> str:
    return re.sub(r"\s+", "_", value.strip().upper())


def normalize_text_for_preview(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


_INSIGHT_LINE_PREFIX = re.compile(r"^(?:[\s]*[-*•●\u2022]\s*|\s*\d+[\.)]\s*)+")


def clean_insight_line(line: str) -> str:
    t = (line or "").replace("\r\n", "\n").strip()
    if not t:
        return ""
    while True:
        n = _INSIGHT_LINE_PREFIX.sub("", t).strip()
        if n == t:
            break
        t = n
    return t


def extract_priority_insight_lines(raw: Any) -> list[str]:
    """Normalise priority insights from list, newline/bullet string, dict, JSON string, tuple, or set."""
    if raw is None:
        return []
    if isinstance(raw, dict):
        def _key_sort(k: Any) -> tuple[int, str]:
            s = str(k)
            return (int(s) if s.isdigit() else 9999, s)

        lines: list[str] = []
        for key in sorted(raw.keys(), key=_key_sort):
            val = raw.get(key)
            if val is None:
                continue
            cleaned = clean_insight_line(str(val))
            if cleaned:
                lines.append(cleaned)
        return lines
    if isinstance(raw, (list, tuple, set)):
        out: list[str] = []
        for item in raw:
            c = clean_insight_line(str(item))
            if c:
                out.append(c)
        return out
    s = str(raw).strip()
    if not s or s.lower() == "none":
        return []
    if s.startswith(("[", "{")):
        try:
            parsed = json.loads(s)
        except json.JSONDecodeError:
            parsed = None
        if parsed is not None:
            return extract_priority_insight_lines(parsed)
    if "\n" in s:
        parts = s.splitlines()
    elif ";" in s and len(s) > 30:
        parts = [p.strip() for p in re.split(r"\s*;\s*", s) if p.strip()]
    else:
        parts = [s]
    return [c for p in parts if (c := clean_insight_line(p))]


def normalize_priority_insights_to_four(
    raw: Any,
    *,
    final_ready: bool = True,
) -> tuple[str, str, str, str]:
    lines = extract_priority_insight_lines(raw)
    if final_ready:
        errors = validate_four_insight_slots(
            insight_1=lines[0] if len(lines) > 0 else "",
            insight_2=lines[1] if len(lines) > 1 else "",
            insight_3=lines[2] if len(lines) > 2 else "",
            insight_4=lines[3] if len(lines) > 3 else "",
        )
        if errors:
            raise InsightValidationError(
                "Final-ready email requires exactly four priority insights.",
                technical_message="; ".join(errors),
            )
        return lines[0], lines[1], lines[2], lines[3]
    slots: list[str] = []
    for i in range(PRIORITY_INSIGHT_COUNT):
        slots.append(lines[i] if i < len(lines) else INSIGHT_SLOT_PAD)
    return slots[0], slots[1], slots[2], slots[3]


def validate_four_insight_slots(
    *,
    insight_1: str,
    insight_2: str,
    insight_3: str,
    insight_4: str,
) -> list[str]:
    errors: list[str] = []
    for slot_name, value in (
        ("insight_1", insight_1),
        ("insight_2", insight_2),
        ("insight_3", insight_3),
        ("insight_4", insight_4),
    ):
        cleaned = (value or "").strip()
        if not cleaned:
            errors.append(f"{slot_name} is missing")
        elif cleaned == NO_ADDITIONAL_INSIGHT_MARKER:
            errors.append(f"{slot_name} uses placeholder padding")
    return errors


def validate_sms_output(sms_version: str, *, limit: int = SMS_CHAR_LIMIT) -> tuple[bool, str]:
    sms = (sms_version or "").strip()
    if not sms:
        return False, "SMS is empty"
    if len(sms) > limit:
        return False, f"SMS exceeds {limit} characters ({len(sms)})"
    return True, ""


def resolve_summary_for_email(
    *,
    summary: str = "",
    report_summary: str = "",
    executive_summary: str = "",
    report_reference: dict[str, Any] | None = None,
) -> str:
    candidates: list[str] = [summary, report_summary, executive_summary]
    if isinstance(report_reference, dict):
        for key in (
            "summary",
            "Summary",
            "report_summary",
            "Report summary",
            "executive_summary",
            "Executive Summary",
        ):
            val = report_reference.get(key)
            if val is not None:
                candidates.append(str(val))
    for cand in candidates:
        t = normalize_text_for_preview(str(cand or ""))
        if t:
            return t
    return ""


def split_priority_insights(priority_insights: Any) -> list[str]:
    return extract_priority_insight_lines(priority_insights)


def validate_rendered_email_mapping(
    *,
    preview_text: str,
    html_body: str,
    insight_1: str,
    insight_2: str,
    insight_3: str,
    insight_4: str,
    summary_resolved: str = "",
    sms_version: str = "",
    omit_sms_from_email: bool = False,
    had_raw_insight_lines: bool = False,
    final_ready: bool = True,
) -> tuple[list[str], bool, bool]:
    """Return (warnings, rendered_contains_insights, rendered_contains_summary)."""
    warnings: list[str] = []
    sms_clean = (sms_version or "").strip()
    if omit_sms_from_email and sms_clean:
        sms_in_plain = sms_clean in preview_text
        sms_in_html = sms_clean in html_body
        if sms_in_plain or sms_in_html:
            logging.warning(
                "SMS content detected in email body when it should be excluded",
            )
        if sms_in_plain:
            warnings.append("SMS text appears in plain-text body but omit_sms_from_email is True")
        if sms_in_html:
            warnings.append("SMS text appears in HTML body but omit_sms_from_email is True")

    if final_ready:
        slot_errors = validate_four_insight_slots(
            insight_1=insight_1,
            insight_2=insight_2,
            insight_3=insight_3,
            insight_4=insight_4,
        )
        warnings.extend(slot_errors)
        sms_ok, sms_err = validate_sms_output(sms_clean)
        if sms_clean and not sms_ok:
            warnings.append(sms_err)

    insights = (insight_1, insight_2, insight_3, insight_4)
    non_pad = [x for x in insights if x.strip() and x != INSIGHT_SLOT_PAD]
    rendered_contains_insights = False
    if non_pad:
        rendered_contains_insights = all(x in preview_text for x in non_pad)
        if had_raw_insight_lines and not rendered_contains_insights:
            warnings.append("Insights were supplied but plain body does not contain all four verbatim (check template variables)")
    m = re.search(r"key insights:\s*", preview_text, flags=re.IGNORECASE)
    if m:
        after = preview_text[m.end() : m.end() + 500]
        if not re.search(r"[A-Za-z0-9]", after[:300]):
            warnings.append("'Key Insights:' is not followed by readable content")

    summ = (summary_resolved or "").strip()
    rendered_contains_summary = False
    if summ:
        rendered_contains_summary = summ in preview_text or summ in html_to_preview_text(html_body)

    return warnings, rendered_contains_insights, rendered_contains_summary


def render_priority_insights_text(priority_insights: str | list[str]) -> str:
    lines = split_priority_insights(priority_insights)
    if not lines:
        return "No priority insights available."
    return "\n".join(f"• {line}" for line in lines)


def text_block_to_html(text: str) -> str:
    cleaned = normalize_text_for_preview(text)
    if not cleaned:
        return ""
    paragraphs = [paragraph.strip() for paragraph in re.split(r"\n\s*\n", cleaned) if paragraph.strip()]
    html_paragraphs = []
    for paragraph in paragraphs:
        escaped_lines = [html.escape(line.strip()) for line in paragraph.splitlines() if line.strip()]
        html_paragraphs.append(f"<p>{'<br>'.join(escaped_lines)}</p>")
    return "\n".join(html_paragraphs)


def render_priority_insights_html(priority_insights: str | list[str]) -> str:
    lines = split_priority_insights(priority_insights)
    if not lines:
        return "<p>No priority insights available.</p>"
    return "<ul>" + "".join(f"<li>{html.escape(line)}</li>" for line in lines) + "</ul>"


def convert_text_to_html(text: str) -> str:
    html_text = text_block_to_html(text)
    return html_text or "<p></p>"


def ensure_outlook_html_document(html: str) -> str:
    """Return HTML suitable for Outlook MailItem.HTMLBody: full document with charset when possible."""
    raw = (html or "").strip()
    if not raw:
        return ""
    low = raw.lower()
    if "<html" not in low:
        return (
            "<!DOCTYPE html>\n"
            '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            "<title></title>\n</head>\n<body>\n"
            f"{raw}\n</body>\n</html>"
        )
    he = low.find("</head>")
    head_chunk = raw[: he if he != -1 else min(len(raw), 8000)]
    if "charset" not in head_chunk.lower() and re.search(r"<head[^>]*>", raw, flags=re.IGNORECASE):
        raw = re.sub(
            r"(<head[^>]*>)",
            r'\1\n<meta charset="utf-8">',
            raw,
            count=1,
            flags=re.IGNORECASE,
        )
    return raw


def html_to_preview_text(html_text: str) -> str:
    preview = html_text
    substitutions = [
        (r"<li[^>]*>", "• "),
        (r"</li>", "\n"),
        (r"<br\s*/?>", "\n"),
        (r"</p>", "\n\n"),
        (r"</div>", "\n"),
        (r"</h[1-6]>", "\n\n"),
        (r"</tr>", "\n"),
        (r"<style.*?>.*?</style>", ""),
        (r"<script.*?>.*?</script>", ""),
    ]
    for pattern, replacement in substitutions:
        preview = re.sub(pattern, replacement, preview, flags=re.IGNORECASE | re.DOTALL)
    preview = re.sub(r"<[^>]+>", "", preview)
    preview = html.unescape(preview)
    return normalize_text_for_preview(preview)


def collect_placeholders(template_text: str) -> list[str]:
    placeholders = [normalize_placeholder_name(match.group(1)) for match in EXPLICIT_PLACEHOLDER_PATTERN.finditer(template_text)]
    placeholders.extend(normalize_placeholder_name(match.group(1)) for match in LEGACY_PLACEHOLDER_PATTERN.finditer(template_text))
    placeholders.extend(normalize_placeholder_name(match.group(1)) for match in SQUARE_BRACKET_PLACEHOLDER_PATTERN.finditer(template_text))
    for match in JINJA_VAR_PATTERN.finditer(template_text):
        placeholders.append(normalize_placeholder_name(match.group(1)))
    seen: set[str] = set()
    ordered: list[str] = []
    for placeholder in placeholders:
        if placeholder in seen:
            continue
        seen.add(placeholder)
        ordered.append(placeholder)
    return ordered


def collect_jinja_identifier_names(template_text: str) -> list[str]:
    """Return ordered unique Jinja `{{ var }}` identifiers (upper-normalised names)."""
    seen: set[str] = set()
    ordered: list[str] = []
    for match in JINJA_VAR_PATTERN.finditer(template_text):
        name = normalize_placeholder_name(match.group(1))
        if name in seen:
            continue
        seen.add(name)
        ordered.append(name)
    return ordered


def template_has_insight_or_summary_slots(body_template: str) -> bool:
    """True when the template body uses Jinja for insights or summary (not static Word layout only)."""
    names = {normalize_placeholder_name(m.group(1)) for m in JINJA_VAR_PATTERN.finditer(body_template)}
    slots = {
        "INSIGHT_1",
        "INSIGHT_2",
        "INSIGHT_3",
        "INSIGHT_4",
        "POINT_1",
        "POINT_2",
        "POINT_3",
        "POINT_4",
        "SUMMARY",
    }
    return bool(names & slots)


def inject_insights_and_summary_into_html_if_missing(
    html_doc: str,
    *,
    insight_1: str,
    insight_2: str,
    insight_3: str,
    insight_4: str,
    summary: str,
) -> str:
    """Insert bullet list + summary after 'Key Insights:' when Word/Outlook HTML has no {{ insight_* }} / {{ summary }}.

    Avoids double-insert if insight text already appears immediately after the heading block.
    """
    if not html_doc or "key insights" not in html_doc.lower():
        return html_doc
    lines = [
        x.strip()
        for x in (insight_1, insight_2, insight_3, insight_4)
        if (x or "").strip() and (x or "").strip() != INSIGHT_SLOT_PAD
    ]
    summ = normalize_text_for_preview(str(summary or ""))
    if not lines and (not summ or summ == "[Missing report summary]"):
        return html_doc
    lower = html_doc.lower()
    idx = lower.find("key insights")
    if idx < 0:
        return html_doc
    colon = lower.find(":", idx)
    if colon < 0:
        return html_doc
    close_p = html_doc.find("</p>", colon)
    if close_p < 0:
        return html_doc
    insert_pos = close_p + len("</p>")
    window = html_doc[insert_pos : insert_pos + 8000]
    if lines and lines[0] in window:
        return html_doc
    logging.warning("Template missing Jinja slots; legacy injection fallback used.")
    items = "".join(f"<li>{html.escape(line)}</li>" for line in lines)
    block = f"<ul>{items}</ul>"
    if summ and summ != "[Missing report summary]":
        block += f"<p>{html.escape(summ)}</p>"
    return html_doc[:insert_pos] + block + html_doc[insert_pos:]


def validate_preferred_body_variables(template_text: str, context: dict[str, str]) -> list[str]:
    """If the template references preferred required variables, ensure each is non-empty after render prep."""
    used = {normalize_placeholder_name(m.group(1)).lower() for m in JINJA_VAR_PATTERN.finditer(template_text)}
    missing: list[str] = []
    for req in PREFERRED_EMAIL_REQUIRED_VARS:
        if req not in used:
            continue
        value = (context.get(req) or context.get(req.upper()) or "").strip()
        if not value:
            missing.append(req)
    return missing


def _convert_legacy_placeholders_to_jinja(template_text: str) -> str:
    replaced = LEGACY_PLACEHOLDER_PATTERN.sub(lambda match: f"{{{{ {normalize_placeholder_name(match.group(1))} }}}}", template_text)
    replaced = SQUARE_BRACKET_PLACEHOLDER_PATTERN.sub(lambda match: f"{{{{ {normalize_placeholder_name(match.group(1))} }}}}", replaced)
    return replaced


def _render_jinja_template(template_text: str, context: dict[str, str]) -> str:
    try:
        environment = Environment(undefined=StrictUndefined, autoescape=False)
        template = environment.from_string(template_text)
        return template.render(**context)
    except TemplateError as exc:
        raise TemplateValidationError(
            "Unable to render the selected email template. Check the template syntax and required fields.",
            technical_message=str(exc),
        ) from exc


def read_template_text(template_path: Path) -> tuple[str, str]:
    errors: list[str] = []
    for encoding in SUPPORTED_TEMPLATE_ENCODINGS:
        try:
            return template_path.read_text(encoding=encoding), encoding
        except UnicodeError as exc:
            errors.append(f"{encoding}: {exc}")
        except OSError as exc:
            logging.exception("Failed reading template file: %s", template_path)
            raise TemplateValidationError(
                "Unable to read the selected email template. Please use a standard .txt, .html, or .msg template, or re-save the file in a compatible format.",
                technical_message=str(exc),
            ) from exc

    logging.warning("Template decode attempts failed for %s | %s", template_path, " | ".join(errors))
    raise TemplateValidationError(
        "Unable to read the selected email template. Please use a standard .txt, .html, or .msg template, or re-save the file in a compatible format.",
        technical_message=" | ".join(errors),
    )


def _decode_template_bytes(raw_value: bytes) -> str:
    for encoding in SUPPORTED_TEMPLATE_ENCODINGS:
        try:
            return raw_value.decode(encoding)
        except UnicodeError:
            continue
    return raw_value.decode("latin-1", errors="replace")


def _load_msg_via_outlook(template_path: Path) -> tuple[str | None, str | None, str]:
    """Try loading the MSG via Outlook to preserve exact HTMLBody.

    Returns (subject, body, type) where type is "html" or "text". Subject/body may be None
    if not available. This gracefully falls back when Outlook/pywin32 is unavailable.
    """
    try:  # pragma: no cover - Windows/Outlook specific
        import win32com.client  # type: ignore[import-not-found]
    except Exception:
        return None, None, "text"

    try:  # pragma: no cover - Windows/Outlook specific
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        item = mapi.OpenSharedItem(str(template_path))
        subject = str(getattr(item, "Subject", "") or "").strip() or None
        html_body = str(getattr(item, "HTMLBody", "") or "").strip() or None
        if html_body:
            return subject, html_body, "html"
        body = str(getattr(item, "Body", "") or "").strip() or None
        if body:
            return subject, body, "text"
        return subject, None, "text"
    except Exception:
        return None, None, "text"


def _clean_msg_plain_body(text: str) -> str:
    """Heuristically remove MSG header noise from a plain text body.

    Some .msg templates store RTF-only content that libraries expose as plain text with
    leading header-like artifacts (e.g., X-MSMail, MicrosoftInternetExplorer4, Clean, false).
    This tries to strip those top-of-file artifacts so the preview matches what Outlook shows.
    """
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    header_prefix = re.compile(r"^[A-Za-z0-9-]+:\s*")
    junk_exact = {"clean", "false", "x-none", "en-au", "microsoftinternetexplorer4"}

    # Find the first likely content line (greeting or sentence case), up to first 100 lines
    anchor = 0
    for i, line in enumerate(lines[:100]):
        low = line.strip().lower()
        if not low:
            continue
        if header_prefix.match(line) or low in junk_exact or low.startswith("x-"):
            continue
        if re.match(r"^(hi|hello|dear|team|good\s+(morning|afternoon|evening)|please\b|attached\b)", low):
            anchor = i
            break
        # If the line has normal sentence punctuation, treat as start of content
        if re.search(r"[a-z][.!?]\s*$", low) or re.search(r"\b(the|and|for|with|of)\b", low):
            anchor = i
            break
        # Otherwise keep scanning

    cleaned = "\n".join(lines[anchor:]) if lines else text
    # Collapse excessive blank lines
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip()
    return cleaned


def _clean_preview_text_outlook_junk(text: str) -> str:
    """Remove common Outlook/Word artifacts that can bleed into preview text.

    Applied to preview_text when the source template is a .msg to give a clean
    editor experience while leaving the HTML body intact for Outlook drafts.
    """
    if not text:
        return text
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    junk_exact = {"clean", "false", "x-none", "en-au", "documentemail", "microsoftinternetexplorer4"}
    header_prefix = re.compile(r"^[A-Za-z0-9-]+:\s*")
    filtered: list[str] = []
    started = False
    for line in lines:
        low = line.strip().lower()
        if not started:
            if not low or low in junk_exact or low.startswith("x-") or header_prefix.match(line):
                continue
            # Heuristic: consider this the start when it looks like human content
            if re.match(r"^(hi|hello|dear|team|good\s+(morning|afternoon|evening)|please\b|attached\b)", low) or re.search(
                r"\b(the|and|for|with|of)\b",
                low,
            ):
                started = True
            else:
                # Skip the first 20 non-empty junk-like lines
                if len(filtered) < 1 and len(low) < 40:
                    continue
                started = True
        filtered.append(line)
    cleaned = "\n".join(filtered)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip()
    return cleaned


def load_msg_template(template_path: Path) -> tuple[str, str, str]:
    # First try reading via Outlook to preserve exact visible HTML
    subject, body, body_type = _load_msg_via_outlook(template_path)
    if body:
        subject_template = (subject or "").strip() or DEFAULT_SUBJECT_TEMPLATE
        if body_type == "html":
            return subject_template, body, "html"
        # Plain text from Outlook — apply light cleanup
        return subject_template, _clean_msg_plain_body(body), "text"

    # Fallback to extract_msg when Outlook isn't available
    try:
        import extract_msg  # type: ignore[import-not-found]
    except ImportError as exc:
        raise TemplateValidationError(
            "MSG template support requires either Outlook (pywin32) or extract-msg. Please install one of them.",
            technical_message=str(exc),
        ) from exc

    try:
        message = extract_msg.Message(str(template_path))
        subject_template = str(getattr(message, "subject", "") or "").strip() or DEFAULT_SUBJECT_TEMPLATE

        html_body = getattr(message, "htmlBody", None)
        if isinstance(html_body, bytes):
            html_text = _decode_template_bytes(html_body).strip()
        else:
            html_text = str(html_body or "").strip()

        if html_text:
            return subject_template, html_text, "html"

        body_template = str(getattr(message, "body", "") or "").strip() or DEFAULT_BODY_TEMPLATE
        body_template = _clean_msg_plain_body(body_template)
        return subject_template, body_template, "text"
    except Exception as exc:  # noqa: BLE001
        logging.exception("Failed reading MSG template file: %s", template_path)
        raise TemplateValidationError(
            "Unable to read the selected MSG email template. Please confirm the file opens in Outlook and contains a readable subject/body.",
            technical_message=str(exc),
        ) from exc


def extract_subject_and_body(raw_text: str, template_type: str) -> tuple[str, str]:
    subject_template = DEFAULT_SUBJECT_TEMPLATE
    body_template = raw_text.strip() or DEFAULT_BODY_TEMPLATE

    if template_type == "html":
        subject_comment_match = HTML_SUBJECT_COMMENT_PATTERN.search(raw_text)
        subject_meta_match = HTML_SUBJECT_META_PATTERN.search(raw_text)
        if subject_comment_match:
            subject_template = subject_comment_match.group(1).strip() or DEFAULT_SUBJECT_TEMPLATE
            body_template = HTML_SUBJECT_COMMENT_PATTERN.sub("", raw_text, count=1).strip() or body_template
        elif subject_meta_match:
            subject_template = subject_meta_match.group(1).strip() or DEFAULT_SUBJECT_TEMPLATE
            body_template = HTML_SUBJECT_META_PATTERN.sub("", raw_text, count=1).strip() or body_template
    else:
        lines = raw_text.splitlines()
        if lines:
            match = TEXT_SUBJECT_LINE_PATTERN.match(lines[0])
            if match:
                subject_template = match.group(1).strip() or DEFAULT_SUBJECT_TEMPLATE
                body_template = "\n".join(lines[1:]).strip() or DEFAULT_BODY_TEMPLATE

    return subject_template, body_template


def load_template_file(template_path: str | Path | None) -> LoadedTemplate:
    if not template_path:
        raw_text = DEFAULT_BODY_TEMPLATE
        placeholders = collect_placeholders(DEFAULT_SUBJECT_TEMPLATE + "\n" + raw_text)
        return LoadedTemplate(
            template_path=None,
            template_type="text",
            subject_template=DEFAULT_SUBJECT_TEMPLATE,
            body_template=raw_text,
            raw_text=raw_text,
            encoding="built-in",
            placeholders_found=placeholders,
        )

    resolved = Path(template_path).expanduser().resolve()
    if not resolved.exists():
        raise TemplateValidationError(
            "The selected email template could not be found. Please choose an existing .txt, .html, .htm, or .msg file.",
            technical_message=f"Template file not found: {resolved}",
        )

    suffix = resolved.suffix.lower()
    if suffix not in SUPPORTED_TEMPLATE_EXTENSIONS:
        raise TemplateValidationError(
            "Unsupported email template format. Please choose a .txt, .html, .htm, or .msg template file.",
            technical_message=f"Unsupported template extension: {suffix}",
        )

    if suffix == ".msg":
        subject_template, body_template, template_type = load_msg_template(resolved)
        raw_text = f"Subject: {subject_template}\n\n{body_template}"
        encoding = "msg"
    else:
        raw_text, encoding = read_template_text(resolved)
        template_type = "html" if suffix in {".html", ".htm"} else "text"
        subject_template, body_template = extract_subject_and_body(raw_text, template_type)

    placeholders = collect_placeholders(subject_template + "\n" + body_template)
    logging.info(
        "Loaded template file %s | type=%s | encoding=%s | placeholders=%s",
        resolved,
        template_type,
        encoding,
        ", ".join(placeholders) or "none",
    )
    return LoadedTemplate(
        template_path=resolved,
        template_type=template_type,
        subject_template=subject_template,
        body_template=body_template,
        raw_text=raw_text,
        encoding=encoding,
        placeholders_found=placeholders,
    )


def validate_template_file(template_path: str | Path | None) -> dict[str, Any]:
    template = load_template_file(template_path)
    if not template.placeholders_found:
        logging.warning("Template file has no placeholders detected: %s", template.template_path)
    return {
        "template_path": str(template.template_path) if template.template_path else "",
        "template_type": template.template_type,
        "encoding": template.encoding,
        "placeholders_found": template.placeholders_found,
    }


def split_month_label(month_label: str) -> tuple[str, str]:
    """Split a "Mon YYYY" label into ("July", "2026").

    Falls back to returning the label as-is with an empty year when it is not in
    the expected form, so a template never renders an empty month.
    """
    cleaned = (month_label or "").strip()
    parts = cleaned.split()
    if len(parts) == 2 and parts[1].isdigit():
        for fmt in ("%b %Y", "%B %Y"):
            try:
                parsed = datetime.strptime(cleaned, fmt)
            except ValueError:
                continue
            return parsed.strftime("%B"), str(parsed.year)
        return parts[0], parts[1]
    return cleaned, ""


def build_email_context_object(
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
    """Canonical lowercase Jinja context for dealership emails (tests and diagnostics)."""
    insight_1, insight_2, insight_3, insight_4 = normalize_priority_insights_to_four(priority_insights)
    sender_name_value = sender_name.strip() or DEFAULT_SENDER_NAME
    sender_title_value = sender_title.strip() or DEFAULT_SENDER_TITLE
    summary_text = normalize_text_for_preview(summary)
    month_name_value, year_value = split_month_label(month_label)
    return {
        "dealership": dealership_name.strip(),
        "name": greeting_name.strip(),
        "email": email.strip(),
        "cc": cc.strip(),
        "month": month_label.strip(),
        # "month" is the short label ("Jul 2026"). Customer-facing copy uses the
        # full month name alone ("...Reports for July"), with the year only in
        # the subject line, so both parts are exposed separately.
        "month_name": month_name_value,
        "year": year_value,
        "insight_1": insight_1,
        "insight_2": insight_2,
        "insight_3": insight_3,
        "insight_4": insight_4,
        "point_1": insight_1,
        "point_2": insight_2,
        "point_3": insight_3,
        "point_4": insight_4,
        "summary": summary_text,
        "sales_report_path": sales_report_path.strip(),
        "service_report_path": service_report_path.strip(),
        "sender_name": sender_name_value,
        "sender_title": sender_title_value,
    }


def build_template_context(
    *,
    dealership_name: str,
    month_label: str,
    executive_summary: str,
    priority_insights: Any,
    summary: str,
    sms_version: str,
    sentiment: str,
    review_reason_tags: list[str],
    to_recipients: str = "",
    cc_recipients: str = "",
    subject: str = "",
    template_type: str = "text",
    sales_report_path: str = "",
    service_report_path: str = "",
    sender_name: str = "",
    sender_title: str = "",
    greeting_name: str = "",
    omit_sms_from_email: bool = True,
) -> dict[str, str]:
    review_reason_text = ", ".join(tag for tag in review_reason_tags if tag.strip()) or "None"
    insight_1, insight_2, insight_3, insight_4 = normalize_priority_insights_to_four(priority_insights)
    point_1, point_2, point_3, point_4 = insight_1, insight_2, insight_3, insight_4
    resolved_greeting = (greeting_name.strip() or dealership_name.strip() or "there").strip()
    sms_effective = "" if omit_sms_from_email else sms_version

    core_email = build_email_context_object(
        dealership_name=dealership_name,
        greeting_name=resolved_greeting,
        month_label=month_label,
        priority_insights=priority_insights,
        summary=summary,
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name,
        sender_title=sender_title,
        email=to_recipients,
        cc=cc_recipients,
    )

    if template_type == "html":
        p1, p2, p3, p4 = point_1, point_2, point_3, point_4
        context = {
            "DEALERSHIP_NAME": html.escape(dealership_name.strip()),
            "DEALERSHIP": html.escape(dealership_name.strip()),
            "NAME": html.escape(resolved_greeting),
            "MONTH": html.escape(month_label.strip()),
            "MONTH_LABEL": html.escape(month_label.strip()),
            "EXECUTIVE_SUMMARY": text_block_to_html(executive_summary),
            "PRIORITY_INSIGHTS": render_priority_insights_html(priority_insights),
            "SUMMARY": text_block_to_html(summary),
            "SMS_VERSION": text_block_to_html(sms_effective),
            "INSIGHT_1": html.escape(insight_1),
            "INSIGHT_2": html.escape(insight_2),
            "INSIGHT_3": html.escape(insight_3),
            "INSIGHT_4": html.escape(insight_4),
            "POINT_1": html.escape(p1),
            "POINT_2": html.escape(p2),
            "POINT_3": html.escape(p3),
            "POINT_4": html.escape(p4),
            "SALES_REPORT_PATH": html.escape(sales_report_path.strip()),
            "SERVICE_REPORT_PATH": html.escape(service_report_path.strip()),
            "SENDER_NAME": html.escape((sender_name.strip() or DEFAULT_SENDER_NAME)),
            "SENDER_TITLE": html.escape((sender_title.strip() or DEFAULT_SENDER_TITLE)),
            "SENTIMENT": html.escape(sentiment.strip()),
            "REVIEW_REASON_TAGS": html.escape(review_reason_text),
            "TO": html.escape(to_recipients.strip()),
            "CC": html.escape(cc_recipients.strip()),
            "SUBJECT": html.escape(subject.strip()),
        }
        context.update({key.lower(): value for key, value in context.items()})
        # Preferred lowercase keys mirror plain-text values where appropriate
        context["dealership"] = html.escape(core_email["dealership"])
        context["name"] = html.escape(core_email["name"])
        context["month"] = html.escape(core_email["month"])
        context["insight_1"] = html.escape(insight_1)
        context["insight_2"] = html.escape(insight_2)
        context["insight_3"] = html.escape(insight_3)
        context["insight_4"] = html.escape(insight_4)
        context["point_1"] = html.escape(point_1)
        context["point_2"] = html.escape(point_2)
        context["point_3"] = html.escape(point_3)
        context["point_4"] = html.escape(point_4)
        context["summary"] = text_block_to_html(summary)
        context["sales_report_path"] = context["SALES_REPORT_PATH"]
        context["service_report_path"] = context["SERVICE_REPORT_PATH"]
        context["sender_name"] = context["SENDER_NAME"]
        context["sender_title"] = context["SENDER_TITLE"]
        return context

    context = {
        "DEALERSHIP_NAME": dealership_name.strip(),
        "DEALERSHIP": dealership_name.strip(),
        "NAME": resolved_greeting,
        "MONTH": month_label.strip(),
        "MONTH_LABEL": month_label.strip(),
        "EXECUTIVE_SUMMARY": normalize_text_for_preview(executive_summary),
        "PRIORITY_INSIGHTS": render_priority_insights_text(priority_insights),
        "SUMMARY": normalize_text_for_preview(summary),
        "SMS_VERSION": normalize_text_for_preview(sms_effective),
        "INSIGHT_1": insight_1,
        "INSIGHT_2": insight_2,
        "INSIGHT_3": insight_3,
        "INSIGHT_4": insight_4,
        "POINT_1": point_1,
        "POINT_2": point_2,
        "POINT_3": point_3,
        "POINT_4": point_4,
        "SALES_REPORT_PATH": sales_report_path.strip(),
        "SERVICE_REPORT_PATH": service_report_path.strip(),
        "SENDER_NAME": sender_name.strip() or DEFAULT_SENDER_NAME,
        "SENDER_TITLE": sender_title.strip() or DEFAULT_SENDER_TITLE,
        "SENTIMENT": sentiment.strip(),
        "REVIEW_REASON_TAGS": review_reason_text,
        "TO": to_recipients.strip(),
        "CC": cc_recipients.strip(),
        "SUBJECT": subject.strip(),
    }
    context.update({key.lower(): value for key, value in context.items()})
    context.update(core_email)
    context["insight_1"] = insight_1
    context["insight_2"] = insight_2
    context["insight_3"] = insight_3
    context["insight_4"] = insight_4
    context["point_1"] = point_1
    context["point_2"] = point_2
    context["point_3"] = point_3
    context["point_4"] = point_4
    return context


def render_template(template_text: str, context: dict[str, str]) -> tuple[str, list[str], list[str]]:
    explicit_found = [normalize_placeholder_name(match.group(1)) for match in EXPLICIT_PLACEHOLDER_PATTERN.finditer(template_text)]
    legacy_found = [normalize_placeholder_name(match.group(1)) for match in LEGACY_PLACEHOLDER_PATTERN.finditer(template_text)]
    bracket_found = [normalize_placeholder_name(match.group(1)) for match in SQUARE_BRACKET_PLACEHOLDER_PATTERN.finditer(template_text)]
    found_placeholders = explicit_found + legacy_found + bracket_found
    normalized_template = _convert_legacy_placeholders_to_jinja(template_text)
    rendered = _render_jinja_template(normalized_template, context)

    replaced = []
    for placeholder in found_placeholders:
        if placeholder in context and placeholder not in replaced:
            replaced.append(placeholder)

    missing = [placeholder for placeholder in found_placeholders if placeholder not in context]
    ordered_missing: list[str] = []
    seen_missing: set[str] = set()
    for placeholder in missing:
        if placeholder in seen_missing:
            continue
        seen_missing.add(placeholder)
        ordered_missing.append(placeholder)
    return rendered, replaced, ordered_missing


def render_email(
    *,
    template_path: str | Path | None,
    dealership_name: str,
    month_label: str,
    executive_summary: str,
    priority_insights: Any,
    summary: str,
    sms_version: str,
    sentiment: str,
    review_reason_tags: list[str],
    to_recipients: str = "",
    cc_recipients: str = "",
    attachments: list[Path] | None = None,
    sales_report_path: str = "",
    service_report_path: str = "",
    sender_name: str = "",
    sender_title: str = "",
    greeting_name: str = "",
    omit_sms_from_email: bool = True,
) -> RenderedEmail:
    loaded_template = load_template_file(template_path)
    validate_template_file(template_path)
    pi_lines = extract_priority_insight_lines(priority_insights)
    logging.info(
        "Jinja email render | template=%s | type=%s | greeting=%r | omit_sms=%s | pi_type=%s | pi_line_count=%s | summary_len=%s",
        loaded_template.template_path or "(built-in)",
        loaded_template.template_type,
        greeting_name,
        omit_sms_from_email,
        type(priority_insights).__name__,
        len(pi_lines),
        len(str(summary or "")),
    )
    if pi_lines:
        logging.debug("Priority insights preview | %r", pi_lines[:6])

    subject_context = build_template_context(
        dealership_name=dealership_name,
        month_label=month_label,
        executive_summary=executive_summary,
        priority_insights=priority_insights,
        summary=summary,
        sms_version=sms_version,
        sentiment=sentiment,
        review_reason_tags=review_reason_tags,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        template_type="text",
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name,
        sender_title=sender_title,
        greeting_name=greeting_name,
        omit_sms_from_email=omit_sms_from_email,
    )
    subject, subject_replaced, subject_missing = render_template(loaded_template.subject_template, subject_context)
    subject = normalize_text_for_preview(subject) or normalize_text_for_preview(
        f"DealerSocket {month_label} Insights - {dealership_name}"
    )

    body_context = build_template_context(
        dealership_name=dealership_name,
        month_label=month_label,
        executive_summary=executive_summary,
        priority_insights=priority_insights,
        summary=summary,
        sms_version=sms_version,
        sentiment=sentiment,
        review_reason_tags=review_reason_tags,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        subject=subject,
        template_type=loaded_template.template_type,
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name,
        sender_title=sender_title,
        greeting_name=greeting_name,
        omit_sms_from_email=omit_sms_from_email,
    )
    rendered_body, body_replaced, body_missing = render_template(loaded_template.body_template, body_context)
    preferred_missing = validate_preferred_body_variables(loaded_template.body_template, body_context)
    if preferred_missing:
        msg = (
            "The email template references one or more preferred fields that are empty: "
            + ", ".join(preferred_missing)
            + ". Update the report content or contacts before sending."
        )
        logging.warning("Preferred template variables empty | %s", msg)

    if loaded_template.template_type == "html":
        html_body = rendered_body.strip()
        if not template_has_insight_or_summary_slots(loaded_template.body_template):
            html_body = inject_insights_and_summary_into_html_if_missing(
                html_body,
                insight_1=str(body_context.get("insight_1", "")),
                insight_2=str(body_context.get("insight_2", "")),
                insight_3=str(body_context.get("insight_3", "")),
                insight_4=str(body_context.get("insight_4", "")),
                summary=str(body_context.get("summary", "")),
            )
        preview_text = html_to_preview_text(html_body)
        body_text = preview_text
    else:
        body_text = normalize_text_for_preview(rendered_body)
        html_body = convert_text_to_html(body_text)
        preview_text = body_text

    # If this came from a .msg template, apply extra cleaning to the preview text only
    if loaded_template.encoding == "msg":
        preview_text = _clean_preview_text_outlook_junk(preview_text)
        if loaded_template.template_type != "html":
            body_text = _clean_msg_plain_body(body_text)

    had_raw_insight_lines = len(pi_lines) > 0
    summary_plain = normalize_text_for_preview(summary)
    audit_warnings, render_contains_insights, render_contains_summary = validate_rendered_email_mapping(
        preview_text=preview_text,
        html_body=html_body,
        insight_1=str(body_context.get("insight_1", "")),
        insight_2=str(body_context.get("insight_2", "")),
        insight_3=str(body_context.get("insight_3", "")),
        insight_4=str(body_context.get("insight_4", "")),
        summary_resolved=summary_plain,
        sms_version=sms_version,
        omit_sms_from_email=omit_sms_from_email,
        had_raw_insight_lines=had_raw_insight_lines,
    )
    if audit_warnings:
        logging.warning("Email render validation | %s", " | ".join(audit_warnings))

    attachment_paths = [Path(path) for path in (attachments or []) if Path(path).exists()]
    placeholders_found = []
    for placeholder in loaded_template.placeholders_found:
        if placeholder not in placeholders_found:
            placeholders_found.append(placeholder)
    replaced_placeholders = []
    for placeholder in [*subject_replaced, *body_replaced]:
        if placeholder not in replaced_placeholders:
            replaced_placeholders.append(placeholder)
    missing_placeholders = []
    for placeholder in [*subject_missing, *body_missing]:
        if placeholder not in missing_placeholders:
            missing_placeholders.append(placeholder)

    return RenderedEmail(
        to=to_recipients.strip(),
        cc=cc_recipients.strip(),
        subject=subject,
        body=preview_text,
        body_text=body_text,
        html_body=html_body,
        preview_text=preview_text,
        attachments=attachment_paths,
        template_type=loaded_template.template_type,
        template_path=str(loaded_template.template_path) if loaded_template.template_path else "",
        template_encoding=loaded_template.encoding,
        placeholders_found=placeholders_found,
        replaced_placeholders=replaced_placeholders,
        missing_placeholders=missing_placeholders,
        render_audit_warnings=audit_warnings,
        render_contains_insights=render_contains_insights,
        render_contains_summary=render_contains_summary,
    )


def build_email_preview(
    *,
    template_path: str | Path | None,
    dealership_name: str,
    month_label: str,
    executive_summary: str,
    priority_insights: Any,
    summary: str,
    sms_version: str,
    sentiment: str,
    review_reason_tags: list[str],
    to_recipients: str = "",
    cc_recipients: str = "",
    attachments: list[Path] | None = None,
    greeting_name: str = "",
    sales_report_path: str = "",
    service_report_path: str = "",
    sender_name: str = "",
    sender_title: str = "",
) -> dict[str, Any]:
    rendered_email = render_email(
        template_path=template_path,
        dealership_name=dealership_name,
        month_label=month_label,
        executive_summary=executive_summary,
        priority_insights=priority_insights,
        summary=summary,
        sms_version=sms_version,
        sentiment=sentiment,
        review_reason_tags=review_reason_tags,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        attachments=attachments,
        sales_report_path=sales_report_path,
        service_report_path=service_report_path,
        sender_name=sender_name,
        sender_title=sender_title,
        greeting_name=greeting_name,
        omit_sms_from_email=True,
    )
    return {
        "to": rendered_email.to,
        "cc": rendered_email.cc,
        "subject": rendered_email.subject,
        "body": rendered_email.body,
        "body_text": rendered_email.body_text,
        "html_body": rendered_email.html_body,
        "preview_text": rendered_email.preview_text,
        "attachments": rendered_email.attachments,
        "template_type": rendered_email.template_type,
        "template_path": rendered_email.template_path,
        "template_encoding": rendered_email.template_encoding,
        "placeholders_found": rendered_email.placeholders_found,
        "replaced_placeholders": rendered_email.replaced_placeholders,
        "missing_placeholders": rendered_email.missing_placeholders,
        "render_audit_warnings": rendered_email.render_audit_warnings,
        "render_contains_insights": rendered_email.render_contains_insights,
        "render_contains_summary": rendered_email.render_contains_summary,
    }



def join_addresses(values: list[str]) -> str:
    """De-duplicate case-insensitively, keep the order they were written in."""
    seen: set[str] = set()
    kept: list[str] = []
    for value in values:
        low = value.lower()
        if low in seen:
            continue
        seen.add(low)
        kept.append(value)
    return "; ".join(kept)


def merge_contact_rows(first: dict[str, str], second: dict[str, str]) -> dict[str, str]:
    """Combine two contacts.csv rows naming the same dealership.

    A second row used to overwrite the first outright, so adding a colleague to a
    dealership silently removed the person already there — no error, and only visible
    once an email went to the wrong person. Every address field is unioned instead.
    """
    merged = dict(first)
    for field_name in ("email", "cc", "bcc"):
        merged[field_name] = join_addresses(
            split_addresses(first.get(field_name, "")) + split_addresses(second.get(field_name, ""))
        )
    merged["to"] = merged["email"]

    # The greeting names whoever the email actually goes to.
    names = [n for n in (str(first.get("contact_name") or "").strip(), str(second.get("contact_name") or "").strip()) if n]
    unique_names: list[str] = []
    for name in names:
        if name.lower() not in {n.lower() for n in unique_names}:
            unique_names.append(name)
    if len(unique_names) == 1:
        merged["contact_name"] = unique_names[0]
    elif len(unique_names) == 2:
        merged["contact_name"] = f"{unique_names[0]} and {unique_names[1]}"
    elif unique_names:
        merged["contact_name"] = ", ".join(unique_names[:-1]) + f" and {unique_names[-1]}"

    # Active if any row says so: the intent of adding a live contact is to send.
    actives = [parse_active_flag(str(row.get("active", "yes"))) for row in (first, second)]
    merged["active"] = "yes" if any(flag is not False for flag in actives) else "no"
    for field_name in ("notes", "last_updated"):
        merged[field_name] = str(first.get(field_name) or "").strip() or str(second.get(field_name) or "").strip()
    return merged


def _parse_mapping_csv(mapping_path: Path) -> dict[str, dict[str, str]]:
    mappings: dict[str, dict[str, str]] = {}
    with mapping_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            row_lower = {str(key).strip().lower(): ("" if value is None else str(value).strip()) for key, value in row.items()}
            dealership_name = row_lower.get("dealership") or row_lower.get("dealership_name") or ""
            if not dealership_name:
                continue
            row_obj = normalize_contact_row(row_lower)
            key = normalize_name(dealership_name)
            existing = mappings.get(key)
            if existing is not None:
                row_obj = merge_contact_rows(existing, row_obj)
            mappings[key] = row_obj
            deal_exact = dealership_name.strip()
            if deal_exact:
                mappings[deal_exact] = row_obj
    return mappings



def _parse_mapping_json(mapping_path: Path) -> dict[str, dict[str, str]]:
    raw = json.loads(mapping_path.read_text(encoding="utf-8"))
    mappings: dict[str, dict[str, str]] = {}
    if isinstance(raw, dict):
        for key, value in raw.items():
            if isinstance(value, dict):
                row_lower = {str(k).lower(): ("" if v is None else str(v).strip()) for k, v in value.items()}
                row_lower.setdefault("dealership", str(key))
                row_obj = normalize_contact_row(row_lower)
                nk = normalize_name(str(key))
                mappings[nk] = row_obj
                ek = str(key).strip()
                if ek:
                    mappings[ek] = row_obj
    elif isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                continue
            row_lower = {str(k).lower(): ("" if v is None else str(v).strip()) for k, v in item.items()}
            dealership_name = str(item.get("dealership") or item.get("dealership_name") or "").strip()
            if not dealership_name:
                continue
            row_obj = normalize_contact_row(row_lower)
            mappings[normalize_name(dealership_name)] = row_obj
            if dealership_name.strip():
                mappings[dealership_name.strip()] = row_obj
    return mappings



def load_contact_rows(mapping_path: str | Path | None) -> list[dict[str, str]]:
    if not mapping_path:
        return []

    resolved = Path(mapping_path).expanduser().resolve()
    if not resolved.exists():
        return []

    suffix = resolved.suffix.lower()
    rows: list[dict[str, str]] = []
    if suffix == ".csv":
        with resolved.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                row_lower = {str(k).strip().lower(): ("" if v is None else str(v).strip()) for k, v in row.items()}
                rows.append(normalize_contact_row(row_lower))
        return rows
    if suffix == ".json":
        raw = json.loads(resolved.read_text(encoding="utf-8"))
        if isinstance(raw, list):
            for item in raw:
                if not isinstance(item, dict):
                    continue
                row_lower = {str(k).lower(): ("" if v is None else str(v).strip()) for k, v in item.items()}
                dealership_name = str(item.get("dealership") or item.get("dealership_name") or "").strip()
                if not dealership_name:
                    continue
                rows.append(normalize_contact_row(row_lower))
        elif isinstance(raw, dict):
            for key, value in raw.items():
                if not isinstance(value, dict):
                    continue
                row_lower = {str(k).lower(): ("" if v is None else str(v).strip()) for k, v in value.items()}
                row_lower.setdefault("dealership", str(key))
                rows.append(normalize_contact_row(row_lower))
        return rows
    raise ValueError("Contact file must be a CSV or JSON file.")


def validate_contact_row_for_save(row: dict[str, str]) -> list[str]:
    errors: list[str] = []
    dealership = str(row.get("dealership") or "").strip()
    contact_name = str(row.get("contact_name") or "").strip()
    email = str(row.get("email") or row.get("to") or "").strip()
    active_raw = str(row.get("active") or "yes").strip() or "yes"
    active = parse_active_flag(active_raw)
    if not dealership:
        errors.append("dealership is required")
    if not contact_name:
        errors.append("contact_name is required")
    if active is None:
        errors.append("active must be yes/no or true/false")
    is_active = active is not False
    if is_active and email and not is_probable_email_list(email):
        errors.append("email must look like a valid address when the contact is active")
    if is_active and not email:
        errors.append("email is required when the contact is active")
    return errors


def save_contact_rows(mapping_path: str | Path, rows: list[dict[str, str]]) -> None:
    resolved = Path(mapping_path).expanduser().resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    normalized_rows: list[dict[str, str]] = []
    for row in rows:
        dealership = str(row.get("dealership") or "").strip()
        if not dealership:
            continue
        email = str(row.get("email") or row.get("to") or "").strip()
        normalized_rows.append(
            {
                "dealership": dealership,
                "contact_name": str(row.get("contact_name") or "").strip(),
                "email": email,
                "cc": str(row.get("cc") or "").strip(),
                "bcc": str(row.get("bcc") or "").strip(),
                "active": str(row.get("active") or "yes").strip() or "yes",
                "notes": str(row.get("notes") or "").strip(),
                "last_updated": str(row.get("last_updated") or "").strip(),
            }
        )
    suffix = resolved.suffix.lower()
    if suffix not in {".csv", ".json"}:
        raise ValueError("Contact file must be a CSV or JSON file.")
    temporary: Path | None = None
    try:
        # Same-directory replacement is atomic. Never truncate the live contacts file
        # before the entire new version has been written successfully.
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="", dir=resolved.parent,
            prefix=f".{resolved.name}.", suffix=".tmp", delete=False,
        ) as handle:
            temporary = Path(handle.name)
            if suffix == ".csv":
                writer = csv.DictWriter(handle, fieldnames=list(CONTACT_CSV_FIELDNAMES))
                writer.writeheader()
                writer.writerows(normalized_rows)
            else:
                json.dump(normalized_rows, handle, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, resolved)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def load_recipient_mappings(mapping_path: str | Path | None) -> dict[str, dict[str, str]]:
    if not mapping_path:
        return {}

    resolved = Path(mapping_path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Recipient mapping file not found: {resolved}")

    suffix = resolved.suffix.lower()
    if suffix == ".csv":
        return _parse_mapping_csv(resolved)
    if suffix == ".json":
        return _parse_mapping_json(resolved)
    raise ValueError("Recipient mapping file must be a CSV or JSON file.")



def lookup_contact_row(dealership_name: str, mapping: dict[str, dict[str, str]] | None) -> dict[str, str] | None:
    """Resolve contacts.csv row: exact key, normalised key, then case-insensitive dealership column match."""
    mapping = mapping or {}
    d = (dealership_name or "").strip()
    if not d:
        return None
    row = mapping.get(d) or mapping.get(normalize_name(d)) or mapping.get(d.lower())
    if row:
        return row
    dn = normalize_name(d)
    d_lower = d.lower()
    seen: set[int] = set()
    for cand in mapping.values():
        rid = id(cand)
        if rid in seen:
            continue
        seen.add(rid)
        rd = str(cand.get("dealership") or "").strip()
        if not rd:
            continue
        if rd == d or normalize_name(rd) == dn or rd.lower() == d_lower:
            return cand
    return None


def resolve_contact_for_dealership(
    dealership_name: str,
    mapping: dict[str, dict[str, str]] | None,
    *,
    manual_to: str = "",
    manual_cc: str = "",
    manual_bcc: str = "",
) -> ResolvedContact:
    d = dealership_name.strip()
    row = lookup_contact_row(dealership_name, mapping)
    cname = str(row.get("contact_name") or "").strip() if row else ""
    row_email = str(row.get("email") or row.get("to") or "").strip() if row else ""
    row_cc = str(row.get("cc") or "").strip() if row else ""
    row_bcc = str(row.get("bcc") or "").strip() if row else ""

    manual_to_s = manual_to.strip()
    manual_cc_s = manual_cc.strip()
    manual_bcc_s = manual_bcc.strip()

    if manual_to_s or manual_cc_s or manual_bcc_s:
        ready = bool(manual_to_s and is_probable_email_list(manual_to_s))
        if not row:
            logging.info(
                "Contact resolution | manual override | no contact row | dealership=%r | ready=%s",
                dealership_name,
                ready,
            )
            return ResolvedContact(manual_to_s, manual_cc_s, "", d, False, ready, False, manual_bcc_s)

        active = parse_active_flag(str(row.get("active", "yes")))
        if active is None:
            logging.warning("Contact row active flag not recognised for %s; defaulting to active.", dealership_name)
            active = True
        if not active:
            logging.info("Contact resolution | manual override | inactive row | dealership=%r", dealership_name)
            return ResolvedContact(manual_to_s, manual_cc_s, cname, d, True, ready, False, manual_bcc_s)
        greeting = (cname or d) if cname else d
        logging.info("Contact resolution | manual override | dealership=%r | ready=%s", dealership_name, ready)
        return ResolvedContact(manual_to_s, manual_cc_s, cname, greeting, True, ready, True, manual_bcc_s)

    if not row:
        logging.info("Contact resolution | no row | dealership=%r", dealership_name)
        return ResolvedContact("", "", "", "", False, False, False, "")

    active = parse_active_flag(str(row.get("active", "yes")))
    if active is None:
        logging.warning("Contact row active flag not recognised for %s; defaulting to active.", dealership_name)
        active = True

    greeting_with_contact = (cname or d) if cname else d
    if not active:
        logging.info("Contact resolution | inactive | dealership=%r", dealership_name)
        return ResolvedContact("", "", cname, d, True, False, False, "")

    if not row_email or not is_probable_email_list(row_email):
        logging.info("Contact resolution | missing or invalid email | dealership=%r", dealership_name)
        return ResolvedContact("", "", cname, greeting_with_contact, True, False, True, "")

    logging.info("Contact resolution | matched | dealership=%r | email=%r", dealership_name, row_email)
    return ResolvedContact(row_email, row_cc, cname, greeting_with_contact, True, True, True, row_bcc)


def resolve_recipients(
    dealership_name: str,
    mapping: dict[str, dict[str, str]] | None,
    *,
    manual_to: str = "",
    manual_cc: str = "",
) -> tuple[str, str]:
    resolved = resolve_contact_for_dealership(
        dealership_name,
        mapping,
        manual_to=manual_to,
        manual_cc=manual_cc,
    )
    return resolved.to, resolved.cc



def create_outlook_draft(
    *,
    to_recipients: str,
    cc_recipients: str,
    # Defaulted so any caller written before BCC existed still works — the desktop app
    # goes through report_engine, but nothing should break for want of one argument.
    bcc_recipients: str = "",
    subject: str,
    body_text: str,
    html_body: str,
    attachments: list[Path],
    inline_cid_attachments: list[tuple[str, Path]] | None = None,
    send: bool = False,
) -> dict[str, Any]:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - Windows-specific integration
        raise RuntimeError("pywin32 is required for Outlook draft creation. Install it and ensure classic Outlook is available.") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail_item = outlook.CreateItem(0)
    mail_item.To = to_recipients
    mail_item.CC = cc_recipients
    mail_item.BCC = bcc_recipients
    mail_item.Subject = subject

    pr_cid = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    for cid, path in inline_cid_attachments or []:
        resolved = Path(path)
        if not resolved.exists():
            logging.warning("Outlook draft | skipping missing inline file | cid=%s path=%s", cid, resolved)
            continue
        att = mail_item.Attachments.Add(str(resolved))
        try:
            pa = att.PropertyAccessor
            pa.SetProperty(pr_cid, (cid or "").strip())
        except Exception:
            logging.warning("Outlook inline Content-ID not set for %s (images may not appear inline).", resolved)

    for attachment in attachments:
        if Path(attachment).exists():
            mail_item.Attachments.Add(str(Path(attachment)))

    final_html = (html_body or "").strip()
    if not final_html:
        logging.warning(
            "Outlook draft | empty HTML supplied; converting plain text only — branded layout will be lost.",
        )
        final_html = convert_text_to_html(body_text)
    final_html = ensure_outlook_html_document(final_html)
    mail_item.HTMLBody = final_html
    # Do not set mail_item.Body after HTMLBody — Outlook derives plain text from HTML.

    logging.info(
        "Outlook draft | HTMLBody length=%s | attachment_count=%s | inline_cid=%s",
        len(str(getattr(mail_item, "HTMLBody", "") or "")),
        len(attachments),
        len(inline_cid_attachments or []),
    )

    if send:
        # Save first so the message survives as a draft if Send is blocked by
        # Outlook's programmatic-access guard or a transport error.
        mail_item.Save()
        mail_item.Send()
        logging.info("Outlook | SENT | to=%s | subject=%s", to_recipients, subject)
    else:
        mail_item.Save()

    return {
        "to": to_recipients,
        "cc": cc_recipients,
        "bcc": bcc_recipients,
        "subject": subject,
        "attachment_count": len(attachments),
        "saved": True,
        "sent": bool(send),
    }
