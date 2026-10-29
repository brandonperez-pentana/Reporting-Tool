"""Write a finished email to disk as .eml, for importing into Outlook by hand.

This is the manual path. Everything else in Reporter reaches Outlook through COM,
which needs Outlook installed, registered and cooperating; when it is not — a new
machine, a broken registration, an Office repair pending — there is no way to get
the month's work out. An .eml is the escape hatch: it is a plain MIME file written
with the standard library, so it is produced identically on any machine with no
Outlook, no pywin32 and no third-party package involved.

The file carries the same body, the same inline branding images and the same PDF
attachments the Outlook draft would have had, so nothing is lost by taking this
route. `X-Unsent: 1` is what makes it useful rather than an archive format: Outlook
reads that header and opens the file as an editable, sendable draft rather than as a
received message, so double-clicking one puts the operator in the compose window
with the recipients, subject, body and attachments already filled in.

Deliberately not .msg: writing a real .msg requires driving Outlook, which is the
thing that is unavailable in every situation this exists for.
"""

from __future__ import annotations

import mimetypes
import re
from dataclasses import dataclass
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

# Windows forbids these outright; the rest is trimmed so a dealership name with a
# slash or a trailing dot cannot produce a file the operator has to rename by hand.
_UNSAFE_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MAX_STEM_LENGTH = 120


@dataclass(frozen=True)
class ExportedEmail:
    path: Path
    attachment_names: list[str]
    inline_asset_count: int
    bytes_written: int


def safe_filename_stem(text: str, fallback: str = "email") -> str:
    """Turn a dealership and month into something Windows will accept as a filename."""
    cleaned = _UNSAFE_FILENAME.sub("-", str(text or "")).strip(" .")
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned[:MAX_STEM_LENGTH].strip(" .")
    return cleaned or fallback


def unique_path(folder: Path, stem: str, suffix: str = ".eml") -> Path:
    """A path that does not exist yet, so re-running never silently overwrites a file.

    The operator may export, fix one email, and export again; the second run must not
    destroy the first without saying so.
    """
    candidate = folder / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    for n in range(2, 1000):
        candidate = folder / f"{stem} ({n}){suffix}"
        if not candidate.exists():
            return candidate
    raise OSError(f"Could not find a free filename for {stem}{suffix} in {folder}")


def _guess_type(path: Path, default: tuple[str, str]) -> tuple[str, str]:
    guessed, _ = mimetypes.guess_type(path.name)
    if not guessed or "/" not in guessed:
        return default
    maintype, subtype = guessed.split("/", 1)
    return maintype, subtype


def build_message(
    *,
    subject: str,
    to_recipients: str,
    cc_recipients: str = "",
    bcc_recipients: str = "",
    body_text: str = "",
    html_body: str = "",
    inline_assets: list[tuple[str, Path]] | None = None,
    attachments: list[Path] | None = None,
    sender: str = "",
) -> tuple[EmailMessage, list[str], int]:
    """Assemble the MIME message. Returns it with the attachment names actually added.

    The names are returned rather than assumed because a path that has gone missing
    since the report ran is skipped, and the operator needs to see that it was.
    """
    message = EmailMessage()
    message["Subject"] = str(subject or "")
    if str(to_recipients or "").strip():
        message["To"] = str(to_recipients).strip()
    if str(cc_recipients or "").strip():
        message["Cc"] = str(cc_recipients).strip()
    if str(bcc_recipients or "").strip():
        # Outlook reads Bcc from a draft it opens and keeps it hidden from recipients,
        # the same as one typed in the compose window.
        message["Bcc"] = str(bcc_recipients).strip()
    if str(sender or "").strip():
        message["From"] = str(sender).strip()
    message["Date"] = formatdate(localtime=True)
    # No Message-ID on purpose. Outlook stamps its own when the draft is actually sent,
    # and the stdlib's make_msgid() resolves the machine's FQDN to build one — a lookup
    # measured at over two seconds per call here, which a 27-email export would pay 27
    # times over, and which puts the machine's hostname in a file for no benefit.
    # The header that makes Outlook open this as a draft to edit and send rather
    # than as a message that has already arrived.
    message["X-Unsent"] = "1"

    message.set_content(str(body_text or ""))

    inline_added = 0
    if str(html_body or "").strip():
        message.add_alternative(str(html_body), subtype="html")
        html_part = message.get_payload()[-1]
        for cid, path in inline_assets or []:
            source = Path(path)
            if not source.is_file():
                continue
            maintype, subtype = _guess_type(source, ("image", "png"))
            # The html refers to these as src="cid:banner_top", so the Content-ID
            # has to be exactly <banner_top> — add_related's angle brackets included.
            html_part.add_related(
                source.read_bytes(),
                maintype=maintype,
                subtype=subtype,
                cid=f"<{cid}>",
                filename=source.name,
            )
            inline_added += 1

    attached: list[str] = []
    for item in attachments or []:
        source = Path(item)
        if not source.is_file():
            continue
        maintype, subtype = _guess_type(source, ("application", "octet-stream"))
        message.add_attachment(
            source.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=source.name,
        )
        attached.append(source.name)

    return message, attached, inline_added


def write_eml(
    folder: str | Path,
    filename_stem: str,
    *,
    subject: str,
    to_recipients: str,
    cc_recipients: str = "",
    bcc_recipients: str = "",
    body_text: str = "",
    html_body: str = "",
    inline_assets: list[tuple[str, Path]] | None = None,
    attachments: list[Path] | None = None,
    sender: str = "",
) -> ExportedEmail:
    """Write one .eml into folder and report what went into it."""
    target_folder = Path(folder).expanduser()
    target_folder.mkdir(parents=True, exist_ok=True)

    message, attached, inline_count = build_message(
        subject=subject,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        bcc_recipients=bcc_recipients,
        body_text=body_text,
        html_body=html_body,
        inline_assets=inline_assets,
        attachments=attachments,
        sender=sender,
    )

    path = unique_path(target_folder, safe_filename_stem(filename_stem))
    raw = message.as_bytes()
    path.write_bytes(raw)
    return ExportedEmail(
        path=path,
        attachment_names=attached,
        inline_asset_count=inline_count,
        bytes_written=len(raw),
    )
