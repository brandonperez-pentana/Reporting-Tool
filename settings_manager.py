from __future__ import annotations

import contextlib
import contextvars
import json
import os
import re
from pathlib import Path
from typing import Any, Iterator

SETTINGS_FILE_NAME = "user_settings.json"
SHARED_SETTINGS_FILE_NAME = "shared_settings.json"
DEFAULT_SETTINGS = {
    "accounts_folder": "",
    # Where the per-dealership folders of exported Sales/Service PDFs live. Empty means
    # they sit beside the workbooks, in the accounts folder.
    "pdf_source_folder": "",
    "output_folder": "",
    "recipient_mapping_path": "",
    "month": "",
    # Where PDF copies of sent emails are archived. Empty means the operator has
    # not chosen one yet and will be prompted at send time.
    "sent_email_pdf_folder": "",
    # Where "Save as .eml" writes approved emails for manual import into Outlook.
    "eml_export_folder": "",
    # "rules" (default) or "funnel". The funnel engine reads the month model and
    # replaces only the client-facing copy; everything else about a run is unchanged,
    # so both can be produced from the same reports and compared.
    "insight_engine": "rules",
    # Whether the Performance Intelligence report is built for each dealership from the
    # workbook's history tab. On by default; needs Chromium (see performance_report.py).
    "performance_report": "on",
    # Whether the optional AI layer may be called. Off in a fresh install. The
    # credentials live in a .env file (see ai_config.py and .env.example), never
    # here: this file gets copied between machines and pasted into tickets.
    "ai_insights_enabled": "off",
    # Addresses BCC'd on every email, on top of whatever a contact row asks for.
    # Empty in a fresh install: these are the operator's own colleagues, and they must
    # not travel inside a package handed to someone else.
    "default_bcc": "",
    # "My dealerships": workbook filenames in the accounts folder this operator reports
    # on. Empty means every workbook in the library.
    "portfolio": [],
    "window": {
        "width": 1600,
        "height": 960,
    },
}

# Saved per Windows user when Reporter runs on a server. Every other key is shared and
# only an administrator may change it.
USER_SETTING_KEYS = frozenset({
    "month",
    "portfolio",
    "output_folder",
    "sent_email_pdf_folder",
    # Keep legacy test-send values personal when migrating old settings to server mode.
    "test_send_address",
    "eml_export_folder",
    "default_bcc",
    "window",
})

_CURRENT_USER: contextvars.ContextVar[str | None] = contextvars.ContextVar("reporter_user", default=None)


def server_mode() -> bool:
    return os.environ.get("REPORTER_MODE", "").strip().lower() == "server"


def data_root() -> Path | None:
    raw = os.environ.get("REPORTER_DATA_ROOT", "").strip()
    return Path(raw).expanduser() if raw else None


def diagnostics_dir() -> Path:
    """Keep server diagnostics with server data, not in the shared code install."""
    root = data_root()
    if server_mode() and root is not None:
        return root / "logs"
    return Path(__file__).resolve().parent / "logs"


def current_user() -> str | None:
    return _CURRENT_USER.get()


@contextlib.contextmanager
def as_user(user: str | None) -> Iterator[None]:
    """Run the enclosed code as `user`, so every load_settings() call sees their settings."""
    token = _CURRENT_USER.set(user.casefold() if user else None)
    try:
        yield
    finally:
        _CURRENT_USER.reset(token)


def user_key(user: str) -> str:
    """A folder name for a login such as CORP\\jane.doe that cannot climb out of users/."""
    key = re.sub(r"[^a-z0-9._-]+", "_", str(user).casefold()).strip("._")
    if not key:
        raise ValueError(f"Unusable user name: {user!r}")
    return key


def is_admin(user: str | None = None) -> bool:
    who = user or current_user()
    if not who:
        # The single-user desktop copy: whoever runs it owns every setting.
        return True
    admins = {name.strip().casefold() for name in os.environ.get("REPORTER_ADMINS", "").split(",") if name.strip()}
    return who.casefold() in admins


def _multi_user() -> bool:
    return data_root() is not None and bool(current_user())


def user_folder(user: str | None = None) -> Path:
    root = data_root()
    who = user or current_user()
    if root is None or not who:
        raise RuntimeError("A per-user folder needs REPORTER_DATA_ROOT and a signed-in user.")
    return root / "users" / user_key(who)


def shared_settings_path() -> Path:
    root = data_root()
    if root is None:
        raise RuntimeError("Shared settings need REPORTER_DATA_ROOT.")
    return root / SHARED_SETTINGS_FILE_NAME


def _defaults() -> dict[str, Any]:
    return json.loads(json.dumps(DEFAULT_SETTINGS))


def _read_json(path: Path) -> dict[str, Any]:
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return loaded if isinstance(loaded, dict) else {}


def _overlay(settings: dict[str, Any], values: dict[str, Any]) -> None:
    for key, value in values.items():
        if key == "window" and isinstance(value, dict):
            settings["window"].update(value)
        else:
            settings[key] = value


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    # Write-then-rename, so two saves landing together cannot leave half a file.
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(temp, path)


def seed_shared_settings() -> Path | None:
    """First start on a server: carry the existing install's shared settings across."""
    root = data_root()
    if root is None:
        return None
    target = shared_settings_path()
    legacy = Path(__file__).resolve().parent / SETTINGS_FILE_NAME
    if target.exists() or not legacy.exists():
        return None
    shared = {key: value for key, value in _read_json(legacy).items() if key not in USER_SETTING_KEYS}
    _write_json(target, shared)
    return target


def user_data_dir() -> Path:
    """Per-user folder for state that must survive a read-only install."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_STATE_HOME")
    root = Path(base) if base else Path.home() / ".local" / "state"
    return root / "Reporter"


def _is_writable(folder: Path) -> bool:
    probe = folder / ".write_probe"
    try:
        folder.mkdir(parents=True, exist_ok=True)
        probe.write_text("", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def get_settings_path(settings_path: str | Path | None = None) -> Path:
    if settings_path is not None:
        return Path(settings_path).expanduser().resolve()
    if _multi_user():
        return user_folder() / SETTINGS_FILE_NAME
    app_folder = Path(__file__).resolve().parent
    beside_app = app_folder / SETTINGS_FILE_NAME
    # Keep settings beside the app where that works, which is the long-standing
    # behaviour and keeps an install self-contained. Fall back to the per-user
    # data folder when the install location is read-only (Program Files, a
    # managed share), rather than failing to save at all.
    if beside_app.exists() or _is_writable(app_folder):
        return beside_app
    return user_data_dir() / SETTINGS_FILE_NAME



def load_settings(settings_path: str | Path | None = None) -> dict[str, Any]:
    settings = _defaults()
    if settings_path is None and _multi_user():
        _overlay(settings, _read_json(shared_settings_path()))
        # A shared output folder would let one person's run clear another's files.
        settings["output_folder"] = ""
        personal = _read_json(get_settings_path())
        _overlay(settings, {key: value for key, value in personal.items() if key in USER_SETTING_KEYS})
        if not str(settings.get("output_folder") or "").strip():
            settings["output_folder"] = str(user_folder() / "Output")
        return settings

    _overlay(settings, _read_json(get_settings_path(settings_path)))
    return settings


def save_settings(
    settings: dict[str, Any],
    settings_path: str | Path | None = None,
    *,
    shared: bool = False,
) -> Path:
    """Merge `settings` into the file on disk.

    On a server, a user save keeps only USER_SETTING_KEYS and a shared save keeps only
    the rest, so neither file can hold the other's values.
    """
    if settings_path is None and _multi_user():
        if shared:
            path = shared_settings_path()
            changes = {key: value for key, value in settings.items() if key not in USER_SETTING_KEYS}
        else:
            path = get_settings_path()
            changes = {key: value for key, value in settings.items() if key in USER_SETTING_KEYS}
        current = _read_json(path)
        current.update(changes)
        _write_json(path, current)
        return path

    resolved_path = get_settings_path(settings_path)
    current = load_settings(resolved_path)
    current.update(settings)
    if "window" in settings and isinstance(settings["window"], dict):
        current["window"].update(settings["window"])

    _write_json(resolved_path, current)
    return resolved_path
