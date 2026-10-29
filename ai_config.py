"""Where the optional AI layer gets its credentials and settings.

Nothing in the reporting pipeline calls this yet. It exists so that when insight
copy starts asking a model for help, the connection details are already in one
place, already off by default, and already impossible to commit by accident.

The split is deliberate:

  * Secrets live in a ``.env`` file (or in real environment variables), never in
    ``user_settings.json``. ``user_settings.json`` travels between machines in
    support conversations and screenshots; an API key must not.
  * The on/off switch lives in ``user_settings.json`` next to every other
    operator preference, so the UI can toggle it without touching a secret.

A run is only allowed to call a model when both halves agree: a key is present
*and* the operator has turned it on. A missing or malformed ``.env`` is never an
error -- it just means the AI layer stays off, which is the shipped default.

Load order, first hit wins per key:

  1. Real environment variables (a CI job or a shell ``set``)
  2. ``.env`` beside the app
  3. ``.env`` in the per-user data folder, for read-only installs

Check what the app can see::

    python ai_config.py

That prints the resolved settings with the key masked, and never prints the key
itself. Use :func:`redact` before putting any AI-related string in a log.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ENV_FILE_NAME = ".env"
PREFIX = "REPORTER_AI_"

# Sensible endpoints per provider so a .env only has to carry a key and a model.
# "custom" covers anything OpenAI-compatible: a gateway, a local llama.cpp server,
# an internal proxy. It has no default endpoint on purpose -- if you pick custom you
# have to say where.
PROVIDER_DEFAULTS: dict[str, dict[str, str]] = {
    "anthropic": {
        "base_url": "https://api.anthropic.com",
        "model": "claude-sonnet-4-5",
        "api_version": "2023-06-01",
    },
    "openai": {
        "base_url": "https://api.openai.com",
        "model": "gpt-4o-mini",
        "api_version": "",
    },
    "azure-openai": {
        "base_url": "",
        "model": "",
        "api_version": "2024-10-21",
    },
    "custom": {
        "base_url": "",
        "model": "",
        "api_version": "",
    },
}

TRUE_WORDS = {"1", "true", "yes", "on", "enabled"}


def _bool(raw: str, default: bool = False) -> bool:
    text = str(raw or "").strip().lower()
    if not text:
        return default
    return text in TRUE_WORDS


def _int(raw: str, default: int) -> int:
    try:
        return int(str(raw).strip())
    except (TypeError, ValueError):
        return default


def _float(raw: str, default: float) -> float:
    try:
        return float(str(raw).strip())
    except (TypeError, ValueError):
        return default


# --- .env reading ---------------------------------------------------------------

def parse_env_text(text: str) -> dict[str, str]:
    """Parse ``.env`` content into a mapping.

    Deliberately small: ``KEY=value`` one per line, ``#`` comments, optional
    ``export`` prefix, optional surrounding quotes. No variable interpolation, no
    multi-line values -- an API key needs none of it, and a parser that does more
    is a parser that surprises someone at 6pm on a Friday.
    """
    values: dict[str, str] = {}
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        if stripped.startswith("export "):
            stripped = stripped[len("export "):].lstrip()
        name, _, raw = stripped.partition("=")
        name = name.strip()
        if not name:
            continue
        value = raw.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        values[name] = value
    return values


def env_file_candidates() -> list[Path]:
    """Every place a ``.env`` is looked for, in precedence order."""
    app_folder = Path(__file__).resolve().parent
    candidates = [app_folder / ENV_FILE_NAME]
    try:
        import settings_manager

        candidates.append(settings_manager.user_data_dir() / ENV_FILE_NAME)
    except Exception:  # noqa: BLE001 - config must load even if settings cannot
        pass
    return candidates


def read_env_files() -> tuple[dict[str, str], list[Path]]:
    """Merge every ``.env`` that exists. Earlier files win over later ones."""
    merged: dict[str, str] = {}
    found: list[Path] = []
    for path in env_file_candidates():
        try:
            if not path.is_file():
                continue
            values = parse_env_text(path.read_text(encoding="utf-8-sig"))
        except OSError:
            continue
        found.append(path)
        for key, value in values.items():
            merged.setdefault(key, value)
    return merged, found


def _lookup(name: str, env_values: dict[str, str], default: str = "") -> str:
    """Real environment first, then whatever the .env files supplied."""
    full = f"{PREFIX}{name}"
    for candidate in (os.environ.get(full), env_values.get(full)):
        if candidate is not None and str(candidate).strip():
            return str(candidate).strip()
    return default


# --- the config object ----------------------------------------------------------

@dataclass(frozen=True)
class AIConfig:
    """Resolved AI connection settings. Never log this object directly."""

    provider: str = "anthropic"
    api_key: str = ""
    model: str = ""
    base_url: str = ""
    api_version: str = ""
    timeout: int = 60
    max_tokens: int = 1024
    temperature: float = 0.2
    enabled: bool = False
    sources: list[str] = field(default_factory=list)

    @property
    def configured(self) -> bool:
        """Enough detail present to attempt a call."""
        return bool(self.api_key and self.model and self.base_url)

    @property
    def active(self) -> bool:
        """Configured *and* switched on. The only state that permits a call."""
        return self.enabled and self.configured

    @property
    def masked_key(self) -> str:
        if not self.api_key:
            return ""
        tail = self.api_key[-4:] if len(self.api_key) > 8 else ""
        return f"{'*' * 8}{tail}"

    def why_inactive(self) -> str:
        """One line an operator can act on, or empty when the layer is usable."""
        if not self.api_key:
            return f"No API key. Set {PREFIX}API_KEY in a .env file beside the app."
        if not self.model:
            return f"No model set. Add {PREFIX}MODEL to the .env file."
        if not self.base_url:
            return f"No endpoint for provider '{self.provider}'. Add {PREFIX}BASE_URL."
        if not self.enabled:
            return "AI assistance is switched off in Settings."
        return ""

    def status(self) -> dict[str, Any]:
        """Safe to return over HTTP, put in a log, or paste into a support ticket."""
        return {
            "provider": self.provider,
            "model": self.model,
            "base_url": self.base_url,
            "api_key_set": bool(self.api_key),
            "api_key_masked": self.masked_key,
            "enabled": self.enabled,
            "configured": self.configured,
            "active": self.active,
            "reason": self.why_inactive(),
            "env_files": list(self.sources),
            "timeout": self.timeout,
            "max_tokens": self.max_tokens,
            "temperature": self.temperature,
        }


def load(*, enabled_override: bool | None = None) -> AIConfig:
    """Resolve the AI configuration from the environment, .env files and settings."""
    env_values, found = read_env_files()

    provider = _lookup("PROVIDER", env_values, "anthropic").lower()
    defaults = PROVIDER_DEFAULTS.get(provider, PROVIDER_DEFAULTS["custom"])

    if enabled_override is not None:
        enabled = enabled_override
    else:
        raw_enabled = _lookup("ENABLED", env_values)
        if raw_enabled:
            enabled = _bool(raw_enabled)
        else:
            enabled = _settings_enabled()

    return AIConfig(
        provider=provider,
        api_key=_lookup("API_KEY", env_values),
        model=_lookup("MODEL", env_values, defaults["model"]),
        base_url=_lookup("BASE_URL", env_values, defaults["base_url"]).rstrip("/"),
        api_version=_lookup("API_VERSION", env_values, defaults["api_version"]),
        timeout=_int(_lookup("TIMEOUT", env_values), 60),
        max_tokens=_int(_lookup("MAX_TOKENS", env_values), 1024),
        temperature=_float(_lookup("TEMPERATURE", env_values), 0.2),
        enabled=enabled,
        sources=[str(path) for path in found],
    )


def _settings_enabled() -> bool:
    """The operator-facing switch, read from user_settings.json.

    Choosing "AI assisted" in Settings > Insight engine *is* the switch: an operator
    who picks it should not then have to find a second toggle to make it do anything.
    ``ai_insights_enabled`` remains as an independent override for anything that wants
    the layer without changing the engine, and for turning it off from a build server.
    """
    try:
        import settings_manager

        settings = settings_manager.load_settings()
    except Exception:  # noqa: BLE001 - a missing settings file must not break config
        return False

    if str(settings.get("insight_engine", "")).strip().lower() == "ai":
        return True
    return _bool(str(settings.get("ai_insights_enabled", "")))


def redact(text: str, config: AIConfig | None = None) -> str:
    """Strip the API key out of a string before it reaches a log or the UI."""
    resolved = config or load()
    if resolved.api_key and resolved.api_key in text:
        return text.replace(resolved.api_key, resolved.masked_key or "***")
    return text


def example_env_path() -> Path:
    return Path(__file__).resolve().parent / ".env.example"


def _cli() -> int:
    config = load()
    status = config.status()
    print("Reporter AI configuration")
    print("-" * 40)
    print(f"  .env files read : {', '.join(status['env_files']) or 'none found'}")
    print(f"  provider        : {status['provider']}")
    print(f"  model           : {status['model'] or '(not set)'}")
    print(f"  endpoint        : {status['base_url'] or '(not set)'}")
    print(f"  api key         : {status['api_key_masked'] or '(not set)'}")
    print(f"  switched on     : {'yes' if status['enabled'] else 'no'}")
    print(f"  usable          : {'yes' if status['active'] else 'no'}")
    if status["reason"]:
        print(f"\n  {status['reason']}")
        example = example_env_path()
        if example.is_file():
            print(f"  Template to copy: {example}  ->  {example.with_suffix('')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_cli())
