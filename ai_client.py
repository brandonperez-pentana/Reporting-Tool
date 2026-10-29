"""A thin call into whichever model ``ai_config`` points at.

Nothing in the reporting pipeline imports this. It is here so the configuration
added alongside it can be proved to work -- a key that has never been used is a
key you find out is wrong during a month-end run.

Two entry points:

  * :func:`check_connection` -- a cheap round trip for a "Test connection" button.
  * :func:`generate` -- send a prompt, get text back.

Both refuse to do anything unless the config is *active*: a key is present and the
operator has switched the layer on. That refusal is a returned message, not an
exception, everywhere it can be, because this must never take a month-end run down
with it. When insight copy eventually calls a model, it calls it as an enrichment
on top of copy that already stands on its own -- if this module fails, the report
is the same report it is today.

Standard library only, so nothing new lands in requirements-web.txt.
"""

from __future__ import annotations

import json
import logging
import urllib.error
import urllib.request
from typing import Any

import ai_config
from ai_config import AIConfig

LOGGER = logging.getLogger("reporter.ai")


class AIUnavailable(RuntimeError):
    """The layer is off, unconfigured, or the provider refused the call."""


def _headers(config: AIConfig) -> dict[str, str]:
    headers = {"content-type": "application/json"}
    if config.provider == "anthropic":
        headers["x-api-key"] = config.api_key
        headers["anthropic-version"] = config.api_version or "2023-06-01"
    elif config.provider == "azure-openai":
        headers["api-key"] = config.api_key
    else:
        headers["authorization"] = f"Bearer {config.api_key}"
    return headers


def _endpoint(config: AIConfig) -> str:
    if config.provider == "anthropic":
        return f"{config.base_url}/v1/messages"
    if config.provider == "azure-openai":
        version = config.api_version or "2024-10-21"
        return (
            f"{config.base_url}/openai/deployments/{config.model}"
            f"/chat/completions?api-version={version}"
        )
    return f"{config.base_url}/v1/chat/completions"


def _payload(config: AIConfig, prompt: str, system: str, max_tokens: int) -> dict[str, Any]:
    if config.provider == "anthropic":
        body: dict[str, Any] = {
            "model": config.model,
            "max_tokens": max_tokens,
            "temperature": config.temperature,
            "messages": [{"role": "user", "content": prompt}],
        }
        if system:
            body["system"] = system
        return body

    messages: list[dict[str, str]] = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    body = {
        "model": config.model,
        "max_tokens": max_tokens,
        "temperature": config.temperature,
        "messages": messages,
    }
    if config.provider == "azure-openai":
        body.pop("model")  # the deployment in the URL is the model
    return body


def _extract_text(config: AIConfig, data: dict[str, Any]) -> str:
    if config.provider == "anthropic":
        blocks = data.get("content") or []
        parts = [str(block.get("text", "")) for block in blocks if block.get("type") == "text"]
        return "\n".join(part for part in parts if part).strip()
    choices = data.get("choices") or []
    if not choices:
        return ""
    return str(choices[0].get("message", {}).get("content", "")).strip()


def generate(
    prompt: str,
    *,
    system: str = "",
    config: AIConfig | None = None,
    max_tokens: int | None = None,
) -> str:
    """Send a prompt and return the model's text.

    Raises :class:`AIUnavailable` when the layer is off or the provider errors, so a
    caller can decide between falling back to deterministic copy (what the pipeline
    should do) and surfacing the problem (what a Settings test button should do).
    """
    resolved = config or ai_config.load()
    if not resolved.active:
        raise AIUnavailable(resolved.why_inactive() or "AI assistance is unavailable.")

    body = _payload(resolved, prompt, system, max_tokens or resolved.max_tokens)
    request = urllib.request.Request(
        _endpoint(resolved),
        data=json.dumps(body).encode("utf-8"),
        headers=_headers(resolved),
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=resolved.timeout) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:400]
        except Exception:  # noqa: BLE001
            pass
        message = f"{resolved.provider} returned HTTP {exc.code}."
        if exc.code in (401, 403):
            message += " The API key was rejected. Check the key in your .env file."
        elif exc.code == 404:
            message += f" Check the model name ('{resolved.model}') and the endpoint."
        elif exc.code == 429:
            message += " Rate limited. Try again shortly."
        if detail:
            message += f" {ai_config.redact(detail, resolved)}"
        raise AIUnavailable(message) from exc
    except urllib.error.URLError as exc:
        raise AIUnavailable(
            f"Could not reach {resolved.base_url}: {exc.reason}. "
            "On a managed network this is usually a proxy or firewall rule."
        ) from exc
    except (TimeoutError, OSError) as exc:
        raise AIUnavailable(f"The request to {resolved.provider} failed: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise AIUnavailable(f"{resolved.provider} returned a response that was not JSON.") from exc

    text = _extract_text(resolved, data)
    if not text:
        raise AIUnavailable(f"{resolved.provider} returned an empty response.")
    return text


def check_connection(config: AIConfig | None = None) -> dict[str, Any]:
    """One cheap round trip. Returns a result rather than raising, for the UI."""
    resolved = config or ai_config.load()
    if not resolved.configured:
        return {"ok": False, "message": resolved.why_inactive(), "status": resolved.status()}

    # Bypass the on/off switch: an operator testing a key has not flipped it yet.
    probe = AIConfig(**{**resolved.__dict__, "enabled": True})
    try:
        reply = generate("Reply with the single word: ready.", config=probe, max_tokens=16)
    except AIUnavailable as exc:
        return {"ok": False, "message": str(exc), "status": resolved.status()}
    return {
        "ok": True,
        "message": f"Connected to {resolved.provider} ({resolved.model}). Replied: {reply[:60]}",
        "status": resolved.status(),
    }


def _cli() -> int:
    result = check_connection()
    print(result["message"])
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(_cli())
