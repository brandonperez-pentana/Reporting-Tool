"""The AI pass: the funnel engine's copy, rewritten, and never trusted.

This layer never decides *what* is true. The month model and
:mod:`reporting.funnel_insights` have already chosen the four findings, their
direction, their size and their order, from the figures. All this does is ask a
model to say those same findings better -- tighter phrasing, a more natural
sentence, less repetition across four bullets that were each written alone.

Everything the model returns is then checked back against what went in, and
rejected whole if it fails:

  * Every number in the output must have been in the input. A model that invents
    a figure for a dealership is the one failure mode that matters here, and it
    is cheap to catch: the deterministic copy is the only source of facts.
  * The bullet count must match, so a finding cannot be dropped or added.
  * The SMS must still fit its limit.
  * The operator's standing rules still apply -- no counterfactuals, no
    instruction, never the word money.

Rejection is silent and total: the caller keeps the funnel copy it already had.
That is the whole safety story. A wrong key, a timeout, a throttled account, a
model having an off day, a locked-down proxy -- every one of them lands in the
same place, which is the report that ships today.
"""

from __future__ import annotations

import json
import logging
import re
from typing import Any

LOGGER = logging.getLogger("reporter.ai.insights")

SMS_LIMIT = 160

# Phrasing the operator has ruled out. Counterfactuals are advice wearing a number as
# a disguise; instruction is redundant because naming the driver *is* the instruction.
BANNED_PATTERNS = (
    r"\bmoney\b",
    r"\bshould\b",
    r"\brecommend",
    r"\bconsider\b",
    r"\bsuggest",
    r"\bwould have\b",
    r"\bcould have\b",
    r"\bif you\b",
    r"\btry to\b",
    r"\bneed to\b",
    r"\bmake sure\b",
)

SYSTEM_PROMPT = """You rewrite monthly performance copy for car dealerships.

You are given findings that have already been calculated. Your only job is to say \
them more clearly. You are not an analyst here and you have no access to the \
underlying data.

Absolute rules:
1. Never introduce a number, percentage, dollar amount, month, campaign or metric \
name that is not already in the text you were given. Not one.
2. Never change what a finding says: direction, size and subject stay exactly as \
supplied. If something went down, it still went down.
3. Return the same number of insights you were given, in the same order.
4. No counterfactuals, no hypotheticals, no advice, no instructions, no questions.
5. Never use the word "money". Never address the reader as "you".
6. Movement first, then the driver, then the driver's share.
7. Plain declarative sentences. No adjectives that grade performance \
("strong", "disappointing", "impressive"). The reader converts figures for a living.

Return only a JSON object, no markdown fence and no preamble:
{"insights": ["...", "..."], "summary": "...", "sms": "..."}"""


def _numbers(text: str) -> set[str]:
    """Every numeric token, normalised so 1,234 and 1234 compare equal."""
    found = set()
    for raw in re.findall(r"\d[\d,]*(?:\.\d+)?", text or ""):
        cleaned = raw.replace(",", "").rstrip(".")
        if cleaned:
            found.add(cleaned.rstrip("0").rstrip(".") if "." in cleaned else cleaned)
    return found


def _banned(text: str) -> str:
    lowered = (text or "").lower()
    for pattern in BANNED_PATTERNS:
        if re.search(pattern, lowered):
            return pattern.strip("\\b")
    return ""


def _build_prompt(bullets: list[str], summary: str, sms: str, month_name: str) -> str:
    numbered = "\n".join(f"{index}. {bullet}" for index, bullet in enumerate(bullets, 1))
    return (
        f"Month: {month_name}\n\n"
        f"INSIGHTS ({len(bullets)} of them, keep the count and the order):\n{numbered}\n\n"
        f"SUMMARY:\n{summary}\n\n"
        f"SMS (must stay under {SMS_LIMIT} characters):\n{sms}\n\n"
        "Rewrite each of these. Same facts, same order, better sentences."
    )


def _parse(reply: str) -> dict[str, Any] | None:
    text = (reply or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        # Some models still wrap the object in a sentence. Take the outermost braces.
        start, end = text.find("{"), text.rfind("}")
        if start == -1 or end <= start:
            return None
        try:
            data = json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            return None
    return data if isinstance(data, dict) else None


def validate(
    data: dict[str, Any],
    bullets: list[str],
    summary: str,
    sms: str,
) -> tuple[list[str], str, str] | None:
    """Accept the rewrite only if it says the same things. Otherwise reject it all."""
    new_bullets = data.get("insights")
    new_summary = str(data.get("summary") or "").strip()
    new_sms = str(data.get("sms") or "").strip()

    if not isinstance(new_bullets, list) or len(new_bullets) != len(bullets):
        LOGGER.info("AI rewrite rejected: insight count changed")
        return None

    new_bullets = [str(item or "").strip() for item in new_bullets]
    if not all(new_bullets) or not new_summary or not new_sms:
        LOGGER.info("AI rewrite rejected: an empty field came back")
        return None

    if len(new_sms) > SMS_LIMIT:
        LOGGER.info("AI rewrite rejected: SMS was %d characters", len(new_sms))
        return None

    # The check that matters. Anything numeric in the rewrite must have been supplied.
    allowed = _numbers(" ".join(bullets) + " " + summary + " " + sms)
    for field_name, text in (
        ("insights", " ".join(new_bullets)),
        ("summary", new_summary),
        ("sms", new_sms),
    ):
        invented = _numbers(text) - allowed
        if invented:
            LOGGER.warning(
                "AI rewrite rejected: %s contained figures that were not supplied (%s)",
                field_name,
                ", ".join(sorted(invented)[:5]),
            )
            return None

    for text in [*new_bullets, new_summary, new_sms]:
        hit = _banned(text)
        if hit:
            LOGGER.info("AI rewrite rejected: banned phrasing (%s)", hit)
            return None

    # A rewrite that runs away in length is not a rewrite.
    for original, rewritten in zip(bullets, new_bullets):
        if len(rewritten) > max(240, int(len(original) * 1.8)):
            LOGGER.info("AI rewrite rejected: an insight nearly doubled in length")
            return None

    return new_bullets, new_summary, new_sms


def refine_client_copy(
    bullets: list[str],
    summary: str,
    sms: str,
    month_name: str,
) -> tuple[list[str], str, str] | None:
    """Ask the model to improve the copy. Returns None whenever anything is off."""
    if not bullets or not summary:
        return None

    try:
        import ai_client

        reply = ai_client.generate(
            _build_prompt(bullets, summary, sms, month_name),
            system=SYSTEM_PROMPT,
        )
    except Exception as exc:  # noqa: BLE001 - never let this path break a month
        LOGGER.info("AI rewrite unavailable, keeping the funnel copy: %s", exc)
        return None

    data = _parse(reply)
    if data is None:
        LOGGER.info("AI rewrite rejected: the reply was not the JSON object asked for")
        return None

    return validate(data, bullets, summary, sms)
