"""The AI engine: it must never change a figure, and never take a month down.

Two properties are worth more than everything else in here:

  * With no key configured -- which is every fresh install -- choosing "ai" produces
    byte-identical copy to choosing "funnel".
  * A model that invents a number is rejected whole, and the funnel copy survives.

Everything else is a variation on those.
"""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import ai_config  # noqa: E402
import dealership_insights_local as core  # noqa: E402
import settings_manager  # noqa: E402
from reporting import ai_insights  # noqa: E402
from tests.test_funnel_engine import EngineSwitch, run  # noqa: E402

BULLETS = [
    "Campaign return reached $44,183, up from $10,059, with Closed RO returning $23,114 of it.",
    "Units sold fell to 38, down from 51.",
    "Sales lead volume fell to 214, down from 402.",
]
SUMMARY = "Campaign return reached $44,183 and units sold fell to 38."
SMS = "July: campaign return $44,183, units sold 38, leads 214."


def reply(**overrides) -> str:
    payload = {"insights": list(BULLETS), "summary": SUMMARY, "sms": SMS}
    payload.update(overrides)
    return json.dumps(payload)


def refine(raw: str):
    with mock.patch("ai_client.generate", return_value=raw):
        return ai_insights.refine_client_copy(list(BULLETS), SUMMARY, SMS, "July")


class EngineSelectionTests(unittest.TestCase):
    def test_ai_is_not_the_shipped_default(self) -> None:
        self.assertEqual(settings_manager.DEFAULT_SETTINGS.get("insight_engine"), "rules")
        self.assertEqual(settings_manager.DEFAULT_SETTINGS.get("ai_insights_enabled"), "off")

    def test_ai_selects_the_funnel_path_too(self) -> None:
        """The AI engine is the funnel engine plus a rewrite, never a separate source."""
        with EngineSwitch("ai"):
            self.assertEqual(core.insight_engine_selected(), "ai")
            self.assertTrue(core.funnel_engine_selected())
            self.assertTrue(core.ai_engine_selected())

    def test_funnel_does_not_switch_on_the_rewrite(self) -> None:
        with EngineSwitch("funnel"):
            self.assertTrue(core.funnel_engine_selected())
            self.assertFalse(core.ai_engine_selected())

    def test_an_unknown_engine_falls_back_to_rules(self) -> None:
        with EngineSwitch("gpt5-turbo-ultra"):
            self.assertEqual(core.insight_engine_selected(), "rules")
            self.assertFalse(core.funnel_engine_selected())

    def test_an_unreadable_settings_file_does_not_switch_engines(self) -> None:
        original = settings_manager.load_settings
        settings_manager.load_settings = lambda *a, **k: (_ for _ in ()).throw(OSError("gone"))
        try:
            self.assertEqual(core.insight_engine_selected(), "rules")
            self.assertFalse(core.ai_engine_selected())
        finally:
            settings_manager.load_settings = original


class UnconfiguredIsIndistinguishableFromFunnelTests(unittest.TestCase):
    def test_no_key_means_the_funnel_copy_ships_unchanged(self) -> None:
        """The property that lets an operator pick "ai" before the key arrives."""
        with mock.patch.object(ai_config, "load", return_value=ai_config.AIConfig()):
            funnel = run("funnel")
            ai = run("ai")
        self.assertEqual(funnel[0], ai[0])   # insights
        self.assertEqual(funnel[1], ai[1])   # summary
        self.assertEqual(funnel[2], ai[2])   # sms

    def test_a_provider_outage_keeps_the_funnel_copy(self) -> None:
        with mock.patch("ai_client.generate", side_effect=OSError("connection reset")):
            funnel = run("funnel")
            ai = run("ai")
        self.assertEqual(funnel[:3], ai[:3])


class RewriteIsCheckedAgainstTheFiguresTests(unittest.TestCase):
    def test_a_faithful_rewrite_is_accepted(self) -> None:
        tightened = [
            "Campaign return reached $44,183, up from $10,059; Closed RO returned $23,114 of that.",
            "Units sold fell to 38 from 51.",
            "Sales lead volume fell to 214 from 402.",
        ]
        result = refine(reply(insights=tightened))
        self.assertIsNotNone(result)
        self.assertEqual(result[0], tightened)

    def test_an_invented_figure_is_rejected(self) -> None:
        """The failure that matters: a number nobody calculated reaching a dealership."""
        invented = list(BULLETS)
        invented[1] = "Units sold fell to 38, down from 51, a 25.5% drop."
        self.assertIsNone(refine(reply(insights=invented)))

    def test_a_dropped_insight_is_rejected(self) -> None:
        self.assertIsNone(refine(reply(insights=BULLETS[:2])))

    def test_an_added_insight_is_rejected(self) -> None:
        self.assertIsNone(refine(reply(insights=[*BULLETS, "Service retention held steady."])))

    def test_advice_is_rejected(self) -> None:
        preachy = list(BULLETS)
        preachy[0] = "Campaign return reached $44,183; you should keep Closed RO running."
        self.assertIsNone(refine(reply(insights=preachy)))

    def test_the_word_money_is_rejected(self) -> None:
        self.assertIsNone(refine(reply(summary="The campaigns brought in more money this month.")))

    def test_an_oversized_sms_is_rejected(self) -> None:
        self.assertIsNone(refine(reply(sms="July: " + "detail, " * 40)))

    def test_an_empty_field_is_rejected(self) -> None:
        self.assertIsNone(refine(reply(summary="")))

    def test_a_runaway_rewrite_is_rejected(self) -> None:
        padded = list(BULLETS)
        padded[1] = "Units sold fell to 38, down from 51. " + "Context follows. " * 20
        self.assertIsNone(refine(reply(insights=padded)))

    def test_prose_instead_of_json_is_rejected(self) -> None:
        self.assertIsNone(refine("Certainly! Here are the improved insights."))

    def test_a_fenced_json_object_is_still_read(self) -> None:
        self.assertIsNotNone(refine("```json\n" + reply() + "\n```"))

    def test_a_json_object_inside_chatter_is_still_read(self) -> None:
        self.assertIsNotNone(refine("Here you go:\n" + reply() + "\nHope that helps."))

    def test_thousands_separators_do_not_count_as_new_figures(self) -> None:
        """44,183 and 44183 are the same number; a rewrite may reformat freely."""
        reformatted = list(BULLETS)
        reformatted[0] = "Campaign return reached $44183, up from $10059, with Closed RO at $23114."
        self.assertIsNotNone(refine(reply(insights=reformatted)))


class ConfigurationTests(unittest.TestCase):
    def test_a_key_is_never_exposed_by_status(self) -> None:
        config = ai_config.AIConfig(api_key="sk-ant-secret-value-1234", model="m", base_url="https://x")
        status = config.status()
        self.assertNotIn("secret-value", json.dumps(status))
        self.assertTrue(status["api_key_masked"].endswith("1234"))
        self.assertTrue(status["api_key_set"])

    def test_redact_strips_the_key_from_text(self) -> None:
        config = ai_config.AIConfig(api_key="sk-ant-secret-value-1234")
        self.assertNotIn("secret-value", ai_config.redact("failed with sk-ant-secret-value-1234", config))

    def test_env_parsing_handles_the_shapes_people_actually_write(self) -> None:
        parsed = ai_config.parse_env_text(
            "\n".join([
                "# a comment",
                "",
                "REPORTER_AI_API_KEY=plain-value",
                'REPORTER_AI_MODEL="quoted-value"',
                "export REPORTER_AI_PROVIDER=openai",
                "not-a-setting-line",
            ])
        )
        self.assertEqual(parsed["REPORTER_AI_API_KEY"], "plain-value")
        self.assertEqual(parsed["REPORTER_AI_MODEL"], "quoted-value")
        self.assertEqual(parsed["REPORTER_AI_PROVIDER"], "openai")

    def test_a_config_without_a_key_is_never_active(self) -> None:
        config = ai_config.AIConfig(model="m", base_url="https://x", enabled=True)
        self.assertFalse(config.active)
        self.assertIn("API key", config.why_inactive())

    def test_the_example_env_ships_without_a_key_in_it(self) -> None:
        example = ROOT / ".env.example"
        self.assertTrue(example.is_file(), ".env.example must ship with the app")
        self.assertIn("REPORTER_AI_API_KEY=\n", example.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
