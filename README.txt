Reporter-Web — insights update
==============================

WHAT THIS IS
    The insight engine rebuilt as four layers instead of one. It reads both reports as a
    single month, works out what happened from the figures, and only then writes
    anything. The engine in use today is untouched and still the default: nothing
    changes for anyone who does not switch it on.

HOW TO INSTALL
    1. Close Reporter if it is running.
    2. Extract this zip straight over C:\PentDev\Reporter-Web, letting it overwrite.
       It contains only the files listed at the bottom. Your contacts.csv, your settings
       and your accounts folder are not in it and are not touched.
    3. Start Reporter again.

HOW TO TURN IT ON
    Settings -> Insight engine -> "funnel". Switch it back to "rules" at any time; the
    setting is read fresh on every run, so nothing needs reinstalling either way.

BEFORE YOU SWITCH, READ BOTH
    From C:\PentDev\Reporter-Web, after running a month:

        .venv\Scripts\python tools\dev\engine_compare.py
        .venv\Scripts\python tools\dev\engine_compare.py "Demo Site 1"
        .venv\Scripts\python tools\dev\engine_compare.py --month "Jun 2026"

    It prints the four insights, the summary and the SMS from both engines side by side
    for the same reports, and writes nothing. That is the comparison worth making on
    your real accounts before this reaches a dealership.

WHAT CHANGED IN THE COPY
    - Order follows the money, not the biggest percentage: units, then what the
      campaigns returned, then where the sales funnel leaked, then service, then a
      campaign that lost money, then follow-up workload. A service metric can no longer
      outrank a collapse in leads because its percentage happened to be larger.
    - Every movement names what drove it where the report says so — which lead source,
      which campaign, and how much of the total that campaign accounts for.
    - Rates read as rates: 4% to 9% is "5 pts", not "125%".
    - The summary frames the month and then states the same movements in short form,
      instead of repeating the four insights word for word above them.
    - Direction words describe the movement, never whether it was welcome. Overdue tasks
      rising is a rise, and the SMS says so.

WHAT IT DOES NOT CHANGE
    Only the four insights, the summary and the SMS. Sentiment, review flags, scoring,
    the DOCX and TXT exports, the emails and the SMS export all still come from the
    existing path. If the funnel engine cannot build a month for a dealership it says
    nothing and that dealership falls back to the rule engine for that run alone.

VERIFY THE SWAP
        .venv\Scripts\python -m pytest -q tests

FILES IN THIS ZIP
    dealership_insights_local.py  (608,050 bytes)
    settings_manager.py  (4,067 bytes)
    reporting/__init__.py  (48 bytes)
    reporting/report_breakdowns.py  (10,155 bytes)
    reporting/month_model.py  (14,738 bytes)
    reporting/funnel_insights.py  (13,912 bytes)
    webapp/api.py  (30,417 bytes)
    webapp/static/index.html  (12,515 bytes)
    webapp/static/app.js  (44,784 bytes)
    tools/dev/engine_compare.py  (4,642 bytes)
    tools/dev/month_model_view.py  (4,590 bytes)
    tests/test_report_breakdowns.py  (6,452 bytes)
    tests/test_month_model.py  (11,087 bytes)
    tests/test_funnel_engine.py  (12,838 bytes)
    tests/test_insight_copy.py  (42,001 bytes)
