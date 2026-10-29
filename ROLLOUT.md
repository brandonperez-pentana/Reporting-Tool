# Rolling Reporter out to other people

This package contains no settings, no API key, no contacts and no client data. Every
person who installs it configures their own on first run, which is deliberate: a
package that arrives pre-filled with someone else's paths is a package that quietly
emails the wrong dealership.

## What each person does

1. Copy the folder somewhere local. `C:\Reporter` or `C:\PentDev\Reporter-Web` both
   work; avoid OneDrive-synced folders, which lock files mid-run.
2. Right-click `setup.ps1` and run with PowerShell, or from a terminal in that folder:

       powershell -ExecutionPolicy Bypass -File .\setup.ps1

   It finds a Python, builds `.venv`, installs the dependencies and writes
   `python-path.txt` for this machine.
3. Double-click `Start Reporter.bat`.
4. The first screen asks for an Accounts folder, an output folder and a contacts file.

If setup cannot find a Python, it says so and stops. Python 3.11 or 3.12 from
python.org, with "Add python.exe to PATH" ticked, is what the app is built against.

## What is per-person and never shipped

| Item                       | Where it lives           | Why it is not in the package     |
| -------------------------- | ------------------------ | -------------------------------- |
| Paths, sender name, engine | `user_settings.json`     | Different per person and machine |
| AI key                     | `.env`                   | A secret; never travels          |
| Interpreter path           | `python-path.txt`        | Different per machine            |
| Contacts                   | `data\contacts.csv`      | Real recipient addresses         |
| Metric history             | `data\metric_history.json` | Real client figures            |
| Past output and logs       | `Insights Output\`, `logs\` | Real client figures           |

All of these are kept out of the repo. `data\contacts.csv` is the one exception: it is
tracked as a header-only starter file, and each clone keeps its real contacts out with
`git update-index --skip-worktree data/contacts.csv`.

## Sending is per-person too

Reporter drives the classic Outlook profile of whoever is running it. Nothing sends
without that person clicking approve, and the address it sends from is theirs. There is
no shared service account and no server-side send.

## Updates

Put the new package on a share and have people run:

    powershell -ExecutionPolicy Bypass -File .\update.ps1 -Source \\server\share\Reporter-Web

or against a zip:

    powershell -ExecutionPolicy Bypass -File .\update.ps1 -Source C:\Downloads\Reporter-Web.zip

It copies the code, protects everything in the table above, clears stale bytecode,
installs any new dependencies, backs up settings first, and verifies the app still
imports before declaring success. Add `-WhatIf` to see what it would touch.

`VERSION` is what the script reports moving between. Bump it whenever you cut a package.

### If you would rather use Git

The repository is already set up for it: `.gitignore` covers every per-person file, so
`git pull` is a safe update on a machine that has the repo cloned. The scripts exist for
machines without Git, which on a managed build is most of them.

## Ingestion (2026.09.2)

`excel_reader.py` reads a month straight out of a DealerSocket master workbook,
replacing the PDF-to-text-to-regex path. It reads tabs 1 and 2 by position, never by
sheet name, and never writes to the file. `tests/test_excel_reader.py` asserts parity:
every metric matches what the PDF path extracted, and the engine produces identical
insight copy from either source.


## Reading from the workbook library (2026.09.5)

The library keeps one workbook per dealership at the top level, with a folder of
exported PDFs beside each. Reporter now reads it that way:

* The workbooks decide which dealerships exist and which month each one holds.
* Every figure comes from the sheet. The PDFs are the attachment and the preview.
* A dealership whose PDFs were never exported still gets a full month of insights.
* A workbook holding a different month says so by name, rather than reporting a
  missing file that is no longer the source of anything.

A library with no workbooks still runs entirely on PDFs, so dealerships can be moved
across one at a time.


## Reporting any month in the workbook (2026.09.7)

A master holds years of history in its third tab, and Reporter now offers every month
it has figures for, not only the latest.

* The latest month is read from the rendered report on tabs 1 and 2, so its figures
  are the ones printed on the PDF attached to the email.
* Any earlier month is read from the history tab, because no PDF of it exists.
* The two agree: a test asserts that reading the same month from either tab produces
  identical figures, dealership by dealership and metric by metric.
* The earliest month in a workbook is not offered -- it has nothing before it to
  compare against, so there is no movement to report.
* The month picker opens on the newest month that has data, instead of last calendar
  month, which is how a fresh run used to land on a month with nothing in it.



## The Performance Intelligence report (2026.09.10)

Reporter now builds the four-page Performance Intelligence report for each dealership
and shows it in the report pane on step 1, instead of relying on PDFs exported from the
spreadsheet by hand.

* The design lives in `templates/performance_intelligence.html` -- the dashboard as it
  was authored, with the file picker and the CDN script removed. Reporter writes the
  figures into it and prints the page, so editing that file changes the report and no
  Python has to know what it looks like.
* The figures come from the workbook's **history tab**, so unlike the spreadsheet's own
  exported report this can be produced for **any month the workbook covers**, with the
  year-on-year column and the thirteen-month charts.
* Rendering uses a browser through Playwright, one instance reused for a whole run --
  about a second per dealership after the first. `setup.ps1` downloads Chromium, and
  where a managed network blocks that download Reporter falls back to Microsoft Edge or
  Google Chrome if either is installed, which needs no download at all. Settings shows
  which one is in use. With none of them the month still runs and the pane says so.
* Reports are written to `<output folder>\Performance Intelligence` and reused until the
  workbook or the template changes. The workbook is never modified.
* The pane keeps the spreadsheet's own Sales and Service exports available beside the
  built report for anyone checking a figure against its source.

Settings > Performance Intelligence report switches it off and shows whether this
machine can render it.
