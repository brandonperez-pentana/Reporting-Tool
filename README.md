# Reporter (browser edition)

Turns each dealership's monthly DealerSocket Sales and Service PDFs into priority
insights, an SMS message, and a branded HTML email — reviewed and sent by an operator,
never automatically. This package runs that workflow in a browser tab.

The workflow is three stages, left to right:

1. **Reporting** — run the month, then read each dealership's insights and SMS against
   its source PDF side by side. Approving a dealership builds its email in the
   background.
2. **Email Queue** — the approved emails. Fix contact details, tick the ones that are
   ready, then save them as Outlook drafts, as .eml files, or as PDF copies.
   **Reporter does not send email.** Sending is done by a person, in Outlook.
3. **SMS** — one document containing every approved dealership's SMS, ready to copy out.

## Setting up on a new machine

Requires **Windows** with **Python 3.11 or 3.12**, and — for drafts and PDF copies —
**classic Outlook** and **Word**. Stages 1 and 3 need neither.

**Double-click `Start Reporter.bat`.** That is the whole setup. It finds a Python, and
the first time it runs it creates the virtual environment and installs the dependencies —
a couple of minutes, once, and it needs internet access for that first run only. After
that it just starts the app.

If it cannot find a Python it says so and stops. Install Python 3.12 from
[python.org](https://www.python.org/downloads/), tick **"Add python.exe to PATH"** on the
first screen of the installer, and double-click the file again.

<details><summary>Doing it by hand instead</summary>

```powershell
Set-Location "C:\Reporter"  # Replace with the downloaded repository folder.
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-web.txt
.\.venv\Scripts\python.exe -m playwright install chromium
.\.venv\Scripts\python.exe .\reporter_web.py
```

If `py` is not recognised, install Python 3.12 using your company's approved process,
or call the installed interpreter by its full path to create the environment:

```powershell
& "C:\Program Files\Python312\python.exe" -m venv .venv
```

That path is an example; use the actual Python installation path on the new machine.
Then run the dependency installation and startup commands above.
Calling `.venv\Scripts\python.exe` directly also sidesteps `Activate.ps1`, which a
managed machine's execution policy often blocks.

Chromium is used for Performance Intelligence PDF reports. If its download is blocked,
Reporter can fall back to installed Microsoft Edge or Google Chrome; Settings shows
whether report rendering is available.

</details>

It prints the URL and opens `http://127.0.0.1:8765` in your default browser. Leave that
PowerShell window open — closing it stops the server. **Ctrl+C** to stop.

```powershell
.\.venv\Scripts\python.exe .\reporter_web.py --port 9000  # if 8765 is taken
.\.venv\Scripts\python.exe .\reporter_web.py --no-browser
```

Nothing here needs `libtkhtml3.0.dll` or any other native UI library: the server is
Python's own `http.server` and the page is plain HTML, CSS and JavaScript with no
framework and no CDN. It works on a machine with no outbound network access.

## Starting on the company server

The Python path must belong to the **server**, not another user's PC. Do not copy a
virtual environment from your workstation: download or clone the repository, install
Python 3.12 on the server, and create its own `.venv` using the setup steps above.
The installation account needs write access to the repository folder and access to
the approved Python package source (and a browser download if using Chromium).

### Local use on the server (for example, through Remote Desktop)

After the one-time setup, run these commands in PowerShell each time:

```powershell
Set-Location "C:\Reporter"  # Replace with the downloaded repository folder.
& ".\.venv\Scripts\python.exe" ".\reporter_web.py" --no-browser
```

Open `http://127.0.0.1:8765` **in a browser on the server**. Alternatively,
double-click `Start Reporter.bat` in that folder to start and open the browser.
Keep the terminal open; Ctrl+C stops Reporter. This is an interactive launch, not
an automatically starting Windows service.

On first run, configure the accounts/library, output and contacts paths for the server.
The Windows account running Reporter needs read access to the source files and write
access to the selected output and contacts locations. Your workstation's paths and
Outlook profile do not automatically exist on the server.

### Shared website accessed from staff PCs

The command above listens only on `127.0.0.1`. Staff cannot reach it by substituting
the server name in the URL, and opening firewall port 8765 does not change that.
Do not change the Python server to listen on `0.0.0.0`.

For shared access, IT should use the existing
[IIS configuration template](deploy/iis/web.config). It keeps Python private and
places IIS with Windows Authentication and HTTPS in front:

1. Install IIS with Windows Authentication and Microsoft's HttpPlatformHandler v1.2.
2. Complete the Python and dependency setup on the server.
3. Use the template as the IIS site's `web.config`. Replace the Python/script paths,
   data and log paths, public HTTPS origin and administrator logins with the server's
   actual values. The template includes the required IIS authentication configuration.
4. Configure an application pool with **No Managed Code** and a service identity
   that can read the application and dealership library and write the configured
   data, logs and output locations. Create the data and log directories first.
5. Configure the HTTPS binding/certificate and Windows sign-in with IT. The template
   sets `REPORTER_MODE=server`, `REPORTER_DATA_ROOT`, `REPORTER_PUBLIC_ORIGINS` and
   `REPORTER_ADMINS`; retain these with your deployment's values.

IIS starts and manages Python through HttpPlatformHandler; do not also run the manual
startup command for that site. Users open the configured HTTPS address rather than
`127.0.0.1:8765`. Shared settings live under `REPORTER_DATA_ROOT`, with personal
settings under its `users` folder.

**Deployment checks:** open the HTTPS site from a staff PC, confirm Windows sign-in
and separation of users' settings, then run a sample report and export an `.eml`.
Local startup alone does not verify the IIS deployment. Outlook drafts and Word-based
PDF copies depend on Office automation under the process account; do not assume an
IIS process can use each visiting user's desktop Outlook profile. Use `.eml` export
for users to open in their own Outlook unless IT has validated the required workflow.

## First run

This copy has never been configured, and it knows it: the first time you open it you get
a **Set up Reporter** screen asking for three paths, each with a **Browse…** button so
you can find them rather than type them. Fill those in and you are running. Everything
else already has a working default.

| Asked for at setup | What it is |
| --- | --- |
| Accounts folder | The parent folder holding one subfolder per dealership, each containing that dealership's Sales/Service PDFs. |
| Output folder | Where generated insights and reports are written. |
| Contacts file | The `contacts.csv` mapping dealerships to contact names and addresses. The blank one that ships here is offered by default. |
| Save .eml files to *(optional)* | Where **Save as .eml** writes emails you want to import into Outlook by hand. |

The rest lives in **⚙ Settings** and can wait until you need it:

| Setting | What it is |
| --- | --- |
| PDF copies folder | Optional. Leave it empty and each PDF copy is written beside that dealership's own source reports. |
| Always BCC | Optional. Addresses BCC'd on every email, on top of anything a contact row asks for. Semicolon-separated. Empty in a fresh copy. |
| Email template assets | The branded images. Points at the bundled `email_assets/` automatically — change it only to use your own. |
| Sign-off name / title | Optional. See below. |
| Template path | Only used if you switch assembly mode to `legacy_msg`. Leave it empty. |

Settings are saved to `user_settings.json` beside the app. If the install location is
read-only (Program Files, a managed share), they move to `%LOCALAPPDATA%\Reporter\`
automatically. Nothing from the machine this package was built on is in it — no paths,
no contacts, no history — so every copy starts as a new install.

### The branded email is already here

`email_assets/` holds the four images the email uses — header banner, Monthly Snapshot
hero, Service Now button, footer logo — and `templates/dealer_snapshot_email.html.j2`
holds the layout. Together they are the whole branded email, so it renders correctly on
a machine with no Outlook, no `.msg` template and no network access. There is nothing to
copy across and nothing to point at.

### contacts.csv

`data/contacts.csv` ships as a header and nothing else. There is no customer data in
this package, and no example row to delete.

```csv
dealership,contact_name,email,cc,bcc,active,notes,last_updated
```

Use **Contacts** in the app rather than editing the file. It groups contacts by
dealership, so several people at one site read as one block, and the dealership itself is
**picked from a list** of the accounts folder's subfolder names — a name cannot be almost
right, which is the failure that otherwise produces an email with no recipient and no
error. A row naming a dealership with no matching folder keeps its value and is flagged
at the top of the panel.

Several contacts at one dealership all receive the email: their addresses are combined
and de-duplicated, and the greeting names them ("Hi Sam and Pat,").

`active` and `notes` are still in the file but no longer shown — existing values are
honoured, so a `contacts.csv` from an older build keeps behaving the way its author
meant. `active` has to be `yes` for a dealership to be sendable, and new contacts added
in the app are active.

### Sign-off

Emails go out **unsigned** unless you say otherwise: no name, no title, no "Best
regards" — the body simply ends. That is deliberate, so a copy handed to someone else
never signs their mail with the previous owner's name.

To add one, fill in **Sign-off name** and **Sign-off title** in Settings. Fill in one and
you get one line; leave both empty and the block is skipped entirely rather than leaving
a gap under "Best regards,".

## What goes in each email

Every email carries that dealership's own two source PDFs — the Sales report and the
Service report for the month, exactly the files stage 1 parsed — as attachments, plus the
four branding images inline. The engine takes them from the dealership's own record, so
one dealership's PDFs cannot reach another's email. Stage 2 names the attached files in
the results after a send or an export, so it is visible rather than assumed.

A dealership whose PDFs were missing that month is refused outright: its report would be
placeholder text rather than real figures, and neither drafting nor exporting will
proceed for it.

## Reporter does not send email

There is no send button, and the API refuses a request to send even if one is
constructed by hand. Every action on stage 2 stops at something a person then acts on:

| Button | What it does |
| --- | --- |
| **Save as drafts** | Puts each email in Outlook's Drafts folder, addressed and attached, for you to read and send. |
| **Save as .eml** | Writes each email to disk as a file Outlook opens ready to send. Needs no Outlook to produce. |
| **Save PDF copies** | Writes a PDF of each email beside that dealership's own source reports. |

Nothing reaches a dealership because of a click in a browser tab.

## PDF copies

**Save PDF copies** renders each ticked email and writes it **into the same folder that
dealership's Sales and Service PDFs came from** — not a central archive, and not a
subfolder, so the copy sits with the paperwork it belongs to:

```
...\Demo Site 1\
   Demo Site 1 - Sales Report - Aug 26.pdf
   Demo Site 1 - Service Report - Aug 26.pdf
   Demo Site 1 - Email - Aug 2026.pdf     <- written by Save PDF copies
```

The folder is taken from the report PDF that was actually parsed, so it works whatever
the folder is called and however you arrange the accounts folder. The month is in the
filename because without a month subfolder there is nothing else stopping one month's
copy from overwriting the last.

Rendering uses **Word**, which is present wherever classic Outlook is. If Word cannot be
started the copy is saved as HTML instead and the results say so — never nothing.

If your accounts folder is read-only, or you would rather keep every copy in one place,
set **PDF copies folder** in Settings; copies then go there, in a subfolder per month.

## Sending by hand: Save as .eml

**Save as .eml** on stage 2 writes each ticked email to disk instead of handing it to
Outlook. It is the way out when Outlook automation will not run — a new machine, a
pending Office repair, an install that has not been activated yet — and it produces the
same email: same body, same branding images, same two PDFs attached, same To/CC/BCC.

Point **Save .eml files to** at a folder in Settings, tick the emails, click **Save as
.eml**. Then in Explorer, double-click one. Outlook opens it as a draft with the
recipients, subject, body and attachments already filled in, ready for you to read and
send. Drag them into a Drafts folder if you would rather deal with them there.

This path uses nothing but the Python standard library — no Outlook, no `pywin32`, no
extra package — so it works on a machine where everything else about Outlook is broken.
Files are never overwritten: exporting twice gives you `Name - Jul 2026 (2).eml`.

## Reading the month as a whole

Most months move both ways at once — leads down while campaign returns climb is
ordinary, not a contradiction. The report is built so that neither half can crowd out
the other:

- **The four insights** must cover both directions whenever the month moved materially
  both ways. If ranking by size alone fills every slot from one side, the weakest slot
  goes to the best insight from the other. The lead is never displaced: the month's
  biggest movement still opens the report, whichever way it went.
- **The summary** keeps three of the four clauses. Where the first two already agree
  with each other, the third is chosen to show the other side.
- **The SMS** takes three signals, one per family, which could all come from one
  department. The weakest now gives up its slot to the missing side, so the message
  reads "deals up 26.7%, but service ROI down 79.6% …" rather than three service
  declines in a month when sales grew.
- **The opening line** cannot contradict the figures underneath it. If it would, it
  falls back to "mixed" rather than claiming the opposite.

None of this invents balance. A month with nothing good in it still says so, and a
trivial 1% drift does not count as the other side of the story — a move has to be
material to earn a slot.

## When one side of the business reports nothing

If a month's service report comes through empty — every figure zero, or the section
missing — the engine has no service evidence to work with and every insight comes from
sales. In a soft month that used to read as an unbroken run of bad news, while the same
dealership with service figures led on a campaign ROI rise.

Two things happen now:

- **The ordering rebalances.** If the report opens on a decline and a genuine
  improvement is further down, that improvement moves to the front. Nothing is added,
  removed or reworded, and every decline still appears in full immediately below.
- **The dealership is flagged** with `service_data_missing` on stage 1 and in the
  review exports, because an empty service report usually means one was missed rather
  than a quiet month. The email itself says nothing about it — that flag is for you,
  before you send.

A month where *neither* side parsed is a different thing: that is a placeholder report,
and it is refused for drafting and export rather than rebalanced.

## Trying it before you have real data

`Accounts/` contains ten synthetic **Demo Site** dealerships with July 2026 PDFs, so you
can point the Accounts folder at it, run the month and see the whole workflow end to end
without any customer data. Everything except Outlook sending works.

## Checking the install

```powershell
python -m unittest discover -s tests -p "test_*.py"
```

347 tests. They need no Outlook, no Word and no source PDFs. Tests covering the desktop
app skip automatically, since it is not part of this package.

## Baseline comparisons

Insights compare each month against the dealership's own record — "the sharpest drop in
5 months", "the lowest in 5 months". That history lives in `data/metric_history.json`,
which every run appends to. **This package ships without one**, so comparisons start
appearing after about five months of runs on this machine.

To start sooner, copy `data/metric_history.json` across from a machine that already has
one, or — if you bring over an `Insights Output/Debug Text` folder — seed it from the
report text already extracted there:

```powershell
python tools\backfill_metric_history.py
```

## Layout

| Path | What it holds |
| --- | --- |
| `Start Reporter.bat` | Double-click launcher: finds Python, sets up on first run, starts the app. |
| `reporter_web.py` | Starts the browser UI. |
| `webapp/` | The web front end: `server.py` (stdlib HTTP server), `api.py` (a JSON wrapper over `report_engine`), `static/` (one page). See `webapp/README.md`. |
| `report_engine.py` | Session state, approval and send gating, working-email edits. |
| `dealership_insights_local.py` | PDF parsing and insight generation. |
| `email_builder.py` / `email_assembler.py` | Email context, branded HTML, inline images. |
| `email_export.py` | Writes an email to disk as `.eml`, standard library only. |
| `email_assets/` | The four branding images and the manifest naming them. |
| `reporting/` | Cross-metric relationships, language map, metric history store. |
| `reporter/` | Report context and DOCX/TXT template rendering. |
| `sent_email_archive.py` | PDF copies of sent emails. |
| `settings_manager.py` | Reads and writes `user_settings.json`. |
| `templates/` | Jinja email templates and the DOCX report template. |
| `run_month.py` | Runs a month headlessly, without the UI. |
| `data/dealership_aliases.json` | Short forms used in PDF filenames and the full dealership name each stands for. Optional; edit without touching code. |
| `data/contacts.csv` | Header-only starter contacts file; each person fills in their own. |
| `tests/` | The test suite. |
| `Accounts/` | Ten synthetic demo dealerships. |

## The AI engine

Settings > Insight engine has a third choice, **AI assisted**. It is the funnel engine
plus a rewrite: the month model still reads both reports, still decides which four
findings matter, and still calculates every figure. A model is asked only to say those
same findings better -- tighter sentences, less repetition across four bullets that were
each written alone.

It never gets to decide what is true, and it is never trusted. Everything it returns is
checked back against the copy that went in, and thrown away whole if:

* a number appears that was not in the input, which is the one failure that matters
* the insight count changed, so a finding cannot be dropped or added
* the SMS no longer fits its limit
* the standing rules are broken -- advice, counterfactuals, or the word money

A rejection is silent: the dealership gets the funnel copy it would have had anyway. So
does a missing key, a timeout, a blocked proxy, or a throttled account. Choosing AI
assisted before a key exists is safe and produces byte-identical output to funnel.

### Giving it a key

The key lives in a `.env` file beside the app, never in `user_settings.json` -- that file
gets copied between machines and pasted into tickets. `.env` is git-ignored.

```
copy .env.example .env
```

Fill in `REPORTER_AI_API_KEY`, and `REPORTER_AI_MODEL` if you do not want the default.
Anthropic, OpenAI, Azure OpenAI and any OpenAI-compatible gateway are supported; set
`REPORTER_AI_PROVIDER` accordingly. Then check what the app can see:

```
python ai_config.py     # prints the settings, key masked
python ai_client.py     # one real round trip to the provider
```

Settings shows the same status, with a Test connection button, whenever AI assisted is
the selected engine. The key is masked everywhere and is never sent to the browser.

Nothing here adds a dependency: the client is standard library only, so
`requirements-web.txt` is unchanged.

## What is not in this package

The desktop app (`reporter_ui.py`) and its Tk dependencies; any `.msg` or `.oft` Outlook
item (those are saved customer emails, and the branded layout is shipped as templates and
images instead); the real contact list; any generated reports; `user_settings.json`; and
the metric history. Each is either customer data or specific to the machine it came from.

The build refuses to produce a zip if any of that is present, so this is enforced rather
than remembered.

## If Outlook will not start

`-2147221005 Invalid class string` means Windows could not resolve the ProgID
`Outlook.Application`. Classic Outlook being installed and open is **not** enough on its
own — the ProgID has to be visible to the process that is asking.

Run the diagnostic first; it reports what is actually true on the machine rather than
guessing:

```powershell
python tools\diagnose_outlook.py
```

It reads the registry and tries the call. It sends nothing and changes nothing. If it
reports that `Dispatch` succeeded, Outlook automation works from that Python and the
problem is elsewhere — click **Re-check** on stage 2, which caches its result for the
life of the page and will otherwise keep showing the old failure.

The causes it distinguishes, in order of how often they are the answer:

1. **Outlook is not running.** Start classic Outlook and leave it open. This is the
   cheapest thing to rule out and, on the machine this was written for, it was the
   answer.
2. **Elevation.** Office can register its COM classes per user. A PowerShell started "as
   administrator" can read a different `HKEY_CURRENT_USER` and not find them. Run
   Reporter from a normal window.
3. **Bitness mismatch.** 32-bit Python cannot drive 64-bit Office, or the reverse. The
   diagnostic prints both, so compare them; matching them means installing the Python
   that matches Office and recreating the venv with it.
4. **Broken registration.** Settings → Apps → Microsoft 365 → Modify → **Quick Repair**.
   Note that `OUTLOOK.EXE /regserver` **no longer exists** in Click-to-Run Office — it
   answers "The command line argument is not valid", which is not a sign of anything
   being wrong.

If Outlook is showing an activation notice, deal with that too. An unlicensed Office can
end up partially registered, and it will stop working at the deadline regardless.

Stage 2 checks all of this when you open it and says so before you tick anything, rather
than failing every email in a batch and telling you afterwards. **Re-check** re-runs the
test once you have changed something.

Everything that does not involve Outlook is unaffected: the month still runs, insights
and reports are still written, the SMS document on stage 3 is still produced, and **Save
as .eml** will still give you every approved email as a file you can open in Outlook
yourself. **Save PDF copies** needs Word rather than Outlook, and falls back
to HTML if Word cannot start.
