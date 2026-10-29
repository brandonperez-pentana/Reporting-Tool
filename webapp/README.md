# Reporter in a browser

The same three-stage workflow as the desktop app, in a browser tab. The desktop app
(`reporter_ui.py`) is untouched and still the supported way to run a month; this is a
second front end over the same engine, for anyone who would rather not use Tk.

## Running it

```powershell
cd <the folder holding reporter_web.py>
.\.venv\Scripts\Activate.ps1
python reporter_web.py
```

It opens http://127.0.0.1:8765 in your default browser. `--port 9000` moves it,
`--no-browser` stops it opening a window.

Stop it with Ctrl+C in that window. Closing the terminal stops the server; the browser
tab will then say it cannot reach it.

## What it needs

Nothing new. It uses `http.server` from the standard library, so there is no extra
`pip install` and no new wheel for endpoint security to block. It does need the same
Python environment the desktop app uses, because it imports the same modules.

Run it **on the machine that has the PDFs and Outlook**. It binds to `127.0.0.1`, so
only that machine can reach it — the same reach the desktop app has.

## How it maps to the desktop app

| Desktop | Browser |
| --- | --- |
| Stage 1 Reporting | **1 · Reporting** — month, Run Reporting, dealership list, insight/summary/SMS editing, Approve & next, source PDF |
| Stage 2 Email Queue | **2 · Email Queue** — ready-to-send ticks, To/CC/Subject, HTML preview, Save as drafts, Send test, Send ticked |
| Stage 3 SMS | **3 · SMS** — the consolidated document, Copy all, Save as .txt, Export CSV breakdown |
| ⚙ Settings | **Settings** — the same five paths plus the test address |
| Contacts manager | **Contacts** |
| Report reference window | **Report reference**, on stage 2 |

Two differences worth knowing:

* **No file-picker dialogs.** A browser cannot open one for a server. Paths are typed,
  and each has a **Check** button that confirms the folder or file exists.
* **The PDF renders in the browser's own viewer** rather than a Tk canvas, so paging and
  zoom are the browser's. **Open in default app** still hands the file to Windows.

## Architecture

```
reporter_web.py          launcher
webapp/server.py         stdlib HTTP server, routing, static files, security guards
webapp/api.py            JSON wrapper over report_engine — no reporting logic of its own
webapp/static/           one page: index.html, app.css, app.js (no framework, no CDN)
```

`webapp/api.py` calls `report_engine` exactly as `reporter_ui.py` does. Nothing in
`webapp/` parses a PDF, builds an email or writes an insight. Keep it that way: both
front ends stay in step only for as long as the engine is the single source of truth.

Two things this layer adds that a desktop UI gets for free:

* **Serialisation.** `report_engine` holds one module-level session. A browser can fire
  overlapping requests in a way a single-threaded Tk loop never does, so every engine
  call goes through one lock.
* **Jobs.** Reporting runs and send batches take minutes. They run on a worker thread and
  the page polls `/api/jobs/<id>` — the web equivalent of the Tk worker queue.

## Security

Localhost only, no login — the same posture as the desktop app. Two checks keep that
true even though a browser is involved: the `Host` header must be a loopback name (which
defeats DNS rebinding), and anything that changes state or sends mail must be
same-origin. Neither replaces authentication, so **do not bind this to `0.0.0.0`**.
The rendered email preview is sandboxed with a Content-Security-Policy that blocks
script and network access.
