"""A dependency-free local web server for Reporter.

Runs on the machine that holds the PDFs and Outlook, binds to the loopback address
only, and serves the same three-stage workflow the desktop app does. Nothing here is
installed: it is http.server from the standard library plus a small router, chosen so
the app adds no wheels to a managed Windows build.

Security posture matches the desktop app's: only this machine can reach it. Two checks
keep that true even though a browser is involved --

* the Host header must be a loopback name, which stops a hostile page from using DNS
  rebinding to point a real domain at 127.0.0.1 and drive this API;
* any request that changes something or sends mail must be same-origin.

Neither is a substitute for a login. Do not bind this to 0.0.0.0.

Server mode (REPORTER_MODE=server) keeps the loopback binding and puts IIS in front.
IIS signs the user in with Windows Authentication and HttpPlatformHandler forwards
their Windows token (X-IIS-WindowsAuthToken), which this process resolves to
DOMAIN\\user. See deploy/iis/web.config.
"""

from __future__ import annotations

import argparse
import hmac
import json
import logging
import mimetypes
import os
import re
import subprocess
import sys
import threading
import traceback
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, quote, unquote, urlparse

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from webapp import api  # noqa: E402
import settings_manager  # noqa: E402

STATIC_DIR = Path(__file__).resolve().parent / "static"
LOGGER = logging.getLogger("reporter.web")

DEFAULT_PORT = 8765
ALLOWED_HOSTS = {"127.0.0.1", "localhost", "[::1]", "::1"}
MAX_BODY_BYTES = 8 * 1024 * 1024

IIS_TOKEN_HEADER = "X-IIS-WindowsAuthToken"
# Only for a proxy that sets it after authenticating; never trust it by default.
USER_HEADER = "X-Reporter-User"
PROXY_SECRET_HEADER = "X-Reporter-Proxy"
# DOMAIN\user or user@domain.
_USER_PATTERN = re.compile(r"^[A-Za-z0-9._ -]{1,64}([\\@][A-Za-z0-9._ -]{1,128})?$")


def _public_origins() -> set[str]:
    raw = os.environ.get("REPORTER_PUBLIC_ORIGINS", "")
    return {item.strip().lower().rstrip("/") for item in raw.split(",") if item.strip()}


def _user_from_iis_token(raw: str) -> str:
    """HttpPlatformHandler duplicates the caller's token into this process and sends the handle."""
    import win32api  # type: ignore[import-not-found]
    import win32security  # type: ignore[import-not-found]

    handle = int(raw, 16)
    try:
        sid = win32security.GetTokenInformation(handle, win32security.TokenUser)[0]
        name, domain, _ = win32security.LookupAccountSid(None, sid)
    finally:
        win32api.CloseHandle(handle)
    return f"{domain}\\{name}"

Handler = Callable[..., Any]
_ROUTES: list[tuple[str, re.Pattern[str], Handler]] = []


def route(method: str, pattern: str) -> Callable[[Handler], Handler]:
    compiled = re.compile(f"^{pattern}$")

    def register(func: Handler) -> Handler:
        _ROUTES.append((method, compiled, func))
        return func

    return register


class Request:
    """One request, with its body already read.

    The body is read up front by the handler rather than lazily here. On a keep-alive
    connection an unread body is still sitting in the socket when the next request
    arrives, so the server parses it as a request line and answers 501 — which is what
    happened to any POST whose route ignored its body, and to every POST rejected by a
    guard before the route ran.
    """

    def __init__(self, handler: "ReporterHandler", query: dict[str, list[str]], body: bytes = b"") -> None:
        self.handler = handler
        self.query = query
        self.body = body

    def param(self, name: str, default: str = "") -> str:
        values = self.query.get(name) or []
        return values[0] if values else default

    def json(self) -> dict[str, Any]:
        if not self.body:
            return {}
        try:
            payload = json.loads(self.body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise api.ApiError("The request body was not valid JSON.") from exc
        return payload if isinstance(payload, dict) else {"value": payload}


class Response:
    def __init__(self, body: bytes, content_type: str, status: int = 200, headers: dict[str, str] | None = None) -> None:
        self.body = body
        self.content_type = content_type
        self.status = status
        self.headers = headers or {}


def json_response(payload: Any, status: int = 200) -> Response:
    return Response(json.dumps(payload, default=str).encode("utf-8"), "application/json; charset=utf-8", status)


# --- page and assets ------------------------------------------------------------

@route("GET", r"/")
def index(request: Request) -> Response:
    return _static_file("index.html")


@route("GET", r"/static/(?P<name>[A-Za-z0-9_.-]+)")
def static_asset(request: Request, name: str) -> Response:
    return _static_file(name)


def _static_file(name: str) -> Response:
    target = (STATIC_DIR / name).resolve()
    if STATIC_DIR.resolve() not in target.parents or not target.is_file():
        raise api.ApiError("Not found.", 404)
    content_type, _ = mimetypes.guess_type(target.name)
    return Response(target.read_bytes(), content_type or "application/octet-stream")


# --- bootstrap and settings -----------------------------------------------------

@route("GET", r"/api/bootstrap")
def bootstrap(request: Request) -> Response:
    return json_response(
        {
            "settings": api.get_settings(),
            "months": api.month_options(),
            "session": api.session_state(),
            "sms_limit": api.reporting_core.SMS_CHAR_LIMIT,
        }
    )


@route("GET", r"/api/settings")
def get_settings(request: Request) -> Response:
    return json_response(api.get_settings())


@route("POST", r"/api/settings")
def post_settings(request: Request) -> Response:
    return json_response(api.save_settings(request.json()))


@route("GET", r"/api/settings/ai")
def ai_settings(request: Request) -> Response:
    return json_response(api.ai_status())


@route("POST", r"/api/settings/ai/test")
def ai_settings_test(request: Request) -> Response:
    return json_response(api.ai_test())


@route("POST", r"/api/settings/check-path")
def check_path(request: Request) -> Response:
    return json_response(api.path_check(request.json()))


@route("POST", r"/api/settings/validate-template")
def validate_template(request: Request) -> Response:
    return json_response(api.validate_template(request.json()))


@route("POST", r"/api/browse")
def browse(request: Request) -> Response:
    return json_response(api.browse(request.json()))


@route("GET", r"/api/months")
def months(request: Request) -> Response:
    return json_response(api.month_options())


@route("GET", r"/api/portfolio")
def portfolio(request: Request) -> Response:
    return json_response(api.portfolio_options())


# --- reporting ------------------------------------------------------------------

@route("POST", r"/api/reporting/run")
def run_reporting(request: Request) -> Response:
    return json_response(api.start_reporting(request.json()))


@route("GET", r"/api/session")
def session(request: Request) -> Response:
    return json_response(api.session_state())


@route("GET", r"/api/jobs/(?P<job_id>[a-f0-9]+)")
def job(request: Request, job_id: str) -> Response:
    return json_response(api.job_status(job_id))


@route("GET", r"/api/jobs/(?P<job_id>[a-f0-9]+)/download")
def job_download(request: Request, job_id: str) -> Response:
    data, content_type, name = api.export_download(job_id)
    ascii_name = name.encode("ascii", "replace").decode("ascii").replace('"', "'")
    return Response(
        data,
        content_type,
        headers={"Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(name)}"},
    )


# --- one dealership -------------------------------------------------------------

@route("GET", r"/api/dealerships/(?P<name>[^/]+)")
def dealership(request: Request, name: str) -> Response:
    return json_response(api.dealership_report(unquote(name)))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/text")
def dealership_text(request: Request, name: str) -> Response:
    return json_response(api.save_dealership_text(unquote(name), request.json()))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/email")
def dealership_email(request: Request, name: str) -> Response:
    return json_response(api.save_working_email(unquote(name), request.json()))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/email/build")
def dealership_email_build(request: Request, name: str) -> Response:
    return json_response(api.start_email_build(unquote(name), request.json()))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/approve")
def dealership_approve(request: Request, name: str) -> Response:
    return json_response(api.set_approved(unquote(name), request.json()))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/skip")
def dealership_skip(request: Request, name: str) -> Response:
    return json_response(api.set_skipped(unquote(name), request.json()))


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/ready")
def dealership_ready(request: Request, name: str) -> Response:
    return json_response(api.set_ready(unquote(name), request.json()))


@route("GET", r"/api/dealerships/(?P<name>[^/]+)/email.html")
def dealership_email_html(request: Request, name: str) -> Response:
    details = api.dealership_report(unquote(name))
    html = str(details.get("html_body") or "")
    if not html.strip():
        html = (
            "<!doctype html><meta charset='utf-8'>"
            "<body style='font:14px system-ui;padding:24px;color:#444'>"
            "No email has been built for this dealership yet.</body>"
        )
    # The email is trusted local output, but it is still third-party markup rendered in
    # an iframe: keep it from reaching the network or running script.
    return Response(
        html.encode("utf-8"),
        "text/html; charset=utf-8",
        headers={"Content-Security-Policy": "default-src 'none'; img-src data: cid: blob:; style-src 'unsafe-inline'"},
    )


@route("GET", r"/api/dealerships/(?P<name>[^/]+)/pdf")
def dealership_pdf(request: Request, name: str) -> Response:
    documents = api.dealership_report(unquote(name)).get("source_documents") or {}
    kind = request.param("kind", "performance")
    path = str(documents.get(f"{kind}_pdf") or "") if kind in {"performance", "sales", "service"} else ""
    if not path:
        raise api.ApiError(f"No {kind} PDF was found for this dealership.", 404)
    try:
        data = Path(path).read_bytes()
    except FileNotFoundError as exc:
        raise api.ApiError(f"The {kind} PDF is no longer at {path}.", 404) from exc
    name = Path(path).name
    # Headers are latin-1 only; names with e.g. an en dash need the RFC 5987 form.
    ascii_name = name.encode("ascii", "replace").decode("ascii").replace('"', "'")
    return Response(
        data,
        "application/pdf",
        headers={"Content-Disposition": f"inline; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(name)}"},
    )


@route("POST", r"/api/dealerships/(?P<name>[^/]+)/open-pdf")
def dealership_open_pdf(request: Request, name: str) -> Response:
    if settings_manager.server_mode():
        raise api.ApiError("Opening files in an app is not available on the server. Use New tab.", 403)
    payload = request.json()
    documents = api.dealership_report(unquote(name)).get("source_documents") or {}
    kind = str(payload.get("kind") or "sales")
    path = str(documents.get(f"{kind}_pdf") or "")
    if not path:
        raise api.ApiError(f"No {kind} PDF was found for this dealership.", 404)
    _open_with_default_app(path)
    return json_response({"opened": path})


def _open_with_default_app(path: str) -> None:
    """Hand a file to the desktop, the way the Tk app's 'Open in Default App' does."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])  # noqa: S603,S607
        else:
            subprocess.Popen(["xdg-open", path])  # noqa: S603,S607
    except Exception as exc:  # noqa: BLE001
        raise api.ApiError(f"Could not open that file: {api.friendly_error(exc)}") from exc


# --- queue, sending, SMS, contacts ----------------------------------------------

@route("GET", r"/api/queue")
def queue(request: Request) -> Response:
    return json_response({"queue": api.session_state()["queue"]})


@route("GET", r"/api/outlook")
def outlook(request: Request) -> Response:
    return json_response(api.outlook_status())


@route("POST", r"/api/drafts")
def drafts(request: Request) -> Response:
    return json_response(api.start_drafts(request.json()))


@route("POST", r"/api/pdf-copies")
def pdf_copies(request: Request) -> Response:
    return json_response(api.start_pdf_copies(request.json()))


@route("POST", r"/api/export")
def export_eml(request: Request) -> Response:
    return json_response(api.start_export(request.json()))


@route("GET", r"/api/sms")
def sms(request: Request) -> Response:
    return json_response(api.sms_document())


@route("POST", r"/api/sms/save")
def sms_save(request: Request) -> Response:
    return json_response(api.save_sms_document(request.json()))


@route("POST", r"/api/sms/export")
def sms_export(request: Request) -> Response:
    return json_response(api.export_sms())


@route("GET", r"/api/dealership-names")
def dealership_names(request: Request) -> Response:
    return json_response({"names": api.dealership_names()})


@route("GET", r"/api/contacts")
def contacts(request: Request) -> Response:
    return json_response(api.contacts(request.param("path") or None))


@route("POST", r"/api/contacts")
def contacts_save(request: Request) -> Response:
    return json_response(api.save_contacts(request.json()))


# --- server ---------------------------------------------------------------------

class ReporterHandler(BaseHTTPRequestHandler):
    server_version = "Reporter"
    sys_version = ""
    protocol_version = "HTTP/1.1"

    def do_GET(self) -> None:  # noqa: N802
        self._dispatch("GET")

    def do_POST(self) -> None:  # noqa: N802
        self._dispatch("POST")

    def _dispatch(self, method: str) -> None:
        try:
            # Drain the body before anything else can fail. See Request's docstring.
            body = self._read_body()
        except api.ApiError as exc:
            self.close_connection = True
            self._write(json_response({"error": exc.message}, exc.status))
            return
        except (BrokenPipeError, ConnectionResetError):
            return

        try:
            # First, so the token handle IIS duplicated into this process is always closed.
            user = self._signed_in_user()
            self._guard_host()
            if method == "POST":
                self._guard_origin()
            parsed = urlparse(self.path)
            path = parsed.path.rstrip("/") or "/"
            request = Request(self, parse_qs(parsed.query), body)
            with settings_manager.as_user(user):
                for route_method, pattern, func in _ROUTES:
                    if route_method != method:
                        continue
                    match = pattern.match(path)
                    if match:
                        response = func(request, **match.groupdict())
                        self._write(response)
                        return
            raise api.ApiError("Not found.", 404)
        except api.ApiError as exc:
            self._write(json_response({"error": exc.message}, exc.status))
        except (BrokenPipeError, ConnectionResetError):
            return
        except Exception as exc:  # noqa: BLE001
            LOGGER.error("Unhandled error on %s %s\n%s", method, self.path, traceback.format_exc())
            self._write(json_response({"error": api.friendly_error(exc)}, 500))

    def _read_body(self) -> bytes:
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            return b""
        if length > MAX_BODY_BYTES:
            # Too big to read and stay in sync, so the connection has to end here.
            raise api.ApiError("That request is too large.", 413)
        return self.rfile.read(length)

    def _guard_host(self) -> None:
        host = (self.headers.get("Host") or "").split(":")[0].strip().lower()
        allowed = {name.strip("[]") for name in ALLOWED_HOSTS}
        if settings_manager.server_mode():
            allowed |= {(urlparse(origin).hostname or "").lower() for origin in _public_origins()}
        if host and host not in allowed:
            raise api.ApiError("This app only answers requests addressed to localhost.", 403)

    def _guard_origin(self) -> None:
        origin = self.headers.get("Origin")
        if origin is None:
            return
        if settings_manager.server_mode() and origin.strip().lower().rstrip("/") in _public_origins():
            return
        hostname = (urlparse(origin).hostname or "").lower()
        if hostname not in {name.strip("[]") for name in ALLOWED_HOSTS}:
            raise api.ApiError("Cross-origin requests are not allowed.", 403)

    def _signed_in_user(self) -> str | None:
        """The caller's Windows login in server mode, or None on the desktop."""
        if not settings_manager.server_mode():
            return None
        token = (self.headers.get(IIS_TOKEN_HEADER) or "").strip()
        secret = os.environ.get("REPORTER_PROXY_SECRET", "")
        if secret and not hmac.compare_digest(self.headers.get(PROXY_SECRET_HEADER) or "", secret):
            if token:
                self._close_token(token)
            raise api.ApiError("Requests must come through the Reporter site.", 403)
        user = ""
        if token:
            try:
                user = _user_from_iis_token(token)
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("Could not resolve the Windows token IIS forwarded: %s", exc)
        elif os.environ.get("REPORTER_TRUST_USER_HEADER") == "1":
            user = (self.headers.get(USER_HEADER) or "").strip()
        if not user and os.environ.get("REPORTER_DEV_USER"):
            # Local testing without IIS: every request is this user. Never set it on a server.
            user = os.environ["REPORTER_DEV_USER"].strip()
        if not _USER_PATTERN.match(user):
            raise api.ApiError("Sign-in required. Open Reporter through its site address.", 401)
        return user

    @staticmethod
    def _close_token(raw: str) -> None:
        try:
            import win32api  # type: ignore[import-not-found]

            win32api.CloseHandle(int(raw, 16))
        except Exception:  # noqa: BLE001
            LOGGER.debug("Could not close a forwarded token handle", exc_info=True)

    def _write(self, response: Response) -> None:
        body = response.body
        self.send_response(response.status)
        self.send_header("Content-Type", response.content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cache-Control", "no-store")
        for key, value in response.headers.items():
            self.send_header(key, value)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def log_message(self, fmt: str, *args: Any) -> None:
        LOGGER.debug("%s - %s", self.address_string(), fmt % args)


def serve(port: int = DEFAULT_PORT, open_browser: bool = True) -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    server = settings_manager.server_mode()
    if server:
        if settings_manager.data_root() is None:
            raise SystemExit("REPORTER_MODE=server needs REPORTER_DATA_ROOT set to a folder for per-user settings.")
        seeded = settings_manager.seed_shared_settings()
        if seeded:
            print(f"Shared settings created from this install's settings: {seeded}")
        open_browser = False
    address = ("127.0.0.1", port)
    httpd = ThreadingHTTPServer(address, ReporterHandler)
    url = f"http://127.0.0.1:{port}/"
    print(f"Reporter (web) is running at {url}")
    print("Server mode: reachable only through the IIS site." if server else "This machine only. Press Ctrl+C to stop.")
    if server and os.environ.get("REPORTER_DEV_USER"):
        print(f"TESTING ONLY: every request without a sign-in is treated as {os.environ['REPORTER_DEV_USER']}.")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping.")
    finally:
        httpd.server_close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Reporter in a browser, on this machine only.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser window on start.")
    args = parser.parse_args(argv)
    serve(port=args.port, open_browser=not args.no_browser)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
