"""Start Reporter in a browser.

The desktop app (reporter_ui.py) is untouched and still the supported way to run a
month. This serves the same workflow over http://127.0.0.1:8765 for anyone who would
rather work in a browser tab. Both drive the same report_engine session, so whichever
you start, the behaviour is the same.

    python reporter_web.py
    python reporter_web.py --port 9000 --no-browser
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from webapp.server import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
