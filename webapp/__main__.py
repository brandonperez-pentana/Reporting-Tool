"""Entry point so `python -m webapp` starts the browser front end."""

from webapp.server import main

if __name__ == "__main__":
    raise SystemExit(main())
