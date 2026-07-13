"""Vercel Python entrypoint. Vercel routes every request to this module and
expects an ASGI-callable named `app` — so this just puts the real app/
directory on sys.path (mirroring `uvicorn --app-dir app main:app` locally)
and re-exports the FastAPI instance from app/main.py.
"""

import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from main import app  # noqa: E402
