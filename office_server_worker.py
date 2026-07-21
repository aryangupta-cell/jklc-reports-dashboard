#!/usr/bin/env python3
"""
Office-server worker -- runs a report's normal, unchanged `process()`
pipeline on this machine instead of on the hosted (Render) instance, for
reports whose files are too large for the free tier's RAM (currently
Report 6 / Battery Disconnection Mail Creation).

Invoked fresh over SSH, once per report generation -- NOT a standing
service. Reuses the exact same `reports.<module>.process()` function the
hosted app would otherwise call, unmodified: this file only handles moving
bytes in and out.

Protocol (paired with app/core/ssh_worker.py on the hosting side):
  stdin  -- a zip containing meta.json ({"dates": {...}, "files":
            {slot_key: filename}}) plus one file per slot_key.
  stdout -- a zip containing exactly one file: the generated report.
  stderr -- logs / error details on failure (non-zero exit code).

Deployment on the office server (same pattern as Anchal's app):
  1. Check out this repo (or copy app/ + this file) to some directory,
     e.g. ~/mtr-report-worker/.
  2. Inside that directory: python3 -m venv venv && venv/bin/pip install
     -r requirements.txt (same requirements.txt as the hosted app).
  3. Add the hosting side's SSH public key to ~/.ssh/authorized_keys here.
  4. On Render, set REMOTE_DIR to this directory's path (e.g.
     /home/<user>/mtr-report-worker) and REMOTE_PYTHON to the venv's own
     interpreter (e.g. /home/<user>/mtr-report-worker/venv/bin/python3) so
     it has the installed packages -- plus SSH_HOST / SSH_USER /
     SSH_KEY_PATH / SSH_HOST_KEY (see app/core/ssh_worker.py).

Every run auto-updates its own checkout first (`git pull --ff-only`, see
_git_pull below) -- unlike the hosted app on Render, which redeploys on
every push automatically, this folder is just a plain clone that would
otherwise silently keep running whatever code was here the last time
someone manually pulled (this bit us once already: a formula/formatting
rewrite was pushed and tested live for a while before anyone noticed this
folder was still several commits behind). If the pull fails (network
blip, merge conflict, etc.) the run aborts with a clear error instead of
silently using stale or partially-updated code -- if that's happening
often enough to be annoying, running `git pull` manually here and
re-running the report is the fallback.

Runs entirely under a temp directory that lives on tmpfs (/dev/shm) when
available, so nothing here touches real disk and everything is cleaned up
before the process exits -- some company servers restrict storage use.
"""

import io
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent

# Mirrors the hosted app's own layout (`uvicorn main:app --app-dir app`) --
# the `app` directory is the import root, so `reports.<module>` resolves
# the same way here as it does on Render.
sys.path.insert(0, str(REPO_DIR / "app"))

# Map of report id -> its module path under `reports/`, and the name of the
# entry-point function to call. Add an entry here for any other report you
# offload later -- no other changes needed in this file.
REPORT_MODULES = {
    "6": ("reports.report_6_battery_disconnection_mail", "process"),
}


def _git_pull():
    result = subprocess.run(
        ["git", "-C", str(REPO_DIR), "pull", "--ff-only"],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30,
    )
    output = (result.stdout + result.stderr).decode(errors="replace")
    print(output, file=sys.stderr)
    if result.returncode != 0:
        print("git pull failed -- aborting rather than run possibly-stale/broken code.", file=sys.stderr)
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("usage: office_server_worker.py <report_id>", file=sys.stderr)
        sys.exit(2)
    report_id = sys.argv[1]

    entry = REPORT_MODULES.get(report_id)
    if entry is None:
        print(f"Unknown or unsupported report id '{report_id}'", file=sys.stderr)
        sys.exit(2)
    module_path, fn_name = entry

    _git_pull()

    import importlib
    module = importlib.import_module(module_path)
    process_fn = getattr(module, fn_name)

    zip_bytes = sys.stdin.buffer.read()
    tmp_base = "/dev/shm" if os.path.isdir("/dev/shm") else None

    with tempfile.TemporaryDirectory(dir=tmp_base) as tmp:
        job_dir = Path(tmp)
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            zf.extractall(job_dir)

        meta = json.loads((job_dir / "meta.json").read_text())
        dates = meta["dates"]
        input_files = {
            slot_key: job_dir / filename for slot_key, filename in meta["files"].items()
        }

        output_path = process_fn(input_files, dates, job_dir)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(output_path, arcname=output_path.name)
        sys.stdout.buffer.write(out_buf.getvalue())


if __name__ == "__main__":
    main()
