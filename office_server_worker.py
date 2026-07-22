#!/usr/bin/env python3
"""
Office-server worker -- runs a report's normal, unchanged `process()`
pipeline on this machine instead of on the hosted (Render) instance, so
report generation uses this server's CPU/RAM instead of Render's free
tier. Covers all 6 currently-implemented reports (see REPORT_MODULES
below); each report's own `process()` logic is completely untouched --
only where it executes changes.

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
    "1": ("reports.report_1_daily_tracker", "process"),
    "2": ("reports.report_2_live_detention", "process"),
    "3": ("reports.report_3_daily_tracking", "process"),
    "4": ("reports.report_4_battery_disconnected", "process"),
    "5": ("reports.report_5_at_fix_ontrip", "process"),
    "6": ("reports.report_6_battery_disconnection_mail", "process"),
    "7": ("reports.report_7_control_tower_tracker", "process"),
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


_REEXEC_FLAG = "_OFFICE_WORKER_PULLED"


def main():
    if len(sys.argv) < 2:
        print("usage: office_server_worker.py <report_id>", file=sys.stderr)
        sys.exit(2)

    # Pull BEFORE anything else -- including before even looking at
    # REPORT_MODULES. A previous version checked the report id against
    # REPORT_MODULES first and only pulled afterwards, so a brand-new
    # report id (added by the very commit this pull would fetch) failed
    # immediately every time, without ever reaching the pull.
    #
    # Pulling alone still isn't enough, though: Python already reads and
    # compiles this entire file -- including REPORT_MODULES's dict literal
    # further down -- before any of its top-level code runs. A git pull
    # only updates the file ON DISK; it can't retroactively change what
    # this already-running process has in memory. So after a pull, re-run
    # this same script as a fresh subprocess -- the new process reads the
    # updated file from disk before running anything, which is what
    # actually makes a newly pulled REPORT_MODULES entry (or any other
    # code change in this file) take effect in the SAME invocation instead
    # of only benefiting the next one. stdin/stdout/stderr are passed
    # through directly so the subprocess is transparent to the SSH caller
    # on the other end (see core/ssh_worker.py's protocol). The env var
    # flag prevents pulling/re-running a second time once already fresh.
    #
    # (Uses subprocess.run rather than os.execv/execve for this -- despite
    # execv being the more "obvious" fit for in-place re-exec, it couldn't
    # be verified reliably in local testing on this dev machine, and a
    # spawned-subprocess-with-inherited-stdio achieves the same effect in a
    # way that's portable and was actually confirmed working end-to-end.)
    if not os.environ.get(_REEXEC_FLAG):
        _git_pull()
        env = os.environ.copy()
        env[_REEXEC_FLAG] = "1"
        result = subprocess.run(
            [sys.executable, os.path.abspath(__file__), *sys.argv[1:]],
            stdin=sys.stdin.buffer, stdout=sys.stdout.buffer, stderr=sys.stderr.buffer,
            env=env,
        )
        sys.exit(result.returncode)

    report_id = sys.argv[1]
    entry = REPORT_MODULES.get(report_id)
    if entry is None:
        print(f"Unknown or unsupported report id '{report_id}'", file=sys.stderr)
        sys.exit(2)
    module_path, fn_name = entry

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
