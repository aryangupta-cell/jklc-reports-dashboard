"""
SSH offload for reports too heavy for Render's free tier.

Same pattern already set up for Anchal's app: instead of running a report's
pandas pipeline on the small hosted instance, stream the input files to a
company server with real RAM/CPU over SSH, run the exact same tested
pipeline there, and stream the generated file back. Env var names below
match that existing setup so the two apps can be configured the same way.

Pairs with `office_server_worker.py` (repo root), which must be deployed
alongside a checkout of this repo on the office server and is invoked fresh
per request -- not a standing service, no process left running 24/7.

Uses the system `ssh` binary via subprocess rather than paramiko: paramiko
unreliably raised "Socket is closed" on large payloads in testing regardless
of chunking; shelling out was solid.

Configuration -- all via environment variables, nothing sensitive hardcoded
or committed:
  SSH_HOST        hostname or IP of the office server. Offload is only
                   attempted when this is set (see is_configured());
                   otherwise callers fall back to local processing, unchanged.
  SSH_USER         SSH username
  SSH_KEY_PATH     path to the private key file -- e.g. a Render "Secret
                   File", read from /etc/secrets/<filename>
  SSH_HOST_KEY     the office server's known_hosts line (e.g. the output of
                   `ssh-keyscan <host>`), used to pin host identity instead
                   of trusting on first use. If unset, falls back to
                   StrictHostKeyChecking=accept-new.
  REMOTE_DIR       directory on the office server holding this repo's
                   office_server_worker.py + its venv
  REMOTE_PYTHON    path to the python interpreter to run it with (typically
                   the venv's own interpreter, so it has pandas/openpyxl)
"""

import io
import json
import logging
import os
import subprocess
import tempfile
import zipfile
from pathlib import Path

from reports.errors import ReportProcessingError

log = logging.getLogger(__name__)

DEFAULT_TIMEOUT_SECONDS = 300


def is_configured() -> bool:
    """True once SSH_HOST is set. Callers use this to decide whether to
    offload to the office server or run the report locally as before --
    with this unset (the default), nothing about existing behavior changes."""
    return bool(os.environ.get("SSH_HOST"))


def _build_input_zip(input_files: dict, dates: dict) -> bytes:
    meta = {"dates": dates, "files": {key: path.name for key, path in input_files.items()}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("meta.json", json.dumps(meta))
        for path in input_files.values():
            zf.write(path, arcname=path.name)
    return buf.getvalue()


def run_remote(report_id: str, input_files: dict, dates: dict, output_dir: Path) -> Path:
    """Runs the given report's process_fn on the office server instead of
    locally. Requires SSH_HOST etc. to be set -- check is_configured() first."""
    host = os.environ["SSH_HOST"]
    user = os.environ.get("SSH_USER", "")
    key_path = os.environ.get("SSH_KEY_PATH")
    host_key = os.environ.get("SSH_HOST_KEY")
    remote_dir = os.environ.get("REMOTE_DIR", "~/mtr-report-worker")
    remote_python = os.environ.get("REMOTE_PYTHON", "python3")

    target = f"{user}@{host}" if user else host
    cmd = ["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=15"]

    known_hosts_path = None
    if host_key:
        fd, known_hosts_path = tempfile.mkstemp(prefix="ssh_known_hosts_")
        with os.fdopen(fd, "w") as f:
            f.write(host_key.strip() + "\n")
        cmd += ["-o", f"UserKnownHostsFile={known_hosts_path}", "-o", "StrictHostKeyChecking=yes"]
    else:
        cmd += ["-o", "StrictHostKeyChecking=accept-new"]

    if key_path:
        cmd += ["-i", key_path]
    cmd += [target, f"{remote_python} {remote_dir}/office_server_worker.py {report_id}"]

    input_zip = _build_input_zip(input_files, dates)

    try:
        try:
            result = subprocess.run(
                cmd,
                input=input_zip,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=DEFAULT_TIMEOUT_SECONDS,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise ReportProcessingError(
                "The office server took too long to respond (report offload timed out)."
            ) from exc
        except FileNotFoundError as exc:
            raise ReportProcessingError(
                "The 'ssh' command isn't available on this host -- can't offload to the office server."
            ) from exc
    finally:
        if known_hosts_path:
            os.remove(known_hosts_path)

    if result.returncode != 0:
        stderr_tail = result.stderr.decode(errors="replace")[-2000:]
        log.error("Office-server worker failed (report %s): %s", report_id, stderr_tail)
        raise ReportProcessingError(
            "Report processing on the office server failed. Details have been logged."
        )

    if not result.stdout:
        raise ReportProcessingError("Office server returned no output for this report.")

    with zipfile.ZipFile(io.BytesIO(result.stdout)) as zf:
        names = zf.namelist()
        if not names:
            raise ReportProcessingError("Office server's response didn't contain a file.")
        output_name = names[0]
        output_path = output_dir / output_name
        output_path.write_bytes(zf.read(output_name))

    return output_path
