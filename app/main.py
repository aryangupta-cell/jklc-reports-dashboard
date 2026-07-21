import logging
import mimetypes
import tempfile
from pathlib import Path

from fastapi import Depends, FastAPI, Request
from fastapi.responses import RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from core.auth import require_settings_password
from core.settings_store import (
    ReportOverride,
    SlotOverride,
    get_effective_reports,
    load_overrides,
    save_overrides,
)
from reports import registry  # noqa: F401 (registers report_1..report_9 on import)
from reports.errors import ReportProcessingError
from reports.registry import REPORTS

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent

# No auth for now (internal tool, per explicit instruction) — this means the
# deployed URL is fully public. To re-enable: add
# `dependencies=[Depends(require_auth)]` back here, using core/auth.py, which
# is still in place (HTTP Basic, reads DASHBOARD_USERNAME/DASHBOARD_PASSWORD).
app = FastAPI(title="ATC Report Dashboard")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")


@app.get("/")
def index(request: Request):
    reports = get_effective_reports(enabled_only=True)
    return templates.TemplateResponse(request, "index.html", {"reports": reports})


@app.get("/report/{report_id}")
def report_page(request: Request, report_id: str):
    meta = next((r for r in get_effective_reports() if r.id == report_id), None)
    if meta is None:
        return templates.TemplateResponse(
            request,
            "report.html",
            {"report": None, "error": f"No report with id '{report_id}'."},
            status_code=404,
        )
    return templates.TemplateResponse(request, "report.html", {"report": meta})


def _collect_dates(form, date_mode: str) -> dict:
    if date_mode == "range":
        start_date = form.get("start_date")
        end_date = form.get("end_date")
        if not start_date or not end_date:
            raise ReportProcessingError("Please pick both a start date and an end date.")
        return {"start_date": start_date, "end_date": end_date}
    if date_mode == "single":
        report_date = form.get("report_date")
        if not report_date:
            raise ReportProcessingError("Please pick a report date.")
        return {"report_date": report_date}
    return {}


def _collect_extra_number_fields(form, extra_number_fields) -> dict:
    """Optional extra number inputs a report declares (e.g. Report 3's "Days
    back") land in the same `dates` dict process_fn already receives, keyed
    by field.key. Blank/invalid input falls back to the field's default
    rather than erroring — these are meant to be optional conveniences, not
    required inputs."""
    result = {}
    for extra_field in extra_number_fields:
        raw = form.get(extra_field.key)
        try:
            value = int(raw)
        except (TypeError, ValueError):
            value = extra_field.default
        if value < extra_field.min_value:
            value = extra_field.min_value
        result[extra_field.key] = value
    return result


@app.post("/report/{report_id}/generate")
async def generate(request: Request, report_id: str):
    meta = next((r for r in get_effective_reports() if r.id == report_id), None)
    if meta is None:
        return templates.TemplateResponse(
            request,
            "report.html",
            {"report": None, "error": f"No report with id '{report_id}'."},
            status_code=404,
        )

    form = await request.form()

    try:
        dates = _collect_dates(form, meta.date_mode)
    except ReportProcessingError as exc:
        return templates.TemplateResponse(request, "report.html", {"report": meta, "error": str(exc)})
    dates.update(_collect_extra_number_fields(form, meta.extra_number_fields))

    # Everything lives in a per-request temp dir and is gone by the time this
    # handler returns — the generated file is read into memory and sent back
    # in THIS response, so there's no separate download step relying on a
    # second request hitting the same server instance (which serverless
    # platforms like Vercel don't guarantee).
    with tempfile.TemporaryDirectory() as tmp:
        job_dir = Path(tmp)

        input_paths = {}
        for slot in meta.input_slots:
            upload = form.get(slot.key)
            if upload is None or not getattr(upload, "filename", ""):
                return templates.TemplateResponse(
                    request,
                    "report.html",
                    {
                        "report": meta,
                        "error": f"Missing file for '{slot.label}'. Please choose a file for every input.",
                    },
                )
            dest = job_dir / upload.filename
            contents = await upload.read()
            dest.write_bytes(contents)
            input_paths[slot.key] = dest

        try:
            output_path = meta.process_fn(input_paths, dates, job_dir)
        except NotImplementedError as exc:
            return templates.TemplateResponse(request, "report.html", {"report": meta, "error": str(exc)})
        except ReportProcessingError as exc:
            return templates.TemplateResponse(request, "report.html", {"report": meta, "error": str(exc)})
        except KeyError as exc:
            # Safety net: most reports already catch KeyError internally and
            # raise a clearer ReportProcessingError (see each report's own
            # process_fn) — this only fires for the ones that don't, or for
            # a KeyError raised outside that internal try block. Without
            # this, the column name gets lost in the generic "something
            # went wrong" message below. exc.args[0] is the missing column
            # name (or a list of names, for a multi-column selection).
            missing = exc.args[0] if exc.args else str(exc)
            log.exception("Missing column in report %s (%s)", report_id, meta.name)
            return templates.TemplateResponse(
                request,
                "report.html",
                {
                    "report": meta,
                    "error": f"'{meta.name}' expected a column that wasn't found in the uploaded "
                    f"file(s): {missing}. Check each file was uploaded to the correct input slot, "
                    "and that it's the right kind of file for this report.",
                },
            )
        except Exception:
            log.exception("Unexpected error processing report %s", report_id)
            return templates.TemplateResponse(
                request,
                "report.html",
                {
                    "report": meta,
                    "error": "Something went wrong while generating this report. "
                    "Details have been logged — contact the report owner.",
                },
            )

        content = output_path.read_bytes()
        output_name = output_path.name

    media_type = mimetypes.guess_type(output_name)[0] or "application/octet-stream"
    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{output_name}"'},
    )


@app.get("/settings", dependencies=[Depends(require_settings_password)])
def settings_page(request: Request):
    overrides = load_overrides()
    reports_by_id = {r.id: r for r in REPORTS.values()}
    # Pair each report's code-defined slots (for keys/order) with its override values.
    rows = []
    for report_id in sorted(reports_by_id, key=int):
        meta = reports_by_id[report_id]
        o = overrides[report_id]
        rows.append({"meta": meta, "override": o})
    rows.sort(key=lambda r: r["override"].order)
    return templates.TemplateResponse(request, "settings.html", {"rows": rows})


@app.post("/settings/save", dependencies=[Depends(require_settings_password)])
async def settings_save(request: Request):
    form = await request.form()
    reports_by_id = {r.id: r for r in REPORTS.values()}

    overrides = {}
    for report_id, meta in reports_by_id.items():
        slots = {
            slot.key: SlotOverride(
                label=form.get(f"label__{report_id}__{slot.key}", slot.label),
                hint=form.get(f"hint__{report_id}__{slot.key}", slot.hint),
            )
            for slot in meta.input_slots
        }
        overrides[report_id] = ReportOverride(
            name=form.get(f"name__{report_id}", meta.name),
            notes=form.get(f"notes__{report_id}", meta.notes),
            output_pattern=form.get(f"output_pattern__{report_id}", meta.output_pattern),
            enabled=form.get(f"enabled__{report_id}") == "on",
            order=int(form.get(f"order__{report_id}", 0) or 0),
            slots=slots,
        )

    save_overrides(overrides)
    return RedirectResponse(url="/settings?saved=1", status_code=303)
