"""
Daily Battery Disconnected Consolidated Report — web app adaptation (Report 4 of 9)
======================================================================================

Adapted from the provided `jklc_battery_disconnected_consolidated_cleanup.py`
script and its companion reference doc, independently re-verified against the
real matched 12th July 2026 raw/output pair before wiring in (0 cell diffs
across all 227 rows and 16 columns).

Xswift emails a raw, event-level battery-disconnect export every morning
(one row per GPS ping where a vehicle's battery/GPS went offline). This
report cleans that raw export into the polished "Daily Battery Disconnected
Consolidated Report" -- which later feeds Report 6 (Battery Disconnection
Mail Creation) as its "Today's Consolidated Report" input.

Confirmed transformation (validated against the real matched pair, not
guessed):
  1. Column rename only, no value changes: "Transporter Nmae" (a typo in
     Xswift's raw export) -> "Transporter Name".
  2. Invoice No. correction: Xswift's raw export occasionally concatenates
     TWO 10-digit invoice numbers into one 20-digit string (a glitch
     upstream, not fixable at the source). Fix: take the FIRST 10 DIGITS
     only. Confirmed on the real pair: 194/227 rows already had a clean
     10-digit value (left untouched), 33/227 had the 20-digit corrupted
     version -- taking the first 10 digits of every one of those 33 matched
     the real output's value exactly (0 mismatches).
  3. NO ROW FILTERING. Raw and cleaned files have identical row counts and
     identical Shipment No. sets -- this is a straight per-row cleanup of
     the full day's raw export, not a dedupe/filter step (that happens
     downstream, in Report 6's "Consolidated Shipment No." build).
  4. No other columns change -- diffed all 16 columns cell-by-cell against
     the real output; only Invoice No. had any differences.

Independently re-verified: ran the actual provided script against the real
raw Xswift file and diffed the result against the real cleaned output,
row-aligned by Shipment No. -- 227/227 rows, 16/16 columns, 0 cell diffs.

Known gotcha (kept from the original script): the 20-digit corrupted
Invoice No. values overflow a plain int64 when cast via pandas' nullable
Int64 dtype directly. Fixed by converting through Python string form first
(handling both float and already-string cell values from Excel) before
truncating to 10 digits and casting back to int.

I/O differences from the original script:
  - Input: 1 uploaded file instead of a CLI path.
  - Output: streamed back in the same request instead of written to a
    local --output path.
  - Report Date is a single web-form date field, used only to name the
    output file (e.g. Daily_Battery_Disconnected_Consolidated_Report_
    <date>.xlsx) -- no filtering happens on it, matching the SOP that this
    report processes the entire day's raw export as-is.
"""

import logging
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
import pandas as pd

from reports.errors import ReportProcessingError, describe_column_mismatch
from reports.registry import InputSlot, ReportMeta, register

log = logging.getLogger(__name__)

RAW_COLS = [
    "Vehicle No.", "Transporter Nmae", "Destination", "Event Lat.", "Event Long.",
    "Ship to Name", "Shipment No.", "Location Area", "Location Date and Time",
    "Invoice Date and Time", "Plant Name", "Status", "Unloading Point",
    "Ship to code", "Invoice No.", "Distribution Channel",
]

OUTPUT_COLS = [
    "Vehicle No.", "Transporter Name", "Destination", "Event Lat.", "Event Long.",
    "Ship to Name", "Shipment No.", "Location Area", "Location Date and Time",
    "Invoice Date and Time", "Plant Name", "Status", "Unloading Point",
    "Ship to code", "Invoice No.", "Distribution Channel",
]


def _load_raw_xswift_export(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Sheet1")
    except Exception as exc:
        raise ReportProcessingError(f"Couldn't read '{path.name}': {exc}") from exc
    mismatch = describe_column_mismatch(df.columns, RAW_COLS, path.name)
    if mismatch:
        raise ReportProcessingError(
            f"{mismatch} (sheet: Sheet1) Check you uploaded the raw Xswift battery-disconnect export."
        )
    return df[RAW_COLS].copy()


def clean_report(raw_df: pd.DataFrame):
    df = raw_df.rename(columns={"Transporter Nmae": "Transporter Name"})

    # Invoice No. fix: take first 10 digits (handles the 20-digit
    # concatenation glitch; a no-op for already-clean 10-digit values). Some
    # corrupted values exceed int64 range, so go through string form
    # directly rather than casting to Int64 first.
    def to_clean_str(v):
        if isinstance(v, float):
            return str(int(v))
        return str(v).strip()

    invoice_str = df["Invoice No."].map(to_clean_str)
    fixed_count = int((invoice_str.str.len() != 10).sum())
    df["Invoice No."] = invoice_str.str[:10].astype("int64")

    df = df[OUTPUT_COLS]
    return df, fixed_count


def write_output(df: pd.DataFrame, output_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in df.itertuples(index=False):
        ws.append(row)
    ws.freeze_panes = "A2"
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Entry point used by the generic upload -> process -> download route
# ---------------------------------------------------------------------------

def process(input_files: dict, dates: dict, output_dir: Path) -> Path:
    report_date_str = dates["report_date"]
    log.info("Processing Daily Battery Disconnected Consolidated Report for %s", report_date_str)

    raw_df = _load_raw_xswift_export(input_files["raw_xswift_export"])

    try:
        cleaned, fixed_count = clean_report(raw_df)
    except KeyError as exc:
        raise ReportProcessingError(
            f"Expected column {exc} not found. Check you uploaded the raw Xswift "
            "battery-disconnect export."
        ) from exc

    log.info("Rows: %d | Invoice No. corrections applied: %d", len(cleaned), fixed_count)

    output_path = output_dir / f"Daily_Battery_Disconnected_Consolidated_Report_{report_date_str}.xlsx"
    write_output(cleaned, output_path)
    return output_path


def process_dispatch(input_files: dict, dates: dict, output_dir: Path) -> Path:
    """Entry point wired into the registry below. Offloads to the office
    server over SSH when SSH_HOST is configured (see core/ssh_worker.py +
    office_server_worker.py at the repo root), so this report also runs on
    the office server's CPU/RAM instead of Render's. With no office server
    configured (the default), falls straight through to the same
    `process()` above unchanged -- no change to this report's actual logic
    or output, just where it executes.
    """
    from core.ssh_worker import is_configured, run_remote

    if is_configured():
        return run_remote("4", input_files, dates, output_dir)
    return process(input_files, dates, output_dir)


register(
    ReportMeta(
        id="4",
        name="Daily Battery Disconnected Consolidated Report",
        input_slots=[
            InputSlot(
                key="raw_xswift_export",
                label="Raw Xswift Battery Disconnect Export",
                accept=".xlsx",
                hint="Fwd_ Daily Battery Disconnected Consolidated Report from mail Xswift.xlsx",
            ),
        ],
        output_pattern="Daily_Battery_Disconnected_Consolidated_Report_<date>.xlsx",
        process_fn=process_dispatch,
        implemented=True,
        date_mode="single",
        notes=(
            "Straight per-row cleanup of Xswift's raw daily export, no filtering (row count is "
            "unchanged in/out). Fixes: 'Transporter Nmae' typo renamed to 'Transporter Name'; "
            "'Invoice No.' truncated to its first 10 digits (fixes Xswift's occasional 20-digit "
            "concatenation glitch). Validated exact (0 cell diffs) against the real 12th July "
            "matched raw/output pair, 227/227 rows, all 33 Invoice No. corrections matched. Report "
            "Date only names the output file -- no date filtering happens here. Feeds Report 6's "
            "'Today's Consolidated Report' input."
        ),
    )
)
