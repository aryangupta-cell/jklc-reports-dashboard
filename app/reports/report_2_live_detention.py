"""
JKLC Live Detention — web app adaptation (Report 2 of 9)
==========================================================

Adapted from the provided `jklc_live_detention.py` script. Business logic
is kept as given except for fixes confirmed against real data:

  1. `build_20km_candidates` — Step 5 date/end-time filter revised.
  2. `apply_detention_rules` — detention-hour rule now floor-only (no upper cap):
     keep rows with Detention Hours >= threshold, however long stuck.
  3. `SUMMARY_INCLUDE_REMARKS` — tightened to ["Detention"] only.

Validated against real 14 July data: 70/71 rows exact (1-row Durg gap
traced to a secondary inclusion rule not yet identified; not blocking).

Output now includes 3 tabs per plant file (matching real master file):
  - MTR: cleaned MTR data (all rows, all columns)
  - 20 KM: filtered candidates (Step 5 output)
  - Summary: final detention report (15 columns)

Each tab has proper formatting (Calibri 11pt, Summary at row 3 cols C-Q).

I/O differences from the original script:
  - Input: 3 uploaded files instead of local file paths.
  - Output: 4 plant-wise .xlsx files (3 tabs each), zipped into one .zip.
  - Start/end dates come from the web form (manual range, matching Report 1's
    UI) instead of a module constant / auto-computed window.
"""

import logging
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, Alignment

from reports.errors import ReportProcessingError
from reports.registry import InputSlot, ReportMeta, register

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# CONFIG — same values as the original script. Duplicated rather than
# imported from Report 1 (each report module stays fully self-contained,
# matching how the original standalone scripts were each independent).
# ---------------------------------------------------------------------------

MTR_DROP_COLUMNS = [
    "Probable Unloading Count",
    "Probable Unloading Detention",
    "1 Km Geofence Start Time",
    "1 Km Geofence End Time",
    "1 Km Geofence Detention",
    "20 Km Geofence Start Time",
    "20 Km Geofence End Time",
    "20 Km Geofence Detention",
    "40 Km Geofence Start Time",
    "40 Km Geofence End Time",
    "40 Km Geofence Detention",
    "Lap Sharable Link",
]

ORG_LOCATION_MAP = {
    "Durg Plant": "JKLC Durg",
    "Jharli Grinding": "JKLC Jharli",
    "JK Lakshmi Cement Limited- Surat Grinding Unit": "JKLC Surat",
    "JK Lakshmi Cement Limited - Surat Grinding Unit": "JKLC Surat",
    "MS JK LAKSHMI CEMENT LIMITED CUTTACK GRINDING UNIT": "JKLC Cuttack",
    "M/S JK Lakshmi Cement Limited Cuttack Grinding Unit": "JKLC Cuttack",
}

ORG_LOCATIONS_ALL = ["JKLC Cuttack", "JKLC Durg", "JKLC Jharli", "JKLC Surat"]

# CONFIRMED 14 July 2026 against real data (see notes below): only
# "Detention" belongs in the final Summary. Previous guess of
# ["Detention", "OUT OF GEOFENCE"] was wrong -- dropped.
SUMMARY_INCLUDE_REMARKS = ["Detention"]

SUMMARY_COLUMNS = [
    "Plant Name", "Invoice Number", "Quantity", "Distribution Channel",
    "Vehicle number", "Transporter", "Ship to Region", "Ship to District",
    "Unloading Point", "Invoice Date", "Detention Since", "Detention (Hours)",
    "Ship to Party", "Lead Distance", "Detention Days",
]


def _compute_date_window(report_date_str: str):
    """Confirmed: 1st of REPORT_DATE's month -> REPORT_DATE itself, NO -3
    day trim (different rule from Report 1). Confirmed both by the SOP text
    and the user's own reference table ("Current Month (1 to today)").

    Kept but NOT called from process() below — per explicit instruction, the
    UI takes a manual start/end range instead (consistent with Report 1),
    so this auto-compute isn't used as the default. Left in place in case a
    "suggest dates" convenience feature is wanted later."""
    report_date = pd.to_datetime(report_date_str)
    start = report_date.replace(day=1)
    end = report_date
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# File I/O
# ---------------------------------------------------------------------------

def _read_any(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            return pd.read_csv(path, dtype=str)
        if suffix in (".xlsx", ".xls"):
            return pd.read_excel(path, dtype=str)
    except Exception as exc:
        raise ReportProcessingError(f"Couldn't read '{path.name}' as a {suffix} file: {exc}") from exc
    raise ReportProcessingError(f"Unsupported file type '{suffix}' for '{path.name}'. Expected .csv or .xlsx.")


# ---------------------------------------------------------------------------
# Clean MTR (same as Report 1 / the original script — unchanged)
# ---------------------------------------------------------------------------

def clean_mtr(df: pd.DataFrame, start_date: str, end_date: str) -> pd.DataFrame:
    df = df.copy()
    df = df.drop(columns=[c for c in MTR_DROP_COLUMNS if c in df.columns])
    df["Org Location"] = df["Org Location"].replace(ORG_LOCATION_MAP)

    df["_inv_dt"] = pd.to_datetime(df["Invoice Date & Time"], errors="coerce")
    mask = (df["_inv_dt"] >= start_date) & (df["_inv_dt"] <= end_date + " 23:59:59")
    df = df[mask]

    return df


# ---------------------------------------------------------------------------
# Step 5: build the "20 KM" candidate sheet
#
# CHANGED from the original script — confirmed with the user 14 July 2026,
# based on direct verification against the real 13 July master output:
#
#   Original filter: Mode in {AT FIX, GPS-API}, Lead Distance > 20,
#   Proximity Start date == REPORT_DATE exactly, Proximity End Time blank.
#   The original script's docstring called this "CONFIRMED EXACT (173/173)."
#
#   Verification against the real file showed that claim was a coincidental
#   ROW-COUNT match, not a real row-identity match: comparing by Invoice No,
#   only 2 of 172 real rows actually matched. The other 170 real rows had
#   Proximity Start dates ranging from 6-12 July (detentions carrying over
#   across multiple days, not just "started today"), and 29/172 real rows
#   (17%) had a non-blank Proximity End Time yet were still included.
#
#   Revised filter: Proximity Start anywhere within [start_date, end_date]
#   (no exact-day requirement), End Time condition dropped entirely (Step 7's
#   Dispatch not-yet-delivered check is the real "still open" signal, not
#   Proximity End Time). This produces 95% overlap (164/172) with the real
#   20 KM sheet, up from ~1% with the original literal-SOP filter.
# ---------------------------------------------------------------------------

def build_20km_candidates(mtr_clean: pd.DataFrame, start_date: str, end_date: str) -> pd.DataFrame:
    df = mtr_clean.copy()
    df["Lead Distance"] = pd.to_numeric(df["Lead Distance"], errors="coerce")
    df["_proximity_start"] = pd.to_datetime(
        df["System Destination Proximity Start Time"], errors="coerce"
    )

    mask = (
        df["Mode"].isin(["AT FIX", "GPS-API"])
        & (df["Lead Distance"] > 20)
        & (df["_proximity_start"] >= start_date)
        & (df["_proximity_start"] <= end_date + " 23:59:59")
    )
    return df[mask].copy()


# ---------------------------------------------------------------------------
# Steps 6-9: Dispatch filter + detention-hour rule + Stamp Status rejected
# (unchanged from the original script)
# ---------------------------------------------------------------------------

def filter_dispatch_for_detention(dispatch_raw: pd.DataFrame) -> pd.DataFrame:
    df = dispatch_raw.copy()
    df["EPOD_Timestamp"] = pd.to_numeric(df["EPOD_Timestamp"], errors="coerce").fillna(0)
    df["MIGO_Timestamp"] = pd.to_numeric(df["MIGO_Timestamp"], errors="coerce").fillna(0)
    return df[(df["EPOD_Timestamp"] == 0) & (df["MIGO_Timestamp"] == 0)]


def apply_detention_rules(candidates: pd.DataFrame, dispatch_filtered: pd.DataFrame,
                           now: pd.Timestamp) -> pd.DataFrame:
    disp_inv = set(dispatch_filtered["Invoice Number"].astype(str).str.strip())
    df = candidates[candidates["Invoice no"].astype(str).str.strip().isin(disp_inv)].copy()

    # Detention (Hours) — NOT a raw MTR column, computed here (see original
    # script's docstring note on this being unconfirmed vs Khagash's exact
    # "NOW" reference point).
    df["Detention (Hours)"] = (now - df["_proximity_start"]).dt.total_seconds() / 3600

    # Detention-hour rule: floor only (no upper cap). Keep trips with
    # detention AT LEAST the threshold, however long they've been stuck.
    # Lead Distance ≤ 200 KM → floor is 24 hours.
    # Lead Distance > 200 KM → floor is 48 hours.
    mask = (
        ((df["Lead Distance"] <= 200) & (df["Detention (Hours)"] >= 24))
        | ((df["Lead Distance"] > 200) & (df["Detention (Hours)"] >= 48))
    )
    df = df[mask]

    df = df[df["Stamp Status"] != "Rejected"]
    return df


# ---------------------------------------------------------------------------
# Merge in the Detention Bot's remark (pass-through, no reclassification,
# unchanged from the original script)
# ---------------------------------------------------------------------------

def merge_bot_remarks(df: pd.DataFrame, bot_output: pd.DataFrame) -> pd.DataFrame:
    bot = bot_output[bot_output["status"] == "ok"][["trip_id", "remark"]].copy()
    bot["trip_id"] = bot["trip_id"].astype(str).str.strip()
    df = df.copy()
    df["Trip ID"] = df["Trip ID"].astype(str).str.strip()
    df = df.merge(bot, left_on="Trip ID", right_on="trip_id", how="inner")
    df = df.rename(columns={"remark": "Remark"})
    return df


# ---------------------------------------------------------------------------
# Build the Summary tab (unchanged from the original script)
# ---------------------------------------------------------------------------

def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # CONFIRMED 14 July 2026 -- see SUMMARY_INCLUDE_REMARKS definition above.
    df = df[df["Remark"].isin(SUMMARY_INCLUDE_REMARKS)]

    df["Detention Days"] = df["Detention (Hours)"] / 24

    out = pd.DataFrame({
        "Plant Name": df["Org Location"],
        "Invoice Number": df["Invoice no"],
        "Quantity": df.get("Quantity", ""),
        "Distribution Channel": df.get("Distribution Channel", ""),
        "Vehicle number": df["Vehicle No."],
        "Transporter": df["Transporter"],
        "Ship to Region": df.get("Ship to District", ""),
        "Ship to District": df.get("Ship to District", ""),
        "Unloading Point": df.get("Destination", ""),
        "Invoice Date": df["Invoice Date & Time"],
        "Detention Since": df["_proximity_start"],
        "Detention (Hours)": df["Detention (Hours)"],
        "Ship to Party": df.get("SOLD TO NM", ""),
        "Lead Distance": df["Lead Distance"],
        "Detention Days": df["Detention Days"],
    })
    return out


# ---------------------------------------------------------------------------
# Format utilities for Excel output (matching the real master file exactly)
# ---------------------------------------------------------------------------

def _style_mtr_headers(ws, n_cols: int, bold: bool = False):
    """Style header row 1 of an MTR/20KM sheet. Data already written by pandas."""
    header_font = Font(name="Calibri", size=11, bold=bold)
    center_align = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(1, col_idx)
        cell.font = header_font
        cell.alignment = center_align


def _write_summary_sheet(ws, df: pd.DataFrame):
    """Write Summary sheet: headers at row 3 cols C-Q, data from row 4, Calibri 11 centered."""
    header_font = Font(name="Calibri", size=11, bold=False)
    data_font = Font(name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center")

    for col_offset, col_name in enumerate(df.columns):
        col_idx = 3 + col_offset  # Column C = 3
        cell = ws.cell(3, col_idx, value=col_name)
        cell.font = header_font
        cell.alignment = center_align

    for row_idx, (_, row) in enumerate(df.iterrows(), start=4):
        for col_offset, value in enumerate(row):
            col_idx = 3 + col_offset
            cell = ws.cell(row_idx, col_idx, value=value)
            cell.font = data_font
            cell.alignment = center_align


# ---------------------------------------------------------------------------
# Split into 4 plant-wise files, zip them for a single download
# ---------------------------------------------------------------------------

def _write_plant_outputs_zip(summary: pd.DataFrame, mtr_clean: pd.DataFrame,
                             candidates_20km: pd.DataFrame, date_label: str,
                             output_dir: Path) -> Path:
    """Write 4 plant-wise Excel files, each with 3 tabs (MTR, 20 KM, Summary)."""
    # Drop internal helper columns before writing
    mtr_write = mtr_clean.drop(columns=["_inv_dt"], errors="ignore")
    km20_write = candidates_20km.drop(columns=["_inv_dt", "_proximity_start"], errors="ignore")

    xlsx_paths = []

    for plant in ORG_LOCATIONS_ALL:
        plant_short = plant.replace("JKLC ", "")
        xlsx_path = output_dir / f"JKLC_Live_Detention_{plant_short}_{date_label}.xlsx"

        plant_summary = summary[summary["Plant Name"] == plant].copy()

        # Use pandas ExcelWriter (fast bulk write) for MTR and 20 KM tabs.
        # Summary is small enough to write cell-by-cell for exact positioning
        # (headers at row 3, col C — not row 1, col A like pandas does).
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            mtr_write.to_excel(writer, sheet_name="MTR", index=False)
            km20_write.to_excel(writer, sheet_name="20 KM", index=False)
            # Summary written as empty placeholder; real write happens below
            pd.DataFrame().to_excel(writer, sheet_name="Summary", index=False)

        # Re-open to apply formatting on headers and write Summary correctly
        from openpyxl import load_workbook as _lw
        wb = _lw(xlsx_path)

        _style_mtr_headers(wb["MTR"], len(mtr_write.columns), bold=True)
        _style_mtr_headers(wb["20 KM"], len(km20_write.columns), bold=False)

        # Clear the placeholder Summary sheet and write it properly
        del wb["Summary"]
        ws_summary = wb.create_sheet("Summary")
        _write_summary_sheet(ws_summary, plant_summary)

        # Reorder sheets: MTR, 20 KM, Summary
        wb._sheets.sort(key=lambda s: ["MTR", "20 KM", "Summary"].index(s.title)
                        if s.title in ["MTR", "20 KM", "Summary"] else 99)

        wb.save(xlsx_path)
        log.info("%s: MTR %d rows, 20KM %d rows, Summary %d rows -> %s",
                 plant, len(mtr_write), len(km20_write), len(plant_summary), xlsx_path.name)
        xlsx_paths.append(xlsx_path)

    zip_path = output_dir / f"JKLC_Live_Detention_{date_label}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in xlsx_paths:
            zf.write(p, arcname=p.name)
    return zip_path


# ---------------------------------------------------------------------------
# Entry point used by the generic upload -> process -> download route
# ---------------------------------------------------------------------------

def process(input_files: dict, dates: dict, output_dir: Path) -> Path:
    mtr_path = input_files["mtr_raw"]
    dispatch_path = input_files["dispatch"]
    bot_path = input_files["detention_bot"]

    # Manual start/end range, per explicit instruction (consistent with
    # Report 1) — NOT run through _compute_date_window(), which is kept but
    # unused (see that function's docstring for the confirmed "1st of month
    # to report date" rule it implements).
    start_date = dates["start_date"]
    end_date = dates["end_date"]
    log.info("Processing JKLC Live Detention for window %s to %s", start_date, end_date)

    mtr_raw = _read_any(mtr_path)
    dispatch_raw = _read_any(dispatch_path)
    bot_output = _read_any(bot_path)

    try:
        mtr_clean = clean_mtr(mtr_raw, start_date, end_date)
        candidates = build_20km_candidates(mtr_clean, start_date, end_date)
        dispatch_filtered = filter_dispatch_for_detention(dispatch_raw)
    except KeyError as exc:
        raise ReportProcessingError(
            f"Expected column {exc} not found. Check you uploaded the correct "
            "file for each slot (MTR Raw / Daily Dispatch / Detention Bot output)."
        ) from exc

    now = pd.Timestamp(datetime.now())
    detained = apply_detention_rules(candidates, dispatch_filtered, now)

    try:
        merged = merge_bot_remarks(detained, bot_output)
    except KeyError as exc:
        raise ReportProcessingError(
            f"Expected column {exc} not found in the Detention Bot output. "
            "Check you uploaded the bot's output CSV (trip_id, remark, status, error)."
        ) from exc

    summary = build_summary(merged)
    log.info("Final Summary rows: %d (%s)", len(summary), dict(summary["Plant Name"].value_counts()))

    date_label = f"{start_date}_to_{end_date}"
    return _write_plant_outputs_zip(summary, mtr_clean, candidates, date_label, output_dir)


register(
    ReportMeta(
        id="2",
        name="JKLC Live Detention",
        input_slots=[
            InputSlot(
                key="mtr_raw",
                label="MTR Raw",
                accept=".csv,.xlsx",
                hint="mtr - 2026-07-13T171021.301.csv",
            ),
            InputSlot(
                key="dispatch",
                label="Daily Dispatch",
                accept=".xlsx",
                hint="daily dispatch 13 Axestrack (75).xlsx",
            ),
            InputSlot(
                key="detention_bot",
                label="Detention Bot Output",
                accept=".csv",
                hint="bot output 13 detention_results.csv",
            ),
        ],
        output_pattern="JKLC_Live_Detention_<start>_to_<end>.zip (4 plants, 3 tabs each: MTR, 20 KM, Summary)",
        process_fn=process,
        implemented=True,
        date_mode="range",
        notes=(
            "3 tabs per plant: MTR (all rows), 20 KM (filtered candidates), Summary (detention report). "
            "Detention-hour rule is floor-only (no upper cap). Formatting matches real master file. "
            "Validated 70/71 rows on real 14 July data; 1-row gap not root-caused."
        ),
    )
)
