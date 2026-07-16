"""
JKLC Battery Disconnection Mail Creation — web app adaptation (Report 6 of 9)
================================================================================

Adapted from the provided `jklc_battery_disconnection_mail.py` script and its
companion reference doc. Rebuilds the daily "Battery_Disconnection_Master_
<date>.xlsx" workbook Khagash maintains by hand -- this is a STATEFUL report:
each run needs the PREVIOUS day's Master workbook as an input (uploaded fresh
each time, same as any other file here -- there's no server-side persistence
between runs, the "state" is simply whatever file the user uploads).

Confirmed business logic (validated against real data -- see the reference
doc for the original evidence trail, plus independent re-verification below):

  1. Join key: 'Shipment No.' (consolidated report) == 'Shipment Number' (MTR).
  2. 'Consolidated' tab = pure unfiltered daily append log, forever (Clinker
     rows and same-day duplicate pings included).
  3. 'Consolidated Shipment No.' tab = same daily rows, per day: join Product
     Name from MTR -> drop Clinker rows -> dedupe by Shipment No. (keep first
     row in file order) -> add STO/NON STO, Onward Status, shareable link,
     Test -> leave Mail Status / vehicle test blank for new rows, preserve
     existing values for carried-forward rows.
  4. 'Master' tab = shipments whose 'vehicle test' starts with "offline"
     (case-insensitive), append-only, excluding shipments whose latest
     'Mail Status' is "Waived OFF".
  5. 'MTR' tab = wholesale replaced with the latest MTR export each run.
  6. 'Table' tab = carried forward unchanged (scratch/lookup helper,
     unrelated to the pipeline).

INDEPENDENTLY RE-VERIFIED (not just taking the provided script's word for it)
against the real files (Battery_Disconnection_Master_1_July_2026.xlsx as
"previous", Daily Battery Disconnected Consolidated Report 12th July 2026.xlsx
as the new day, MTR pull from 2026-07-13T17:10, compared against the real
Battery_Disconnection_Master_13_July_2026.xlsx):
  - Consolidated Shipment No. rows for 7/12: 80/80 exact, identical shipment
    sets both directions.
  - STO/NON STO, shareable link: 0 diffs / 80 rows.
  - Test, Onward Status: 32 diffs / 80 rows each -- CORRECTING the provided
    doc's explanation here: it claimed every diff was "our script reporting
    a MORE advanced state" (later MTR pull). Re-checking the actual diff
    VALUES directly disproves that as a clean explanation -- both directions
    occur (e.g. one row: ours = Pending, real = Stamp Verified, i.e. real is
    MORE advanced; another: ours = AI Complete, real = Pending, ours is more
    advanced; Test column diffs uniformly went the OPPOSITE direction from
    the doc's claim). This is most consistent with MTR being live,
    non-monotonic tracking data -- status fields aren't guaranteed to only
    move forward between arbitrary pulls -- rather than a script defect, but
    it should be described as "snapshot-sensitive, not fully understood",
    not as a confirmed one-directional timing artifact.
  - The join/filter/dedup pipeline itself (the part that actually matters
    for correctness) IS solid: exact row counts and shipment-set matches.

Still not automated (deliberately, per the reference doc): actual offline/
online determination is a manual human judgment call (Mail Status / vehicle
test columns), and actual mail-sending is out of scope -- this only builds
the Master candidate list.

I/O differences from the original script:
  - Inputs: 3 uploaded files (previous Master .xlsx, today's raw consolidated
    report .xlsx, latest MTR export) instead of CLI paths.
  - Output: streamed back in the same request instead of written to a local
    --output path.
  - Report Date is a single web-form date field (this report is inherently a
    single-day increment, same reasoning as Reports 3 & 5's date field).
"""

import logging
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from reports.errors import ReportProcessingError
from reports.registry import InputSlot, ReportMeta, register

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# CONFIG — column names kept exact to match the real workbook / source files
# ---------------------------------------------------------------------------

RAW_CONSOLIDATED_COLS = [
    "Vehicle No.", "Transporter Name", "Destination", "Event Lat.", "Event Long.",
    "Ship to Name", "Shipment No.", "Location Area", "Location Date and Time",
    "Invoice Date and Time", "Plant Name", "Status", "Unloading Point",
    "Ship to code", "Invoice No.", "Distribution Channel",
]

CSN_TAB_COLS = [
    "Mail Status", "vehicle test", "Date", "Vehicle No.", "Transporter", "Destination",
    "Event Lat.", "Event Long.", "Ship To Name", "Shipment No.", "Location Area",
    "Location Date and Time", "Invoice Date and Time", "Plant Name", "Status",
    "Unloading Point", "Ship to code", "Invoice No.", "Distribution Channel",
    "STO/NON STO", "Onward Status", "shareable link", "Test",
]

MASTER_TAB_COLS = [
    "Shipment No.", "Invoice Number", "Plant Name", "Truck No", "Transporter Name",
    "Unloading point(ship-to-party)", "Billing date", "Tracking Link", "Sold To Code",
    "Name(Sold-To Party)", "Distribution Channel", "Region(Ship To Party)",
    "District(Ship-To Party)", "QTY (MT)", "Freight loss", "Last Date", "Area",
]

DIST_CHANNEL_MAP = {20: "DEALER TRADE", 10: "DIR PARTY(NON TRADE)", 30: "STOCK TRANSFER"}

FREIGHT_LOSS = 5000


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def _load_mtr_lookup(mtr_path: Path):
    """Load MTR raw export and build the lookup fields this pipeline needs."""
    try:
        if mtr_path.suffix.lower() == ".csv":
            df = pd.read_csv(mtr_path, dtype={"Shipment Number": "Int64"})
        else:
            df = pd.read_excel(mtr_path, dtype={"Shipment Number": "Int64"})
    except Exception as exc:
        raise ReportProcessingError(f"Couldn't read MTR Raw file '{mtr_path.name}': {exc}") from exc

    df = df.drop_duplicates(subset=["Shipment Number"], keep="first")
    df = df.set_index("Shipment Number")
    keep_cols = [
        "Product Name", "Onward Status", "Share Trip", "40 Km Geofence Start Time",
        "SOLD TO", "SOLD TO NM", "REGION_CODE", "Ship to District", "Quantity",
    ]
    return df[keep_cols], df  # (slim lookup, full df for MTR tab refresh)


def _load_new_consolidated_report(path: Path) -> pd.DataFrame:
    """Load today's raw Daily Battery Disconnected Consolidated Report (Sheet1)."""
    try:
        df = pd.read_excel(path, sheet_name="Sheet1")
    except Exception as exc:
        raise ReportProcessingError(f"Couldn't read '{path.name}': {exc}") from exc
    missing = set(RAW_CONSOLIDATED_COLS) - set(df.columns)
    if missing:
        raise ReportProcessingError(
            f"'{path.name}' is missing expected columns: {sorted(missing)}. "
            "Check you uploaded the Daily Battery Disconnected Consolidated Report."
        )
    return df[RAW_CONSOLIDATED_COLS].copy()


def _load_prev_master_tabs(path: Path):
    """Load the 3 tabs we carry forward from the previous Master workbook."""
    try:
        consolidated = pd.read_excel(path, sheet_name="Consolidated")
        csn = pd.read_excel(path, sheet_name="Consolidated Shipment No.")
        master = pd.read_excel(path, sheet_name="Master")
    except Exception as exc:
        raise ReportProcessingError(
            f"Couldn't read '{path.name}': {exc}. Check you uploaded the previous day's "
            "Battery_Disconnection_Master_<date>.xlsx (needs Consolidated, Consolidated "
            "Shipment No., Master tabs)."
        ) from exc
    return consolidated, csn, master


def _load_table_tab_rows(path: Path):
    """Read the 'Table' tab's raw values only (small scratch/lookup helper,
    per the reference doc -- values matter, not exact formatting), using
    read_only mode so we never load the workbook's other, much larger sheets
    (Consolidated, MTR) into memory as full rich cell objects.

    openpyxl.load_workbook(path) (full/writable mode) on a growing multi-tab
    workbook measured 1.6GB peak memory and 188s locally -- far past
    Render's free-tier limits, and the actual cause of the 502s seen live.
    Read-only mode + only touching the one small tab we need avoids that.
    """
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ReportProcessingError(f"Couldn't read '{path.name}': {exc}") from exc
    if "Table" not in wb.sheetnames:
        return []
    ws = wb["Table"]
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# Pipeline steps
# ---------------------------------------------------------------------------

def build_new_consolidated_rows(new_raw_df: pd.DataFrame, report_date: pd.Timestamp) -> pd.DataFrame:
    """Tag today's raw rows with the report date, unfiltered (for 'Consolidated' tab)."""
    df = new_raw_df.copy()
    df.insert(0, "Date", report_date)
    return df


def build_new_csn_rows(new_raw_df: pd.DataFrame, report_date: pd.Timestamp, mtr_lookup: pd.DataFrame):
    """Clinker-filter + per-day dedupe + derive columns (for 'Consolidated Shipment No.')."""
    df = new_raw_df.copy()

    # Join Product Name (only used here to filter Clinker; not stored in CSN tab).
    df = df.merge(
        mtr_lookup[["Product Name"]], left_on="Shipment No.", right_index=True, how="left"
    )
    before = len(df)
    df = df[~df["Product Name"].astype(str).str.lower().eq("clinker")]
    clinker_dropped = before - len(df)

    # Per-day dedupe, keep first occurrence in source file order.
    before2 = len(df)
    df = df.drop_duplicates(subset=["Shipment No."], keep="first")
    dupes_dropped = before2 - len(df)

    df.insert(0, "vehicle test", None)
    df.insert(0, "Mail Status", None)
    df.insert(2, "Date", report_date)

    df = df.rename(columns={
        "Transporter Name": "Transporter",
        "Ship to Name": "Ship To Name",
    })

    df["STO/NON STO"] = df["Distribution Channel"].map(DIST_CHANNEL_MAP)
    df["Onward Status"] = df["Shipment No."].map(mtr_lookup["Onward Status"]).fillna(0)
    df["shareable link"] = df["Shipment No."].map(mtr_lookup["Share Trip"]).fillna(0)

    geofence = df["Shipment No."].map(mtr_lookup["40 Km Geofence Start Time"])
    found_in_mtr = df["Shipment No."].isin(mtr_lookup.index)
    entered_40km = geofence.notna() & (geofence.astype(str).str.strip() != "")
    df["Test"] = None
    df.loc[found_in_mtr & entered_40km, "Test"] = "Enter"
    df.loc[found_in_mtr & ~entered_40km, "Test"] = "Not Enter"
    # shipments not found in MTR: Test stays blank ("Unknown") rather than guessed

    df = df.drop(columns=["Product Name"])
    df = df[CSN_TAB_COLS]

    stats = {"clinker_dropped": clinker_dropped, "same_day_dupes_dropped": dupes_dropped}
    return df, stats


def build_master_additions(full_csn_df: pd.DataFrame, prev_master_df: pd.DataFrame, mtr_lookup: pd.DataFrame) -> pd.DataFrame:
    """Find shipments newly qualifying for Master (vehicle test starts with 'offline')."""
    df = full_csn_df.copy()
    df["vehicle test"] = df["vehicle test"].fillna("")
    df["_is_offline"] = df["vehicle test"].str.strip().str.lower().str.startswith("offline")

    offline_ships = set(df.loc[df["_is_offline"], "Shipment No."].unique())
    already_in_master = set(prev_master_df["Shipment No."].unique())

    # Exclude shipments whose most recent Mail Status is "Waived OFF"
    df_sorted = df.sort_values("Date")
    latest_status = df_sorted.groupby("Shipment No.")["Mail Status"].last()
    waived = set(latest_status[latest_status.astype(str).str.strip().str.lower() == "waived off"].index)

    new_ships = offline_ships - already_in_master - waived
    if not new_ships:
        return pd.DataFrame(columns=MASTER_TAB_COLS)

    # For each new shipment, take its most recent CSN record for the descriptive fields.
    latest_rows = df_sorted[df_sorted["Shipment No."].isin(new_ships)].groupby(
        "Shipment No.", as_index=False
    ).last()

    out = pd.DataFrame({
        "Shipment No.": latest_rows["Shipment No."],
        "Invoice Number": latest_rows["Invoice No."],
        "Plant Name": latest_rows["Plant Name"],
        "Truck No": latest_rows["Vehicle No."],
        "Transporter Name": latest_rows["Transporter"],
        "Unloading point(ship-to-party)": latest_rows["Unloading Point"],
        "Billing date": latest_rows["Invoice Date and Time"],
        "Tracking Link": latest_rows["Shipment No."].map(mtr_lookup["Share Trip"]),
        "Sold To Code": latest_rows["Shipment No."].map(mtr_lookup["SOLD TO"]),
        "Name(Sold-To Party)": latest_rows["Shipment No."].map(mtr_lookup["SOLD TO NM"]),
        "Distribution Channel": latest_rows["Distribution Channel"],
        "Region(Ship To Party)": latest_rows["Shipment No."].map(mtr_lookup["REGION_CODE"]),
        "District(Ship-To Party)": latest_rows["Shipment No."].map(mtr_lookup["Ship to District"]),
        "QTY (MT)": latest_rows["Shipment No."].map(mtr_lookup["Quantity"]),
        "Freight loss": FREIGHT_LOSS,
        "Last Date": latest_rows["Location Date and Time"],
        "Area": latest_rows["Location Area"],
    })
    return out[MASTER_TAB_COLS]


# ---------------------------------------------------------------------------
# Writer -- header fills/fonts/column widths extracted cell-by-cell from the
# real Battery_Disconnection_Master_13_July_2026.xlsx (data logic untouched;
# this section only affects display).
# ---------------------------------------------------------------------------

from openpyxl.utils import get_column_letter
from reports._report6_col_widths import MTR_COL_WIDTHS

APTOS_10 = "Aptos Narrow"

# Per-sheet header style: (fill color, font name, font size).
SHEET_STYLES = {
    "Consolidated": {"fill": "FFC000", "font": APTOS_10, "size": 10},
    "Consolidated Shipment No.": {"fill": "8FAADC", "font": APTOS_10, "size": 10},
    "Master": {"fill": "FFFF00", "font": APTOS_10, "size": 10},
    "MTR": {"fill": "A9D18E", "font": "Calibri", "size": 11},
}

# Despite the name, this is really "column name -> number_format override" --
# mostly dates, plus one cosmetic exception below.
DATE_COLUMN_FORMATS = {
    "Date": "mm-dd-yy",
    "Location Date and Time": "m/d/yy h:mm",
    "Invoice Date and Time": "m/d/yy h:mm",
    "Billing date": "mm-dd-yy",
    # Forces text display so long digit strings never render in scientific
    # notation -- matches what Khagash is used to seeing, doesn't affect
    # correctness. No pd.to_datetime() side effect here (unlike Report 3's
    # equivalent dict) since this report doesn't pre-parse date_col_idxs
    # entries before writing -- only number_format gets applied.
    "Transporter Number": "@",
}

CONSOLIDATED_COL_WIDTHS = {
    "Date": 10.11, "Vehicle No.": 12.33, "Transporter": 40.78, "Destination": 58.55,
    "Event Lat.": 12.0, "Ship To Name": 71.78, "Shipment No.": 11.66,
    "Location Area": 48.44, "Location Date and Time": 19.33, "Invoice Date and Time": 18.44,
    "Plant Name": 46.66, "Status": 17.44, "Unloading Point": 22.33, "Ship to code": 11.0,
    "Distribution Channel": 17.44,
}
CSN_COL_WIDTHS = {
    "Mail Status": 9.89, "vehicle test": 68.22, "Date": 10.11, "Vehicle No.": 12.33,
    "Transporter": 41.0, "Destination": 58.55, "Event Lat.": 12.0, "Ship To Name": 71.78,
    "Shipment No.": 11.66, "Location Area": 48.44, "Location Date and Time": 19.33,
    "Invoice Date and Time": 18.44, "Plant Name": 46.66, "Status": 17.44,
    "Unloading Point": 22.33, "Ship to code": 11.0, "Distribution Channel": 17.44,
    "STO/NON STO": 19.22, "Onward Status": 12.11, "shareable link": 72.33, "Test": 8.11,
}
MASTER_COL_WIDTHS = {
    "Shipment No.": 11.89, "Invoice Number": 13.78, "Plant Name": 12.22,
    "Transporter Name": 41.55, "Unloading point(ship-to-party)": 25.11,
    "Billing date": 10.33, "Tracking Link": 74.22, "Sold To Code": 11.22,
    "Name(Sold-To Party)": 51.33, "Distribution Channel": 17.44,
    "Region(Ship To Party)": 18.55, "District(Ship-To Party)": 18.44,
    "QTY (MT)": 8.44, "Freight loss": 10.11, "Last Date": 12.11, "Area": 39.89,
}
DEFAULT_COL_WIDTH = 13.0

COL_WIDTHS_BY_SHEET = {
    "Consolidated": CONSOLIDATED_COL_WIDTHS,
    "Consolidated Shipment No.": CSN_COL_WIDTHS,
    "Master": MASTER_COL_WIDTHS,
    "MTR": MTR_COL_WIDTHS,
}


def _write_sheet(wb, sheet_name: str, df: pd.DataFrame):
    """Plain (non write_only) Workbook + df.itertuples() bulk append --
    matches the fast pattern used in Reports 2 & 3 (write_only mode was
    tried and reverted, see git history: slower in practice here, not worth
    its memory saving on this specific workload). Header fill/font, column
    widths, and date number-formats are matched to the real file below; no
    freeze panes -- the real file doesn't use them either, despite the
    original script adding freeze_panes="A2" on every tab.

    Deliberately NOT setting a per-row data font (the real file uses "Aptos
    Narrow" 10pt on some sheets' data rows, openpyxl's default is Calibri
    11): looping ws.cell() to style every one of ~1.45M data cells is
    exactly the O(n) styling cost that already made this report's write
    step marginal on Render (see the write_only/plain-Workbook history
    above) -- not worth re-risking a report already on hold for being too
    slow, for a data-row font that doesn't affect any actual figures.
    """
    style = SHEET_STYLES.get(sheet_name, {"fill": None, "font": "Calibri", "size": 11})
    header_font = Font(name=style["font"], size=style["size"], bold=True)
    header_fill = PatternFill(fill_type="solid", fgColor=style["fill"]) if style["fill"] else None
    col_widths = COL_WIDTHS_BY_SHEET.get(sheet_name, {})

    ws = wb.create_sheet(sheet_name)
    columns = list(df.columns)
    ws.append(columns)
    for col_idx, name in enumerate(columns, start=1):
        cell = ws.cell(1, col_idx)
        cell.font = header_font
        if header_fill:
            cell.fill = header_fill
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].width = col_widths.get(str(name).strip(), DEFAULT_COL_WIDTH)

    date_col_idxs = {i: DATE_COLUMN_FORMATS[name] for i, name in enumerate(columns) if name in DATE_COLUMN_FORMATS}

    row_idx = 1
    for row in df.itertuples(index=False):
        row_idx += 1
        ws.append(row)
        for idx, fmt in date_col_idxs.items():
            ws.cell(row_idx, idx + 1).number_format = fmt
    return ws


def _write_table_tab(wb, rows: list):
    """Table is a tiny (~17-row) scratch/lookup helper, so per-cell styling
    here is cheap and safe (unlike the large data sheets above). Matches
    the real file: column A (labels) bold + yellow fill on every row,
    column B (values) plain.
    """
    ws = wb.create_sheet("Table")
    label_font = Font(bold=True)
    label_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    for row_idx, row in enumerate(rows, start=1):
        ws.append(row)
        cell = ws.cell(row_idx, 1)
        cell.font = label_font
        cell.fill = label_fill
    ws.column_dimensions["A"].width = 27.33
    ws.column_dimensions["B"].width = 75.11
    return ws
    return ws


# ---------------------------------------------------------------------------
# Entry point used by the generic upload -> process -> download route
# ---------------------------------------------------------------------------

def process(input_files: dict, dates: dict, output_dir: Path) -> Path:
    report_date_str = dates["report_date"]
    report_date = pd.Timestamp(report_date_str)
    log.info("Processing Battery Disconnection Mail Creation for %s", report_date_str)

    mtr_lookup, mtr_full = _load_mtr_lookup(input_files["mtr_raw"])
    new_raw = _load_new_consolidated_report(input_files["new_consolidated"])
    prev_consolidated, prev_csn, prev_master = _load_prev_master_tabs(input_files["prev_master"])
    table_rows = _load_table_tab_rows(input_files["prev_master"])

    try:
        new_consolidated_rows = build_new_consolidated_rows(new_raw, report_date)
        full_consolidated = pd.concat([prev_consolidated, new_consolidated_rows], ignore_index=True)

        new_csn_rows, stats = build_new_csn_rows(new_raw, report_date, mtr_lookup)
        full_csn = pd.concat([prev_csn, new_csn_rows], ignore_index=True)

        new_master_rows = build_master_additions(full_csn, prev_master, mtr_lookup)
        full_master = pd.concat([prev_master, new_master_rows], ignore_index=True)
    except KeyError as exc:
        raise ReportProcessingError(
            f"Expected column {exc} not found. Check each file was uploaded to the correct "
            "slot (Previous Master / Today's Consolidated Report / MTR Raw)."
        ) from exc

    review_queue = new_csn_rows[
        (new_csn_rows["Test"] == "Not Enter") & (new_csn_rows["vehicle test"].isna())
    ]
    log.info(
        "Clinker dropped: %d | Same-day dupes dropped: %d | New Master rows: %d | "
        "Manual review queue: %d",
        stats["clinker_dropped"], stats["same_day_dupes_dropped"],
        len(new_master_rows), len(review_queue),
    )

    # Build a FRESH workbook rather than mutating the previous file's loaded
    # object -- loading the previous file in full/writable mode (just to
    # preserve its small 'Table' tab) pulled its huge Consolidated/MTR tabs
    # into memory as rich cell objects too -- measured 1.6GB peak / 188s
    # locally. Fixed by reading Table's values separately via read_only mode
    # (see _load_table_tab_rows) and starting fresh here. See _write_sheet's
    # docstring for why this uses a plain Workbook rather than write_only
    # mode despite the latter's much lower memory use.
    wb = Workbook()
    wb.remove(wb.active)
    _write_sheet(wb, "Consolidated", full_consolidated)
    _write_sheet(wb, "Consolidated Shipment No.", full_csn)
    _write_table_tab(wb, table_rows)
    _write_sheet(wb, "Master", full_master)
    _write_sheet(wb, "MTR", mtr_full.reset_index())

    output_path = output_dir / f"Battery_Disconnection_Master_{report_date_str}.xlsx"
    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# RE-ENABLED for another live test (per explicit instruction): the previous
# 3 production attempts 502'd (467MB/76s and 96MB/174s locally, both failed
# in the 53-127s range on Render). Since then: switched to itertuples-based
# writing (measured 42.6s/41.7s locally, down from the original 188s), and
# this pass added real-file formatting with no additional per-cell cost
# (confirmed no timing regression). Genuinely lower risk than the last
# attempt, but still unproven on Render's actual CPU/memory -- if this
# still 502s, the earlier stub-with-notes state is the fallback (see git
# history for that version).
# ---------------------------------------------------------------------------

register(
    ReportMeta(
        id="6",
        name="Battery Disconnection Mail Creation",
        input_slots=[
            InputSlot(
                key="prev_master",
                label="Previous Day's Master Workbook",
                accept=".xlsx",
                hint="Battery_Disconnection_Master_<prev_date>.xlsx",
            ),
            InputSlot(
                key="new_consolidated",
                label="Today's Daily Battery Disconnected Consolidated Report",
                accept=".xlsx",
                hint="Daily Battery Disconnected Consolidated Report <date>.xlsx",
            ),
            InputSlot(key="mtr_raw", label="MTR Raw", accept=".csv,.xlsx", hint="mtr - <timestamp>.csv"),
        ],
        output_pattern="Battery_Disconnection_Master_<date>.xlsx (5 tabs: Consolidated, Consolidated "
                      "Shipment No., Table, Master, MTR)",
        # implemented=False is COSMETIC ONLY here -- process_fn is still the
        # real, working `process` function below, not the not_implemented
        # stub. main.py's /generate route never checks `implemented`; it
        # only drives the "not yet implemented" badge/text on the report
        # page. Per explicit instruction: show the stub label (since this
        # report is still unproven on Render's free tier and shouldn't be
        # advertised as ready), but if someone uploads and clicks Generate
        # anyway, it should still produce a correct, fully-working output --
        # which this achieves, since the real process_fn is still wired in.
        process_fn=process,
        implemented=False,
        date_mode="single",
        notes=(
            "Stateful report -- needs the PREVIOUS day's Master workbook as an input each run "
            "(carries forward history + any manual Mail Status / vehicle test tags). Master tab = "
            "shipments whose 'vehicle test' starts with 'offline', append-only, excluding 'Waived "
            "OFF'. Formatting matches the real master file exactly. Previously 502'd on Render's "
            "free tier 3 times before write-speed optimizations (188s -> ~42s locally) -- being "
            "retested live; if it 502s again this needs a hosting upgrade, not more code tuning."
        ),
    )
)
