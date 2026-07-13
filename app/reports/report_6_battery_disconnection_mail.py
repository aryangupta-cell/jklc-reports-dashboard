from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="6",
        name="Battery Disconnection Mail Creation",
        input_slots=[
            InputSlot(key="mtr_raw", label="MTR Raw", accept=".csv,.xlsx", hint="raw_mtr_-_<timestamp>.csv"),
            InputSlot(
                key="route_tracker_check",
                label="Route Tracker Check",
                accept=".csv,.xlsx",
                hint="route_tracker_check_<date>.xlsx",
            ),
        ],
        output_pattern="battery_disconnection_mail_drafts_<date>.eml",
        process_fn=not_implemented_process("Battery Disconnection Mail Creation"),
        implemented=False,
        notes="Output is individual mail draft(s) per shipment, not a single spreadsheet.",
    )
)
