from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="4",
        name="Daily Battery Disconnected Consolidated Report",
        input_slots=[
            InputSlot(
                key="mtr_raw_yesterday",
                label="MTR Raw (yesterday)",
                accept=".csv,.xlsx",
                hint="raw_mtr_-_<timestamp>.csv",
            ),
        ],
        output_pattern="Daily_Battery_Disconnected_Consolidated_Report_<date>.xlsx",
        process_fn=not_implemented_process("Daily Battery Disconnected Consolidated Report"),
        implemented=False,
    )
)
