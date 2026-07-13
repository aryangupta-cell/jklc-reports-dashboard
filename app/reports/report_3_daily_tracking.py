from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="3",
        name="JKLC Daily Tracking Report",
        input_slots=[
            InputSlot(key="mtr_raw", label="MTR Raw", accept=".csv,.xlsx", hint="raw_mtr_-_<timestamp>.csv"),
            InputSlot(
                key="at_dashboard",
                label="AT Dashboard / Device Status Export",
                accept=".csv,.xlsx",
                hint="device_status_<date>.xlsx",
            ),
            InputSlot(
                key="master_entry_durg",
                label="Master Entry Data (Durg)",
                accept=".csv,.xlsx",
                hint="master_entry_durg_<date>.xlsx",
            ),
        ],
        output_pattern="JKLC_Daily_Tracking_Report_<date>.xlsx",
        process_fn=not_implemented_process("JKLC Daily Tracking Report"),
        implemented=False,
    )
)
