from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="5",
        name="JKLC AT Fix On-Trip Vehicle Status",
        input_slots=[
            InputSlot(
                key="daily_tracking_output",
                label="Daily Tracking Report Output",
                accept=".xlsx",
                hint="JKLC_Daily_Tracking_Report_<date>.xlsx",
            ),
            InputSlot(
                key="whatsapp_remarks",
                label="WhatsApp Technician Remarks",
                accept=".csv,.xlsx",
                hint="technician_remarks_<date>.xlsx",
            ),
        ],
        output_pattern="JKLC_AT_Fix_Ontrip_Vehicles_Status_<date>.xlsx",
        process_fn=not_implemented_process("JKLC AT Fix On-Trip Vehicle Status"),
        implemented=False,
    )
)
