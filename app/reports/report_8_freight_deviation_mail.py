from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="8",
        name="Freight Deviation Mail",
        input_slots=[
            InputSlot(
                key="control_tower_output",
                label="Control Tower Tracker Output",
                accept=".xlsx",
                hint="Control_Tower_Tracker_<date>.xlsx",
            ),
            InputSlot(
                key="client_exception_list",
                label="Client Exception List",
                accept=".csv,.xlsx",
                hint="client_exception_list_<date>.xlsx",
            ),
        ],
        output_pattern="freight_deviation_mail_drafts_<date>.eml",
        process_fn=not_implemented_process("Freight Deviation Mail"),
        implemented=False,
        notes="Output is mail draft(s), a filtered subset of Report 7.",
    )
)
