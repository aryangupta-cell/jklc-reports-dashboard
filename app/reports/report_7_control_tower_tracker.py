from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="7",
        name="Control Tower Tracker",
        input_slots=[
            InputSlot(
                key="mtr_raw_yesterday",
                label="MTR Raw (yesterday)",
                accept=".csv,.xlsx",
                hint="raw_mtr_-_<timestamp>.csv",
            ),
            InputSlot(
                key="dump_code_list",
                label="Client Dump-Code List",
                accept=".csv,.xlsx",
                hint="dump_code_list_<date>.xlsx",
            ),
            InputSlot(
                key="allowance_40km",
                label="40km Allowance Report",
                accept=".csv,.xlsx",
                hint="40km_allowance_<date>.xlsx",
            ),
        ],
        output_pattern="Control_Tower_Tracker_<date>.xlsx",
        process_fn=not_implemented_process("Control Tower Tracker"),
        implemented=False,
    )
)
