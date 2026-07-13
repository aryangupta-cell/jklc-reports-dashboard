from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="9",
        name="Durg Daily Installation Tracker",
        input_slots=[
            InputSlot(
                key="whatsapp_ops_export",
                label="WhatsApp Ops Team Export",
                accept=".csv,.xlsx",
                hint="ops_team_export_<date>.xlsx",
            ),
        ],
        output_pattern="JKLC_Durg_Daily_Installation_<date>.xlsx",
        process_fn=not_implemented_process("Durg Daily Installation Tracker"),
        implemented=False,
    )
)
