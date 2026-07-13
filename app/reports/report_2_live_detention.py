from reports._stub_helpers import not_implemented_process
from reports.registry import InputSlot, ReportMeta, register

register(
    ReportMeta(
        id="2",
        name="JKLC Live Detention",
        input_slots=[
            InputSlot(key="mtr_raw", label="MTR Raw", accept=".csv,.xlsx", hint="raw_mtr_-_<timestamp>.csv"),
            InputSlot(
                key="detention_bot",
                label="Detention Bot Output",
                accept=".csv,.xlsx",
                hint="detention_bot_output_<date>.csv",
            ),
        ],
        output_pattern="JKLC_Live_Detention_Master_<date>.xlsx",
        process_fn=not_implemented_process("JKLC Live Detention"),
        implemented=False,
        notes="Plant-wise: Durg, Jharli, Surat, Cuttack. Output tabs: MTR, 20KM, Summary, Dispatch, Mail ID.",
    )
)
