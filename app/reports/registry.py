from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Literal

# "range"  -> report page shows Start Date + End Date (blank, no auto-fill)
# "single" -> report page shows one Report Date
# "none"   -> report needs no date input at all
DateMode = Literal["range", "single", "none"]


@dataclass
class InputSlot:
    key: str  # stable identifier process_fn reads input_files[key] by — NOT editable via Settings
    label: str  # UI-editable default (overridden by app/config/report_settings.json if set)
    accept: str  # comma-separated extensions, e.g. ".csv,.xlsx" — technical, not exposed in Settings
    hint: str = ""  # UI-editable default example filename, for reference only — not used for matching


@dataclass
class ExtraNumberField:
    """An optional extra number input on a report's page, alongside its date
    field(s) — e.g. Report 3's "Days back" control. Rendered generically by
    report.html; process_fn reads it from the `dates` dict by `key`, same
    dict the date fields land in. Not file-based, so it's separate from
    InputSlot; not date-typed, so it's separate from date_mode."""
    key: str  # stable identifier — process_fn reads dates[key] by this
    label: str
    default: int
    min_value: int = 1
    hint: str = ""  # short explanatory text shown under the field


@dataclass
class ReportMeta:
    id: str  # stable identifier, matches process_fn wiring — NOT editable via Settings
    name: str  # UI-editable default
    input_slots: List[InputSlot]
    output_pattern: str  # UI-editable default
    process_fn: Callable[[Dict[str, Path], dict, Path], Path]  # (input_files, dates, output_dir) -> output_path
    implemented: bool = True  # code-level: does process_fn do real work yet?
    notes: str = ""  # UI-editable default
    date_mode: DateMode = "range"
    extra_number_fields: List[ExtraNumberField] = field(default_factory=list)


REPORTS: Dict[str, ReportMeta] = {}


def register(meta: ReportMeta) -> None:
    REPORTS[meta.id] = meta
