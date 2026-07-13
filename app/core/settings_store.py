"""Editable report metadata (name, notes, output pattern, enabled, order,
per-slot label/hint) persisted to JSON and overlaid on top of each report
module's code-defined defaults at request time.

Deliberately NOT editable here: report id, input slot keys, accepted file
extensions, process_fn wiring, date_mode — those are load-bearing for the
actual processing logic and can only change in code.
"""

import json
from dataclasses import asdict, dataclass, field, replace
from pathlib import Path
from typing import Dict, List

from reports.registry import REPORTS, ReportMeta

SETTINGS_PATH = Path(__file__).resolve().parent.parent / "config" / "report_settings.json"


@dataclass
class SlotOverride:
    label: str = ""
    hint: str = ""


@dataclass
class ReportOverride:
    name: str = ""
    notes: str = ""
    output_pattern: str = ""
    enabled: bool = True
    order: int = 0
    slots: Dict[str, SlotOverride] = field(default_factory=dict)


def _default_overrides() -> Dict[str, ReportOverride]:
    """Seed overrides from each report's code defaults, in registration order."""
    overrides = {}
    for i, meta in enumerate(sorted(REPORTS.values(), key=lambda m: int(m.id)), start=1):
        overrides[meta.id] = ReportOverride(
            name=meta.name,
            notes=meta.notes,
            output_pattern=meta.output_pattern,
            enabled=True,
            order=i,
            slots={slot.key: SlotOverride(label=slot.label, hint=slot.hint) for slot in meta.input_slots},
        )
    return overrides


def load_overrides() -> Dict[str, ReportOverride]:
    if not SETTINGS_PATH.exists():
        return _default_overrides()

    raw = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    overrides = {}
    for report_id, data in raw.items():
        slots = {key: SlotOverride(**slot_data) for key, slot_data in data.get("slots", {}).items()}
        overrides[report_id] = ReportOverride(
            name=data.get("name", ""),
            notes=data.get("notes", ""),
            output_pattern=data.get("output_pattern", ""),
            enabled=data.get("enabled", True),
            order=data.get("order", 0),
            slots=slots,
        )

    # Any report registered in code but missing from the saved file yet
    # (e.g. a newly added report module) still needs sane defaults.
    defaults = _default_overrides()
    for report_id, default_override in defaults.items():
        if report_id not in overrides:
            overrides[report_id] = default_override

    return overrides


def save_overrides(overrides: Dict[str, ReportOverride]) -> None:
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    raw = {
        report_id: {
            "name": o.name,
            "notes": o.notes,
            "output_pattern": o.output_pattern,
            "enabled": o.enabled,
            "order": o.order,
            "slots": {key: asdict(slot) for key, slot in o.slots.items()},
        }
        for report_id, o in overrides.items()
    }
    SETTINGS_PATH.write_text(json.dumps(raw, indent=2), encoding="utf-8")


def get_effective_reports(enabled_only: bool = False) -> List[ReportMeta]:
    """Merge each report's code-defined ReportMeta with its saved overrides."""
    overrides = load_overrides()  # guaranteed to have an entry for every id in REPORTS
    effective = []
    for meta in REPORTS.values():
        o = overrides[meta.id]
        if enabled_only and not o.enabled:
            continue
        new_slots = [
            replace(slot, label=o.slots[slot.key].label, hint=o.slots[slot.key].hint)
            if slot.key in o.slots
            else slot
            for slot in meta.input_slots
        ]
        merged = replace(
            meta,
            name=o.name or meta.name,
            notes=o.notes if o.notes else meta.notes,
            output_pattern=o.output_pattern or meta.output_pattern,
            input_slots=new_slots,
        )
        merged._order = o.order  # type: ignore[attr-defined]
        merged._enabled = o.enabled  # type: ignore[attr-defined]
        effective.append(merged)

    effective.sort(key=lambda m: getattr(m, "_order", int(m.id)))
    return effective
