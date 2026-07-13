"""Importing this package registers all 9 reports into `registry.REPORTS`."""

from reports import registry  # noqa: F401

from reports import report_1_daily_tracker  # noqa: F401
from reports import report_2_live_detention  # noqa: F401
from reports import report_3_daily_tracking  # noqa: F401
from reports import report_4_battery_disconnected  # noqa: F401
from reports import report_5_at_fix_ontrip  # noqa: F401
from reports import report_6_battery_disconnection_mail  # noqa: F401
from reports import report_7_control_tower_tracker  # noqa: F401
from reports import report_8_freight_deviation_mail  # noqa: F401
from reports import report_9_durg_installation  # noqa: F401
