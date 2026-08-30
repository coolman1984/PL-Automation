"""Stable V1 constants and version/period derivation."""

from __future__ import annotations

from .errors import UnsupportedScopeError
from .models import RunCodes

TARGET_SHEETS = ("VD Total", "MX Total", "DA Total")
TOTAL_SHEET = "Total PL"
HEADER_FIRST_ROW = 1
HEADER_LAST_ROW = 30
XLSB_EXTENSION = ".xlsb"

# Excel COM constants used when generated constants are unavailable.
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_SEMIAUTOMATIC = 2
XL_CALCULATION_DONE = 0
XL_CELL_TYPE_FORMULAS = -4123
XL_CELL_TYPE_CONSTANTS = 2
XL_TO_RIGHT = -4161
XL_FORMAT_FROM_LEFT_OR_ABOVE = 0
XL_ERRORS = 16
XL_DATABASE_SOURCE_TYPE = 1  # xlDatabase: a worksheet range or Excel Table source
XL_EXCEL12 = 50
# Office MsoAutomationSecurityForceDisable.  Set only on automation-owned
# isolated Excel instances, never on the user's attached instance.
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


def _validate_month(month: int) -> None:
    if not isinstance(month, int) or isinstance(month, bool) or not 1 <= month <= 12:
        raise UnsupportedScopeError(f"Month must be an integer from 1 through 12; received {month!r}")


def format_period(year: int, month: int) -> str:
    _validate_month(month)
    if not isinstance(year, int) or isinstance(year, bool) or year < 1900 or year > 9999:
        raise UnsupportedScopeError(f"Year must be a four-digit integer; received {year!r}")
    return f"{year:04d}.{month:03d}"


def _period_suffix(month: int) -> str:
    _validate_month(month)
    return f"{month:02d}" if month <= 9 else {10: "0A", 11: "0B", 12: "0C"}[month]


def target_version_for_month(month: int) -> str:
    return f"T{_period_suffix(month)}"


def forecast_version_for_month(month: int) -> str:
    return f"S{_period_suffix(month)}"


def actual_version_for_month(month: int) -> str:
    _validate_month(month)
    if month > 9:
        raise UnsupportedScopeError(
            "Actual version naming for October through December is not defined in Phase 1"
        )
    return f"A{month:02d}"


def resolve_run_codes(year: int, month: int, *, execution: bool) -> RunCodes:
    _validate_month(month)
    if execution and month != 8:
        raise UnsupportedScopeError(
            f"Phase 1 execution supports August only; received month {month}"
        )
    return RunCodes(
        year=year,
        month=month,
        month_name=MONTH_NAMES[month - 1],
        period=format_period(year, month),
        target_version=target_version_for_month(month),
        forecast_version=forecast_version_for_month(month),
        actual_version=actual_version_for_month(month),
    )
