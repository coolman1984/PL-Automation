import pytest

from src.constants import (
    actual_version_for_month,
    format_period,
    forecast_version_for_month,
    resolve_run_codes,
    target_version_for_month,
)
from src.errors import UnsupportedScopeError


def test_august_codes_are_resolved_from_year_and_month():
    codes = resolve_run_codes(2026, 8, execution=True)

    assert codes.month_name == "August"
    assert codes.period == "2026.008"
    assert codes.target_version == "T08"
    assert codes.forecast_version == "S08"
    assert codes.actual_version == "A08"


def test_period_formatting_is_three_digit_month():
    assert format_period(2026, 1) == "2026.001"
    assert format_period(2026, 9) == "2026.009"


def test_target_and_forecast_codes_support_defined_months():
    assert target_version_for_month(10) == "T0A"
    assert target_version_for_month(12) == "T0C"
    assert forecast_version_for_month(11) == "S0B"


def test_actual_codes_are_not_invented_for_q4():
    with pytest.raises(UnsupportedScopeError):
        actual_version_for_month(10)


def test_execute_mode_is_august_only_in_phase_one():
    with pytest.raises(UnsupportedScopeError):
        resolve_run_codes(2026, 7, execution=True)

