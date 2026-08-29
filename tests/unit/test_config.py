"""Unit tests for strict configuration loading and Phase 1 scope locks."""

from __future__ import annotations

import pytest

from src.config import build_config
from src.errors import ConfigurationError, UnsupportedScopeError
from src.constants import (
    actual_version_for_month,
    forecast_version_for_month,
    format_period,
    resolve_run_codes,
    target_version_for_month,
)


def test_defaults_match_phase_one_scope():
    config = build_config({}, year=2026, month=8, execution=True)
    assert config.target_sheets == ("VD Total", "MX Total", "DA Total")
    assert config.total_sheet == "Total PL"
    assert config.codes.period == "2026.008"
    assert config.codes.target_version == "T08"
    assert config.codes.forecast_version == "S08"
    assert config.codes.actual_version == "A08"
    assert config.safety.update_external_links is False
    assert config.safety.refresh_pivots is False
    assert config.validation.numeric_tolerance == 0.01
    assert config.validation.calculation_timeout_seconds == 1800


def test_unknown_top_level_key_rejected():
    with pytest.raises(ConfigurationError):
        build_config({"mystery": 1}, year=2026, month=8, execution=False)


def test_unknown_nested_key_rejected():
    with pytest.raises(ConfigurationError):
        build_config({"safety": {"delete_source": True}}, year=2026, month=8, execution=False)


def test_external_link_updates_cannot_be_enabled():
    with pytest.raises(ConfigurationError):
        build_config(
            {"safety": {"update_external_links": True}}, year=2026, month=8, execution=True
        )


def test_pivot_refresh_cannot_be_enabled():
    with pytest.raises(ConfigurationError):
        build_config({"safety": {"refresh_pivots": True}}, year=2026, month=8, execution=True)


def test_overwrite_existing_actual_rejected():
    with pytest.raises(ConfigurationError):
        build_config(
            {"safety": {"overwrite_existing_actual": True}},
            year=2026,
            month=8,
            execution=True,
        )


def test_reopen_validation_is_mandatory():
    with pytest.raises(ConfigurationError):
        build_config(
            {"safety": {"reopen_after_save": False}}, year=2026, month=8, execution=True
        )


def test_negative_tolerance_rejected():
    with pytest.raises(ConfigurationError):
        build_config(
            {"validation": {"numeric_tolerance": -0.5}}, year=2026, month=8, execution=True
        )


def test_execution_month_locked_to_august():
    with pytest.raises(UnsupportedScopeError):
        resolve_run_codes(2026, 9, execution=True)
    codes = resolve_run_codes(2026, 9, execution=False)
    assert codes.period == "2026.009"


def test_version_suffixes_follow_spreadsheet_convention():
    assert target_version_for_month(8) == "T08"
    assert target_version_for_month(10) == "T0A"
    assert forecast_version_for_month(12) == "S0C"
    assert actual_version_for_month(8) == "A08"
    with pytest.raises(UnsupportedScopeError):
        actual_version_for_month(10)


def test_period_format_is_zero_padded():
    assert format_period(2026, 8) == "2026.008"
    assert format_period(2026, 11) == "2026.011"
    with pytest.raises(UnsupportedScopeError):
        format_period(2026, 13)
    with pytest.raises(UnsupportedScopeError):
        format_period(2026, 0)
