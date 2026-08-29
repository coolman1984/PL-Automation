"""Strict YAML/CLI configuration loading."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any

import yaml

from .constants import TARGET_SHEETS, TOTAL_SHEET, resolve_run_codes
from .errors import ConfigurationError
from .models import AppConfig, SafetyConfig, ValidationConfig


_TOP_LEVEL_KEYS = {"workbook", "run", "safety", "validation"}
_WORKBOOK_KEYS = {"target_sheets", "total_sheet"}
_RUN_KEYS = {"year", "month"}
_SAFETY_KEYS = {
    "update_external_links",
    "refresh_pivots",
    "overwrite_existing_actual",
    "keep_failed_workbook",
    "reopen_after_save",
}
_VALIDATION_KEYS = {
    "numeric_tolerance",
    "require_all_sheets",
    "require_unique_header_match",
    "calculation_timeout_seconds",
}


def load_yaml(path: Path) -> dict[str, object]:
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
    except FileNotFoundError as exc:
        raise ConfigurationError(f"Configuration file was not found: {path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Configuration YAML is invalid: {exc}") from exc
    if not isinstance(data, dict):
        raise ConfigurationError("Configuration root must be a mapping")
    return data


def _mapping(value: Any, name: str) -> Mapping[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise ConfigurationError(f"Configuration section {name!r} must be a mapping")
    return value


def _reject_unknown(mapping: Mapping[str, Any], allowed: set[str], name: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {name}: {', '.join(unknown)}")


def build_config(
    raw: Mapping[str, object], *, year: int | None, month: int | None, execution: bool
) -> AppConfig:
    _reject_unknown(raw, _TOP_LEVEL_KEYS, "root")
    workbook = _mapping(raw.get("workbook"), "workbook")
    run = _mapping(raw.get("run"), "run")
    safety = _mapping(raw.get("safety"), "safety")
    validation = _mapping(raw.get("validation"), "validation")
    _reject_unknown(workbook, _WORKBOOK_KEYS, "workbook")
    _reject_unknown(run, _RUN_KEYS, "run")
    _reject_unknown(safety, _SAFETY_KEYS, "safety")
    _reject_unknown(validation, _VALIDATION_KEYS, "validation")

    resolved_year = int(year if year is not None else run.get("year", 2026))
    resolved_month = int(month if month is not None else run.get("month", 8))
    codes = resolve_run_codes(resolved_year, resolved_month, execution=execution)

    target_sheets = tuple(workbook.get("target_sheets", TARGET_SHEETS))
    total_sheet = str(workbook.get("total_sheet", TOTAL_SHEET))
    if target_sheets != TARGET_SHEETS or total_sheet != TOTAL_SHEET:
        raise ConfigurationError(
            f"Phase 1 scope requires target sheets {TARGET_SHEETS!r} and total sheet {TOTAL_SHEET!r}"
        )

    safety_config = SafetyConfig(
        update_external_links=bool(safety.get("update_external_links", False)),
        refresh_pivots=bool(safety.get("refresh_pivots", False)),
        overwrite_existing_actual=bool(safety.get("overwrite_existing_actual", False)),
        keep_failed_workbook=bool(safety.get("keep_failed_workbook", True)),
        reopen_after_save=bool(safety.get("reopen_after_save", True)),
    )
    validation_config = ValidationConfig(
        numeric_tolerance=float(validation.get("numeric_tolerance", 0.01)),
        require_all_sheets=bool(validation.get("require_all_sheets", True)),
        require_unique_header_match=bool(validation.get("require_unique_header_match", True)),
        calculation_timeout_seconds=int(validation.get("calculation_timeout_seconds", 1800)),
    )
    config = AppConfig(
        target_sheets=target_sheets,
        total_sheet=total_sheet,
        codes=codes,
        safety=safety_config,
        validation=validation_config,
    )
    validate_config(config)
    return config


def validate_config(config: AppConfig) -> None:
    if config.safety.update_external_links:
        raise ConfigurationError("External-link updates must remain disabled in Phase 1")
    if config.safety.refresh_pivots:
        raise ConfigurationError("Pivot refresh must remain disabled in Phase 1")
    if config.safety.overwrite_existing_actual:
        raise ConfigurationError("Replacing an existing A08 pair is not implemented in Phase 1")
    if not config.safety.reopen_after_save:
        raise ConfigurationError("Post-save reopen validation is mandatory")
    if config.validation.numeric_tolerance < 0:
        raise ConfigurationError("Numeric tolerance cannot be negative")
    if config.validation.calculation_timeout_seconds <= 0:
        raise ConfigurationError("Calculation timeout must be positive")


def load_config(
    path: Path, *, year: int | None, month: int | None, execution: bool
) -> AppConfig:
    return build_config(load_yaml(path), year=year, month=month, execution=execution)

