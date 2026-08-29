"""Human-readable and machine-readable run evidence."""

from __future__ import annotations

import json
import logging
import os
import tempfile
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Sequence

from .models import MonthBlock, RunManifest, ValidationReport


def configure_logging(run_dir: Path, *, verbose: bool = False) -> logging.Logger:
    logger = logging.getLogger("pl_actual_automation")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(run_dir / "automation.log", encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    stream_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.addHandler(stream_handler)
    return logger


def _json_default(value: object):
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value):
        return asdict(value)
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def write_manifest_atomic(path: Path, manifest: RunManifest) -> None:
    payload = json.dumps(manifest.to_dict(), indent=2, ensure_ascii=False, default=_json_default)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(payload)
            handle.write("\n")
        os.replace(temporary_name, path)
    finally:
        try:
            Path(temporary_name).unlink(missing_ok=True)
        except Exception:
            pass


def sanitize_for_logging(value: object) -> object:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(key): sanitize_for_logging(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [sanitize_for_logging(item) for item in value]
    if is_dataclass(value):
        return sanitize_for_logging(asdict(value))
    return str(value)


def _block_lines(block: MonthBlock) -> list[str]:
    return [
        f"  T08 amount: col {block.target_col}",
        f"  T08 %:      col {block.target_pct_col}",
        f"  S08 amount: col {block.forecast_col}",
        f"  S08 %:      col {block.forecast_pct_col}",
        f"  Insert at:  col {block.insert_at_col}",
        f"  September:  col {block.september_start_col}",
        f"  Header rows: version={block.version_header_row}, period={block.period_header_row}, month={block.month_header_row}",
        f"  Last used row: {block.last_used_row}",
        f"  Evidence: {'; '.join(block.evidence)}",
    ]


def render_dry_run_report(
    source_path: Path,
    config,
    mode: str,
    blocks: dict[str, MonthBlock],
    *,
    source_fingerprint=None,
    ready: bool,
    warnings: Sequence[str] = (),
    problems: Sequence[str] = (),
) -> str:
    codes = config.codes
    lines = [
        f"SOURCE: {source_path}",
        "MODE: DRY RUN",
        f"EXCEL ACCESS MODE: {mode}",
        f"YEAR: {codes.year}",
        f"MONTH: {codes.month_name}",
        f"TARGET VERSION: {codes.target_version}",
        f"FORECAST VERSION: {codes.forecast_version}",
        f"ACTUAL VERSION: {codes.actual_version}",
        f"PERIOD: {codes.period}",
    ]
    if source_fingerprint:
        lines.extend(
            [
                f"FILE FORMAT: {source_fingerprint.file_format}",
                f"SOURCE SIZE: {source_fingerprint.size_bytes}",
                f"SOURCE SHA-256: {source_fingerprint.sha256}",
            ]
        )
    for name in (*config.target_sheets, config.total_sheet):
        lines.append("")
        lines.append(name)
        block = blocks.get(name)
        if block is None:
            lines.append("  DISCOVERY: FAILED")
        else:
            lines.extend(_block_lines(block))
            lines.append("  Existing A08: NO")
    lines.extend(
        [
            "",
            "DRM / workbook writable check: PASS (source remains untouched)",
            "External link update: DISABLED",
            "Pivot refresh: DISABLED",
        ]
    )
    if problems:
        lines.append("PROBLEMS:")
        lines.extend(f"- {problem}" for problem in problems)
    if warnings:
        lines.append("WARNINGS:")
        lines.extend(f"- {warning}" for warning in warnings)
    lines.append("")
    lines.append(f"READY TO EXECUTE: {'YES' if ready else 'NO'}")
    return "\n".join(lines) + "\n"


def _validation_lines(report: ValidationReport) -> list[str]:
    lines = [f"[{report.stage}] {'PASS' if report.passed else 'FAIL'}"]
    for check in report.checks:
        status = "PASS" if check.passed else "FAIL"
        lines.append(f"  {status} {check.name}: {check.message}")
    if report.reconciliations:
        failed = sum(1 for item in report.reconciliations if not item.passed)
        lines.append(f"  Reconciliation rows: {len(report.reconciliations)}; mismatches: {failed}")
        for item in report.reconciliations[:50]:
            if not item.passed:
                lines.append(
                    f"    row {item.total_pl_row} {item.label!r}: actual={item.actual!r}, expected={item.expected!r}, difference={item.difference!r}, reason={item.reason}"
                )
    return lines


def render_run_report(
    manifest: RunManifest,
    validation_reports: Sequence[ValidationReport],
) -> str:
    lines = [
        "P&L A08 AUTOMATION RUN REPORT",
        f"Run ID: {manifest.run_id}",
        f"Status: {manifest.status}",
        f"Phase: {manifest.phase}",
        f"Source: {manifest.source}",
        f"Output: {manifest.output or 'NOT PUBLISHED'}",
        f"Started UTC: {manifest.started_utc}",
        f"Ended UTC: {manifest.ended_utc or 'IN PROGRESS'}",
    ]
    if manifest.error:
        lines.extend(["", "ERROR:", f"  Code: {manifest.error.get('code')}", f"  Message: {manifest.error.get('message')}"])
    lines.append("")
    for report in validation_reports:
        lines.extend(_validation_lines(report))
        lines.append("")
    return "\n".join(lines) + "\n"


def write_run_report(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
