"""Controlled workbook calculation and timeout handling."""

from __future__ import annotations

import time

from .constants import XL_CALCULATION_DONE
from .errors import CalculationTimeoutError


def wait_for_calculation(
    app: object, timeout_seconds: int, poll_seconds: float = 0.25
) -> float:
    started = time.monotonic()
    deadline = started + timeout_seconds
    while True:
        try:
            state = int(app.CalculationState)
        except Exception:
            state = XL_CALCULATION_DONE
        if state == XL_CALCULATION_DONE:
            return time.monotonic() - started
        if time.monotonic() >= deadline:
            raise CalculationTimeoutError(
                f"Excel calculation did not finish within {timeout_seconds} seconds"
            )
        time.sleep(min(poll_seconds, max(0.01, deadline - time.monotonic())))


def calculate_workbook_once(
    app: object,
    workbook: object,
    timeout_seconds: int,
    *,
    full_rebuild: bool = True,
) -> float:
    try:
        if full_rebuild:
            app.CalculateFullRebuild()
        else:
            workbook.Calculate()
    except Exception as exc:  # pragma: no cover - requires Excel
        raise CalculationTimeoutError(f"Excel calculation could not start: {exc}") from exc
    return wait_for_calculation(app, timeout_seconds)

