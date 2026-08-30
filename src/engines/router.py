"""Central, deterministic engine-selection policy."""

from __future__ import annotations

from ..engine_contract import EngineCapabilities, EngineDecision
from ..file_probe import FileProbeResult


EXCEL_COM = EngineCapabilities(
    "excel_com",
    can_read=True,
    can_write=True,
    supports_xlsb=True,
    supports_macros=True,
    supports_charts=True,
    supports_pivots=True,
    supports_external_links=True,
    notes=("Requires desktop Excel and authorized access.",),
)
FAST_OOXML = EngineCapabilities(
    "fast_ooxml",
    can_read=True,
    can_write=True,
    notes=("Only eligible after the feature-preservation gate passes.",),
)


def decide_engine(probe: FileProbeResult) -> EngineDecision:
    """Translate the dependency-free probe into one explicit engine decision."""
    if not probe.recognized:
        return EngineDecision(
            "none",
            "stop",
            "The file signature is not a recognized Excel workbook.",
            capabilities=EngineCapabilities("none", can_read=False),
        )
    if probe.recommended_engine == "fast_ooxml_candidate" and probe.fast_edit_candidate:
        return EngineDecision(
            "fast_ooxml",
            "candidate",
            "The package looks simple enough for a future fast engine; preservation gates are still required.",
            warnings=("No fast mutation is allowed until feature and round-trip tests pass.",),
            capabilities=FAST_OOXML,
        )
    if probe.protection in {"nasca_drm", "office_encrypted"} or probe.recommended_engine == "excel_com":
        return EngineDecision(
            "excel_com",
            probe.recommended_com_mode or "open",
            probe.reason,
            warnings=probe.warnings,
            capabilities=EXCEL_COM,
        )
    return EngineDecision(
        "none",
        "stop",
        "No safe write engine was selected for this workbook.",
        warnings=probe.warnings,
        capabilities=EngineCapabilities("none", can_read=False),
    )

