"""Engine-independent protocol for workbook adapters.

The planner and tools depend on this contract, not on pywin32 or a particular
file format.  COM and future fast engines implement it separately.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Protocol, Sequence

from .agent_contracts import TargetRef


@dataclass(frozen=True)
class EngineCapabilities:
    name: str
    can_read: bool = True
    can_write: bool = False
    supports_xlsb: bool = False
    supports_macros: bool = False
    supports_charts: bool = False
    supports_pivots: bool = False
    supports_external_links: bool = False
    notes: tuple[str, ...] = ()


@dataclass(frozen=True)
class EngineDecision:
    engine: str
    mode: str
    reason: str
    warnings: tuple[str, ...] = ()
    capabilities: EngineCapabilities = field(
        default_factory=lambda: EngineCapabilities("unknown", can_read=False)
    )


class WorkbookEngine(Protocol):
    """Minimum interface required by generic tools.

    Concrete methods may raise domain errors.  They must never return COM
    objects to the agent layer.
    """

    @property
    def capabilities(self) -> EngineCapabilities: ...

    def inspect(self) -> dict[str, Any]: ...

    def read_values(self, target: TargetRef) -> Sequence[Sequence[Any]]: ...

    def read_formulas(self, target: TargetRef) -> Sequence[Sequence[Any]]: ...

    def write_values(self, target: TargetRef, values: Sequence[Sequence[Any]]) -> None: ...

    def write_formulas(self, target: TargetRef, formulas: Sequence[Sequence[Any]]) -> None: ...

    def copy_range(self, source: TargetRef, destination: TargetRef, *, mode: str) -> None: ...

    def close(self, *, save: bool = False) -> None: ...

