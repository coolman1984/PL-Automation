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

    def clear_range(self, target: TargetRef) -> None: ...

    def resolve_bounds(self, target: TargetRef) -> tuple[int, int, int, int]:
        """Return (first_row, first_col, last_row, last_col), 1-based."""
        ...

    def fill_formula_down(self, template: TargetRef, target: TargetRef) -> None: ...

    def insert_columns(self, target: TargetRef, count: int) -> None: ...

    def count_formula_errors(self, sheet: str) -> int:
        """Count cells whose formula currently evaluates to an Excel error."""
        ...

    def inspect_pivot_table(self, sheet: str, name: str) -> dict[str, Any]:
        """Structural facts about one named PivotTable; never a COM object."""
        ...

    def resolve_source_bounds(self, address: str) -> Any:
        """An opaque, comparable identity for a source address string.

        Must be directly comparable (``==``) with the ``source_bounds`` value
        returned by :meth:`inspect_pivot_table`, regardless of which address
        notation the caller used.
        """
        ...

    def update_pivot_source(self, sheet: str, name: str, new_source_address: str) -> None:
        """Change one PivotTable's source and perform a targeted refresh."""
        ...

    def close(self, *, save: bool = False) -> None: ...

