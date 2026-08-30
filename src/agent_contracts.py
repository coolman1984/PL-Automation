"""Stable, JSON-safe contracts shared by the agent planner and tool runner.

These models deliberately contain no Excel or COM objects.  They are the small
language between an AI agent and the deterministic safety layer.
"""

from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping


SCHEMA_VERSION = "1.0"


def _json_safe(value: Any) -> Any:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(item) for item in value]
    if hasattr(value, "to_dict"):
        return _json_safe(value.to_dict())
    raise TypeError(f"Value is not JSON safe: {type(value).__name__}")


def as_json(value: Any) -> str:
    """Serialize a contract deterministically for logs and agent hand-offs."""
    return json.dumps(_json_safe(value), ensure_ascii=False, indent=2, sort_keys=True)


def _require_text(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} must be a non-empty string")
    return value.strip()


def _mapping(value: Any, field_name: str) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise ValueError(f"{field_name} must be an object")
    return dict(value)


@dataclass(frozen=True)
class TargetRef:
    """A resolved workbook target; never an active Excel selection."""

    workbook_id: str
    sheet: str | None = None
    address: str | None = None
    object_name: str | None = None

    def __post_init__(self) -> None:
        _require_text(self.workbook_id, "target.workbook_id")
        if self.sheet is not None:
            _require_text(self.sheet, "target.sheet")
        if self.address is not None:
            _require_text(self.address, "target.address")
        if self.object_name is not None:
            _require_text(self.object_name, "target.object_name")
        if not any((self.sheet, self.address, self.object_name)):
            raise ValueError("target must identify a sheet, address, or object")

    def to_dict(self) -> dict[str, Any]:
        return {
            key: value
            for key, value in asdict(self).items()
            if value is not None
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "TargetRef":
        data = _mapping(value, "target")
        allowed = {"workbook_id", "sheet", "address", "object_name"}
        unknown = sorted(set(data) - allowed)
        if unknown:
            raise ValueError(f"target contains unknown fields: {', '.join(unknown)}")
        return cls(**data)


@dataclass(frozen=True)
class ToolRequest:
    schema_version: str
    transaction_id: str
    tool: str
    target: TargetRef | None = None
    arguments: dict[str, Any] = field(default_factory=dict)
    preconditions: tuple[dict[str, Any], ...] = ()
    expected_effect: dict[str, Any] = field(default_factory=dict)
    dry_run: bool = True

    def __post_init__(self) -> None:
        if self.schema_version != SCHEMA_VERSION:
            raise ValueError(f"Unsupported schema_version: {self.schema_version}")
        _require_text(self.transaction_id, "transaction_id")
        _require_text(self.tool, "tool")
        if not isinstance(self.arguments, dict):
            raise ValueError("arguments must be an object")
        if not isinstance(self.expected_effect, dict):
            raise ValueError("expected_effect must be an object")
        if not isinstance(self.dry_run, bool):
            raise ValueError("dry_run must be boolean")

    @classmethod
    def new(
        cls,
        tool: str,
        *,
        target: TargetRef | None = None,
        arguments: Mapping[str, Any] | None = None,
        preconditions: tuple[Mapping[str, Any], ...] | list[Mapping[str, Any]] = (),
        expected_effect: Mapping[str, Any] | None = None,
        dry_run: bool = True,
    ) -> "ToolRequest":
        return cls(
            schema_version=SCHEMA_VERSION,
            transaction_id=f"run-{uuid.uuid4().hex}",
            tool=_require_text(tool, "tool"),
            target=target,
            arguments=_mapping(arguments, "arguments"),
            preconditions=tuple(_mapping(item, "precondition") for item in preconditions),
            expected_effect=_mapping(expected_effect, "expected_effect"),
            dry_run=dry_run,
        )

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "ToolRequest":
        data = _mapping(value, "request")
        allowed = {
            "schema_version", "transaction_id", "tool", "target", "arguments",
            "preconditions", "expected_effect", "dry_run",
        }
        unknown = sorted(set(data) - allowed)
        if unknown:
            raise ValueError(f"request contains unknown fields: {', '.join(unknown)}")
        target = data.get("target")
        preconditions = data.get("preconditions", ())
        if not isinstance(preconditions, (list, tuple)):
            raise ValueError("preconditions must be an array")
        return cls(
            schema_version=data.get("schema_version", SCHEMA_VERSION),
            transaction_id=_require_text(data.get("transaction_id"), "transaction_id"),
            tool=_require_text(data.get("tool"), "tool"),
            target=TargetRef.from_dict(target) if target is not None else None,
            arguments=_mapping(data.get("arguments"), "arguments"),
            preconditions=tuple(_mapping(item, "precondition") for item in preconditions),
            expected_effect=_mapping(data.get("expected_effect"), "expected_effect"),
            dry_run=data.get("dry_run", True),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "transaction_id": self.transaction_id,
            "tool": self.tool,
            "target": self.target.to_dict() if self.target else None,
            "arguments": _json_safe(self.arguments),
            "preconditions": _json_safe(self.preconditions),
            "expected_effect": _json_safe(self.expected_effect),
            "dry_run": self.dry_run,
        }


@dataclass(frozen=True)
class ToolMetrics:
    elapsed_ms: int = 0
    cells_touched: int = 0
    objects_touched: int = 0
    bytes_read: int = 0
    bytes_written: int = 0

    def to_dict(self) -> dict[str, int]:
        return asdict(self)


@dataclass(frozen=True)
class ToolError:
    code: str
    message: str
    recoverable: bool = True
    details: dict[str, Any] = field(default_factory=dict)
    suggested_action: str | None = None

    def __post_init__(self) -> None:
        _require_text(self.code, "error.code")
        _require_text(self.message, "error.message")

    def to_dict(self) -> dict[str, Any]:
        return _json_safe(asdict(self))


@dataclass(frozen=True)
class ToolResult:
    ok: bool
    tool: str
    changed: bool = False
    affected_ranges: tuple[str, ...] = ()
    before_evidence: dict[str, Any] = field(default_factory=dict)
    after_evidence: dict[str, Any] = field(default_factory=dict)
    warnings: tuple[str, ...] = ()
    metrics: ToolMetrics = field(default_factory=ToolMetrics)
    error: ToolError | None = None

    def __post_init__(self) -> None:
        _require_text(self.tool, "tool")
        if self.ok and self.error is not None:
            raise ValueError("successful result cannot contain an error")
        if not self.ok and self.error is None:
            raise ValueError("failed result must contain an error")

    @classmethod
    def success(
        cls,
        tool: str,
        *,
        changed: bool = False,
        affected_ranges: tuple[str, ...] | list[str] = (),
        before_evidence: Mapping[str, Any] | None = None,
        after_evidence: Mapping[str, Any] | None = None,
        warnings: tuple[str, ...] | list[str] = (),
        metrics: ToolMetrics | None = None,
    ) -> "ToolResult":
        return cls(
            ok=True,
            tool=tool,
            changed=changed,
            affected_ranges=tuple(affected_ranges),
            before_evidence=_mapping(before_evidence, "before_evidence"),
            after_evidence=_mapping(after_evidence, "after_evidence"),
            warnings=tuple(str(item) for item in warnings),
            metrics=metrics or ToolMetrics(),
        )

    @classmethod
    def failure(
        cls,
        tool: str,
        error: ToolError,
        *,
        warnings: tuple[str, ...] | list[str] = (),
        metrics: ToolMetrics | None = None,
    ) -> "ToolResult":
        return cls(
            ok=False,
            tool=tool,
            warnings=tuple(str(item) for item in warnings),
            metrics=metrics or ToolMetrics(),
            error=error,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "tool": self.tool,
            "changed": self.changed,
            "affected_ranges": list(self.affected_ranges),
            "before_evidence": _json_safe(self.before_evidence),
            "after_evidence": _json_safe(self.after_evidence),
            "warnings": list(self.warnings),
            "metrics": self.metrics.to_dict(),
            "error": self.error.to_dict() if self.error else None,
        }


@dataclass(frozen=True)
class PlanStep:
    step_id: str
    tool: str
    purpose: str
    request: ToolRequest
    validation: tuple[dict[str, Any], ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "step_id": self.step_id,
            "tool": self.tool,
            "purpose": self.purpose,
            "request": self.request.to_dict(),
            "validation": _json_safe(self.validation),
        }


@dataclass(frozen=True)
class OperationPlan:
    transaction_id: str
    intent: str
    assumptions: tuple[str, ...] = ()
    unresolved: tuple[str, ...] = ()
    risk: str = "low"
    requires_approval: bool = False
    steps: tuple[PlanStep, ...] = ()

    def __post_init__(self) -> None:
        _require_text(self.transaction_id, "transaction_id")
        _require_text(self.intent, "intent")
        if self.risk not in {"none", "low", "medium", "high", "critical"}:
            raise ValueError(f"Unknown plan risk: {self.risk}")
        if self.unresolved and self.requires_approval is False:
            raise ValueError("A plan with unresolved items cannot be approval-free")

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": SCHEMA_VERSION,
            "transaction_id": self.transaction_id,
            "intent": self.intent,
            "assumptions": list(self.assumptions),
            "unresolved": list(self.unresolved),
            "risk": self.risk,
            "requires_approval": self.requires_approval,
            "steps": [step.to_dict() for step in self.steps],
        }

