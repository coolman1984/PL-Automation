"""Generic fail-closed transaction state and journal.

The P&L workflow has its own detailed orchestration.  This small state machine
is the reusable safety boundary for future universal tools and recipes.
"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Mapping


class TransactionState(str, Enum):
    RECEIVED = "received"
    PROBED = "probed"
    BACKED_UP = "backed_up"
    SNAPSHOTTED = "snapshotted"
    PLANNED = "planned"
    APPROVED = "approved"
    WORKING_COPY_READY = "working_copy_ready"
    EXECUTING = "executing"
    SAVED = "saved"
    REOPENED = "reopened"
    VALIDATED = "validated"
    PUBLISHED = "published"
    FAILED_SAFE = "failed_safe"


_ALLOWED: dict[TransactionState, frozenset[TransactionState]] = {
    TransactionState.RECEIVED: frozenset({TransactionState.PROBED, TransactionState.FAILED_SAFE}),
    TransactionState.PROBED: frozenset({TransactionState.BACKED_UP, TransactionState.FAILED_SAFE}),
    TransactionState.BACKED_UP: frozenset({TransactionState.SNAPSHOTTED, TransactionState.FAILED_SAFE}),
    TransactionState.SNAPSHOTTED: frozenset({TransactionState.PLANNED, TransactionState.FAILED_SAFE}),
    TransactionState.PLANNED: frozenset({TransactionState.APPROVED, TransactionState.FAILED_SAFE}),
    TransactionState.APPROVED: frozenset({TransactionState.WORKING_COPY_READY, TransactionState.FAILED_SAFE}),
    TransactionState.WORKING_COPY_READY: frozenset({TransactionState.EXECUTING, TransactionState.FAILED_SAFE}),
    TransactionState.EXECUTING: frozenset({TransactionState.SAVED, TransactionState.FAILED_SAFE}),
    TransactionState.SAVED: frozenset({TransactionState.REOPENED, TransactionState.FAILED_SAFE}),
    TransactionState.REOPENED: frozenset({TransactionState.VALIDATED, TransactionState.FAILED_SAFE}),
    TransactionState.VALIDATED: frozenset({TransactionState.PUBLISHED, TransactionState.FAILED_SAFE}),
    TransactionState.PUBLISHED: frozenset(),
    TransactionState.FAILED_SAFE: frozenset(),
}


@dataclass
class TransactionContext:
    source_path: Path
    artifact_root: Path
    transaction_id: str = field(default_factory=lambda: f"run-{uuid.uuid4().hex}")
    state: TransactionState = TransactionState.RECEIVED
    source_sha256: str | None = None
    events: list[dict[str, Any]] = field(default_factory=list)
    started_utc: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    def __post_init__(self) -> None:
        self.source_path = Path(self.source_path).expanduser().resolve()
        self.artifact_root = Path(self.artifact_root).expanduser().resolve()
        self._record("created", {"source": str(self.source_path)})

    def _record(self, event: str, evidence: Mapping[str, Any] | None = None) -> None:
        self.events.append({
            "sequence": len(self.events) + 1,
            "event": event,
            "state": self.state.value,
            "utc": datetime.now(timezone.utc).isoformat(),
            "evidence": dict(evidence or {}),
        })

    def transition(self, next_state: TransactionState, *, evidence: Mapping[str, Any] | None = None) -> None:
        next_state = TransactionState(next_state)
        if next_state not in _ALLOWED[self.state]:
            raise ValueError(f"Invalid transaction transition: {self.state.value} -> {next_state.value}")
        self.state = next_state
        self._record("state_changed", {"from": self.events[-1]["state"] if self.events else None, **dict(evidence or {})})

    def fail_safe(self, error: BaseException | str, *, evidence: Mapping[str, Any] | None = None) -> None:
        if self.state not in {TransactionState.PUBLISHED, TransactionState.FAILED_SAFE}:
            self.state = TransactionState.FAILED_SAFE
        self._record("failed_safe", {"error": str(error), **dict(evidence or {})})

    def manifest(self) -> dict[str, Any]:
        return {
            "schema_version": "1.0",
            "transaction_id": self.transaction_id,
            "source": str(self.source_path),
            "artifact_root": str(self.artifact_root),
            "state": self.state.value,
            "source_sha256": self.source_sha256,
            "started_utc": self.started_utc,
            "events": list(self.events),
        }

    def write_journal(self, path: Path | None = None) -> Path:
        destination = Path(path or (self.artifact_root / self.transaction_id / "transaction_manifest.json"))
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(destination.suffix + ".tmp")
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(self.manifest(), handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
        temporary.replace(destination)
        return destination
