"""Generic fail-closed transaction state and journal.

The P&L workflow has its own detailed orchestration.  This small state machine
is the reusable safety boundary for future universal tools and recipes.
"""

from __future__ import annotations

import hashlib
import json
import os
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Mapping

from .file_transaction import sha256_file


class TransactionLockError(RuntimeError):
    """Raised when a source workspace is already locked by another run."""


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

# States where an Excel mutation may have been interrupted mid-flight. On
# resume the on-disk working copy cannot be trusted to match the journal, so
# these fail closed instead of resuming.
_UNSAFE_TO_RESUME: frozenset[TransactionState] = frozenset({
    TransactionState.EXECUTING,
    TransactionState.SAVED,
    TransactionState.REOPENED,
})

_HASH_STAGES = ("source", "backup", "working", "output")


def _pid_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        import ctypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


@dataclass
class TransactionContext:
    source_path: Path
    artifact_root: Path
    transaction_id: str = field(default_factory=lambda: f"run-{uuid.uuid4().hex}")
    state: TransactionState = TransactionState.RECEIVED
    source_sha256: str | None = None
    backup_sha256: str | None = None
    working_sha256: str | None = None
    output_sha256: str | None = None
    events: list[dict[str, Any]] = field(default_factory=list)
    started_utc: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    _workspace_locked: bool = field(default=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        self.source_path = Path(self.source_path).expanduser().resolve()
        self.artifact_root = Path(self.artifact_root).expanduser().resolve()
        self._record("created", {"source": str(self.source_path)})

    @property
    def journal_path(self) -> Path:
        return self.artifact_root / self.transaction_id / "transaction_manifest.json"

    def _source_lock_path(self) -> Path:
        digest = hashlib.sha256(str(self.source_path).encode("utf-8")).hexdigest()[:24]
        return self.artifact_root / "locks" / f"{digest}.lock"

    @staticmethod
    def _read_lock(lock_path: Path) -> dict[str, Any] | None:
        try:
            return json.loads(lock_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None

    def acquire_workspace_lock(self) -> None:
        """Lock the source workspace so only this transaction may mutate it."""
        if self._workspace_locked:
            return
        lock_path = self._source_lock_path()
        lock_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "transaction_id": self.transaction_id,
            "pid": os.getpid(),
            "source": str(self.source_path),
            "acquired_utc": datetime.now(timezone.utc).isoformat(),
        }
        for _attempt in range(2):
            try:
                fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                holder = self._read_lock(lock_path)
                if holder and holder.get("transaction_id") == self.transaction_id:
                    self._workspace_locked = True
                    return
                if holder and _pid_is_running(int(holder.get("pid", -1))):
                    raise TransactionLockError(
                        "Source workspace is locked by transaction "
                        f"{holder.get('transaction_id')} (pid {holder.get('pid')})"
                    )
                # Lock owner is gone (dead pid or unreadable lock file): reclaim it.
                try:
                    lock_path.unlink()
                except FileNotFoundError:
                    pass
                continue
            else:
                with os.fdopen(fd, "w", encoding="utf-8") as handle:
                    json.dump(payload, handle)
                self._workspace_locked = True
                self._record("workspace_locked", {"lock_path": str(lock_path)})
                return
        raise TransactionLockError(f"Could not acquire source workspace lock: {lock_path}")

    def release_workspace_lock(self) -> None:
        if not self._workspace_locked:
            return
        lock_path = self._source_lock_path()
        holder = self._read_lock(lock_path)
        if holder and holder.get("transaction_id") == self.transaction_id:
            try:
                lock_path.unlink()
            except FileNotFoundError:
                pass
        self._workspace_locked = False
        self._record("workspace_unlocked", {"lock_path": str(lock_path)})

    def record_hash(self, stage: str, path: Path) -> str:
        if stage not in _HASH_STAGES:
            raise ValueError(f"Unknown hash stage: {stage} (expected one of {_HASH_STAGES})")
        digest = sha256_file(Path(path))
        setattr(self, f"{stage}_sha256", digest)
        self._record(f"{stage}_hash_recorded", {"path": str(path), "sha256": digest})
        return digest

    def _record(self, event: str, evidence: Mapping[str, Any] | None = None) -> None:
        self.events.append({
            "sequence": len(self.events) + 1,
            "event": event,
            "state": self.state.value,
            "utc": datetime.now(timezone.utc).isoformat(),
            "evidence": dict(evidence or {}),
        })
        self.write_journal()

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
            "backup_sha256": self.backup_sha256,
            "working_sha256": self.working_sha256,
            "output_sha256": self.output_sha256,
            "started_utc": self.started_utc,
            "events": list(self.events),
        }

    @classmethod
    def resume(cls, journal_path: Path) -> "TransactionContext":
        """Reconstruct a transaction from its journal for safe resumption.

        Sequence numbers are preserved exactly so a resumed transaction never
        repeats a sequence number already written to disk. States where an
        Excel mutation may have been interrupted mid-flight fail closed
        instead of resuming.
        """
        payload = json.loads(Path(journal_path).read_text(encoding="utf-8"))
        context = cls.__new__(cls)
        context.source_path = Path(payload["source"])
        context.artifact_root = Path(payload["artifact_root"])
        context.transaction_id = payload["transaction_id"]
        context.state = TransactionState(payload["state"])
        context.source_sha256 = payload.get("source_sha256")
        context.backup_sha256 = payload.get("backup_sha256")
        context.working_sha256 = payload.get("working_sha256")
        context.output_sha256 = payload.get("output_sha256")
        context.started_utc = payload["started_utc"]
        context.events = list(payload["events"])
        context._workspace_locked = False
        if context.state in _UNSAFE_TO_RESUME:
            context.fail_safe(
                f"Resumed transaction found in unsafe-to-resume state: {context.state.value}",
                evidence={"resumed_from": str(journal_path)},
            )
        else:
            context._record("resumed", {"resumed_from": str(journal_path), "state": context.state.value})
        return context

    def write_journal(self, path: Path | None = None) -> Path:
        destination = Path(path or self.journal_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(destination.suffix + ".tmp.write")
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(self.manifest(), handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
        # Every state change replaces this file (Task 4: automatic atomic
        # writes), so a transient Windows file-lock (antivirus scanning a
        # just-written file, a lagging previous handle close) can make a
        # single os.replace briefly raise PermissionError. Retry briefly
        # before giving up; the write itself is still atomic either way.
        last_error: OSError | None = None
        for attempt in range(5):
            try:
                temporary.replace(destination)
                return destination
            except PermissionError as exc:
                last_error = exc
                if attempt < 4:
                    time.sleep(0.05 * (attempt + 1))
        assert last_error is not None
        raise last_error
