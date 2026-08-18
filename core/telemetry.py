"""Privacy-safe, bounded local JSONL diagnostics for ER_Predictor."""
from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
import threading
from typing import Any, Mapping
from uuid import uuid4

_FORBIDDEN_KEYS = frozenset({
    "smiles", "raw_smiles", "model_smiles", "cas", "cas_number", "name",
    "chemical_name", "structure", "prediction_row", "input_row",
})


class EventLog:
    def __init__(self, state_root: str | Path, *, max_bytes: int = 2_000_000, backups: int = 3) -> None:
        self.path = Path(state_root) / "logs" / "events.jsonl"
        self.max_bytes = max(1024, int(max_bytes))
        self.backups = max(1, int(backups))
        self._lock = threading.Lock()
        self.disabled_reason = ""
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as error:
            self.disabled_reason = type(error).__name__

    @property
    def enabled(self) -> bool:
        return not self.disabled_reason

    def emit(self, event: str, **fields: Any) -> str:
        event_id = uuid4().hex
        if not self.enabled:
            return event_id
        normalized = {str(key): value for key, value in fields.items()}
        forbidden = sorted(key for key in normalized if key.lower() in _FORBIDDEN_KEYS)
        if forbidden:
            raise ValueError(f"Chemical input fields are forbidden in telemetry: {', '.join(forbidden)}")
        record: Mapping[str, Any] = {
            "timestamp": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
            "event_id": event_id,
            "event": str(event),
            **normalized,
        }
        payload = json.dumps(record, ensure_ascii=False, sort_keys=True, default=str)
        try:
            with self._lock:
                self._rotate_if_needed(len(payload.encode("utf-8")) + 1)
                with self.path.open("a", encoding="utf-8", newline="\n") as handle:
                    handle.write(payload + "\n")
        except Exception as error:
            self.disabled_reason = type(error).__name__
        return event_id

    def _rotate_if_needed(self, incoming_bytes: int) -> None:
        if not self.path.exists() or self.path.stat().st_size + incoming_bytes <= self.max_bytes:
            return
        oldest = self.path.with_suffix(self.path.suffix + f".{self.backups}")
        oldest.unlink(missing_ok=True)
        for index in range(self.backups - 1, 0, -1):
            source = self.path.with_suffix(self.path.suffix + f".{index}")
            if source.exists():
                source.replace(self.path.with_suffix(self.path.suffix + f".{index + 1}"))
        self.path.replace(self.path.with_suffix(self.path.suffix + ".1"))
