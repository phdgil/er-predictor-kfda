from __future__ import annotations

import json

import pytest

from core.telemetry import EventLog


def test_event_log_is_privacy_safe_and_records_correlation_fields(tmp_path):
    log = EventLog(tmp_path)
    event_id = log.emit(
        "model.inference",
        task="classification",
        subtype="er_alpha",
        model_id="model-v1",
        duration_ms=12,
        status="ok",
    )

    record = json.loads(log.path.read_text(encoding="utf-8").strip())
    assert record["event_id"] == event_id
    assert record["event"] == "model.inference"
    assert record["task"] == "classification"
    with pytest.raises(ValueError, match="forbidden"):
        log.emit("bad", smiles="CCO")


def test_event_log_rotates_and_degrades_when_state_root_is_unwritable(tmp_path):
    log = EventLog(tmp_path / "state", max_bytes=1024, backups=2)
    for index in range(30):
        log.emit("batch.progress", index=index, evidence="x" * 80)
    assert log.path.exists()
    assert log.path.with_suffix(".jsonl.1").exists()

    file_root = tmp_path / "not-a-directory"
    file_root.write_text("occupied", encoding="utf-8")
    disabled = EventLog(file_root)
    assert not disabled.enabled
    assert disabled.emit("app.ready")
