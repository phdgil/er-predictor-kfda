from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

import pytest

from modeling.kfda_redevelopment.hash_inputs import (
    IntakeError,
    build_kfda_input_spec,
    create_manifest,
    load_input_spec,
    main,
    verify_manifest,
)


def _spec(path: Path, *, recursive: bool = False) -> dict[str, object]:
    return {
        "inputs": [
            {
                "path": str(path),
                "role": "test-evidence",
                "authority_status": "reference",
                "privacy_class": "confidential",
                "recursive": recursive,
            }
        ]
    }


def test_cli_hashes_an_explicit_json_spec_deterministically(tmp_path: Path):
    source = tmp_path / "source.bin"
    source.write_bytes(b"immutable evidence\x00")
    spec_path = tmp_path / "inputs.json"
    spec_path.write_text(json.dumps(_spec(source)), encoding="utf-8")
    output = tmp_path / "manifest.json"

    assert main(["create", "--input-spec", str(spec_path), "--output", str(output)]) == 0
    first = output.read_bytes()
    assert main(["create", "--input-spec", str(spec_path), "--output", str(output)]) == 0
    manifest = json.loads(output.read_text(encoding="utf-8"))

    assert output.read_bytes() == first
    assert manifest["files"][0]["sha256"] == hashlib.sha256(source.read_bytes()).hexdigest()
    assert manifest["files"][0]["path"] == os.path.normcase(str(source.resolve()))
    assert main(["verify", "--manifest", str(output)]) == 0


def test_verify_rejects_content_or_metadata_drift(tmp_path: Path):
    source = tmp_path / "source.bin"
    source.write_bytes(b"before")
    output = tmp_path / "manifest.json"
    create_manifest(load_input_spec(_write_spec(tmp_path, _spec(source))), output)

    source.write_bytes(b"after!")

    with pytest.raises(IntakeError, match="changed"):
        verify_manifest(output)


def test_recursive_inventory_is_sorted_and_preserves_labels(tmp_path: Path):
    root = tmp_path / "tree"
    nested = root / "nested"
    nested.mkdir(parents=True)
    (root / "z.bin").write_bytes(b"z")
    (nested / "a.bin").write_bytes(b"a")
    output = tmp_path / "manifest.json"

    manifest = create_manifest(load_input_spec(_write_spec(tmp_path, _spec(root, recursive=True))), output)

    assert [Path(record["path"]).name for record in manifest["files"]] == ["a.bin", "z.bin"]
    assert {record["role"] for record in manifest["files"]} == {"test-evidence"}


def test_duplicate_canonical_paths_and_output_recursion_are_rejected(tmp_path: Path):
    source = tmp_path / "source.bin"
    source.write_bytes(b"evidence")
    duplicate = _spec(source)
    duplicate["inputs"].append(dict(duplicate["inputs"][0]))
    with pytest.raises(IntakeError, match="duplicate canonical"):
        create_manifest(load_input_spec(_write_spec(tmp_path, duplicate)), tmp_path / "manifest.json")

    root = tmp_path / "root"
    root.mkdir()
    (root / "evidence.bin").write_bytes(b"evidence")
    with pytest.raises(IntakeError, match="inside recursive input root"):
        create_manifest(load_input_spec(_write_spec(tmp_path, _spec(root, recursive=True))), root / "manifest.json")


def test_symlink_inputs_fail_closed_when_supported(tmp_path: Path):
    target = tmp_path / "target.bin"
    target.write_bytes(b"evidence")
    link = tmp_path / "link.bin"
    try:
        link.symlink_to(target)
    except (NotImplementedError, OSError):
        pytest.skip("symlink creation is unavailable on this platform")

    with pytest.raises(IntakeError, match="symlink or reparse point"):
        create_manifest(load_input_spec(_write_spec(tmp_path, _spec(link))), tmp_path / "manifest.json")


def test_kfda_spec_labels_the_authoritative_77_and_superseded_74_files(tmp_path: Path):
    specification = build_kfda_input_spec(tmp_path)

    assert len(specification["inputs"]) == 6
    by_role = {entry["role"]: entry for entry in specification["inputs"]}
    assert by_role["scope-substance-list-final-77"]["authority_status"] == "authoritative"
    assert by_role["scope-substance-list-original-74"]["authority_status"] == "superseded"


def _write_spec(tmp_path: Path, specification: dict[str, object]) -> Path:
    path = tmp_path / "inputs.json"
    path.write_text(json.dumps(specification), encoding="utf-8")
    return path
