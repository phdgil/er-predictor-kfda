from __future__ import annotations

from pathlib import Path
import sys

import pytest

from core.paths import (
    IMMUTABLE_ARCHIVED_ERTA_ROOT,
    IMMUTABLE_ERTA_ROOT,
    resolve_runtime_paths,
    validate_mutable_directory,
)


def test_default_mutable_roots_are_outside_the_install_tree(monkeypatch, tmp_path):
    local = tmp_path / "local"
    profile = tmp_path / "profile"
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    monkeypatch.setenv("USERPROFILE", str(profile))
    monkeypatch.delenv("ER_PREDICTOR_PORTABLE", raising=False)
    monkeypatch.delattr(sys, "frozen", raising=False)
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)

    paths = resolve_runtime_paths()

    assert paths.state_root == (local / "ER_Predictor" / "v1").resolve()
    assert paths.export_root == (profile / "Documents" / "ER_Predictor" / "Exports").resolve()
    assert paths.state_root.is_dir()
    assert paths.export_root.is_dir()
    assert not paths.portable


def test_portable_mode_is_confined_to_the_new_install_root(monkeypatch, tmp_path):
    install = tmp_path / "ER_Predictor"
    resources = tmp_path / "bundle"
    install.mkdir()
    resources.mkdir()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(install / "ER_Predictor.exe"))
    monkeypatch.setattr(sys, "_MEIPASS", str(resources), raising=False)
    monkeypatch.setenv("ER_PREDICTOR_PORTABLE", "1")

    paths = resolve_runtime_paths()

    assert paths.portable
    assert paths.state_root == (install / "ER_Predictor_UserData" / "state").resolve()
    assert paths.export_root == (install / "ER_Predictor_UserData" / "Exports").resolve()


def test_old_erta_package_is_rejected_before_writes(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(IMMUTABLE_ERTA_ROOT / "ERTA_Predictor.exe"))
    monkeypatch.setattr(sys, "_MEIPASS", str(IMMUTABLE_ERTA_ROOT / "_internal"), raising=False)
    monkeypatch.delenv("ER_PREDICTOR_PORTABLE", raising=False)

    with pytest.raises(RuntimeError, match="immutable ERTA_Predictor"):
        resolve_runtime_paths()


def test_user_selected_output_rejects_read_only_and_old_package_roots(tmp_path):
    resources = tmp_path / "resources"
    resources.mkdir()
    with pytest.raises(RuntimeError, match="read-only resources"):
        validate_mutable_directory(resources / "exports", forbidden_roots=(resources,))

    with pytest.raises(RuntimeError, match="immutable ERTA_Predictor"):
        validate_mutable_directory(IMMUTABLE_ERTA_ROOT / "exports")


@pytest.mark.parametrize(
    "legacy_root",
    (IMMUTABLE_ERTA_ROOT, IMMUTABLE_ARCHIVED_ERTA_ROOT),
)
@pytest.mark.parametrize("descendant", (Path(), Path("exports") / "batch"))
def test_legacy_package_roots_are_rejected_before_probe_creates_them(
    tmp_path, legacy_root, descendant
):
    target = tmp_path / legacy_root / descendant

    with pytest.raises(RuntimeError, match="immutable ERTA_Predictor"):
        validate_mutable_directory(target)

    assert not target.exists()


def test_archived_legacy_package_match_is_case_insensitive_and_component_exact(tmp_path):
    archived_root = (
        tmp_path
        / "fDa_EnDoCrInE_dIsRuPtIoN"
        / "_ArChIvE"
        / "LeGaCy_ApPs"
        / "erta_predictor"
    )
    forbidden_target = archived_root / "Exports"

    with pytest.raises(RuntimeError, match="immutable ERTA_Predictor"):
        validate_mutable_directory(forbidden_target)

    assert not forbidden_target.exists()

    sibling_target = archived_root.parent / "ERTA_Predictor_Backup" / "Exports"
    assert validate_mutable_directory(sibling_target) == sibling_target.resolve()
    assert sibling_target.is_dir()
