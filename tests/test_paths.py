from __future__ import annotations

from pathlib import Path
import sys

import pytest

import core.paths as paths_module
from core.paths import (
    IMMUTABLE_ARCHIVED_ERTA_ROOT,
    IMMUTABLE_ERTA_ROOT,
    RuntimePaths,
    resolve_shared_example_input,
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


def _shared_example_runtime(monkeypatch, tmp_path):
    resource_root = (tmp_path / "bundle").resolve()
    install_root = (tmp_path / "install" / "ER_Predictor").resolve()
    resource_root.mkdir(parents=True)
    install_root.mkdir(parents=True)
    paths = RuntimePaths(
        resource_root=resource_root,
        install_root=install_root,
        state_root=(tmp_path / "state").resolve(),
        export_root=(tmp_path / "exports").resolve(),
        portable=False,
    )
    monkeypatch.setattr(paths_module, "resolve_runtime_paths", lambda: paths)
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "profile"))
    monkeypatch.delattr(sys, "frozen", raising=False)
    return paths


def test_shared_example_is_initialized_once_and_reused_without_overwrite(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    template = paths.resource_root / "templates" / "test.xlsx"
    template.parent.mkdir()
    template.write_bytes(b"bundled workbook bytes")

    first_tab = resolve_shared_example_input()
    assert first_tab == (
        tmp_path
        / "profile"
        / "Documents"
        / "ER_Predictor"
        / "Examples"
        / "test.xlsx"
    ).resolve()
    assert first_tab.read_bytes() == b"bundled workbook bytes"

    first_tab.write_bytes(b"user-edited workbook bytes")
    template.write_bytes(b"replacement bundle bytes")
    second_tab = resolve_shared_example_input()

    assert second_tab == first_tab
    assert second_tab.read_bytes() == b"user-edited workbook bytes"
    assert not tuple(second_tab.parent.glob(".test.xlsx.*.tmp"))


def test_shared_example_explicit_user_file_wins_and_installed_copy_is_rejected(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    supplied = tmp_path / "user-inputs" / "test.xlsx"
    supplied.parent.mkdir()
    supplied.write_bytes(b"user supplied")

    assert resolve_shared_example_input(supplied) == supplied.resolve()

    missing = tmp_path / "user-inputs" / "missing.xlsx"
    with pytest.raises(RuntimeError, match="Shared example workbook is missing"):
        resolve_shared_example_input(missing)
    assert not missing.exists()

    installed = paths.resource_root / "templates" / "test.xlsx"
    installed.parent.mkdir()
    installed.write_bytes(b"read only bundle")
    with pytest.raises(RuntimeError, match="installed resources"):
        resolve_shared_example_input(installed)


def test_frozen_distribution_uses_only_its_writable_parent_container_example(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    app_container = (tmp_path / "published" / "ER_Predictor").resolve()
    install_root = app_container / "ER_Predictor"
    install_root.mkdir(parents=True)
    paths = RuntimePaths(
        resource_root=paths.resource_root,
        install_root=install_root,
        state_root=paths.state_root,
        export_root=paths.export_root,
        portable=False,
    )
    monkeypatch.setattr(paths_module, "resolve_runtime_paths", lambda: paths)
    executable = install_root / "ER_Predictor.exe"
    published = app_container / "test.xlsx"
    published.write_bytes(b"published user workbook")
    bundled = paths.resource_root / "templates" / "test.xlsx"
    bundled.parent.mkdir()
    bundled.write_bytes(b"bundled workbook")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(executable))

    resolved = resolve_shared_example_input()

    assert resolved == published.resolve()
    assert resolved.read_bytes() == b"published user workbook"
    assert not (
        tmp_path
        / "profile"
        / "Documents"
        / "ER_Predictor"
        / "Examples"
        / "test.xlsx"
    ).exists()


def test_unwritable_published_example_errors_without_initializing_alternate(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    app_container = (tmp_path / "published" / "ER_Predictor").resolve()
    install_root = app_container / "ER_Predictor"
    install_root.mkdir(parents=True)
    paths = RuntimePaths(
        resource_root=paths.resource_root,
        install_root=install_root,
        state_root=paths.state_root,
        export_root=paths.export_root,
        portable=False,
    )
    monkeypatch.setattr(paths_module, "resolve_runtime_paths", lambda: paths)
    published = app_container / "test.xlsx"
    published.write_bytes(b"user-edited published workbook")
    bundled = paths.resource_root / "templates" / "test.xlsx"
    bundled.parent.mkdir()
    bundled.write_bytes(b"different bundled workbook")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        sys, "executable", str(install_root / "ER_Predictor.exe")
    )

    def reject_published_directory(directory):
        assert Path(directory) == app_container
        raise RuntimeError(f"ER_Predictor path is not writable: {directory}")

    monkeypatch.setattr(paths_module, "_probe_writable", reject_published_directory)

    with pytest.raises(
        RuntimeError,
        match=r"ER_Predictor path is not writable: .*published.*ER_Predictor",
    ):
        resolve_shared_example_input()

    alternate = (
        tmp_path
        / "profile"
        / "Documents"
        / "ER_Predictor"
        / "Examples"
        / "test.xlsx"
    )
    assert not alternate.exists()
    assert published.read_bytes() == b"user-edited published workbook"


def test_frozen_distribution_ignores_test_file_outside_actual_parent_container(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    actual_container = (tmp_path / "actual-app-container").resolve()
    install_root = actual_container / "ER_Predictor"
    install_root.mkdir(parents=True)
    unrelated = tmp_path / "unrelated-user-folder" / "test.xlsx"
    unrelated.parent.mkdir()
    unrelated.write_bytes(b"unrelated workbook")
    paths = RuntimePaths(
        resource_root=paths.resource_root,
        install_root=install_root,
        state_root=paths.state_root,
        export_root=paths.export_root,
        portable=False,
    )
    monkeypatch.setattr(paths_module, "resolve_runtime_paths", lambda: paths)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        sys, "executable", str(install_root / "ER_Predictor.exe")
    )
    bundled = paths.resource_root / "templates" / "test.xlsx"
    bundled.parent.mkdir()
    bundled.write_bytes(b"bundled workbook")

    resolved = resolve_shared_example_input()

    assert resolved.read_bytes() == b"bundled workbook"
    assert resolved != unrelated.resolve()
    assert unrelated.read_bytes() == b"unrelated workbook"


def test_source_execution_requires_explicit_path_to_use_a_sibling_example(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    sibling = paths.install_root.parent / "test.xlsx"
    sibling.write_bytes(b"source sibling")
    bundled = paths.resource_root / "templates" / "test.xlsx"
    bundled.parent.mkdir()
    bundled.write_bytes(b"bundled workbook")

    default = resolve_shared_example_input()

    assert default != sibling.resolve()
    assert default.read_bytes() == b"bundled workbook"
    assert resolve_shared_example_input(sibling) == sibling.resolve()


def test_archived_frozen_distribution_preserves_its_sibling_example(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    archived_install = (
        tmp_path / "_archive" / "ER_Predictor.rollback" / "ER_Predictor"
    ).resolve()
    archived_install.mkdir(parents=True)
    archived_example = archived_install.parent / "test.xlsx"
    archived_example.write_bytes(b"archived workbook")
    paths = RuntimePaths(
        resource_root=paths.resource_root,
        install_root=archived_install,
        state_root=paths.state_root,
        export_root=paths.export_root,
        portable=False,
    )
    monkeypatch.setattr(paths_module, "resolve_runtime_paths", lambda: paths)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        sys, "executable", str(archived_install / "ER_Predictor.exe")
    )
    bundled = paths.resource_root / "templates" / "test.xlsx"
    bundled.parent.mkdir()
    bundled.write_bytes(b"bundled workbook")

    resolved = resolve_shared_example_input()

    assert resolved.read_bytes() == b"bundled workbook"
    assert resolved != archived_example.resolve()
    assert archived_example.read_bytes() == b"archived workbook"


def test_shared_example_missing_bundle_has_clear_error_without_partial_output(
    monkeypatch, tmp_path
):
    _shared_example_runtime(monkeypatch, tmp_path)

    with pytest.raises(RuntimeError, match="Bundled shared example workbook is missing"):
        resolve_shared_example_input()

    destination = (
        tmp_path
        / "profile"
        / "Documents"
        / "ER_Predictor"
        / "Examples"
        / "test.xlsx"
    )
    assert not destination.exists()


def test_shared_example_atomic_race_does_not_overwrite_user_file(
    monkeypatch, tmp_path
):
    paths = _shared_example_runtime(monkeypatch, tmp_path)
    template = paths.resource_root / "templates" / "test.xlsx"
    template.parent.mkdir()
    template.write_bytes(b"bundled workbook")

    def publish_user_edit_first(_source, destination):
        Path(destination).write_bytes(b"user won the race")
        raise FileExistsError

    monkeypatch.setattr(paths_module.os, "link", publish_user_edit_first)

    resolved = resolve_shared_example_input()

    assert resolved.read_bytes() == b"user won the race"
    assert not tuple(resolved.parent.glob(".test.xlsx.*.tmp"))
