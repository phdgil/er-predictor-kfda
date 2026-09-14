"""Runtime roots for the separate ER_Predictor product.

Bundled resources and the install tree are read-only. Mutable state and exports live
outside the installation unless the user explicitly opts into portable mode.
"""
from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import shutil
import sys
import tempfile

PRODUCT = "ER_Predictor"
PRODUCT_MAJOR = "v1"
SHARED_EXAMPLE_FILENAME = "test.xlsx"
IMMUTABLE_ERTA_ROOT = Path("FDA_endocrine_disruption") / "ERTA_Predictor"
IMMUTABLE_ARCHIVED_ERTA_ROOT = (
    Path("FDA_endocrine_disruption") / "_archive" / "legacy_apps" / "ERTA_Predictor"
)
IMMUTABLE_ERTA_ROOTS = (IMMUTABLE_ERTA_ROOT, IMMUTABLE_ARCHIVED_ERTA_ROOT)


@dataclass(frozen=True)
class RuntimePaths:
    resource_root: Path
    install_root: Path
    state_root: Path
    export_root: Path
    portable: bool


def _resolved(path: Path) -> Path:
    return path.expanduser().resolve(strict=False)


def _is_within(path: Path, parent: Path) -> bool:
    try:
        _resolved(path).relative_to(_resolved(parent))
        return True
    except ValueError:
        return False


def _is_old_package(path: Path) -> bool:
    parts = tuple(part.casefold() for part in _resolved(path).parts)
    for immutable_root in IMMUTABLE_ERTA_ROOTS:
        suffix = tuple(part.casefold() for part in immutable_root.parts)
        if any(
            parts[index : index + len(suffix)] == suffix
            for index in range(len(parts) - len(suffix) + 1)
        ):
            return True
    return False


def _reject_old_package(path: Path, label: str) -> None:
    if _is_old_package(path):
        raise RuntimeError(f"{label} must not resolve inside the immutable ERTA_Predictor package.")


def _is_archive_path(path: Path) -> bool:
    return any(part.casefold() == "_archive" for part in _resolved(path).parts)


def validate_mutable_directory(
    path: str | Path,
    *,
    forbidden_roots: tuple[str | Path, ...] = (),
) -> Path:
    """Resolve and probe a user-selected write directory outside read-only trees."""
    directory = _resolved(Path(path))
    _reject_old_package(directory, "Mutable directory")
    for root in forbidden_roots:
        if _is_within(directory, Path(root)):
            raise RuntimeError(f"Mutable directory must not resolve inside read-only resources: {root}")
    _probe_writable(directory)
    return directory


def _probe_writable(directory: Path) -> None:
    probe_path: Path | None = None
    try:
        directory.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix=".er_predictor_probe_", dir=directory, delete=False) as probe:
            probe_path = Path(probe.name)
        probe_path.unlink()
    except Exception as error:
        if probe_path is not None:
            try:
                probe_path.unlink()
            except OSError:
                pass
        raise RuntimeError(f"ER_Predictor path is not writable: {directory}") from error


def _validate_shared_example_location(path: Path, paths: RuntimePaths) -> Path:
    workbook = _resolved(path)
    _reject_old_package(workbook, "Shared example workbook")
    if _is_archive_path(workbook):
        raise RuntimeError("Shared example workbook must not resolve inside an archive.")
    for root in (paths.resource_root, paths.install_root):
        if _is_within(workbook, root):
            raise RuntimeError(
                f"Shared example workbook must not resolve inside installed resources: {root}"
            )
    return workbook


def _existing_shared_example(path: Path, paths: RuntimePaths) -> Path:
    workbook = _validate_shared_example_location(path, paths)
    if not workbook.exists():
        raise RuntimeError(f"Shared example workbook is missing: {workbook}")
    if not workbook.is_file():
        raise RuntimeError(f"Shared example workbook is not a file: {workbook}")
    _probe_writable(workbook.parent)
    return workbook


def _copy_exclusive_atomic(source: Path, destination: Path) -> None:
    """Publish a byte-exact copy atomically without replacing an existing file."""
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{destination.name}.",
            suffix=".tmp",
            dir=destination.parent,
            delete=False,
        ) as target:
            temporary = Path(target.name)
            with source.open("rb") as bundled:
                shutil.copyfileobj(bundled, target)
            target.flush()
            os.fsync(target.fileno())

        try:
            os.link(temporary, destination)
        except FileExistsError:
            pass
        except OSError as link_error:
            if os.name != "nt":
                raise RuntimeError(
                    f"Could not initialize the shared example workbook atomically: {destination}"
                ) from link_error
            try:
                os.rename(temporary, destination)
            except OSError as rename_error:
                if not destination.exists():
                    raise RuntimeError(
                        f"Could not initialize the shared example workbook atomically: {destination}"
                    ) from rename_error
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError(
            f"Could not initialize the shared example workbook: {destination}"
        ) from error
    finally:
        if temporary is not None:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass


def resolve_shared_example_input(
    example_path: str | Path | None = None,
) -> Path:
    """Return one writable shared batch input without mutating bundled resources.

    An explicit path must already be a user-owned file. Frozen published builds may
    reuse the workbook beside their actual app directory; all other first runs copy
    the bundled workbook to the user's Documents tree without replacing a file.
    """
    paths = resolve_runtime_paths()
    if example_path is not None:
        return _existing_shared_example(Path(example_path), paths)

    if bool(getattr(sys, "frozen", False)):
        executable = _resolved(Path(sys.executable))
        if executable.parent == paths.install_root:
            app_container = paths.install_root.parent
            published_example = app_container / SHARED_EXAMPLE_FILENAME
            if published_example.exists() and not _is_archive_path(app_container):
                if not published_example.is_file():
                    raise RuntimeError(
                        f"Shared example workbook is not a file: {published_example}"
                    )
                published_example = _validate_shared_example_location(
                    published_example, paths
                )
                _probe_writable(published_example.parent)
                return published_example

    profile = os.environ.get("USERPROFILE", "").strip()
    documents = (Path(profile) if profile else Path.home()) / "Documents"
    destination = _resolved(
        documents / PRODUCT / "Examples" / SHARED_EXAMPLE_FILENAME
    )
    destination = _validate_shared_example_location(destination, paths)
    if destination.exists():
        return _existing_shared_example(destination, paths)

    bundled = paths.resource_root / "templates" / SHARED_EXAMPLE_FILENAME
    if not bundled.is_file():
        raise RuntimeError(f"Bundled shared example workbook is missing: {bundled}")
    validate_mutable_directory(
        destination.parent,
        forbidden_roots=(paths.resource_root, paths.install_root),
    )
    _copy_exclusive_atomic(bundled, destination)
    return _existing_shared_example(destination, paths)


def resolve_runtime_paths() -> RuntimePaths:
    source_root = Path(__file__).resolve().parents[1]
    frozen = bool(getattr(sys, "frozen", False))
    resource_root = _resolved(Path(getattr(sys, "_MEIPASS", source_root)))
    install_root = _resolved(Path(sys.executable).parent if frozen else source_root)
    _reject_old_package(resource_root, "Resource root")
    _reject_old_package(install_root, "Install root")

    portable = os.environ.get("ER_PREDICTOR_PORTABLE", "").strip() == "1"
    if portable:
        mutable_root = install_root / "ER_Predictor_UserData"
        state_root = mutable_root / "state"
        export_root = mutable_root / "Exports"
    else:
        local_app_data = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
        documents = Path(os.environ.get("USERPROFILE", Path.home())) / "Documents"
        state_root = local_app_data / PRODUCT / PRODUCT_MAJOR
        export_root = documents / PRODUCT / "Exports"

    state_root = _resolved(state_root)
    export_root = _resolved(export_root)
    _reject_old_package(state_root, "State root")
    _reject_old_package(export_root, "Export root")
    _probe_writable(state_root)
    _probe_writable(export_root)
    return RuntimePaths(resource_root, install_root, state_root, export_root, portable)
