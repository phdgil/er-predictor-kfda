"""Runtime roots for the separate ER_Predictor product.

Bundled resources and the install tree are read-only. Mutable state and exports live
outside the installation unless the user explicitly opts into portable mode.
"""
from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import sys
import tempfile

PRODUCT = "ER_Predictor"
PRODUCT_MAJOR = "v1"
IMMUTABLE_ERTA_ROOT = Path("FDA_endocrine_disruption") / "ERTA_Predictor"


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
    suffix = tuple(part.casefold() for part in IMMUTABLE_ERTA_ROOT.parts)
    return any(parts[index:index + len(suffix)] == suffix for index in range(len(parts) - len(suffix) + 1))




def _reject_old_package(path: Path, label: str) -> None:
    if _is_old_package(path):
        raise RuntimeError(f"{label} must not resolve inside the immutable ERTA_Predictor package.")

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
    directory.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.NamedTemporaryFile(prefix=".er_predictor_probe_", dir=directory, delete=False) as probe:
            probe_path = Path(probe.name)
        probe_path.unlink()
    except Exception as error:
        raise RuntimeError(f"ER_Predictor path is not writable: {directory}") from error


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
