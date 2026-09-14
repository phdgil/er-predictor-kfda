# -*- mode: python ; coding: utf-8 -*-
"""Auditable one-folder ER_Predictor package specification."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys

PROJECT_ROOT = Path(SPECPATH).resolve()
sys.path.insert(0, str(PROJECT_ROOT))

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules
from core.erba_release_allowlist import TRUSTED_ERBA_ARTIFACTS

MODELS_ROOT = PROJECT_ROOT / "models"
ERBA_ROOT = MODELS_ROOT / "erba"
ERBA_CATALOG = ERBA_ROOT / "catalog.v2.json"
ERBA_AD_ROOT = ERBA_ROOT / "ad"
TEMPLATES_ROOT = PROJECT_ROOT / "templates"
SHARED_EXAMPLE_TEMPLATE = TEMPLATES_ROOT / "test.xlsx"
SHARED_EXAMPLE_TEMPLATE_SHA256 = "5a1f569f8f6a5cd47bff67a189645c3f9461fcf07bd24f4e5b4f83197f3350aa"
ERBA_AD_STEMS = (
    "classification_er_alpha_reference",
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def catalogued_erba_datas() -> list[tuple[str, str]]:
    """Return the catalog itself plus only released, integrity-checked ERBA artifacts."""
    if not ERBA_CATALOG.is_file():
        raise SystemExit(f"Required ERBA catalog is missing: {ERBA_CATALOG}")
    try:
        catalog = json.loads(ERBA_CATALOG.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise SystemExit(f"Required ERBA catalog cannot be read: {error}") from error
    if not isinstance(catalog, dict) or catalog.get("schema_version") != 2:
        raise SystemExit("ERBA catalog schema_version must be 2.")
    routes = catalog.get("routes")
    if not isinstance(routes, list):
        raise SystemExit("ERBA catalog routes are missing.")

    datas = [(str(ERBA_CATALOG), "models/erba")]
    seen_paths: set[Path] = set()
    seen_route_ids: set[str] = set()
    seen_release_manifests: set[Path] = set()
    released = 0
    for route in routes:
        if not isinstance(route, dict) or route.get("release_status") != "released":
            continue
        try:
            route_id = f"{route['task']}:{route['subtype']}"
        except KeyError as error:
            raise SystemExit("A released ERBA route lacks task/subtype identity.") from error
        trusted = TRUSTED_ERBA_ARTIFACTS.get(route_id)
        if not isinstance(trusted, dict):
            raise SystemExit(f"Released route is absent from executable trust allowlist: {route_id}")
        artifact = route.get("artifact", route)
        if not isinstance(artifact, dict):
            raise SystemExit("A released ERBA catalog route has no artifact record.")
        try:
            relative = Path(str(artifact["relative_path"]))
            expected_size = int(artifact["size_bytes"])
            expected_sha256 = str(artifact["sha256"]).lower()
        except (KeyError, TypeError, ValueError) as error:
            raise SystemExit("A released ERBA catalog route is incomplete.") from error
        if (
            trusted.get("model_id") != artifact.get("model_id")
            or trusted.get("size_bytes") != expected_size
            or trusted.get("sha256") != expected_sha256
            or trusted.get("release_status") != "released"
        ):
            raise SystemExit(f"Released route disagrees with executable trust allowlist: {route_id}")
        manifest_relative = Path(str(trusted.get("release_manifest_relative_path", "")))
        expected_manifest_sha256 = str(trusted.get("release_manifest_sha256", "")).lower()
        if (
            manifest_relative.is_absolute()
            or ".." in manifest_relative.parts
            or manifest_relative == Path(".")
            or len(expected_manifest_sha256) != 64
        ):
            raise SystemExit(f"Unsafe release-manifest trust entry: {route_id}")
        release_manifest = (ERBA_ROOT / manifest_relative).resolve()
        if (
            ERBA_ROOT.resolve() not in release_manifest.parents
            or not release_manifest.is_file()
            or sha256(release_manifest).lower() != expected_manifest_sha256
        ):
            raise SystemExit(f"Release manifest failed executable trust audit: {route_id}")
        if relative.is_absolute() or ".." in relative.parts or relative == Path("."):
            raise SystemExit(f"Unsafe ERBA artifact path in catalog: {relative}")
        source = (ERBA_ROOT / relative).resolve()
        if ERBA_ROOT.resolve() not in source.parents or not source.is_file():
            raise SystemExit(f"Required catalogued ERBA artifact is missing: {relative}")
        if source.stat().st_size != expected_size or sha256(source).lower() != expected_sha256:
            raise SystemExit(f"Catalogued ERBA artifact failed integrity audit: {relative}")
        if source in seen_paths:
            raise SystemExit(f"Catalogued ERBA artifact is duplicated: {relative}")
        seen_paths.add(source)
        datas.append((str(source), str(Path("models/erba") / relative.parent)))
        if release_manifest not in seen_release_manifests:
            datas.append((str(release_manifest), str(Path("models/erba") / manifest_relative.parent)))
            seen_release_manifests.add(release_manifest)
        seen_route_ids.add(route_id)
        released += 1
    if released == 0:
        raise SystemExit("ERBA catalog has no released artifacts to package.")
    if seen_route_ids != set(TRUSTED_ERBA_ARTIFACTS):
        raise SystemExit("Executable trust allowlist contains an unbundled or duplicate route.")
    return datas


def legacy_model_datas() -> list[tuple[str, str]]:
    """Preserve the legacy ERTA model tree without admitting non-catalogued ERBA files."""
    if not MODELS_ROOT.is_dir():
        raise SystemExit(f"Required legacy model directory is missing: {MODELS_ROOT}")
    return [
        (str(path), str(Path("models") / path.relative_to(MODELS_ROOT).parent))
        for path in MODELS_ROOT.rglob("*")
        if path.is_file() and ERBA_ROOT not in path.parents
    ]


def erba_ad_datas() -> list[tuple[str, str]]:
    """Bundle only the three route-specific AD references and their safe caches."""
    expected = {
        ERBA_AD_ROOT / f"{stem}{suffix}"
        for stem in ERBA_AD_STEMS
        for suffix in (".xlsx", "_ad_cache.npz")
    }
    missing = sorted(str(path) for path in expected if not path.is_file())
    if missing:
        raise SystemExit(f"Required ERBA AD assets are missing: {missing}")
    found = {path for path in ERBA_AD_ROOT.rglob("*") if path.is_file()}
    if found != expected:
        unexpected = sorted(str(path.relative_to(ERBA_AD_ROOT)) for path in found - expected)
        raise SystemExit(f"Unlisted ERBA AD assets are present: {unexpected}")
    return [(str(path), "models/erba/ad") for path in sorted(expected)]


for required_directory in (PROJECT_ROOT / "data", TEMPLATES_ROOT):
    if not required_directory.is_dir():
        raise SystemExit(f"Required legacy resource directory is missing: {required_directory}")
if (
    not SHARED_EXAMPLE_TEMPLATE.is_file()
    or sha256(SHARED_EXAMPLE_TEMPLATE) != SHARED_EXAMPLE_TEMPLATE_SHA256
):
    raise SystemExit(
        f"Required shared example template is missing or altered: {SHARED_EXAMPLE_TEMPLATE}"
    )


def is_runtime_module(name: str) -> bool:
    return not ({"test", "tests", "testing"} & set(name.split(".")))


hiddenimports = []
for package in ("rdkit", "sklearn", "scipy", "PIL", "h5py", "xgboost", "joblib"):
    hiddenimports += collect_submodules(package, filter=is_runtime_module, on_error="warn")
hiddenimports += [
    "core.erba_features",
    "core.paths",
    "openpyxl",
    "requests",
    "rdkit.Avalon",
    "rdkit.Avalon.pyAvalonTools",
    "rdkit.Chem.rdMolDescriptors",
    "tensorflow.keras",
    "tensorflow.keras.models",
    "xgboost",
    "joblib",
]

# Explicit package resources and native extension libraries are required at runtime.
datas = legacy_model_datas() + catalogued_erba_datas() + erba_ad_datas()
datas += [(str(PROJECT_ROOT / "data"), "data"), (str(TEMPLATES_ROOT), "templates")]
for package in ("rdkit", "h5py", "openpyxl", "PIL", "xgboost"):
    datas += collect_data_files(package)
binaries = []
for package in ("rdkit", "h5py", "scipy", "xgboost"):
    binaries += collect_dynamic_libs(package)

block_cipher = None

a = Analysis(
    ["app.py"],
    pathex=[str(PROJECT_ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ER_Predictor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ER_Predictor",
)
