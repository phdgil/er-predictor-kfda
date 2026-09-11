"""Atomically publish the audited ER_Predictor package without touching ERTA."""
from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
import uuid

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "dist" / "ER_Predictor"
ARTIFACTS = ROOT / "artifacts"
BUILD_MANIFEST = ARTIFACTS / "ER_Predictor-build-manifest.json"
BASELINE_MANIFEST = Path(
    r"D:/research/FDA_endocrine_disruption/ER_Predictor_evidence/baseline/prechange_manifest.json"
)
BASELINE_ORIGINAL_OLD = Path(r"D:/research/FDA_endocrine_disruption/ERTA_Predictor")
IMMUTABLE_OLD = Path(
    r"D:/research/FDA_endocrine_disruption/_archive/legacy_apps/ERTA_Predictor"
)
PUBLISH_ROOT = Path(r"D:/research/FDA_endocrine_disruption/ER_Predictor")
PUBLISHED = PUBLISH_ROOT / "ER_Predictor"
ROLLBACK_ROOT = PUBLISH_ROOT / "_archive"
RECEIPT = ARTIFACTS / "ER_Predictor-publication-receipt.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def inventory(root: Path, *, size_key: str = "size_bytes") -> list[dict]:
    if not root.is_dir():
        raise RuntimeError(f"required tree is missing: {root}")
    return [
        {
            "path": path.relative_to(root).as_posix(),
            size_key: path.stat().st_size,
            "sha256": sha256(path),
        }
        for path in sorted(
            (item for item in root.rglob("*") if item.is_file()),
            key=lambda item: item.relative_to(root).as_posix(),
        )
    ]


def canonical(items: list[dict], *, size_key: str) -> list[tuple[str, int, str]]:
    return sorted(
        (str(item["path"]).replace("\\", "/"), int(item[size_key]), str(item["sha256"]).lower())
        for item in items
    )


def expected_distribution() -> list[tuple[str, int, str]]:
    manifest = json.loads(BUILD_MANIFEST.read_text(encoding="utf-8"))
    if manifest.get("product") != "ER_Predictor" or manifest.get("executable") != "dist/ER_Predictor/ER_Predictor.exe":
        raise RuntimeError("build manifest identity is invalid")
    prefix = "dist/ER_Predictor/"
    transformed = []
    for item in manifest.get("inventory", []):
        path = str(item.get("path", "")).replace("\\", "/")
        if not path.startswith(prefix):
            raise RuntimeError(f"build inventory path is outside release unit: {path}")
        transformed.append(
            {
                "path": path[len(prefix):],
                "size_bytes": item["size_bytes"],
                "sha256": item["sha256"],
            }
        )
    if not transformed:
        raise RuntimeError("build manifest has no distribution inventory")
    return canonical(transformed, size_key="size_bytes")


def approved_old_baseline() -> list[tuple[str, int, str]]:
    baseline = json.loads(BASELINE_MANIFEST.read_text(encoding="utf-8"))
    approved_root = BASELINE_ORIGINAL_OLD.resolve(strict=False)
    declared_root = Path(baseline.get("existing_package_root", "")).resolve(strict=False)
    if declared_root != approved_root:
        raise RuntimeError("approved baseline does not identify the immutable ERTA tree")
    scoped_items = []
    for item in baseline.get("existing_package_files", []):
        declared_path = Path(str(item.get("path", "")))
        candidate = declared_path if declared_path.is_absolute() else approved_root / declared_path
        try:
            relative_path = candidate.resolve(strict=False).relative_to(approved_root)
        except ValueError as error:
            raise RuntimeError(
                f"approved baseline path is outside the original immutable ERTA root: {declared_path}"
            ) from error
        if relative_path == Path("."):
            raise RuntimeError("approved baseline contains an empty file path")
        scoped_items.append({**item, "path": relative_path.as_posix()})
    return canonical(scoped_items, size_key="size")


def main() -> None:
    roots = [ROOT.resolve(), SOURCE.resolve(), IMMUTABLE_OLD.resolve(), PUBLISH_ROOT.resolve()]
    if roots[0] == roots[2] or roots[1] == roots[2] or roots[3] == roots[2]:
        raise RuntimeError("immutable ERTA path guard rejected publication")
    if not (SOURCE / "ER_Predictor.exe").is_file():
        raise RuntimeError("audited source package executable is missing")
    catalog_path = SOURCE / "_internal" / "models" / "erba" / "catalog.v2.json"
    if not catalog_path.is_file():
        raise RuntimeError("audited source package ERBA catalog is missing")
    try:
        catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(f"audited source package ERBA catalog cannot be read: {error}") from error
    if not isinstance(catalog, dict) or catalog.get("schema_version") != 2:
        raise RuntimeError("audited source package ERBA catalog schema_version must be 2")
    if not BUILD_MANIFEST.is_file() or not BASELINE_MANIFEST.is_file():
        raise RuntimeError("build or approved baseline manifest is missing")

    expected = expected_distribution()
    source_inventory = canonical(inventory(SOURCE), size_key="size_bytes")
    if source_inventory != expected:
        raise RuntimeError("source distribution differs from the audited build manifest")
    old_baseline = approved_old_baseline()
    old_before = canonical(inventory(IMMUTABLE_OLD, size_key="size"), size_key="size")
    if old_before != old_baseline:
        raise RuntimeError("immutable ERTA tree differs from the approved pre-change baseline")

    ARTIFACTS.mkdir(parents=True, exist_ok=True)
    PUBLISH_ROOT.mkdir(parents=True, exist_ok=True)
    ROLLBACK_ROOT.mkdir(parents=True, exist_ok=True)
    stage = PUBLISH_ROOT / f".ER_Predictor.staging.{uuid.uuid4().hex}"
    backup = ROLLBACK_ROOT / (
        "ER_Predictor.rollback." + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    )
    promoted = False
    try:
        shutil.copytree(SOURCE, stage)
        stage_inventory = canonical(inventory(stage), size_key="size_bytes")
        if stage_inventory != expected:
            raise RuntimeError("staged distribution differs from the build manifest")
        if PUBLISHED.exists():
            if backup.exists():
                raise RuntimeError(f"publication backup already exists: {backup}")
            os.replace(PUBLISHED, backup)
        try:
            os.replace(stage, PUBLISHED)
            promoted = True
        except Exception:
            if backup.exists() and not PUBLISHED.exists():
                os.replace(backup, PUBLISHED)
            raise
        published_inventory = canonical(inventory(PUBLISHED), size_key="size_bytes")
        if published_inventory != expected or not (PUBLISHED / "ER_Predictor.exe").is_file():
            raise RuntimeError("published distribution differs from the build manifest")
    finally:
        if stage.exists():
            shutil.rmtree(stage)

    old_after = canonical(inventory(IMMUTABLE_OLD, size_key="size"), size_key="size")
    old_equal = old_before == old_after == old_baseline
    if not old_equal:
        raise RuntimeError("immutable ERTA tree changed during publication")
    receipt = {
        "schema_version": 1,
        "published_at": datetime.now(timezone.utc).isoformat(),
        "source": str(SOURCE),
        "published": str(PUBLISHED),
        "published_executable": str(PUBLISHED / "ER_Predictor.exe"),
        "promoted": promoted,
        "source_matches_build_manifest": source_inventory == expected,
        "stage_matches_build_manifest": stage_inventory == expected,
        "published_matches_build_manifest": published_inventory == expected,
        "immutable_old_matches_approved_baseline_before": old_before == old_baseline,
        "immutable_old_matches_approved_baseline_after": old_after == old_baseline,
        "immutable_old_unchanged": old_equal,
        "build_manifest_sha256": sha256(BUILD_MANIFEST),
        "approved_baseline_manifest_sha256": sha256(BASELINE_MANIFEST),
        "published_inventory": inventory(PUBLISHED),
    }
    RECEIPT.write_text(
        json.dumps(receipt, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    print(f"Published atomically: {PUBLISHED / 'ER_Predictor.exe'}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Publication failed: {error}", file=sys.stderr)
        raise SystemExit(1) from error
