"""Create and compare reproducibility manifests with format-aware normalization."""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
from pathlib import Path
import platform
import sys
import struct
import zipfile


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def normalized_zip(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    normalized_size = 0
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            payload = archive.read(name)
            encoded = name.encode("utf-8")
            normalized_size += len(encoded) + len(payload)
            digest.update(struct.pack("<Q", len(encoded)))
            digest.update(encoded)
            digest.update(struct.pack("<Q", len(payload)))
            digest.update(hashlib.sha256(payload).digest())
    return digest.hexdigest(), normalized_size


def normalized_pe(path: Path) -> tuple[str, int]:
    payload = bytearray(path.read_bytes())
    if payload[:2] != b"MZ" or len(payload) < 0x40:
        raise RuntimeError(f"expected PE executable: {path}")
    pe_offset = struct.unpack_from("<I", payload, 0x3C)[0]
    if payload[pe_offset:pe_offset + 4] != b"PE\0\0":
        raise RuntimeError(f"invalid PE signature: {path}")
    section_count = struct.unpack_from("<H", payload, pe_offset + 6)[0]
    optional_header_size = struct.unpack_from("<H", payload, pe_offset + 20)[0]
    optional_header = pe_offset + 24
    section_table = optional_header + optional_header_size
    image_end = section_table
    for index in range(section_count):
        section = section_table + index * 40
        raw_size = struct.unpack_from("<I", payload, section + 16)[0]
        raw_offset = struct.unpack_from("<I", payload, section + 20)[0]
        image_end = max(image_end, raw_offset + raw_size)
    image = payload[:image_end]
    image[pe_offset + 8:pe_offset + 12] = b"\0" * 4
    image[optional_header + 64:optional_header + 68] = b"\0" * 4
    return sha256_bytes(image), len(image)


def normalized_sha256(path: Path, relative: str) -> tuple[str, str, int]:
    raw = sha256_file(path)
    if relative == "ER_Predictor.exe":
        normalized, normalized_size = normalized_pe(path)
        return raw, normalized, normalized_size
    if relative == "_internal/base_library.zip":
        normalized, normalized_size = normalized_zip(path)
        return raw, normalized, normalized_size
    return raw, raw, path.stat().st_size


def runtime_source_sha256(root: Path) -> str:
    paths = [root / "app.py"]
    paths.extend(sorted((root / "core").glob("*.py")))
    paths.extend(sorted((root / "gui").glob("*.py")))
    digest = hashlib.sha256()
    for path in sorted(paths, key=lambda item: str(item.relative_to(root))):
        relative = str(path.relative_to(root)).replace("\\", "/").encode("utf-8")
        payload = path.read_bytes()
        digest.update(struct.pack("<Q", len(relative)))
        digest.update(relative)
        digest.update(struct.pack("<Q", len(payload)))
        digest.update(hashlib.sha256(payload).digest())
    return digest.hexdigest()


def build_manifest(root: Path, destination: Path) -> None:
    dist = root / "dist" / "ER_Predictor"
    files = sorted((path for path in dist.rglob("*") if path.is_file()), key=lambda path: str(path))
    catalog_path = root / "models" / "erba" / "catalog.v2.json"
    try:
        catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(f"required ERBA catalog cannot be read: {error}") from error
    if not isinstance(catalog, dict) or catalog.get("schema_version") != 2:
        raise RuntimeError("ERBA catalog schema_version must be 2.")
    inventory = []
    for path in files:
        relative = str(path.relative_to(dist)).replace("\\", "/")
        raw, normalized, normalized_size = normalized_sha256(path, relative)
        inventory.append({
            "path": relative,
            "size_bytes": path.stat().st_size,
            "sha256": raw,
            "normalized_sha256": normalized,
            "normalized_size_bytes": normalized_size,
        })
    components = sorted(
        ({"name": distribution.metadata["Name"].lower(), "version": distribution.version}
         for distribution in importlib.metadata.distributions()),
        key=lambda item: item["name"],
    )
    source = {
        "app_spec_sha256": sha256_file(root / "app.spec"),
        "requirements_lock_sha256": sha256_file(root / "requirements-lock.txt"),
        "wheelhouse_inventory_sha256": sha256_file(root / "artifacts" / "ER_Predictor-wheelhouse-inventory.json"),
        "catalog_sha256": sha256_file(catalog_path),
        "runtime_source_sha256": runtime_source_sha256(root),
        "python_runtime": {
            "implementation": platform.python_implementation(),
            "version": platform.python_version(),
            "cache_tag": sys.implementation.cache_tag,
        },
    }
    payload = {
        "schema_version": 1,
        "product": "ER_Predictor",
        "executable": "ER_Predictor.exe",
        "components": components,
        "source": source,
        "distribution_inventory": inventory,
        "normalization": {
            "ER_Predictor.exe": "raw hash retained; excluded from equality because PyInstaller CArchive embeds build-environment bytecode paths; semantics bound by runtime source, app.spec, Python runtime, and component inventory",
            "_internal/base_library.zip": "raw and normalized hashes retained; excluded from equality because compiled standard-library code objects embed build-environment paths; semantics bound by Python runtime and component inventory",
        },
    }
    destination.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def normalized_inventory(manifest: dict) -> list[dict]:
    semantic_containers = {"ER_Predictor.exe", "_internal/base_library.zip"}
    return [
        {
            "path": item["path"],
            "normalized_size_bytes": item["normalized_size_bytes"],
            "normalized_sha256": item["normalized_sha256"],
        }
        for item in manifest["distribution_inventory"]
        if item["path"] not in semantic_containers
    ]


def compare(first_path: Path, second_path: Path, destination: Path) -> None:
    first = json.loads(first_path.read_text(encoding="utf-8"))
    second = json.loads(second_path.read_text(encoding="utf-8"))
    same = (
        first["components"] == second["components"]
        and first["source"] == second["source"]
        and normalized_inventory(first) == normalized_inventory(second)
    )
    raw_differences = [
        left["path"]
        for left, right in zip(first["distribution_inventory"], second["distribution_inventory"])
        if left["sha256"] != right["sha256"]
    ]
    receipt = {
        "schema_version": 1,
        "first_manifest": first_path.name,
        "second_manifest": second_path.name,
        "normalized_artifacts_equal": same,
        "raw_container_differences": raw_differences,
    }
    destination.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    if not same:
        raise RuntimeError("independent clean-build normalized inventories differ")


def main() -> None:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    manifest_parser = subparsers.add_parser("manifest")
    manifest_parser.add_argument("--root", type=Path, required=True)
    manifest_parser.add_argument("--output", type=Path, required=True)
    compare_parser = subparsers.add_parser("compare")
    compare_parser.add_argument("--first", type=Path, required=True)
    compare_parser.add_argument("--second", type=Path, required=True)
    compare_parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.command == "manifest":
        build_manifest(args.root.resolve(), args.output.resolve())
    else:
        compare(args.first.resolve(), args.second.resolve(), args.output.resolve())


if __name__ == "__main__":
    main()
