"""Create and verify hash-bound manifests without interpreting source contents.

This module only reads source bytes.  In particular, it never opens Office files with
an Office application or parser, so embedded active content cannot execute.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable, Mapping, Sequence

MANIFEST_SCHEMA = "kfda-intake-manifest/v1"
PARSER_VERSION = "1.0.0"
CHUNK_SIZE = 1024 * 1024
_REPARSE_POINT = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)

# The two scope spreadsheets are intentionally distinguished: the 77-substance
# revision is authoritative; the earlier 74-substance sheet remains evidence only.
KFDA_ATTACHMENT_LABELS: tuple[dict[str, str], ...] = (
    {
        "filename": "☆적용범위설정물질(인체독성참고치 미설정 물질 포함)_송부용_260826.xlsx",
        "role": "scope-substance-list-final-77",
        "authority_status": "authoritative",
        "privacy_class": "confidential",
    },
    {
        "filename": "식품의약품안전평가원_GPU 등.jpg",
        "role": "gpu-reference-image",
        "authority_status": "reference",
        "privacy_class": "confidential",
    },
    {
        "filename": "☆25191위평연003 및 인체독성참고치 미설정 물질.xlsx",
        "role": "scope-substance-list-original-74",
        "authority_status": "superseded",
        "privacy_class": "confidential",
    },
    {
        "filename": "식약처 데이터_(ERTA_ERBA) 최종.xlsx",
        "role": "erta-erba-final-data",
        "authority_status": "authoritative",
        "privacy_class": "confidential",
    },
    {
        "filename": "☆예측모델개발현황(공유용)_260820.hwpx",
        "role": "model-development-status",
        "authority_status": "reference",
        "privacy_class": "confidential",
    },
    {
        "filename": "식약처 로고.jpg",
        "role": "agency-logo",
        "authority_status": "reference",
        "privacy_class": "confidential",
    },
)


class IntakeError(ValueError):
    """An input violates the immutable, fail-closed intake contract."""


def build_kfda_input_spec(attachments_root: str | os.PathLike[str]) -> dict[str, Any]:
    """Return the explicit, labelled specification for the six supplied attachments."""
    root = Path(attachments_root)
    return {
        "inputs": [
            {
                "path": str(root / label["filename"]),
                "role": label["role"],
                "authority_status": label["authority_status"],
                "privacy_class": label["privacy_class"],
                "recursive": False,
            }
            for label in KFDA_ATTACHMENT_LABELS
        ]
    }


def _canonical(path: Path) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(os.fspath(path))))


def _is_within(candidate: str, root: str) -> bool:
    try:
        return os.path.commonpath((candidate, root)) == root
    except ValueError:  # Different Windows drives.
        return False


def _is_reparse_or_symlink(path: Path) -> bool:
    try:
        status = os.lstat(path)
    except OSError as error:
        raise IntakeError(f"missing or unreadable path: {path}") from error
    attributes = getattr(status, "st_file_attributes", 0)
    return stat.S_ISLNK(status.st_mode) or bool(attributes & _REPARSE_POINT)


def _reject_reparse_components(path: Path) -> None:
    """Reject an input reached through a link, junction, or other reparse point."""
    absolute = Path(_canonical(path))
    for component in (absolute, *absolute.parents):
        if _is_reparse_or_symlink(component):
            raise IntakeError(f"symlink or reparse point is not allowed: {component}")


def _require_regular_file(path: Path) -> os.stat_result:
    _reject_reparse_components(path)
    try:
        status = os.stat(path, follow_symlinks=False)
    except OSError as error:
        raise IntakeError(f"missing or unreadable path: {path}") from error
    if not stat.S_ISREG(status.st_mode):
        raise IntakeError(f"input is not a regular file: {path}")
    return status


def _require_directory(path: Path) -> None:
    _reject_reparse_components(path)
    if not path.is_dir():
        raise IntakeError(f"recursive input is not a directory: {path}")


def _parse_spec(spec: Mapping[str, Any], base: Path) -> list[dict[str, Any]]:
    entries = spec.get("inputs")
    if not isinstance(entries, list) or not entries:
        raise IntakeError("input spec must contain a non-empty 'inputs' list")
    parsed: list[dict[str, Any]] = []
    required = ("path", "role", "authority_status", "privacy_class", "recursive")
    for index, entry in enumerate(entries):
        if not isinstance(entry, Mapping) or any(key not in entry for key in required):
            raise IntakeError(f"input spec entry {index} is missing required fields")
        if not all(isinstance(entry[key], str) and entry[key] for key in required[:-1]):
            raise IntakeError(f"input spec entry {index} has invalid labels")
        if not isinstance(entry["recursive"], bool) or not isinstance(entry["path"], str) or not entry["path"]:
            raise IntakeError(f"input spec entry {index} has invalid path or recursive flag")
        path = Path(entry["path"])
        if not path.is_absolute():
            path = base / path
        parsed.append(
            {
                "path": Path(_canonical(path)),
                "role": entry["role"],
                "authority_status": entry["authority_status"],
                "privacy_class": entry["privacy_class"],
                "recursive": entry["recursive"],
            }
        )
    return parsed


def load_input_spec(path: str | os.PathLike[str]) -> list[dict[str, Any]]:
    """Load a JSON-only input specification; no source content is interpreted."""
    spec_path = Path(path)
    if _is_reparse_or_symlink(spec_path):
        raise IntakeError(f"symlink or reparse point is not allowed: {spec_path}")
    try:
        with spec_path.open("r", encoding="utf-8") as handle:
            spec = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        raise IntakeError(f"cannot read JSON input spec: {spec_path}") from error
    if not isinstance(spec, Mapping):
        raise IntakeError("input spec must be a JSON object")
    return _parse_spec(spec, Path(_canonical(spec_path)).parent)


def _iter_files(entry: Mapping[str, Any]) -> Iterable[tuple[Path, Mapping[str, Any]]]:
    root = Path(entry["path"])
    if not entry["recursive"]:
        _require_regular_file(root)
        yield root, entry
        return

    _require_directory(root)
    root_name = _canonical(root)
    pending = [root]
    while pending:
        directory = pending.pop()
        try:
            children = sorted(os.scandir(directory), key=lambda item: os.path.normcase(item.name))
        except OSError as error:
            raise IntakeError(f"cannot list input directory: {directory}") from error
        for child in children:
            child_path = Path(child.path)
            child_name = _canonical(child_path)
            if not _is_within(child_name, root_name):
                raise IntakeError(f"path escaped recursive root: {child_path}")
            if _is_reparse_or_symlink(child_path):
                raise IntakeError(f"symlink or reparse point is not allowed: {child_path}")
            try:
                if child.is_dir(follow_symlinks=False):
                    pending.append(child_path)
                elif child.is_file(follow_symlinks=False):
                    yield child_path, entry
                else:
                    raise IntakeError(f"input is not a regular file: {child_path}")
            except OSError as error:
                raise IntakeError(f"cannot inspect input path: {child_path}") from error


def _mtime_utc(status: os.stat_result) -> str:
    return datetime.fromtimestamp(status.st_mtime_ns / 1_000_000_000, tz=timezone.utc).isoformat(
        timespec="microseconds"
    ).replace("+00:00", "Z")


def _hash_record(path: Path, labels: Mapping[str, Any]) -> dict[str, Any]:
    before = _require_regular_file(path)
    digest = hashlib.sha256()
    try:
        with path.open("rb") as handle:
            while chunk := handle.read(CHUNK_SIZE):
                digest.update(chunk)
    except OSError as error:
        raise IntakeError(f"cannot read input file: {path}") from error
    after = _require_regular_file(path)
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise IntakeError(f"input changed while hashing: {path}")
    return {
        "path": _canonical(path),
        "role": labels["role"],
        "authority_status": labels["authority_status"],
        "privacy_class": labels["privacy_class"],
        "size_bytes": before.st_size,
        "sha256": digest.hexdigest(),
        "mtime_utc": _mtime_utc(before),
        "mtime_ns": before.st_mtime_ns,
    }


def _validate_output(output: Path, entries: Sequence[Mapping[str, Any]]) -> Path:
    target = Path(_canonical(output))
    for entry in entries:
        source = _canonical(Path(entry["path"]))
        if entry["recursive"] and _is_within(_canonical(target), source):
            raise IntakeError(f"output is inside recursive input root: {target}")
        if not entry["recursive"] and _canonical(target) == source:
            raise IntakeError(f"output overwrites an input file: {target}")
    return target


def _atomic_write(path: Path, payload: Mapping[str, Any]) -> None:
    parent = path.parent
    if not parent.is_dir():
        raise IntakeError(f"output directory does not exist: {parent}")
    encoded = (json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2) + "\n").encode("utf-8")
    try:
        with NamedTemporaryFile("wb", dir=parent, prefix=f".{path.name}.", suffix=".tmp", delete=False) as handle:
            temporary = Path(handle.name)
            handle.write(encoded)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except OSError as error:
        try:
            temporary.unlink(missing_ok=True)
        except UnboundLocalError:
            pass
        raise IntakeError(f"cannot atomically publish manifest: {path}") from error


def create_manifest(entries: Sequence[Mapping[str, Any]], output: str | os.PathLike[str]) -> dict[str, Any]:
    """Hash inputs, publish a deterministic manifest atomically, and return it."""
    if not entries:
        raise IntakeError("at least one input is required")
    target = _validate_output(Path(output), entries)
    files: list[dict[str, Any]] = []
    seen: set[str] = set()
    for entry in entries:
        for path, labels in _iter_files(entry):
            canonical = _canonical(path)
            if canonical in seen:
                raise IntakeError(f"duplicate canonical input path: {canonical}")
            seen.add(canonical)
            files.append(_hash_record(path, labels))
    files.sort(key=lambda item: item["path"])
    manifest = {
        "manifest_schema": MANIFEST_SCHEMA,
        "parser_version": PARSER_VERSION,
        "files": files,
    }
    _atomic_write(target, manifest)
    return manifest


def verify_manifest(path: str | os.PathLike[str]) -> dict[str, Any]:
    """Verify every recorded source byte-for-byte and metadata-for-metadata."""
    manifest_path = Path(path)
    if _is_reparse_or_symlink(manifest_path):
        raise IntakeError(f"symlink or reparse point is not allowed: {manifest_path}")
    try:
        with manifest_path.open("r", encoding="utf-8") as handle:
            manifest = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        raise IntakeError(f"cannot read manifest: {manifest_path}") from error
    if not isinstance(manifest, Mapping) or manifest.get("manifest_schema") != MANIFEST_SCHEMA:
        raise IntakeError("unsupported manifest schema")
    if manifest.get("parser_version") != PARSER_VERSION or not isinstance(manifest.get("files"), list):
        raise IntakeError("unsupported or malformed manifest")

    observed: list[dict[str, Any]] = []
    seen: set[str] = set()
    for record in manifest["files"]:
        if not isinstance(record, Mapping) or not isinstance(record.get("path"), str):
            raise IntakeError("malformed manifest file record")
        source = Path(record["path"])
        canonical = _canonical(source)
        if canonical != record["path"] or canonical in seen:
            raise IntakeError("manifest contains duplicate or non-normalized paths")
        seen.add(canonical)
        labels = {key: record.get(key) for key in ("role", "authority_status", "privacy_class")}
        if not all(isinstance(value, str) and value for value in labels.values()):
            raise IntakeError("malformed manifest file labels")
        observed.append(_hash_record(source, labels))
    observed.sort(key=lambda item: item["path"])
    expected = sorted(manifest["files"], key=lambda item: item.get("path", ""))
    if observed != expected:
        raise IntakeError("manifest verification failed: source file changed")
    return dict(manifest)


def _argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create or verify a K-FDA SHA-256 intake manifest.")
    commands = parser.add_subparsers(dest="command", required=True)
    create = commands.add_parser("create", help="hash an explicit JSON input specification")
    create.add_argument("--input-spec", required=True, help="JSON specification of exact files or recursive roots")
    create.add_argument("--output", required=True, help="new manifest path outside all recursive input roots")
    verify = commands.add_parser("verify", help="verify all sources recorded in a manifest")
    verify.add_argument("--manifest", required=True, help="manifest created by this tool")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _argument_parser().parse_args(argv)
    try:
        if args.command == "create":
            create_manifest(load_input_spec(args.input_spec), args.output)
        else:
            verify_manifest(args.manifest)
    except IntakeError as error:
        print(f"error: {error}", file=__import__("sys").stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
