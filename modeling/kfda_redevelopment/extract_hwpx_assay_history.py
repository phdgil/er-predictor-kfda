"""Read-only, deterministic HWPX evidence extraction.

This module treats HWPX as hostile ZIP/XML input.  It never invokes an Office
application and records source anchors rather than interpreting assay meaning.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from tempfile import NamedTemporaryFile
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile, ZipInfo


class UnsafeArchiveError(ValueError):
    """Raised when an archive cannot be safely treated as evidence."""


@dataclass(frozen=True)
class ArchiveLimits:
    max_members: int = 2_048
    max_total_uncompressed: int = 512 * 1024 * 1024
    max_member_uncompressed: int = 128 * 1024 * 1024
    max_compression_ratio: float = 100.0


DEFAULT_LIMITS = ArchiveLimits()
_IMAGE_EXTENSIONS = {".bmp", ".gif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".webp", ".wmf", ".emf"}
_CRITICAL_ASSAY_FIELDS = (
    "assay_identifier_and_version",
    "test_system_and_species",
    "receptor_target_and_endpoint_definition",
    "test_article_identity_and_concentration_units",
    "positive_negative_and_vehicle_controls",
    "replicate_count_and_acceptance_criteria",
    "data_normalization_and_calculation_method",
    "result_interpretation_thresholds",
    "deviations_and_quality_control",
    "source_document_provenance",
)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _safe_member_name(name: str) -> str:
    if not name or "\x00" in name or "\\" in name:
        raise UnsafeArchiveError("archive member has an unsafe name")
    path = PurePosixPath(name)
    if path.is_absolute() or any(part in ("", ".", "..") for part in path.parts):
        raise UnsafeArchiveError("archive member path is absolute or traverses directories")
    # A colon is not meaningful in an OOXML/HWPX member and avoids drive/ADS ambiguity.
    if any(":" in part for part in path.parts):
        raise UnsafeArchiveError("archive member has an unsafe path component")
    return path.as_posix()


def _checked_members(archive_path: str | Path, limits: ArchiveLimits) -> tuple[dict[str, bytes], list[dict[str, Any]]]:
    try:
        with ZipFile(archive_path) as archive:
            infos = archive.infolist()
            if len(infos) > limits.max_members:
                raise UnsafeArchiveError("archive member count exceeds limit")
            names: set[str] = set()
            folded_names: set[str] = set()
            total = 0
            validated: list[tuple[str, ZipInfo]] = []
            for info in infos:
                name = _safe_member_name(info.filename)
                if name in names or name.casefold() in folded_names:
                    raise UnsafeArchiveError("archive contains duplicate or case-colliding members")
                names.add(name)
                folded_names.add(name.casefold())
                if info.flag_bits & 0x1:
                    raise UnsafeArchiveError("encrypted archive members are not permitted")
                if info.file_size > limits.max_member_uncompressed:
                    raise UnsafeArchiveError("archive member exceeds uncompressed-size limit")
                if info.compress_size == 0 and info.file_size:
                    raise UnsafeArchiveError("archive member has an unsafe compression ratio")
                if info.compress_size and info.file_size / info.compress_size > limits.max_compression_ratio:
                    raise UnsafeArchiveError("archive member compression ratio exceeds limit")
                total += info.file_size
                if total > limits.max_total_uncompressed:
                    raise UnsafeArchiveError("archive total uncompressed size exceeds limit")
                validated.append((name, info))

            contents: dict[str, bytes] = {}
            manifest: list[dict[str, Any]] = []
            for name, info in sorted(validated, key=lambda item: (item[0].casefold(), item[0])):
                # Stream rather than trusting metadata or using a decompression-sized allocation.
                remaining = limits.max_member_uncompressed
                chunks: list[bytes] = []
                with archive.open(info, "r") as source:
                    while True:
                        chunk = source.read(min(1024 * 1024, remaining + 1))
                        if not chunk:
                            break
                        remaining -= len(chunk)
                        if remaining < 0:
                            raise UnsafeArchiveError("archive member exceeds extraction-size limit")
                        chunks.append(chunk)
                data = b"".join(chunks)
                if len(data) != info.file_size:
                    raise UnsafeArchiveError("archive member size differs from its declared size")
                contents[name] = data
                manifest.append({
                    "compressed_size": info.compress_size,
                    "name": name,
                    "sha256": hashlib.sha256(data).hexdigest(),
                    "uncompressed_size": len(data),
                })
    except (BadZipFile, OSError) as error:
        raise UnsafeArchiveError("input is not a readable ZIP archive") from error
    _reject_external_relationships(contents)
    return contents, manifest


def _reject_external_relationships(contents: dict[str, bytes]) -> None:
    for name in sorted(contents, key=lambda value: (value.casefold(), value)):
        if not name.lower().endswith(".rels"):
            continue
        root = _parse_xml(contents[name], name)
        for element in root.iter():
            if _local_name(element.tag) == "Relationship" and element.attrib.get("TargetMode", "").casefold() == "external":
                raise UnsafeArchiveError("external relationships are not permitted")


def _parse_xml(data: bytes, member: str) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise UnsafeArchiveError(f"XML declarations that can load entities are not permitted: {member}")
    try:
        return ET.fromstring(data)
    except ET.ParseError as error:
        raise UnsafeArchiveError(f"malformed XML member: {member}") from error


def _element_text(element: ET.Element) -> str:
    return "".join(element.itertext())


def _atomic_json(output_path: str | Path, payload: dict[str, Any]) -> Path:
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8")
    with NamedTemporaryFile("wb", dir=destination.parent, prefix=f".{destination.name}.", suffix=".tmp", delete=False) as temporary:
        temporary.write(encoded)
        temporary.flush()
        os.fsync(temporary.fileno())
        temporary_name = temporary.name
    try:
        os.replace(temporary_name, destination)
    except BaseException:
        Path(temporary_name).unlink(missing_ok=True)
        raise
    return destination


def _file_sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def extract_hwpx_assay_history(archive_path: str | Path, output_path: str | Path | None = None, *, limits: ArchiveLimits = DEFAULT_LIMITS) -> dict[str, Any]:
    """Extract HWPX evidence without interpreting document formatting as assay facts."""
    contents, manifest = _checked_members(archive_path, limits)
    text_runs: list[dict[str, Any]] = []
    tables: list[dict[str, Any]] = []
    for member in sorted((name for name in contents if name.lower().endswith(".xml")), key=lambda value: (value.casefold(), value)):
        root = _parse_xml(contents[member], member)
        elements = list(root.iter())
        indices = {id(element): index for index, element in enumerate(elements)}
        for element in elements:
            if element.text:
                text_runs.append({"member": member, "source_index": indices[id(element)], "xml_element": _local_name(element.tag), "text": element.text, "text_location": "text"})
            if element.tail:
                text_runs.append({"member": member, "source_index": indices[id(element)], "xml_element": _local_name(element.tag), "text": element.tail, "text_location": "tail"})
        for table in (element for element in elements if _local_name(element.tag).casefold() == "table"):
            cells = []
            for cell in table.iter():
                if _local_name(cell.tag).casefold() in {"cell", "tc"}:
                    cells.append({"member": member, "source_index": indices[id(cell)], "text": _element_text(cell), "xml_element": _local_name(cell.tag)})
            tables.append({"member": member, "source_index": indices[id(table)], "xml_element": _local_name(table.tag), "cells": cells})
    images = [
        {"member": name, "sha256": hashlib.sha256(contents[name]).hexdigest(), "size": len(contents[name])}
        for name in sorted(contents, key=lambda value: (value.casefold(), value))
        if Path(name).suffix.casefold() in _IMAGE_EXTENSIONS
    ]
    payload: dict[str, Any] = {
        "archive_sha256": _file_sha256(archive_path),
        "embedded_images": images,
        "extraction_review_template": {
            "instruction": "Each field requires scientific confirmation from source evidence; formatting alone is not evidence.",
            "required_scientific_confirmation": [{"field": field, "status": "unconfirmed"} for field in _CRITICAL_ASSAY_FIELDS],
        },
        "format": "hwpx",
        "member_manifest": manifest,
        "privacy_classification": "confidential_source_evidence",
        "tables": tables,
        "text_runs": text_runs,
    }
    if output_path is not None:
        _atomic_json(output_path, payload)
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description="Safely extract HWPX evidence into deterministic JSON.")
    parser.add_argument("archive")
    parser.add_argument("output")
    args = parser.parse_args()
    extract_hwpx_assay_history(args.archive, args.output)


if __name__ == "__main__":
    main()
