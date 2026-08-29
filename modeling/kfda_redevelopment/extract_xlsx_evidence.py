"""Read-only, deterministic XLSX evidence extraction; formulas are never evaluated."""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

from .extract_hwpx_assay_history import (
    ArchiveLimits,
    DEFAULT_LIMITS,
    UnsafeArchiveError,
    _atomic_json,
    _checked_members,
    _file_sha256,
    _local_name,
    _parse_xml,
)

_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_DDE_PATTERN = re.compile(r"\b(?:DDE|WEBSERVICE|HYPERLINK)\s*\(", re.IGNORECASE)
_RED_VALUES = {"FF0000", "FFFF0000", "00FF0000"}


def _relationship_map(contents: dict[str, bytes], member: str) -> dict[str, str]:
    if member not in contents:
        return {}
    root = _parse_xml(contents[member], member)
    result: dict[str, str] = {}
    for relationship in root.iter():
        if _local_name(relationship.tag) == "Relationship":
            identifier = relationship.attrib.get("Id")
            target = relationship.attrib.get("Target")
            if identifier and target:
                result[identifier] = target
    return result


def _shared_strings(contents: dict[str, bytes]) -> list[str]:
    member = "xl/sharedStrings.xml"
    if member not in contents:
        return []
    root = _parse_xml(contents[member], member)
    return ["".join(item.itertext()) for item in root.iter() if _local_name(item.tag) == "si"]


def _red_font_styles(contents: dict[str, bytes]) -> tuple[set[str], list[dict[str, Any]]]:
    member = "xl/styles.xml"
    if member not in contents:
        return set(), []
    root = _parse_xml(contents[member], member)
    fonts = [element for element in root.iter() if _local_name(element.tag) == "font"]
    red_font_ids: set[int] = set()
    for index, font in enumerate(fonts):
        for child in font:
            if _local_name(child.tag) != "color":
                continue
            rgb = child.attrib.get("rgb", "").upper()
            if rgb in _RED_VALUES:
                red_font_ids.add(index)
    cell_xfs = next((element for element in root.iter() if _local_name(element.tag) == "cellXfs"), None)
    red_styles: set[str] = set()
    if cell_xfs is not None:
        for index, xf in enumerate(child for child in cell_xfs if _local_name(child.tag) == "xf"):
            try:
                font_id = int(xf.attrib.get("fontId", "0"))
            except ValueError as error:
                raise UnsafeArchiveError("style has an invalid font identifier") from error
            if font_id in red_font_ids:
                red_styles.add(str(index))
    return red_styles, [{"font_id": index} for index in sorted(red_font_ids)]


def _worksheet_member(target: str) -> str:
    # Workbook relationships target paths relative to xl/.
    if target.startswith("/") or "\\" in target or any(part in ("", ".", "..") for part in target.split("/")):
        raise UnsafeArchiveError("workbook relationship has an unsafe internal target")
    return f"xl/{target}" if not target.startswith("xl/") else target


def extract_xlsx_evidence(archive_path: str | Path, output_path: str | Path | None = None, *, limits: ArchiveLimits = DEFAULT_LIMITS) -> dict[str, Any]:
    """Extract workbook source values and metadata without opening or calculating it."""
    contents, manifest = _checked_members(archive_path, limits)
    workbook_member = "xl/workbook.xml"
    if workbook_member not in contents:
        raise UnsafeArchiveError("archive is not an XLSX workbook")
    workbook = _parse_xml(contents[workbook_member], workbook_member)
    rels = _relationship_map(contents, "xl/_rels/workbook.xml.rels")
    shared_strings = _shared_strings(contents)
    red_styles, red_fonts = _red_font_styles(contents)
    sheets: list[dict[str, Any]] = []
    formula_findings: list[dict[str, Any]] = []
    for sheet in (element for element in workbook.iter() if _local_name(element.tag) == "sheet"):
        relationship_id = sheet.attrib.get(f"{_REL_NS}id")
        target = rels.get(relationship_id or "")
        if not target:
            raise UnsafeArchiveError("worksheet relationship is missing")
        member = _worksheet_member(target)
        if member not in contents:
            raise UnsafeArchiveError("worksheet relationship points outside the archive")
        root = _parse_xml(contents[member], member)
        merged_ranges = sorted(
            (element.attrib["ref"] for element in root.iter() if _local_name(element.tag) == "mergeCell" and "ref" in element.attrib),
            key=str.casefold,
        )
        cells: list[dict[str, Any]] = []
        for cell in (element for element in root.iter() if _local_name(element.tag) == "c"):
            coordinate = cell.attrib.get("r")
            if not coordinate:
                raise UnsafeArchiveError("worksheet cell has no coordinate")
            data_type = cell.attrib.get("t", "n")
            style_id = cell.attrib.get("s")
            formula_element = next((child for child in cell if _local_name(child.tag) == "f"), None)
            value_element = next((child for child in cell if _local_name(child.tag) == "v"), None)
            inline_element = next((child for child in cell if _local_name(child.tag) == "is"), None)
            raw_value = value_element.text if value_element is not None else None
            inline_text = "".join(inline_element.itertext()) if inline_element is not None else None
            shared_string = None
            if data_type == "s" and raw_value is not None:
                try:
                    shared_string = shared_strings[int(raw_value)]
                except (IndexError, ValueError):
                    raise UnsafeArchiveError("shared-string cell has an invalid index") from None
            formula = formula_element.text if formula_element is not None else None
            entry = {
                "coordinate": coordinate,
                "data_type": data_type,
                "formula": formula,
                "inline_text": inline_text,
                "raw_value": raw_value,
                "shared_string": shared_string,
                "style_id": style_id,
            }
            cells.append(entry)
            if formula is not None:
                formula_findings.append({
                    "coordinate": coordinate,
                    "formula": formula,
                    "member": member,
                    "potential_dde_or_link": bool(_DDE_PATTERN.search(formula)),
                })
        cells.sort(key=lambda value: value["coordinate"])
        sheets.append({
            "member": member,
            "merged_ranges": merged_ranges,
            "name": sheet.attrib.get("name", ""),
            "sheet_id": sheet.attrib.get("sheetId"),
            "cells": cells,
        })
    sheets.sort(key=lambda value: (value["name"].casefold(), value["name"]))
    macro_members = sorted((name for name in contents if name.casefold().endswith("vbaproject.bin")), key=str.casefold)
    external_link_members = sorted((name for name in contents if name.casefold().startswith("xl/externallinks/")), key=str.casefold)
    relationship_findings = []
    for name in sorted((item for item in contents if item.casefold().endswith(".rels")), key=str.casefold):
        root = _parse_xml(contents[name], name)
        relationship_findings.extend(
            {"id": element.attrib.get("Id"), "member": name, "target": element.attrib.get("Target"), "type": element.attrib.get("Type")}
            for element in root.iter() if _local_name(element.tag) == "Relationship"
        )
    payload: dict[str, Any] = {
        "archive_sha256": _file_sha256(archive_path),
        "export_safety": {
            "csv_xlsx": {
                "leading_formula_characters": ["=", "+", "-", "@"],
                "required_action": "Later CSV/XLSX exporters must escape values beginning with these characters.",
            }
        },
        "external_link_members": external_link_members,
        "format": "xlsx",
        "formula_findings": formula_findings,
        "macro_members": macro_members,
        "member_manifest": manifest,
        "privacy_classification": "confidential_source_evidence",
        "red_font_annotations": {"red_fonts": red_fonts, "red_style_ids": sorted(red_styles), "cells": [
            {"coordinate": cell["coordinate"], "sheet": sheet["name"], "style_id": cell["style_id"]}
            for sheet in sheets for cell in sheet["cells"] if cell["style_id"] in red_styles
        ]},
        "relationship_findings": sorted(relationship_findings, key=lambda value: (value["member"], value["id"] or "")),
        "sheets": sheets,
    }
    if output_path is not None:
        _atomic_json(output_path, payload)
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description="Safely extract XLSX evidence into deterministic JSON.")
    parser.add_argument("archive")
    parser.add_argument("output")
    args = parser.parse_args()
    extract_xlsx_evidence(args.archive, args.output)


if __name__ == "__main__":
    main()
