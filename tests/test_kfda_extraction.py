from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from modeling.kfda_redevelopment.extract_hwpx_assay_history import ArchiveLimits, UnsafeArchiveError, extract_hwpx_assay_history
from modeling.kfda_redevelopment.extract_xlsx_evidence import extract_xlsx_evidence


def _zip(path: Path, members: dict[str, str], *, compression: int = ZIP_DEFLATED) -> None:
    with ZipFile(path, "w", compression=compression) as archive:
        for name, value in members.items():
            archive.writestr(name, value)


def _xlsx_members(*, formula: str = "SUM(A1:A1)") -> dict[str, str]:
    return {
        "[Content_Types].xml": "<Types/>",
        "xl/workbook.xml": (
            '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Evidence" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>'
        ),
        "xl/styles.xml": (
            '<styleSheet><fonts count="1"><font><color rgb="FFFF0000"/></font></fonts>'
            '<cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>'
        ),
        "xl/worksheets/sheet1.xml": (
            '<worksheet><mergeCells><mergeCell ref="A1:B1"/></mergeCells><sheetData><row r="1">'
            '<c r="A1" t="str" s="0"><f>' + formula + '</f><v>not-calculated</v></c>'
            '</row></sheetData></worksheet>'
        ),
    }


def test_hwpx_extracts_source_anchors_images_and_deterministic_json(tmp_path: Path) -> None:
    source = tmp_path / "source.hwpx"
    _zip(source, {
        "Contents/content.xml": "<document><p>assay source text</p><table><cell>cell evidence</cell></table></document>",
        "BinData/image.png": "not an executed image",
    })
    first = tmp_path / "first.json"
    second = tmp_path / "second.json"

    result = extract_hwpx_assay_history(source, first)
    extract_hwpx_assay_history(source, second)

    assert first.read_bytes() == second.read_bytes()
    assert result["text_runs"][0]["member"] == "Contents/content.xml"
    assert result["tables"][0]["cells"][0]["text"] == "cell evidence"
    assert result["embedded_images"][0]["member"] == "BinData/image.png"
    assert all(item["status"] == "unconfirmed" for item in result["extraction_review_template"]["required_scientific_confirmation"])


def test_xlsx_preserves_formula_without_evaluation_and_records_red_annotation(tmp_path: Path) -> None:
    source = tmp_path / "evidence.xlsx"
    _zip(source, _xlsx_members())

    result = extract_xlsx_evidence(source)

    cell = result["sheets"][0]["cells"][0]
    assert cell["formula"] == "SUM(A1:A1)"
    assert cell["raw_value"] == "not-calculated"
    assert result["red_font_annotations"]["cells"] == [{"coordinate": "A1", "sheet": "Evidence", "style_id": "0"}]
    assert result["export_safety"]["csv_xlsx"]["leading_formula_characters"] == ["=", "+", "-", "@"]


def test_xlsx_flags_dde_formula_without_executing_it(tmp_path: Path) -> None:
    source = tmp_path / "dde.xlsx"
    _zip(source, _xlsx_members(formula='DDE("cmd", "calc")'))

    result = extract_xlsx_evidence(source)

    assert result["formula_findings"] == [{
        "coordinate": "A1", "formula": 'DDE("cmd", "calc")', "member": "xl/worksheets/sheet1.xml", "potential_dde_or_link": True,
    }]


@pytest.mark.parametrize("member", ["../escape.xml", "/absolute.xml", "C:/drive.xml"])
def test_unsafe_member_paths_fail_closed(tmp_path: Path, member: str) -> None:
    source = tmp_path / "unsafe.zip"
    _zip(source, {member: "payload"})

    with pytest.raises(UnsafeArchiveError):
        extract_hwpx_assay_history(source)


def test_zip_bomb_limit_and_external_relationship_fail_closed(tmp_path: Path) -> None:
    oversized = tmp_path / "oversized.hwpx"
    _zip(oversized, {"Contents/content.xml": "x" * 64})
    with pytest.raises(UnsafeArchiveError):
        extract_hwpx_assay_history(oversized, limits=ArchiveLimits(max_member_uncompressed=32))

    external = tmp_path / "external.hwpx"
    _zip(external, {
        "Contents/content.xml": "<document/>",
        "_rels/document.xml.rels": '<Relationships><Relationship Id="rId1" Target="https://example.invalid" TargetMode="External"/></Relationships>',
    })
    with pytest.raises(UnsafeArchiveError, match="external relationships"):
        extract_hwpx_assay_history(external)


def test_json_output_is_utf8_and_machine_readable(tmp_path: Path) -> None:
    source = tmp_path / "source.hwpx"
    output = tmp_path / "evidence.json"
    _zip(source, {"Contents/content.xml": "<document><p>한글</p></document>"})

    extract_hwpx_assay_history(source, output)

    assert json.loads(output.read_text(encoding="utf-8"))["format"] == "hwpx"
