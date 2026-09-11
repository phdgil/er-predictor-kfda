"""Static contracts for the isolated, reproducible ER_Predictor package."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_spec_targets_only_the_new_product_name_and_erba_catalog():
    spec = (PROJECT_ROOT / "app.spec").read_text(encoding="utf-8")
    assert 'name="ER_Predictor"' in spec
    assert "name='ERTA_Predictor'" not in spec
    assert 'ERBA_CATALOG = ERBA_ROOT / "catalog.v2.json"' in spec
    assert '(str(ERBA_CATALOG), "models/erba")' in spec
    assert "catalogued_erba_datas()" in spec
    assert "legacy_model_datas()" in spec
    assert "ERBA_ROOT not in path.parents" in spec
    assert "TRUSTED_ERBA_ARTIFACTS" in spec
    assert "release_manifest_sha256" in spec
    assert 'catalog.get("schema_version") != 2' in spec


def test_package_pipeline_requires_catalog_v2_schema():
    for path in ("build_exe.bat", "build_reproducibility_manifest.py", "publish_exe.py"):
        source = (PROJECT_ROOT / path).read_text(encoding="utf-8")
        assert "catalog.v2.json" in source
        assert "schema_version" in source



def test_spec_declares_required_hidden_imports_and_native_package_collection():
    spec = (PROJECT_ROOT / "app.spec").read_text(encoding="utf-8")
    for required in (
        '"core.erba_features"',
        '"joblib"',
        '"xgboost"',
        '"rdkit.Avalon.pyAvalonTools"',
        '"tensorflow.keras"',
        '"openpyxl"',
        '"requests"',
        '"PIL"',
        '"scipy"',
        '"h5py"',
        "collect_dynamic_libs(package)",
        "collect_data_files(package)",
    ):
        assert required in spec


def test_build_uses_two_independent_offline_hash_locked_venvs():
    script = (PROJECT_ROOT / "build_exe.bat").read_text(encoding="utf-8")
    assert "ER_Predictor-build-first.venv" in script
    assert "ER_Predictor-build-second.venv" in script
    assert 'call :create_venv "%FIRST_VENV%"' in script
    assert 'call :create_venv "%SECOND_VENV%"' in script
    assert "sys.version_info[:2] == (3, 10)" in script
    assert "sys.maxsize > 2**32" in script
    assert "sys.prefix != sys.base_prefix" in script
    assert "PYTHONNOUSERSITE=1" in script
    assert "PIP_REQUIRE_VIRTUALENV=true" in script
    assert "PIP_NO_INDEX=1" in script
    assert "--no-index" in script
    assert "--require-hashes" in script
    assert "--trusted-host" not in script
    assert "SOURCE_DATE_EPOCH" in script
    assert "PYTHONHASHSEED=0" in script
    assert "dist\\ER_Predictor\\ER_Predictor.exe" in script


def test_build_has_source_distribution_and_reproducibility_audits():
    script = (PROJECT_ROOT / "build_exe.bat").read_text(encoding="utf-8")
    for required in (
        "prepare_build_wheelhouse.py",
        "ERBA source artifact audit passed",
        "catalogued ERBA artifact integrity failure",
        "ER_Predictor distribution artifact audit passed",
        "unlisted or missing ERBA joblib",
        "ER_Predictor-build-manifest.json",
        "ER_Predictor-wheelhouse-inventory.json",
        "ER_Predictor-resolved-requirements.txt",
        "ER_Predictor-reproducibility-manifest.first.json",
        "ER_Predictor-reproducibility-manifest.second.json",
        "ER_Predictor-clean-build-comparison.json",
        "independent clean-build normalized inventories differ",
    ):
        assert required in script


def test_generated_lock_is_complete_hash_pinned_for_every_wheel():
    lock = (PROJECT_ROOT / "requirements-lock.txt").read_text(encoding="utf-8")
    entries = [line for line in lock.splitlines() if line and not line.startswith("#")]
    assert entries, "wheelhouse preparation must generate the release lock before testing"
    pattern = re.compile(r"^[A-Za-z0-9_.-]+==[^ ]+ --hash=sha256:[0-9a-f]{64}$")
    assert all(pattern.fullmatch(line) for line in entries)
    lowered = "\n".join(entries).lower()
    for distribution in (
        "tensorflow==",
        "scikit-learn==",
        "joblib==",
        "xgboost==",
        "rdkit==",
        "pyinstaller==",
        "pytest==",
        "pywinauto==",
        "pip==",
        "setuptools==",
    ):
        assert distribution in lowered


def test_wheelhouse_helper_binds_lock_wheels_and_exact_toolchain():
    helper = (PROJECT_ROOT / "prepare_build_wheelhouse.py").read_text(encoding="utf-8")
    for required in (
        "CPython 3.10 x64",
        "executable_sha256",
        "requirements_input_sha256",
        "--only-binary=:all:",
        "--require-hashes",
        "--dry-run",
        "lock and wheelhouse distribution sets differ",
    ):
        assert required in helper


def test_publication_is_fixed_atomic_and_baseline_gated():
    batch = (PROJECT_ROOT / "publish_exe.bat").read_text(encoding="utf-8")
    helper = (PROJECT_ROOT / "publish_exe.py").read_text(encoding="utf-8")
    assert "accepts no destination or positional arguments" in batch
    for required in (
        "D:/research/FDA_endocrine_disruption/ER_Predictor",
        'BASELINE_ORIGINAL_OLD = Path(r"D:/research/FDA_endocrine_disruption/ERTA_Predictor")',
        "D:/research/FDA_endocrine_disruption/_archive/legacy_apps/ERTA_Predictor",
        'ROLLBACK_ROOT = PUBLISH_ROOT / "_archive"',
        "backup = ROLLBACK_ROOT / (",
        "prechange_manifest.json",
        "existing_package_files",
        "relative_to(approved_root)",
        "old_before == old_after == old_baseline",
        "source distribution differs from the audited build manifest",
        "staged distribution differs from the build manifest",
        "published distribution differs from the build manifest",
        "immutable ERTA tree differs from the approved pre-change baseline",
        "os.replace",
        "ER_Predictor-publication-receipt.json",
    ):
        assert required in helper


def test_docs_state_product_routes_writable_roots_and_evidence_caveat():
    readme = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")
    deploy = (PROJECT_ROOT / "README_DEPLOY_KR.md").read_text(encoding="utf-8")
    for document in (readme, deploy):
        assert "ER_Predictor" in document
        assert "ERalpha" in document
        assert "Guide" in document
        assert "historically exposed" in document or "과거에 노출된" in document
        assert "regulatory" in document or "규제" in document
        assert "ERTA_Predictor" in document
        assert "prepare_build_wheelhouse.py prepare" in document
        assert "두" in document or "two" in document.lower()
        assert "ERβ" not in document or "제외" in document
    assert "ER_Predictor_UserData" in readme
    assert "ERTA_ERBA_example.xlsx" in readme


def test_combined_template_preserves_separate_erta_and_erba_input_sheets():
    import zipfile

    template = PROJECT_ROOT / "templates" / "ERTA_ERBA_example.xlsx"
    assert template.is_file()
    with zipfile.ZipFile(template) as workbook:
        workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        first_sheet = workbook_xml.index('name="ERBA_Input"')
        second_sheet = workbook_xml.index('name="ERTA_Input"')
        worksheet_text = "\n".join(
            workbook.read(name).decode("utf-8")
            for name in workbook.namelist()
            if name.startswith("xl/worksheets/")
        )
    assert first_sheet < second_sheet
    for header in ("Row_ID", "CAS", "SMILES", "CID", "label"):
        assert header in worksheet_text

def test_external_runner_rejects_stale_receipt_and_binds_nonce_and_pid(tmp_path):
    from qa.external_qa_runner import wait_for_receipt

    receipt = tmp_path / "internal-receipt.json"
    receipt.write_text(
        json.dumps(
            {
                "run_nonce": "stale",
                "target_pid": 123,
                "passed": True,
                "checks": [{"check": "stale", "passed": True}],
            }
        ),
        encoding="utf-8",
    )
    with pytest.raises(AssertionError, match="nonce mismatch"):
        wait_for_receipt(receipt, 0.1, "current", 123)

    receipt.write_text(
        json.dumps(
            {
                "run_nonce": "current",
                "target_pid": 456,
                "passed": True,
                "checks": [{"check": "fresh", "passed": True}],
            }
        ),
        encoding="utf-8",
    )
    payload = wait_for_receipt(receipt, 0.1, "current", 456)
    assert payload["checks"][0]["check"] == "fresh"


def test_external_runner_requires_fresh_root_and_exact_round_trip():
    source = (PROJECT_ROOT / "qa" / "external_qa_runner.py").read_text(encoding="utf-8")
    assert "output root must not already exist" in source
    assert 'after["sha256"] != before["sha256"]' in source
    assert '"run_nonce": run_nonce' in source
    assert '"target_pid": process.pid' in source
    assert '"terminated_by_runner"' in source
