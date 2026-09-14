from copy import deepcopy

import pytest

from core.contracts import (
    ERBA_CLASSIFICATION_DIAGNOSTIC_COLUMNS,
    ERBA_CLASSIFICATION_METADATA_COLUMNS,
)
from core.native_qa import (
    BATCH_RECOGNITION_CRITERIA,
    BATCH_RECOGNITION_GATE_ID,
    BATCH_UNAVAILABLE_LABEL,
    ERALPHA_PRIMARY_FORBIDDEN_COLUMNS,
    ERALPHA_PRIMARY_TRUSTED_COLUMNS,
    ERALPHA_WORKBOOK_SHEET_ORDER,
    NativePackageQa,
    evaluate_batch_recognition_contract,
)


_NORMAL_CONTROLS = {
    "input": "normal",
    "template": "normal",
    "run": "normal",
}
_DISABLED_CONTROLS = {
    "input": "disabled",
    "template": "disabled",
    "run": "disabled",
}


def test_native_qa_locks_v3_primary_contract_and_gate_counts():
    assert ERALPHA_WORKBOOK_SHEET_ORDER == (
        "Predictions",
        "Guide",
        "Diagnostics",
        "Input",
        "Metadata",
    )
    assert ERALPHA_PRIMARY_TRUSTED_COLUMNS == (
        "CAS",
        "SMILES",
        "Canonical_SMILES",
        "Mol_valid",
        "Probability_Negative_0",
        "Probability_Positive_1",
        "Prediction",
        "Prediction_label",
        "AD",
        "AD_MeanDistance",
        "AD_DistanceThreshold",
        "AD_Distance_InDomain",
        "AD_SimilarityMax",
        "AD_SimilarityThreshold",
        "AD_Similarity_InDomain",
        "AD_PC1",
        "AD_PC2",
        "PubChem_CID",
        "PubChem_status",
    )
    assert ERBA_CLASSIFICATION_DIAGNOSTIC_COLUMNS == (
        "Row_ID",
        "CAS",
        "row_index",
        "Status_Code",
        "Status_Message",
        "SMILES_Provenance",
        "Result_Status",
        "Reason_Category",
        "Reason_Description",
        "Recommended_Action",
    )
    assert ERBA_CLASSIFICATION_METADATA_COLUMNS == (
        "Excel_Contract_ID",
        "Workflow",
        "Task",
        "Subtype",
        "Decision_rule",
        "Model_ID",
        "Model_SHA256",
        "Preprocessing_Policy_ID",
        "Protocol_SHA256",
        "Source_Manifest_SHA256",
        "Historical_Exposure_Manifest_SHA256",
        "Split_Manifest_SHA256",
        "Nested_CV_SHA256",
        "Internal_Resplit_SHA256",
        "Preprocessing_Parity_SHA256",
        "Report_SHA256",
        "Caveat_SHA256",
        "Performance_Evidence_Scope",
        "Evidence_Caveat",
    )
    assert len(ERALPHA_PRIMARY_TRUSTED_COLUMNS) == len(
        set(ERALPHA_PRIMARY_TRUSTED_COLUMNS)
    )
    assert set(ERALPHA_PRIMARY_TRUSTED_COLUMNS).isdisjoint(
        ERBA_CLASSIFICATION_METADATA_COLUMNS
    )
    assert set(ERALPHA_PRIMARY_TRUSTED_COLUMNS).isdisjoint(
        set(ERBA_CLASSIFICATION_DIAGNOSTIC_COLUMNS) - {"CAS"}
    )
    assert {
        "model_sha256",
        "evidence_caveat",
        "Model_SHA256",
        "Protocol_SHA256",
        "Evidence_Caveat",
        "Status_Code",
        "Status_Message",
        "Reason_Category",
    } <= set(ERALPHA_PRIMARY_FORBIDDEN_COLUMNS)
    assert len(NativePackageQa.BASELINE_CHECKS) == 22
    assert len(NativePackageQa.SHARED_EXAMPLE_BATCH_CHECKS) == 2
    assert len(NativePackageQa.BATCH_FEEDBACK_CHECKS) == 2
    assert len(
        NativePackageQa.BASELINE_CHECKS
        | NativePackageQa.SHARED_EXAMPLE_BATCH_CHECKS
        | NativePackageQa.BATCH_FEEDBACK_CHECKS
    ) == 26
    assert len(BATCH_RECOGNITION_CRITERIA) == 9


def _dialog(kind: str, title: str, message: str) -> dict:
    function, icon = {
        "info": ("showinfo", "blue-information"),
        "warning": ("showwarning", "yellow-warning"),
        "error": ("showerror", "red-error"),
    }[kind]
    return {
        "kind": kind,
        "function": function,
        "icon_semantic": icon,
        "title": title,
        "message": message,
        "batch_controls_at_dialog": {
            "erta": deepcopy(_NORMAL_CONTROLS),
            "eralpha": deepcopy(_NORMAL_CONTROLS),
        },
    }


def _widget_contract() -> dict:
    endpoint_receipt = {
        "criterion_id": "BRG-01-FEEDBACK-PLACEMENT",
        "passed": True,
        "reading_order": [
            "run_batch",
            "aggregate_progress",
            "batch_result",
            "per_row_detail",
        ],
        "bindings": {
            "run_command": "123batch_predict_clicked",
            "run_callback": "batch_predict_clicked",
            "progress_value_variable": "PY_VAR_PROGRESS_VALUE",
            "progress_text_variable": "PY_VAR_PROGRESS_TEXT",
            "per_row_detail_variable": "PY_VAR_DETAIL_TEXT",
        },
    }
    return {
        "criterion_id": "BRG-01-FEEDBACK-PLACEMENT",
        "passed": True,
        "exact_layout_match": True,
        "endpoints": {
            "erta": deepcopy(endpoint_receipt),
            "eralpha": deepcopy(endpoint_receipt),
        },
    }


def _unsupported_own_window_capture() -> dict:
    return {
        "supported": False,
        "method": "PrintWindow(PW_RENDERFULLCONTENT)",
        "reason": "own-window capture is unavailable in this unit-test process",
        "capture_scope": "own Tk HWND only",
    }


def _publication_record(
    endpoint: str,
    phase: str,
    *,
    total: int,
    unavailable: int,
    screenshot: bool = False,
) -> dict:
    output = f"C:/qa/{endpoint}-{phase}.xlsx"
    if unavailable == 0:
        availability = "All rows were predicted."
    elif unavailable == total:
        availability = (
            "No rows could be predicted. "
            "Row-level reasons are saved in the workbook."
        )
    else:
        availability = (
            f"{unavailable} row(s) could not be predicted. "
            "Row-level reasons are saved in the workbook."
        )
    message = "\n".join(
        (
            f"Saved:\n{output}",
            "",
            f"Total rows: {total}",
            f"{BATCH_UNAVAILABLE_LABEL}: {unavailable}",
            "",
            availability,
        )
    )
    result = "\n".join(
        (
            "Batch job completed and workbook saved.",
            "",
            f"Total rows: {total}",
            f"{BATCH_UNAVAILABLE_LABEL}: {unavailable}",
            "",
            availability,
            "",
            f"Output workbook: {output}",
        )
    )
    record = {
        "endpoint": endpoint,
        "phase": phase,
        "scenario": phase,
        "progress": f"100% - {total}/{total} - Completed",
        "status": f"Batch prediction completed: {output}",
        "result": result,
        "controls": deepcopy(_NORMAL_CONTROLS),
        "dialog": _dialog("info", "Batch prediction done", message),
        "output": output,
        "output_exists": True,
        "fresh_output_count": 1,
        "total_count": total,
        "available_count": total - unavailable,
        "unavailable_count": unavailable,
        "workbook_unavailable_count": unavailable,
    }
    if screenshot:
        record["screenshot"] = _unsupported_own_window_capture()
    return record


def _records() -> list[dict]:
    records = []
    endpoint_counts = {
        "erta": {
            "success": (3, 0),
            "partial_unavailable": (4, 1),
            "all_unavailable": (2, 2),
        },
        "eralpha": {
            "success": (5, 0),
            "partial_unavailable": (5, 2),
            "all_unavailable": (3, 3),
        },
    }
    for endpoint, counts in endpoint_counts.items():
        records.append(
            {
                "endpoint": endpoint,
                "phase": "new_run",
                "scenario": "partial_unavailable",
                "progress": "0% - 0/0 - Reading input workbook",
                "status": "Batch prediction started.",
                "result": (
                    "Batch prediction is running.\n\n"
                    "Completion details will appear after the workbook is saved."
                ),
                "controls": deepcopy(_DISABLED_CONTROLS),
                "dialog": None,
                "prior_progress": "100% - 3/3 - Completed",
                "prior_status": (
                    "Batch prediction completed: C:/qa/prior.xlsx"
                ),
                "prior_result": (
                    "Batch job completed and workbook saved.\n\n"
                    "Output workbook: C:/qa/prior.xlsx"
                ),
                "screenshot": _unsupported_own_window_capture(),
            }
        )
        records.append(
            {
                "endpoint": endpoint,
                "phase": "resolving",
                "scenario": "partial_unavailable",
                "transitions": [
                    {
                        "sequence": 1,
                        "channel": "progress",
                        "text": "30% - 1/4 - Resolving CAS/SMILES",
                    },
                    {
                        "sequence": 2,
                        "channel": "detail",
                        "text": (
                            "Fetching SMILES from PubChem: "
                            "1 / 4 (50-00-0)"
                        ),
                    },
                ],
            }
        )
        for phase, (total, unavailable) in counts.items():
            records.append(
                _publication_record(
                    endpoint,
                    phase,
                    total=total,
                    unavailable=unavailable,
                    screenshot=phase == "success",
                )
            )
        records.append(
            {
                "endpoint": endpoint,
                "phase": "no_output_failure",
                "scenario": "no_output_failure",
                "progress": "100% - 0/0 - Failed",
                "status": "Batch prediction failed: write failed",
                "result": (
                    "Batch prediction failed.\n\n"
                    "Technical details: write failed"
                ),
                "controls": deepcopy(_NORMAL_CONTROLS),
                "dialog": _dialog(
                    "error",
                    "Batch prediction failed",
                    "write failed",
                ),
                "output": "",
                "output_exists": False,
                "fresh_output_count": 0,
                "total_count": 0,
                "available_count": 0,
                "unavailable_count": 0,
                "workbook_unavailable_count": 0,
            }
        )
    return records


def test_gate_accepts_targeted_endpoint_and_prediction_count_differences():
    receipt = evaluate_batch_recognition_contract(
        _records(),
        _widget_contract(),
    )

    assert receipt["gate_id"] == BATCH_RECOGNITION_GATE_ID
    assert receipt["passed"] is True
    assert receipt["criteria_ids"] == list(BATCH_RECOGNITION_CRITERIA)
    assert all(item["passed"] for item in receipt["criteria"])
    assert receipt["normalization"] == {
        "normalized_fields": [
            "endpoint display label",
            "endpoint-specific prediction-label counts",
        ],
        "whole_messages_normalized": False,
        "raw_displayed_strings_retained": True,
    }
    assert (
        receipt["comparison"]["partial_unavailable"]["erta"]
        ["unavailable_count"]
        == 1
    )
    assert (
        receipt["comparison"]["partial_unavailable"]["eralpha"]
        ["unavailable_count"]
        == 2
    )
    for phase in ("success", "partial_unavailable", "all_unavailable"):
        assert {
            (
                endpoint_receipt["dialog_function"],
                endpoint_receipt["icon_semantic"],
            )
            for endpoint_receipt in receipt["comparison"][phase].values()
        } == {("showinfo", "blue-information")}
    assert {
        (
            endpoint_receipt["dialog_function"],
            endpoint_receipt["icon_semantic"],
        )
        for endpoint_receipt in receipt["comparison"][
            "no_output_failure"
        ].values()
    } == {("showerror", "red-error")}
    assert (
        receipt["validation_scope"]
        ["automated_recognition_contract_checks"]
        is True
    )
    assert (
        receipt["validation_scope"]
        ["real_human_usability_study_performed"]
        is False
    )


def test_gate_rejects_stale_completed_and_yellow_success_semantics():
    stale = _records()
    stale_run = next(
        record
        for record in stale
        if record["endpoint"] == "erta"
        and record["phase"] == "new_run"
    )
    stale_run["progress"] = "100% - 4/4 - Completed"
    with pytest.raises(AssertionError, match="retained terminal success"):
        evaluate_batch_recognition_contract(stale, _widget_contract())

    warned = _records()
    partial = next(
        record
        for record in warned
        if record["endpoint"] == "eralpha"
        and record["phase"] == "partial_unavailable"
    )
    partial["dialog"] = _dialog(
        "warning",
        "Batch prediction completed with warnings",
        partial["dialog"]["message"],
    )
    with pytest.raises(AssertionError, match="successful publication dialog"):
        evaluate_batch_recognition_contract(warned, _widget_contract())


def test_gate_rejects_hidden_unavailable_rows_and_false_failure_outputs():
    hidden = _records()
    all_unavailable = next(
        record
        for record in hidden
        if record["endpoint"] == "erta"
        and record["phase"] == "all_unavailable"
    )
    all_unavailable["workbook_unavailable_count"] -= 1
    with pytest.raises(AssertionError, match="differs from workbook count"):
        evaluate_batch_recognition_contract(hidden, _widget_contract())

    false_failure = _records()
    failure = next(
        record
        for record in false_failure
        if record["endpoint"] == "eralpha"
        and record["phase"] == "no_output_failure"
    )
    failure["output"] = "C:/qa/should-not-exist.xlsx"
    failure["output_exists"] = True
    failure["fresh_output_count"] = 1
    with pytest.raises(AssertionError, match="created or claimed an output"):
        evaluate_batch_recognition_contract(
            false_failure,
            _widget_contract(),
        )
