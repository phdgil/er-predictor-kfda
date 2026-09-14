import sys
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

import pytest

from core.native_qa import (
    BATCH_RECOGNITION_CRITERIA,
    inspect_endpoint_geometry_parity,
    inspect_shared_example_parity,
)
from gui.erba_tab import ErbaTab, SharedExampleInput
from gui.main_window import MainWindow


PROJECT_ROOT = Path(__file__).resolve().parents[1]
pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="real Tk control geometry is a Windows release-surface contract",
)


@pytest.fixture
def real_tk_window(tmp_path, monkeypatch):
    output_root = tmp_path / "output"
    state_root = tmp_path / "state"
    user_profile = tmp_path / "user-profile"
    output_root.mkdir()
    state_root.mkdir()
    user_profile.mkdir()
    monkeypatch.setenv("USERPROFILE", str(user_profile))
    example_directory = tmp_path / "shared-example"
    example_directory.mkdir()
    example_workbook = example_directory / "test.xlsx"
    example_workbook.write_bytes(
        (PROJECT_ROOT / "templates" / "test.xlsx").read_bytes()
    )
    shared_example = SharedExampleInput(
        workbook=example_workbook,
        cas="6422-86-2",
        smiles="",
    )
    try:
        with patch(
            "gui.main_window.load_shared_example_input",
            return_value=shared_example,
        ), patch(
            "gui.erba_tab.load_shared_example_input",
            return_value=shared_example,
        ), patch.object(
            MainWindow,
            "load_defaults_on_startup",
        ), patch.object(
            ErbaTab,
            "load_defaults_on_startup",
        ):
            window = MainWindow(
                project_root=str(PROJECT_ROOT),
                install_root=str(PROJECT_ROOT),
                output_root=str(output_root),
                state_root=str(state_root),
            )
    except tk.TclError as error:
        pytest.skip(f"Windows Tk display is unavailable: {error}")
    window.state("normal")
    window.geometry("1180x800+20+20")
    window.update()
    try:
        yield window
    finally:
        try:
            window.destroy()
        except tk.TclError:
            pass


def test_real_endpoint_controls_have_exact_geometry_in_every_surface_state(
    real_tk_window,
):
    window = real_tk_window
    original = {
        "geometry": window.geometry(),
        "endpoint": window.notebook.select(),
        "erta_mode": window.parity_widgets["mode_notebook"].select(),
        "eralpha_mode": window.eralpha_tab.parity_widgets[
            "mode_notebook"
        ].select(),
        "erta_options": window.options_visible,
        "eralpha_options": window.eralpha_tab.options_visible,
    }

    receipt = inspect_endpoint_geometry_parity(window)

    assert receipt["coordinate_space"] == "relative to each selected endpoint_tab"
    assert len(receipt["compared_widget_keys"]) == 28
    assert receipt["compared_container_keys"] == [
        "batch_input_box",
        "batch_progress_box",
        "batch_result_box",
        "header_box",
        "single_input_box",
    ]
    recognition = receipt["batch_recognition_widgets"]
    assert recognition["criterion_id"] == "BRG-01-FEEDBACK-PLACEMENT"
    assert (
        recognition["criterion_id"]
        in BATCH_RECOGNITION_CRITERIA
    )
    assert recognition["passed"] is True
    assert recognition["exact_layout_match"] is True
    assert set(recognition["endpoints"]) == {"erta", "eralpha"}
    for endpoint, owner in (
        ("erta", window),
        ("eralpha", window.eralpha_tab),
    ):
        endpoint_receipt = recognition["endpoints"][endpoint]
        assert endpoint_receipt["passed"] is True
        assert endpoint_receipt["reading_order"] == [
            "run_batch",
            "aggregate_progress",
            "batch_result",
            "per_row_detail",
        ]
        bindings = endpoint_receipt["bindings"]
        assert bindings["run_callback"] == "batch_predict_clicked"
        assert "batch_predict_clicked" in bindings["run_command"]
        assert (
            bindings["progress_value_variable"]
            == str(owner.batch_progress_value)
        )
        assert (
            bindings["progress_text_variable"]
            == str(owner.batch_progress_var)
        )
        detail_var = getattr(owner, "status_var", None)
        if detail_var is None:
            detail_var = owner.batch_status_var
        assert (
            bindings["per_row_detail_variable"]
            == str(detail_var)
        )
        boxes = endpoint_receipt["boxes"]
        run_bottom = (
            boxes["run_batch"]["y"]
            + boxes["run_batch"]["height"]
        )
        progress_bottom = (
            boxes["aggregate_progress"]["y"]
            + boxes["aggregate_progress"]["height"]
        )
        result_bottom = (
            boxes["batch_result"]["y"]
            + boxes["batch_result"]["height"]
        )
        assert run_bottom <= boxes["aggregate_progress"]["y"]
        assert progress_bottom <= boxes["batch_result"]["y"]
        assert result_bottom <= boxes["per_row_detail"]["y"]
    assert (
        recognition["endpoints"]["erta"]["boxes"]
        == recognition["endpoints"]["eralpha"]["boxes"]
    )
    assert receipt["state_count"] == 8
    assert receipt["toggle_invocations"]["erta"] > 0
    assert receipt["toggle_invocations"]["eralpha"] > 0
    observed = set()
    sizes = {}
    for state in receipt["states"]:
        key = (
            state["window_geometry"],
            state["mode"],
            state["options_open"],
        )
        observed.add(key)
        sizes.setdefault(state["window_geometry"], tuple(state["actual_size"]))
        assert state["exact_match"] is True
        assert state["actual_size"] == state["requested_size"]
        assert state["controls"]["erta"] == state["controls"]["eralpha"]
        assert state["containers"]["erta"] == state["containers"]["eralpha"]
        for endpoint in ("erta", "eralpha"):
            settling = state["settling"][endpoint]
            assert settling["stable_samples"] == 2
            assert settling["root_size"] == state["requested_size"]
            assert settling["cycles"] <= 12
            assert settling["idle_reached"] is True
            assert {
                f"container:{name}"
                for name in receipt["compared_container_keys"]
            } <= settling["boxes"].keys()
        controls = state["controls"]["erta"]
        containers = state["containers"]["erta"]
        assert controls["options_frame"]["mapped"] is state["options_open"]
        assert controls["single_tab"]["mapped"] is (state["mode"] == "single")
        assert controls["batch_tab"]["mapped"] is (state["mode"] == "batch")
        assert controls["example_button"]["text"] == "Example input"
        assert "text" not in controls["cas_entry"]
        assert "text" not in controls["batch_status"]
        assert containers["single_input_box"]["mapped"] is (
            state["mode"] == "single"
        )
        assert containers["batch_input_box"]["mapped"] is (
            state["mode"] == "batch"
        )
    assert observed == {
        (size, mode, options_open)
        for size in ("initial", "resized")
        for mode in ("single", "batch")
        for options_open in (False, True)
    }
    assert sizes["initial"] != sizes["resized"]
    assert window.geometry() == original["geometry"]
    assert window.notebook.select() == original["endpoint"]
    assert window.parity_widgets["mode_notebook"].select() == original["erta_mode"]
    assert (
        window.eralpha_tab.parity_widgets["mode_notebook"].select()
        == original["eralpha_mode"]
    )
    assert window.options_visible is original["erta_options"]
    assert window.eralpha_tab.options_visible is original["eralpha_options"]


def test_real_example_buttons_restore_identical_preset_and_batch_workbook(
    real_tk_window,
    tmp_path,
):
    window = real_tk_window

    receipt = inspect_shared_example_parity(
        window,
        tmp_path / "fresh-shared-example-profile",
    )

    assert receipt["path_equal"] is True
    assert receipt["content_equal"] is True
    assert receipt["default_writable"] is True
    assert receipt["default_matches_bundled"] is True
    assert receipt["fresh_copy_writable"] is True
    assert receipt["fresh_copy_matches_bundled"] is True
    assert Path(receipt["fresh_copy"]) == (
        tmp_path
        / "fresh-shared-example-profile"
        / "Documents"
        / "ER_Predictor"
        / "Examples"
        / "test.xlsx"
    ).resolve()
    assert Path(receipt["fresh_copy"]).is_file()
    assert Path(receipt["workbook"]).name == "test.xlsx"
    assert Path(receipt["bundled_workbook"]).name == "test.xlsx"
    assert receipt["bundled_headers"] == ["CARSRN"]
    assert receipt["cas_header"] == "CARSRN"
    assert receipt["smiles_header"] is None
    assert receipt["bundled_row_count"] == 25
    assert receipt["workbook_row_count"] == 25
    assert receipt["smiles"] == ""
    assert receipt["cas"] == window.example_cas == window.eralpha_tab.example_cas
    assert (
        receipt["smiles"]
        == window.example_smiles
        == window.eralpha_tab.example_smiles
    )
    assert Path(receipt["workbook"]) == Path(window.batch_input_var.get()).resolve()
    assert Path(receipt["workbook"]) == Path(
        window.eralpha_tab.batch_input_var.get()
    ).resolve()
    assert (
        receipt["batch_display"]
        == window.batch_input_display_var.get()
        == window.eralpha_tab.batch_input_display_var.get()
    )
    assert (
        receipt["batch_destination"]
        == str(Path(receipt["workbook"]).parent)
        == window.batch_destination_var.get()
        == window.eralpha_tab.batch_destination_var.get()
    )
    assert receipt["workbook_size_bytes"] > 0
    assert len(receipt["workbook_sha256"]) == 64
    assert receipt["bundled_sha256"] == (
        "5a1f569f8f6a5cd47bff67a189645c3f9461fcf07bd24f4e5b4f83197f3350aa"
    )
    assert receipt["fresh_copy_sha256"] == receipt["bundled_sha256"]
    assert receipt["fresh_copy_size_bytes"] == receipt["bundled_size_bytes"]
    assert [row["endpoint"] for row in receipt["callbacks"]] == [
        "erta",
        "eralpha",
    ]
    assert all(row["other_endpoint_unchanged"] for row in receipt["callbacks"])
    assert all(
        row["restored_batch"]
        == [
            receipt["workbook"],
            "test.xlsx",
            str(Path(receipt["workbook"]).parent),
        ]
        for row in receipt["callbacks"]
    )
    assert (window.cas_var.get(), window.smiles_var.get()) == (
        receipt["cas"],
        receipt["smiles"],
    )
    assert (
        window.eralpha_tab.cas_var.get(),
        window.eralpha_tab.smiles_var.get(),
    ) == (receipt["cas"], receipt["smiles"])
