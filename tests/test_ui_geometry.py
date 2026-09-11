import sys
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

import pytest

from core.native_qa import (
    inspect_endpoint_geometry_parity,
    inspect_shared_example_parity,
)
from gui.erba_tab import ErbaTab
from gui.main_window import MainWindow


PROJECT_ROOT = Path(__file__).resolve().parents[1]
pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="real Tk control geometry is a Windows release-surface contract",
)


@pytest.fixture
def real_tk_window(tmp_path):
    output_root = tmp_path / "output"
    state_root = tmp_path / "state"
    output_root.mkdir()
    state_root.mkdir()
    try:
        with patch.object(MainWindow, "load_defaults_on_startup"), patch.object(
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
):
    window = real_tk_window

    receipt = inspect_shared_example_parity(window)

    assert receipt["path_equal"] is True
    assert receipt["content_equal"] is True
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
    assert receipt["workbook_size_bytes"] > 0
    assert len(receipt["workbook_sha256"]) == 64
    assert [row["endpoint"] for row in receipt["callbacks"]] == [
        "erta",
        "eralpha",
    ]
    assert all(row["other_endpoint_unchanged"] for row in receipt["callbacks"])
    assert (window.cas_var.get(), window.smiles_var.get()) == (
        receipt["cas"],
        receipt["smiles"],
    )
    assert (
        window.eralpha_tab.cas_var.get(),
        window.eralpha_tab.smiles_var.get(),
    ) == (receipt["cas"], receipt["smiles"])
