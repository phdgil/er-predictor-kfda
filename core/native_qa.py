"""Opt-in native-package acceptance driver used only by release verification."""
from __future__ import annotations

from copy import copy
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import time
import traceback

from openpyxl import Workbook, load_workbook

from core.contracts import (
    ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID,
    ERBASubtype,
    ERBATask,
    HISTORICAL_EXPOSURE_CAVEAT_COMPACT,
)

DISTRIBUTED_TEST_XLSX_SHA256 = (
    "5a1f569f8f6a5cd47bff67a189645c3f9461fcf07bd24f4e5b4f83197f3350aa"
)


def _sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _normalized_example_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().casefold())


def _read_shared_example_workbook(path: Path) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        normalized = {
            _normalized_example_header(value): index
            for index, value in enumerate(headers)
        }
        cas_index = next(
            (
                normalized[alias]
                for alias in (
                    "cas",
                    "casno",
                    "casrn",
                    "carsrn",
                    "casnumber",
                )
                if alias in normalized
            ),
            None,
        )
        if cas_index is None:
            raise AssertionError(
                f"shared example workbook lacks a recognized CAS column: {headers}"
            )
        smiles_index = next(
            (
                normalized[alias]
                for alias in (
                    "smiles",
                    "canonicalsmiles",
                    "isomericsmiles",
                )
                if alias in normalized
            ),
            None,
        )
        values = list(sheet.iter_rows(min_row=2, values_only=True))
        cas_values = [
            str(row[cas_index] or "").strip()
            for row in values
        ]
        smiles_values = [
            (
                str(row[smiles_index] or "").strip()
                if smiles_index is not None
                else ""
            )
            for row in values
        ]
        return {
            "sheet": sheet.title,
            "headers": headers,
            "row_count": len(values),
            "cas_header": headers[cas_index],
            "cas_values": cas_values,
            "smiles_header": (
                headers[smiles_index] if smiles_index is not None else None
            ),
            "smiles_values": smiles_values,
        }
    finally:
        workbook.close()


def _openpyxl_cell_kind(cell) -> str:
    if cell.row == 1 and cell.value is not None:
        return "header"
    if cell.value is None:
        return "blank"
    if isinstance(cell.value, bool):
        return "boolean"
    if cell.is_date:
        return "date"
    if isinstance(cell.value, (int, float)):
        return "number"
    return "text"


def _openpyxl_format_signature(cell) -> dict:
    """Return the complete value-independent openpyxl cell-format signature."""
    return {
        "font": repr(copy(cell.font)),
        "fill": repr(copy(cell.fill)),
        "border": repr(copy(cell.border)),
        "alignment": repr(copy(cell.alignment)),
        "number_format": str(cell.number_format),
        "protection": repr(copy(cell.protection)),
        "quote_prefix": bool(cell.quotePrefix),
        "pivot_button": bool(cell.pivotButton),
    }


def _worksheet_format_inventory(sheet) -> dict:
    inventories: dict[str, dict[str, dict]] = {}
    counts: dict[str, int] = {}
    for row in sheet.iter_rows():
        for cell in row:
            kind = _openpyxl_cell_kind(cell)
            signature = _openpyxl_format_signature(cell)
            canonical = json.dumps(signature, sort_keys=True)
            inventories.setdefault(kind, {}).setdefault(
                canonical,
                {
                    "signature": signature,
                    "example_coordinate": cell.coordinate,
                },
            )
            counts[kind] = counts.get(kind, 0) + 1
    return {
        kind: {
            "cell_count": counts[kind],
            "signatures": [
                receipt
                for _, receipt in sorted(inventories[kind].items())
            ],
        }
        for kind in sorted(inventories)
    }


def compare_plain_primary_workbook_formatting(
    erta_workbook: str | Path,
    eralpha_workbook: str | Path,
) -> dict:
    """Compare exact cell formatting by kind, never endpoint data semantics."""
    erta = load_workbook(erta_workbook, read_only=False, data_only=False)
    eralpha = load_workbook(eralpha_workbook, read_only=False, data_only=False)
    try:
        erta_sheet = erta.active
        if "Predictions" not in eralpha.sheetnames:
            raise AssertionError("ERalpha workbook lacks the Predictions sheet")
        eralpha_sheet = eralpha["Predictions"]
        inventories = {
            "erta": _worksheet_format_inventory(erta_sheet),
            "eralpha": _worksheet_format_inventory(eralpha_sheet),
        }
        compared_kinds = sorted(
            set(inventories["erta"]) & set(inventories["eralpha"])
        )
        required_kinds = {"header", "text", "number", "boolean"}
        missing = required_kinds - set(compared_kinds)
        if missing:
            raise AssertionError(
                "primary worksheet formatting comparison lacks common cell "
                f"kinds: {sorted(missing)}"
            )
        non_uniform = {
            f"{endpoint}/{kind}": len(
                inventories[endpoint][kind]["signatures"]
            )
            for endpoint in ("erta", "eralpha")
            for kind in required_kinds
            if len(inventories[endpoint][kind]["signatures"]) != 1
        }
        if non_uniform:
            raise AssertionError(
                "plain primary worksheets have non-uniform formatting within "
                f"a cell kind: {non_uniform}"
            )
        mismatches = {}
        for kind in compared_kinds:
            signatures = {
                endpoint: [
                    item["signature"]
                    for item in inventories[endpoint][kind]["signatures"]
                ]
                for endpoint in ("erta", "eralpha")
            }
            if signatures["erta"] != signatures["eralpha"]:
                mismatches[kind] = signatures
        if mismatches:
            raise AssertionError(
                "ERalpha Predictions formatting differs from plain ERTA cells "
                "of the same kind: "
                + json.dumps(mismatches, sort_keys=True)
            )
        return {
            "comparison": (
                "exact openpyxl font/fill/border/alignment/number-format/"
                "protection/quote-prefix/pivot-button signature by cell kind"
            ),
            "data_semantics_compared": False,
            "cell_kind_policy": (
                "header row, then blank/boolean/date/number/text by openpyxl "
                "cell value type; values and column meanings are ignored"
            ),
            "erta_sheet": erta_sheet.title,
            "eralpha_sheet": eralpha_sheet.title,
            "compared_kinds": compared_kinds,
            "required_kinds": sorted(required_kinds),
            "inventories": inventories,
            "uniform_within_required_kinds": True,
            "exact_match": True,
        }
    finally:
        erta.close()
        eralpha.close()


def _widget_geometry_receipt(widget, endpoint_tab) -> dict:
    mapped = bool(widget.winfo_ismapped())
    receipt = {
        "mapped": mapped,
        "class": str(widget.winfo_class()),
        "manager": str(widget.winfo_manager()),
    }
    try:
        receipt["state"] = str(widget.cget("state"))
    except Exception:
        pass
    # Tcl accepts abbreviated option names.  Asking an Entry for "text" can
    # therefore return its endpoint-local -textvariable name (PY_VAR...), which
    # is widget plumbing rather than visible content.  Compare static button
    # captions, while input/status values are checked through their variables
    # by the focused semantic inspections.
    if receipt["class"] in {"Button", "TButton"}:
        receipt["text"] = str(widget.cget("text"))
    if mapped:
        width = int(widget.winfo_width())
        height = int(widget.winfo_height())
        if width <= 1 or height <= 1:
            raise AssertionError(
                f"mapped {widget!s} has an unresolved {width}x{height} bounding box"
            )
        receipt["box"] = {
            "x": int(widget.winfo_rootx() - endpoint_tab.winfo_rootx()),
            "y": int(widget.winfo_rooty() - endpoint_tab.winfo_rooty()),
            "width": width,
            "height": height,
        }
    return receipt


def _capture_own_tk_window(window, destination: Path) -> dict:
    """Capture only this Tk HWND via PrintWindow; never fall back to the desktop."""
    if os.name != "nt":
        return {
            "supported": False,
            "method": "PrintWindow(PW_RENDERFULLCONTENT)",
            "reason": "own-window native capture is only available on Windows",
        }
    try:
        import ctypes

        from PIL import Image, ImageStat
        import win32gui
        import win32ui

        handle = int(window.winfo_id())
        handle = int(win32gui.GetAncestor(handle, 2) or handle)
        left, top, right, bottom = win32gui.GetWindowRect(handle)
        width, height = right - left, bottom - top
        if width <= 1 or height <= 1:
            raise RuntimeError(f"invalid own-window bounds {width}x{height}")

        window_dc = win32gui.GetWindowDC(handle)
        source_dc = win32ui.CreateDCFromHandle(window_dc)
        memory_dc = source_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(source_dc, width, height)
        memory_dc.SelectObject(bitmap)
        try:
            if not ctypes.windll.user32.PrintWindow(
                handle,
                memory_dc.GetSafeHdc(),
                2,
            ):
                raise RuntimeError("PrintWindow failed")
            image = Image.frombuffer(
                "RGB",
                (width, height),
                bitmap.GetBitmapBits(True),
                "raw",
                "BGRX",
                0,
                1,
            ).copy()
        finally:
            win32gui.DeleteObject(bitmap.GetHandle())
            memory_dc.DeleteDC()
            source_dc.DeleteDC()
            win32gui.ReleaseDC(handle, window_dc)

        destination.parent.mkdir(parents=True, exist_ok=True)
        image.save(destination)
        extrema = ImageStat.Stat(image).extrema
        if not any(low != high for low, high in extrema):
            raise RuntimeError("PrintWindow returned a uniform image")
        return {
            "supported": True,
            "method": "PrintWindow(PW_RENDERFULLCONTENT)",
            "path": str(destination),
            "sha256": _sha256_path(destination),
            "size": [width, height],
            "rgb_extrema": [list(channel) for channel in extrema],
            "target_hwnd": handle,
        }
    except Exception as error:
        try:
            destination.unlink(missing_ok=True)
        except OSError:
            pass
        return {
            "supported": False,
            "method": "PrintWindow(PW_RENDERFULLCONTENT)",
            "reason": f"{type(error).__name__}: {error}",
        }


def _alternate_window_extent(
    current: int,
    minimum: int,
    maximum: int,
    delta: int,
) -> int:
    smaller = current - delta
    if smaller >= minimum:
        return smaller
    larger = current + delta
    if maximum <= 0 or larger <= maximum:
        return larger
    if minimum != current:
        return minimum
    if maximum != current:
        return maximum
    raise AssertionError(f"Tk window extent {current} cannot be resized")


def _layout_box(widget, root) -> dict:
    mapped = bool(widget.winfo_ismapped())
    result = {"mapped": mapped}
    if mapped:
        result["box"] = {
            "x": int(widget.winfo_rootx() - root.winfo_rootx()),
            "y": int(widget.winfo_rooty() - root.winfo_rooty()),
            "width": int(widget.winfo_width()),
            "height": int(widget.winfo_height()),
        }
    return result


def _settle_tk_layout(
    app,
    watched_widgets: dict,
    *,
    expected_root_size: tuple[int, int] | None,
    context: str,
    max_cycles: int = 12,
) -> dict:
    """Drain real Tk events until root and watched boxes repeat unchanged."""
    previous = None
    stable_samples = 0
    observations = []
    for cycle in range(1, max_cycles + 1):
        idle_reached = []
        app.after_idle(lambda: idle_reached.append(True))
        app.update()
        app.update_idletasks()
        root_size = (int(app.winfo_width()), int(app.winfo_height()))
        boxes = {
            name: _layout_box(widget, app)
            for name, widget in sorted(watched_widgets.items())
        }
        snapshot = {
            "root_size": list(root_size),
            "idle_reached": bool(idle_reached),
            "boxes": boxes,
        }
        observations.append(snapshot)
        signature = json.dumps(snapshot, sort_keys=True)
        root_matches = (
            expected_root_size is None or root_size == expected_root_size
        )
        if root_matches and idle_reached and signature == previous:
            stable_samples += 1
        elif root_matches and idle_reached:
            stable_samples = 1
        else:
            stable_samples = 0
        if stable_samples >= 2:
            return {
                "context": context,
                "cycles": cycle,
                "stable_samples": stable_samples,
                **snapshot,
            }
        previous = signature

    requested = (
        list(expected_root_size) if expected_root_size is not None else None
    )
    raise AssertionError(
        f"Tk layout did not settle for {context} after {max_cycles} cycles: "
        + json.dumps(
            {
                "requested_root_size": requested,
                "last_observations": observations[-3:],
            },
            sort_keys=True,
        )
    )


def inspect_endpoint_geometry_parity(
    app,
    screenshot_directory: str | Path | None = None,
) -> dict:
    """Inspect the real mapped ERTA/ERalpha controls at identical Tk geometries."""
    endpoints = {
        "erta": app,
        "eralpha": app.eralpha_tab,
    }
    maps = {
        name: dict(getattr(owner, "parity_widgets", {}))
        for name, owner in endpoints.items()
    }
    if not maps["erta"] or not maps["eralpha"]:
        raise AssertionError("both endpoints must expose a non-empty parity_widgets map")
    if maps["erta"].keys() != maps["eralpha"].keys():
        raise AssertionError(
            "endpoint parity widget keys differ: "
            f"ERTA={sorted(maps['erta'])}, ERalpha={sorted(maps['eralpha'])}"
        )
    required_keys = {
        "endpoint_tab",
        "options_button",
        "options_frame",
        "model_entry",
        "model_browse_button",
        "model_reload_button",
        "ad_entry",
        "ad_browse_button",
        "ad_reload_button",
        "mode_notebook",
        "single_tab",
        "batch_tab",
        "cas_label",
        "cas_entry",
        "pubchem_button",
        "smiles_label",
        "smiles_entry",
        "example_button",
        "single_predict_button",
        "draw_structure_button",
        "batch_input_entry",
        "batch_browse_button",
        "batch_destination_entry",
        "batch_example_button",
        "batch_predict_button",
        "batch_progress",
        "batch_status",
        "batch_result",
    }
    missing = required_keys - maps["erta"].keys()
    if missing:
        raise AssertionError(
            f"endpoint parity map is missing required controls: {sorted(missing)}"
        )
    containers = {
        name: {
            "header_box": controls["options_button"].master,
            "single_input_box": controls["cas_entry"].master,
            "batch_input_box": controls["batch_input_entry"].master,
            "batch_progress_box": controls["batch_progress"].master,
            "batch_result_box": controls["batch_result"].master,
        }
        for name, controls in maps.items()
    }

    app.update()
    app.update_idletasks()
    original_geometry = app.geometry()
    original_state = str(app.state())
    original_root_size = (int(app.winfo_width()), int(app.winfo_height()))
    original_endpoint = app.notebook.select()
    original_modes = {
        name: controls["mode_notebook"].select()
        for name, controls in maps.items()
    }
    original_options = {
        name: bool(owner.options_visible)
        for name, owner in endpoints.items()
    }
    screenshot_root = (
        Path(screenshot_directory)
        if screenshot_directory is not None
        else None
    )
    screenshot_available = screenshot_root is not None
    screenshots = []
    states = []
    toggle_invocations = {"erta": 0, "eralpha": 0}
    window_size_settling = {}

    def select_surface(
        name: str,
        mode: str,
        options_open: bool,
        expected_root_size: tuple[int, int],
    ) -> dict:
        nonlocal toggle_invocations
        owner = endpoints[name]
        controls = maps[name]
        watched = {
            **{
                f"control:{key}": widget
                for key, widget in controls.items()
            },
            **{
                f"container:{key}": widget
                for key, widget in containers[name].items()
            },
        }
        app.notebook.select(controls["endpoint_tab"])
        controls["mode_notebook"].select(controls[f"{mode}_tab"])
        _settle_tk_layout(
            app,
            watched,
            expected_root_size=expected_root_size,
            context=f"{name}/{mode}/before Options transition",
        )
        actual_open = bool(controls["options_frame"].winfo_ismapped())
        if bool(owner.options_visible) != actual_open:
            raise AssertionError(
                f"{name} Options state flag and mapped state disagree"
            )
        if actual_open != options_open:
            controls["options_button"].invoke()
            toggle_invocations[name] += 1
        settled = _settle_tk_layout(
            app,
            watched,
            expected_root_size=expected_root_size,
            context=(
                f"{name}/{mode}/Options "
                f"{'open' if options_open else 'closed'}"
            ),
        )
        if (
            bool(owner.options_visible) != options_open
            or bool(controls["options_frame"].winfo_ismapped()) != options_open
            or bool(controls["options_frame"].grid_info()) != options_open
        ):
            raise AssertionError(
                f"{name} Options button did not produce the requested "
                f"{'open' if options_open else 'closed'} state"
            )
        selected = controls["mode_notebook"].select()
        if selected != str(controls[f"{mode}_tab"]):
            raise AssertionError(f"{name} did not select its {mode} tab")
        return settled

    try:
        if original_state != "normal":
            app.state("normal")
        initial_settling = _settle_tk_layout(
            app,
            {"top_notebook": app.notebook},
            expected_root_size=None,
            context="initial root geometry",
        )
        initial_size = (int(app.winfo_width()), int(app.winfo_height()))
        minimum = tuple(int(value) for value in app.minsize())
        maximum = tuple(int(value) for value in app.maxsize())
        resized_size = (
            _alternate_window_extent(
                initial_size[0],
                minimum[0],
                maximum[0],
                73,
            ),
            _alternate_window_extent(
                initial_size[1],
                minimum[1],
                maximum[1],
                47,
            ),
        )
        if resized_size == initial_size:
            raise AssertionError("the alternate Tk geometry did not resize the window")

        for size_name, requested_size in (
            ("initial", initial_size),
            ("resized", resized_size),
        ):
            app.geometry(f"{requested_size[0]}x{requested_size[1]}")
            size_settling = _settle_tk_layout(
                app,
                {"top_notebook": app.notebook},
                expected_root_size=requested_size,
                context=f"{size_name} requested root geometry",
            )
            window_size_settling[size_name] = size_settling
            actual_size = tuple(size_settling["root_size"])
            if size_name == "resized" and actual_size == initial_size:
                raise AssertionError("Tk ignored the requested resize")

            for mode in ("single", "batch"):
                for options_open in (False, True):
                    endpoint_receipts = {}
                    container_receipts = {}
                    endpoint_sizes = {}
                    endpoint_settling = {}
                    for endpoint_name in endpoints:
                        endpoint_settling[endpoint_name] = select_surface(
                            endpoint_name,
                            mode,
                            options_open,
                            requested_size,
                        )
                        endpoint_sizes[endpoint_name] = [
                            int(app.winfo_width()),
                            int(app.winfo_height()),
                        ]
                        endpoint_tab = maps[endpoint_name]["endpoint_tab"]
                        endpoint_receipts[endpoint_name] = {
                            key: _widget_geometry_receipt(widget, endpoint_tab)
                            for key, widget in maps[endpoint_name].items()
                        }
                        container_receipts[endpoint_name] = {
                            key: _widget_geometry_receipt(widget, endpoint_tab)
                            for key, widget in containers[endpoint_name].items()
                        }
                        if screenshot_available and screenshot_root is not None:
                            option_name = "options-open" if options_open else "options-closed"
                            screenshot = _capture_own_tk_window(
                                app,
                                screenshot_root
                                / (
                                    f"{size_name}-{mode}-{option_name}-"
                                    f"{endpoint_name}.png"
                                ),
                            )
                            screenshot.update(
                                {
                                    "window_geometry": size_name,
                                    "mode": mode,
                                    "options_open": options_open,
                                    "endpoint": endpoint_name,
                                }
                            )
                            screenshots.append(screenshot)
                            if not screenshot["supported"]:
                                screenshot_available = False

                    if endpoint_sizes["erta"] != endpoint_sizes["eralpha"]:
                        raise AssertionError(
                            f"root geometry changed between endpoints: {endpoint_sizes}"
                        )
                    if endpoint_receipts["erta"] != endpoint_receipts["eralpha"]:
                        different = [
                            key
                            for key in maps["erta"]
                            if endpoint_receipts["erta"][key]
                            != endpoint_receipts["eralpha"][key]
                        ]
                        raise AssertionError(
                            "actual ERTA/ERalpha control geometry differs for "
                            f"{size_name}/{mode}/Options "
                            f"{'open' if options_open else 'closed'}: "
                            + json.dumps(
                                {
                                    key: {
                                        endpoint: endpoint_receipts[endpoint][key]
                                        for endpoint in endpoints
                                    }
                                    for key in different
                                },
                                sort_keys=True,
                            )
                        )
                    if container_receipts["erta"] != container_receipts["eralpha"]:
                        different = [
                            key
                            for key in containers["erta"]
                            if container_receipts["erta"][key]
                            != container_receipts["eralpha"][key]
                        ]
                        raise AssertionError(
                            "actual ERTA/ERalpha container boxes differ for "
                            f"{size_name}/{mode}/Options "
                            f"{'open' if options_open else 'closed'}: "
                            + json.dumps(
                                {
                                    key: {
                                        endpoint: container_receipts[endpoint][key]
                                        for endpoint in endpoints
                                    }
                                    for key in different
                                },
                                sort_keys=True,
                            )
                        )
                    states.append(
                        {
                            "window_geometry": size_name,
                            "requested_size": list(requested_size),
                            "actual_size": endpoint_sizes["erta"],
                            "mode": mode,
                            "options_open": options_open,
                            "exact_match": True,
                            "controls": endpoint_receipts,
                            "containers": container_receipts,
                            "settling": endpoint_settling,
                        }
                    )
    finally:
        for name, owner in endpoints.items():
            controls = maps[name]
            app.notebook.select(controls["endpoint_tab"])
            controls["mode_notebook"].select(original_modes[name])
            app.update()
            app.update_idletasks()
            if bool(owner.options_visible) != original_options[name]:
                controls["options_button"].invoke()
                app.update()
                app.update_idletasks()
        app.notebook.select(original_endpoint)
        app.geometry(original_geometry)
        if original_state != "normal":
            app.state(original_state)
        if original_state == "normal":
            _settle_tk_layout(
                app,
                {"top_notebook": app.notebook},
                expected_root_size=original_root_size,
                context="restored root geometry",
            )
        else:
            app.update_idletasks()

    expected_state_count = 2 * 2 * 2
    if len(states) != expected_state_count:
        raise AssertionError(
            f"expected {expected_state_count} geometry comparisons, found {len(states)}"
        )
    return {
        "coordinate_space": "relative to each selected endpoint_tab",
        "compared_widget_keys": sorted(maps["erta"]),
        "compared_container_keys": sorted(containers["erta"]),
        "state_count": len(states),
        "toggle_invocations": toggle_invocations,
        "initial_settling": initial_settling,
        "window_size_settling": window_size_settling,
        "states": states,
        "screenshots": screenshots,
        "screenshot_policy": (
            "Own Tk HWND only via PrintWindow; no desktop-capture fallback."
        ),
    }


def inspect_shared_example_parity(
    app,
    isolated_profile_root: str | Path | None = None,
) -> dict:
    """Verify both endpoints share the writable, byte-faithful test.xlsx."""
    from core.paths import (
        resolve_shared_example_input,
        validate_mutable_directory,
    )

    endpoints = {
        "erta": app,
        "eralpha": app.eralpha_tab,
    }
    expected = {
        name: {
            "cas": str(owner.example_cas),
            "smiles": str(owner.example_smiles),
            "workbook": Path(owner.example_workbook).expanduser().resolve(
                strict=True
            ),
            "batch_workbook": Path(owner.batch_input_var.get()).expanduser().resolve(
                strict=True
            ),
            "batch_display": owner.batch_input_display_var.get(),
            "batch_destination": owner.batch_destination_var.get(),
        }
        for name, owner in endpoints.items()
    }
    for name, owner in endpoints.items():
        if not expected[name]["cas"]:
            raise AssertionError(f"{name} preset CAS is blank")
        if (
            owner.cas_var.get() != expected[name]["cas"]
            or owner.smiles_var.get() != expected[name]["smiles"]
        ):
            raise AssertionError(f"{name} did not start with its preset example")
        if expected[name]["workbook"] != expected[name]["batch_workbook"]:
            raise AssertionError(
                f"{name} batch example path differs from its preset workbook"
            )
        if expected[name]["batch_display"] != expected[name]["workbook"].name:
            raise AssertionError(
                f"{name} batch example display does not name the preset workbook"
            )
        if expected[name]["batch_destination"] != str(
            expected[name]["workbook"].parent
        ):
            raise AssertionError(
                f"{name} batch destination does not name the writable preset parent"
            )

    if (
        expected["erta"]["cas"],
        expected["erta"]["smiles"],
    ) != (
        expected["eralpha"]["cas"],
        expected["eralpha"]["smiles"],
    ):
        raise AssertionError("ERTA and ERalpha preset CAS/SMILES values differ")
    if expected["erta"]["workbook"] != expected["eralpha"]["workbook"]:
        raise AssertionError("ERTA and ERalpha batch example paths differ")
    if expected["erta"]["batch_display"] != expected["eralpha"]["batch_display"]:
        raise AssertionError("ERTA and ERalpha batch example displays differ")
    if (
        expected["erta"]["batch_destination"]
        != expected["eralpha"]["batch_destination"]
    ):
        raise AssertionError("ERTA and ERalpha batch destinations differ")
    if expected["erta"]["workbook"].name.casefold() != "test.xlsx":
        raise AssertionError("shared batch default is not named test.xlsx")

    erta_bytes = expected["erta"]["workbook"].read_bytes()
    eralpha_bytes = expected["eralpha"]["workbook"].read_bytes()
    if erta_bytes != eralpha_bytes:
        raise AssertionError("ERTA and ERalpha batch example workbook bytes differ")

    default_metadata = _read_shared_example_workbook(
        expected["erta"]["workbook"]
    )
    if not default_metadata["cas_values"]:
        raise AssertionError("shared example workbook has no data rows")
    source_cas = default_metadata["cas_values"][0]
    source_smiles = default_metadata["smiles_values"][0]
    if (source_cas, source_smiles) != (
        expected["erta"]["cas"],
        expected["erta"]["smiles"],
    ):
        raise AssertionError(
            "preset CAS/SMILES does not equal the first shared workbook example"
        )

    default_hash_before = hashlib.sha256(erta_bytes).hexdigest()
    validate_mutable_directory(
        expected["erta"]["workbook"].parent,
        forbidden_roots=(app.project_root, app.install_root),
    )
    if not os.access(expected["erta"]["workbook"], os.W_OK):
        raise AssertionError("resolved shared test.xlsx is not writable")
    if hashlib.sha256(
        expected["erta"]["workbook"].read_bytes()
    ).hexdigest() != default_hash_before:
        raise AssertionError(
            "writability inspection changed the shared default workbook"
        )

    bundled_workbook = (
        Path(app.project_root).expanduser().resolve(strict=True)
        / "templates"
        / "test.xlsx"
    )
    if not bundled_workbook.is_file():
        raise AssertionError(
            f"bundled shared example is missing: {bundled_workbook}"
        )
    bundled_bytes = bundled_workbook.read_bytes()
    bundled_sha256 = hashlib.sha256(bundled_bytes).hexdigest()
    if bundled_sha256 != DISTRIBUTED_TEST_XLSX_SHA256:
        raise AssertionError(
            "bundled test.xlsx bytes differ from the supplied distribution source"
        )
    bundled_metadata = _read_shared_example_workbook(bundled_workbook)
    if (
        bundled_metadata["headers"] != ["CARSRN"]
        or bundled_metadata["cas_header"] != "CARSRN"
        or bundled_metadata["smiles_header"] is not None
        or bundled_metadata["row_count"] != 25
        or not all(bundled_metadata["cas_values"])
    ):
        raise AssertionError(
            "bundled test.xlsx does not preserve the supplied CARSRN/25-row contract"
        )

    fresh_copy = None
    if isolated_profile_root is not None:
        qa_profile = Path(isolated_profile_root).expanduser().resolve(
            strict=False
        )
        requested_fresh_copy = (
            qa_profile
            / "Documents"
            / "ER_Predictor"
            / "Examples"
            / "test.xlsx"
        )
        if qa_profile.exists():
            raise AssertionError(
                f"fresh shared-example QA profile already exists: {qa_profile}"
            )
        original_environment = {
            name: os.environ.get(name)
            for name in (
                "USERPROFILE",
                "LOCALAPPDATA",
                "ER_PREDICTOR_PORTABLE",
            )
        }
        had_frozen = hasattr(os.sys, "frozen")
        original_frozen = getattr(os.sys, "frozen", None)
        try:
            os.environ["USERPROFILE"] = str(qa_profile)
            os.environ["LOCALAPPDATA"] = str(qa_profile / "LocalAppData")
            os.environ["ER_PREDICTOR_PORTABLE"] = "0"
            setattr(os.sys, "frozen", False)
            fresh_copy = resolve_shared_example_input()
        finally:
            if had_frozen:
                setattr(os.sys, "frozen", original_frozen)
            else:
                delattr(os.sys, "frozen")
            for name, value in original_environment.items():
                if value is None:
                    os.environ.pop(name, None)
                else:
                    os.environ[name] = value
        if fresh_copy != requested_fresh_copy or not fresh_copy.is_file():
            raise AssertionError(
                "shared-example resolver did not create the isolated first-run copy"
            )
        if fresh_copy.read_bytes() != bundled_bytes:
            raise AssertionError(
                "newly created shared default does not byte-match bundled test.xlsx"
            )
        validate_mutable_directory(
            fresh_copy.parent,
            forbidden_roots=(app.project_root, app.install_root),
        )
        if not os.access(fresh_copy, os.W_OK):
            raise AssertionError("newly created shared-example QA copy is not writable")

    original_endpoint = app.notebook.select()
    original_modes = {
        name: owner.parity_widgets["mode_notebook"].select()
        for name, owner in endpoints.items()
    }
    callback_receipts = []
    try:
        for name, owner in endpoints.items():
            other_name = "eralpha" if name == "erta" else "erta"
            other = endpoints[other_name]
            other_values = (
                other.cas_var.get(),
                other.smiles_var.get(),
                other.batch_input_var.get(),
                other.batch_input_display_var.get(),
                other.batch_destination_var.get(),
            )
            owner.cas_var.set(f"{name}-changed-cas")
            owner.smiles_var.set(f"{name}-changed-smiles")
            owner.batch_input_var.set(
                str(expected[name]["workbook"].parent / f"{name}-changed.xlsx")
            )
            owner.batch_input_display_var.set(f"{name}-changed.xlsx")
            owner.batch_destination_var.set(f"{name}-changed-destination")
            app.notebook.select(owner.parity_widgets["endpoint_tab"])
            owner.parity_widgets["mode_notebook"].select(
                owner.parity_widgets["single_tab"]
            )
            app.update_idletasks()
            owner.parity_widgets["example_button"].invoke()
            app.update_idletasks()
            restored = (owner.cas_var.get(), owner.smiles_var.get())
            expected_values = (expected[name]["cas"], expected[name]["smiles"])
            restored_batch = (
                Path(owner.batch_input_var.get()).expanduser().resolve(
                    strict=True
                ),
                owner.batch_input_display_var.get(),
                owner.batch_destination_var.get(),
            )
            expected_batch = (
                expected[name]["workbook"],
                expected[name]["workbook"].name,
                str(expected[name]["workbook"].parent),
            )
            if restored != expected_values:
                raise AssertionError(
                    f"{name} Example input button did not restore the preset"
                )
            if restored_batch != expected_batch:
                raise AssertionError(
                    f"{name} Example input button did not restore the batch default"
                )
            if (
                other.cas_var.get(),
                other.smiles_var.get(),
                other.batch_input_var.get(),
                other.batch_input_display_var.get(),
                other.batch_destination_var.get(),
            ) != other_values:
                raise AssertionError(
                    f"{name} Example input callback leaked into {other_name}"
                )
            callback_receipts.append(
                {
                    "endpoint": name,
                    "widget": "example_button",
                    "invoked": True,
                    "restored": list(restored),
                    "restored_batch": [
                        str(restored_batch[0]),
                        restored_batch[1],
                        restored_batch[2],
                    ],
                    "other_endpoint_unchanged": True,
                }
            )
    finally:
        for name, owner in endpoints.items():
            owner.parity_widgets["mode_notebook"].select(original_modes[name])
        app.notebook.select(original_endpoint)
        app.update_idletasks()

    return {
        "cas": expected["erta"]["cas"],
        "smiles": expected["erta"]["smiles"],
        "workbook": str(expected["erta"]["workbook"]),
        "workbook_size_bytes": len(erta_bytes),
        "workbook_sha256": default_hash_before,
        "workbook_sheet": default_metadata["sheet"],
        "workbook_headers": default_metadata["headers"],
        "workbook_row_count": default_metadata["row_count"],
        "cas_header": default_metadata["cas_header"],
        "smiles_header": default_metadata["smiles_header"],
        "batch_display": expected["erta"]["batch_display"],
        "batch_destination": expected["erta"]["batch_destination"],
        "path_equal": True,
        "content_equal": True,
        "default_writable": True,
        "default_matches_bundled": erta_bytes == bundled_bytes,
        "bundled_workbook": str(bundled_workbook),
        "bundled_size_bytes": len(bundled_bytes),
        "bundled_sha256": bundled_sha256,
        "bundled_headers": bundled_metadata["headers"],
        "bundled_row_count": bundled_metadata["row_count"],
        "bundled_cas_values": bundled_metadata["cas_values"],
        "fresh_copy": str(fresh_copy) if fresh_copy is not None else None,
        "fresh_copy_writable": fresh_copy is not None,
        "fresh_copy_matches_bundled": fresh_copy is not None,
        "fresh_copy_sha256": (
            _sha256_path(fresh_copy) if fresh_copy is not None else None
        ),
        "fresh_copy_size_bytes": (
            fresh_copy.stat().st_size if fresh_copy is not None else None
        ),
        "fresh_copy_resolution": (
            "resolve_shared_example_input() under a QA-only USERPROFILE and "
            "LOCALAPPDATA with frozen layout temporarily disabled"
            if fresh_copy is not None
            else None
        ),
        "callbacks": callback_receipts,
    }


class NativePackageQa:
    """Drive real Tk callbacks and record fail-closed packaged-app evidence."""

    BASELINE_CHECKS = {
        "erta_eralpha_exact_ui_geometry_and_examples",
        "eralpha_real_resource_buttons_release_enforcement_and_isolation",
        "erta_startup_resources_and_controls",
        "erta_options_and_resource_controls",
        "erta_exact_valid_single",
        "erta_legacy_invalid_direct_smiles",
        "erta_mocked_pubchem_callback",
        "erta_mocked_pubchem_prediction",
        "erta_model_ad_browse_callbacks",
        "erta_template_callback_and_dialog",
        "erta_batch_workbook_graphs_and_order",
        "erba_evidence_caveat_and_dialog",
        "erba_route",
        "erba_invalid_wildcard",
        "erba_invalid_pipe",
        "erba_invalid_metal",
        "erba_invalid_blank",
        "erba_invalid_invalid_cas",
        "erba_eralpha_batch_workbook_ad_graphs_and_summary",
        "fixed_eralpha_classification_tab",
        "erta_erba_erta_state_isolation",
        "read_only_install_and_writable_state",
    }
    BATCH_FEEDBACK_CHECKS = {
        "erta_batch_failure_dialog_and_control_recovery",
        "eralpha_batch_failure_dialog_and_control_recovery",
    }
    SHARED_EXAMPLE_BATCH_CHECKS = {
        "erta_shared_test_batch_success_dialog_and_predictions",
        "eralpha_shared_test_batch_success_dialog_and_predictions",
    }
    VALID_SMILES = "C[C@]12CC[C@H]3[C@@H]([C@@H]1CC[C@@H]2O)CCC4=CC(=CC=C34)O"
    LEGACY_INVALID_SMILES = "not a SMILES"
    MOCK_CAS = "50-00-0"
    MOCK_SMILES = "C=O"
    ERBA_BATCH_AD_COLUMNS = (
        "AD",
        "AD_MeanDistance",
        "AD_DistanceThreshold",
        "AD_Distance_InDomain",
        "AD_SimilarityMax",
        "AD_SimilarityThreshold",
        "AD_Similarity_InDomain",
        "AD_PC1",
        "AD_PC2",
    )
    ERBA_BATCH_GRAPH_FILES = {
        "binding_class_count.png",
        "binding_probability_histogram.png",
        "top_binding_chemicals.png",
        "ad_pca_plot.png",
        "ad_decision_plot.png",
    }
    ERBA_PROVENANCE_COLUMNS = (
        "Protocol_SHA256",
        "Source_Manifest_SHA256",
        "Historical_Exposure_Manifest_SHA256",
        "Split_Manifest_SHA256",
        "Nested_CV_SHA256",
        "Internal_Resplit_SHA256",
        "Report_SHA256",
        "Caveat_SHA256",
    )

    def __init__(self, app, destination: str | Path) -> None:
        self.app = app
        self.destination = Path(destination)
        self.destination.parent.mkdir(parents=True, exist_ok=True)
        self.root = Path(app.output_root) / "native-package-qa"
        self.root.mkdir(parents=True, exist_ok=True)
        self.run_root = self.root / f"run-{os.getpid()}-{time.time_ns()}"
        self.run_root.mkdir(parents=True, exist_ok=False)
        self.stage = -4
        self._tick_active = False
        self.route_index = 0
        self.invalid_index = 0
        self.started_at = time.monotonic()
        self.deadline = self.started_at + 900
        self.erta_summary = ""
        self.erta_probability = ""
        self.erta_batch_input_path: Path | None = None
        self.erta_batch_outputs_before: set[Path] = set()
        self.erba_batch_input_path: Path | None = None
        self.erba_batch_outputs_before: set[Path] = set()
        self.erba_batch_graph_dirs_before: set[Path] = set()
        self.dialogs: list[dict] = []
        self.patches: list[tuple[object, str, object]] = []
        self.mock_cas_enabled = False
        self.mock_cas_values: set[str] = {self.MOCK_CAS}
        self.shared_default_path: Path | None = None
        self.shared_default_sha256 = ""
        self.shared_bundle_path: Path | None = None
        self.shared_bundle_sha256 = ""
        self.qa_shared_example_path: Path | None = None
        self.erta_batch_output_path: Path | None = None
        gate_value = os.environ.get("ER_PREDICTOR_EXTERNAL_DRIVER_GATE", "").strip()
        self.external_driver_gate = Path(gate_value) if gate_value else None
        self.run_nonce = os.environ.get("ER_PREDICTOR_AUTOMATION_NONCE", "").strip()
        self.target_pid = os.getpid()
        self.routes = [
            (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA),
        ]
        self.invalid_cases = [
            ("wildcard", "[*]CC", "", "wildcard_smiles"),
            ("pipe", "CC|CC", "", "cxsmiles_not_allowed"),
            ("metal", "[Na+]", "", "no_carbon"),
            ("blank", "", "", "blank_smiles"),
            ("invalid_cas", "", "123-45-6", "CAS is invalid"),
        ]
        self.transcript = {
            "schema_version": 2,
            "kind": "native-app-automation-transcript",
            "surface": "native-desktop",
            "executable": os.fspath(Path(os.sys.executable).resolve()),
            "run_nonce": self.run_nonce,
            "target_pid": self.target_pid,
            "checks": [],
            "automation_scopes": [],
        }

    def start(self) -> None:
        self._install_dialog_observer()
        callback = self._wait_for_external_driver if self.external_driver_gate else self.tick
        self.app.after(100, callback)

    def _wait_for_external_driver(self) -> None:
        try:
            if time.monotonic() >= self.deadline:
                raise TimeoutError("external pywinauto driver did not release the native QA gate")
            if not self.external_driver_gate.is_file():
                self.app.after(100, self._wait_for_external_driver)
                return
            payload = json.loads(self.external_driver_gate.read_text(encoding="utf-8"))
            self._require(payload.get("passed") is True, "external pywinauto driver gate is not passing")
            self._require(bool(self.run_nonce), "automation run nonce is missing")
            self._require(payload.get("run_nonce") == self.run_nonce, "external driver gate nonce mismatch")
            self._require(payload.get("target_pid") == self.target_pid, "external driver gate target PID mismatch")
            self._record(
                "external_pywinauto_driver_gate",
                driver_kind=payload.get("driver_kind"),
                run_nonce=self.run_nonce,
                target_pid=self.target_pid,
                real_input_actions=payload.get("real_input_actions"),
                evidence_path=str(self.external_driver_gate),
            )
            self.app.after(100, self.tick)
        except Exception:
            self._finish(False, traceback.format_exc())

    def _body(self, tab=None) -> str:
        tab = tab or self.app.erba_tab
        fields = (
            tab.single_prediction_summary_var,
            tab.single_negative_probability_var,
            tab.single_positive_probability_var,
            tab.single_pic50_var,
            tab.single_ic50_var,
            tab.single_ad_domain_var,
            tab.single_detail_var,
            tab.single_status_var,
        )
        return "\n".join(value.get().strip() for value in fields if value.get().strip())

    def _record(self, check: str, **evidence) -> None:
        self.transcript["checks"].append({"check": check, **evidence, "passed": True})

    def _require(self, condition: bool, message: str) -> None:
        if not condition:
            raise AssertionError(message)

    def _patch(self, owner: object, name: str, replacement: object) -> None:
        self.patches.append((owner, name, getattr(owner, name)))
        setattr(owner, name, replacement)

    def _restore_patches(self) -> None:
        while self.patches:
            owner, name, original = self.patches.pop()
            setattr(owner, name, original)

    def _install_dialog_observer(self) -> None:
        # Native message boxes block unattended Tk automation; record their exact
        # visible title/body while preserving the callback path that invokes them.
        from gui import erba_tab, main_window

        def observe(kind):
            def handler(title, message, **_kwargs):
                controls_at_dialog = {}
                for endpoint, owner in (
                    ("erta", self.app),
                    ("eralpha", self.app.eralpha_tab),
                ):
                    try:
                        controls_at_dialog[endpoint] = (
                            self._batch_control_states(owner)
                        )
                    except Exception as error:
                        controls_at_dialog[endpoint] = {
                            "unavailable": f"{type(error).__name__}: {error}"
                        }
                self.dialogs.append(
                    {
                        "kind": kind,
                        "title": str(title),
                        "message": str(message),
                        "batch_controls_at_dialog": controls_at_dialog,
                    }
                )
                return "ok"
            return handler

        messagebox_modules = {
            id(main_window.messagebox): main_window.messagebox,
            id(erba_tab.messagebox): erba_tab.messagebox,
        }
        for messagebox_module in messagebox_modules.values():
            self._patch(messagebox_module, "showerror", observe("error"))
            self._patch(messagebox_module, "showinfo", observe("info"))
            self._patch(messagebox_module, "showwarning", observe("warning"))
        self.transcript["automation_scopes"].append({
            "scope": "messagebox_observer",
            "functions": ["showinfo", "showwarning", "showerror"],
            "reason": "record dialog-visible semantics without blocking opt-in automation",
        })

    def _install_mock_cas(self) -> None:
        if self.mock_cas_enabled:
            return
        from gui import erba_tab, main_window

        def mocked_cas_to_smiles(cas: str) -> dict:
            normalized = str(cas).strip()
            if normalized not in self.mock_cas_values:
                raise RuntimeError(f"unexpected deterministic QA CAS: {cas}")
            return {"CanonicalSMILES": self.MOCK_SMILES, "PubChem_CID": "712"}

        self._patch(main_window, "cas_to_smiles", mocked_cas_to_smiles)
        self._patch(erba_tab, "cas_to_smiles", mocked_cas_to_smiles)
        self.mock_cas_enabled = True
        self.transcript["automation_scopes"].append({
            "scope": "mocked_pubchem",
            "cas_values": sorted(self.mock_cas_values),
            "canonical_smiles": self.MOCK_SMILES,
            "cid": "712",
            "reason": (
                "deterministic opt-in UI callback and bundled 25-CAS batch "
                "coverage; every CAS is deliberately substituted with the same "
                "valid SMILES and no network request is made"
            ),
            "production_equivalent": False,
            "online_verification": (
                "Run the packaged app without native-QA substitution to resolve "
                "the actual 25 CAS values through PubChem."
            ),
        })

    def _install_picker(self, save_path: Path) -> None:
        from gui import main_window

        def choose_open_file(**kwargs):
            title = str(kwargs.get("title", ""))
            return self.app.ad_ref_path_var.get() if "AD reference" in title else self.app.model_path_var.get()

        self._patch(main_window.filedialog, "askopenfilename", choose_open_file)
        self._patch(main_window.filedialog, "asksaveasfilename", lambda **_kwargs: str(save_path))
        self.transcript["automation_scopes"].append({
            "scope": "filedialog_selection",
            "path": str(save_path),
            "reason": "exercise callback state transitions without coordinate-driven native dialogs",
        })

    def _finish(self, passed: bool, error: str = "") -> None:
        self._restore_patches()
        self.transcript["passed"] = passed
        self.transcript["duration_seconds"] = time.monotonic() - self.started_at
        if error:
            self.transcript["error"] = error
        self.destination.write_text(
            json.dumps(self.transcript, indent=2, ensure_ascii=False, default=str) + "\n",
            encoding="utf-8",
        )
        hold_seconds = max(1, int(os.environ.get("ER_PREDICTOR_AUTOMATION_HOLD_SECONDS", "30")))
        self.app.after(hold_seconds * 1000, self.app.destroy)

    def _reschedule(self, delay: int = 100) -> None:
        if time.monotonic() >= self.deadline:
            raise TimeoutError(f"native package QA timed out at stage {self.stage}")
        self.app.after(delay, self.tick)

    def _wait_erta_result(self, previous: str = "") -> bool:
        return (
            self.app.status_var.get().startswith("Single prediction done:")
            and self.app.prediction_summary_var.get() != "Prediction: -"
            and (not previous or self.app.result_text.get("1.0", "end").strip() != previous)
        )

    def _wait_dialog(self, count: int) -> bool:
        return len(self.dialogs) > count

    @staticmethod
    def _batch_control_states(owner) -> dict:
        return {
            "input": str(owner.batch_input_button.cget("state")),
            "template": str(owner.download_template_button.cget("state")),
            "run": str(owner.run_batch_button.cget("state")),
        }

    def _terminal_dialog(
        self,
        initial_count: int,
        *,
        kind: str,
        title: str,
        context: str,
    ) -> dict:
        observed = self.dialogs[initial_count:]
        self._require(
            len(observed) == 1,
            f"{context} emitted {len(observed)} dialogs instead of exactly one: "
            f"{observed}",
        )
        dialog = observed[0]
        self._require(
            dialog["kind"] == kind and dialog["title"] == title,
            f"{context} dialog contract drift: {dialog}",
        )
        return dialog

    def _copy_shared_example_for_batch(self, endpoint: str) -> Path:
        self._require(
            self.qa_shared_example_path is not None
            and self.qa_shared_example_path.is_file(),
            "fresh bundled shared-example QA source is unavailable",
        )
        source = self.qa_shared_example_path
        source_hash = _sha256_path(source)
        destination = (
            self.run_root / "shared-example-success" / endpoint / "test.xlsx"
        )
        destination.parent.mkdir(parents=True, exist_ok=False)
        self._require(
            destination.resolve(strict=False) != source.resolve(strict=True),
            f"{endpoint} QA batch input unexpectedly targets the shared source",
        )
        shutil.copyfile(source, destination)
        self._require(
            destination.read_bytes() == source.read_bytes()
            and _sha256_path(source) == source_hash,
            f"{endpoint} QA copy changed the shared example bytes",
        )
        return destination.resolve(strict=True)

    def _make_corrupt_batch_input(self, endpoint: str) -> Path:
        destination = (
            self.run_root / "deterministic-batch-failure" / endpoint / "test.xlsx"
        )
        destination.parent.mkdir(parents=True, exist_ok=False)
        destination.write_bytes(
            b"ER_Predictor native QA deterministic invalid xlsx fixture\n"
        )
        return destination.resolve(strict=True)

    @staticmethod
    def _same_path(first: str | Path, second: str | Path) -> bool:
        try:
            return (
                Path(first).expanduser().resolve(strict=True)
                == Path(second).expanduser().resolve(strict=True)
            )
        except (OSError, RuntimeError):
            return False

    def _parity_startup_ready(self) -> bool:
        app = self.app
        tab = app.eralpha_tab
        statuses = (app.status_var.get(), tab.single_status_var.get())
        for status in statuses:
            if status.startswith("Startup failed:"):
                raise RuntimeError(status)
        return bool(
            app.predictor.is_loaded()
            and app.ad_calculator.fitted
            and app.loaded_model_path_var.get()
            and app.loaded_ad_ref_path_var.get()
            and tab.loaded_model_path_var.get()
            and tab.loaded_ad_ref_path_var.get()
            and tab.predictor is not None
        )

    def _invoke_open_file_button(
        self,
        button,
        selected_path: str | Path,
        *,
        purpose: str,
    ) -> dict:
        """Invoke a real widget command while replacing only the picker result."""
        from gui import erba_tab

        calls = []
        original = erba_tab.filedialog.askopenfilename

        def selected(**kwargs):
            calls.append(dict(kwargs))
            return os.fspath(selected_path)

        self._require(
            str(button.cget("state")) != "disabled",
            f"{purpose} widget is disabled",
        )
        erba_tab.filedialog.askopenfilename = selected
        try:
            button.invoke()
        finally:
            erba_tab.filedialog.askopenfilename = original
        self._require(
            len(calls) == 1,
            f"{purpose} did not invoke exactly one real filedialog callback",
        )
        receipt = {
            "purpose": purpose,
            "widget_text": str(button.cget("text")),
            "selected_path": os.fspath(selected_path),
            "dialog_title": str(calls[0].get("title", "")),
            "callback_invoked": True,
        }
        self.transcript["automation_scopes"].append(
            {
                "scope": "filedialog_selection_return_only",
                **receipt,
                "reason": (
                    "invoke the actual mapped Tk button/callback without "
                    "coordinate-driving a native chooser"
                ),
            }
        )
        return receipt

    @staticmethod
    def _erta_resource_state(app) -> dict:
        model_path = Path(app.model_path_var.get()).expanduser().resolve(
            strict=True
        )
        ad_path = Path(app.ad_ref_path_var.get()).expanduser().resolve(
            strict=True
        )
        return {
            "model_selection": str(model_path),
            "model_loaded": app.loaded_model_path_var.get(),
            "model_sha256": _sha256_path(model_path),
            "ad_selection": str(ad_path),
            "ad_loaded": app.loaded_ad_ref_path_var.get(),
            "ad_sha256": _sha256_path(ad_path),
            "predictor_identity": id(app.predictor),
            "predictor_model_path": str(
                getattr(app.predictor, "model_path", "")
            ),
            "ad_calculator_identity": id(app.ad_calculator),
            "status": app.status_var.get(),
        }

    @staticmethod
    def _erba_model_state(tab) -> dict:
        loaded = getattr(tab.predictor, "_loaded", {})
        return {
            "model_selection": tab.model_path_var.get(),
            "model_loaded": tab.loaded_model_path_var.get(),
            "predictor_identity": id(tab.predictor),
            "model_generation": int(tab._model_generation),
            "selected_model_id": (
                tab.selected_model_spec.model_id
                if tab.selected_model_spec is not None
                else ""
            ),
            "loaded_route_keys": sorted(
                f"{task.value}/{subtype.value}" for task, subtype in loaded
            ),
            "loaded_artifact_identities": sorted(
                id(value[0]) for value in loaded.values()
            ),
        }

    def _start_exact_ui_parity_stage(self) -> None:
        app = self.app
        tab = app.eralpha_tab
        example_receipt = inspect_shared_example_parity(
            app,
            self.run_root / "fresh-default-profile",
        )
        self.shared_default_path = Path(example_receipt["workbook"])
        self.shared_default_sha256 = example_receipt["workbook_sha256"]
        self.shared_bundle_path = Path(example_receipt["bundled_workbook"])
        self.shared_bundle_sha256 = example_receipt["bundled_sha256"]
        self.qa_shared_example_path = Path(example_receipt["fresh_copy"])
        self.mock_cas_values.update(example_receipt["bundled_cas_values"])
        self.transcript["automation_scopes"].append(
            {
                "scope": "isolated_first_run_shared_example_resolution",
                "profile_root": str(
                    self.run_root / "fresh-default-profile"
                ),
                "temporary_environment": {
                    "USERPROFILE": "QA-only profile_root",
                    "LOCALAPPDATA": "QA-only profile_root/LocalAppData",
                    "ER_PREDICTOR_PORTABLE": "0",
                    "sys.frozen": False,
                },
                "reason": (
                    "exercise resolve_shared_example_input() creation without "
                    "reading, replacing, or publishing beside the user's "
                    "distributed test.xlsx"
                ),
            }
        )
        geometry_receipt = inspect_endpoint_geometry_parity(
            app,
            self.run_root / "ui-parity-screenshots",
        )
        self._record(
            "erta_eralpha_exact_ui_geometry_and_examples",
            examples=example_receipt,
            geometry=geometry_receipt,
        )

        route = tab._selected_route("single")
        spec = tab.selected_model_spec
        self._require(
            route == (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
            and spec is not None,
            "ERalpha does not expose its released classification model",
        )
        released_model = (
            tab.catalog_path.parent / spec.relative_path
        ).resolve(strict=True)
        released_ad = Path(tab.ad_ref_path_var.get()).resolve(strict=True)
        self._require(
            self._same_path(tab.model_path_var.get(), released_model)
            and self._same_path(tab.loaded_model_path_var.get(), released_model),
            "ERalpha startup did not select and load the released current model",
        )
        self._require(
            self._same_path(tab.ad_ref_path_var.get(), released_ad)
            and self._same_path(tab.loaded_ad_ref_path_var.get(), released_ad),
            "ERalpha startup did not select and load its approved AD reference",
        )
        self._require(
            _sha256_path(released_model) == spec.sha256.lower(),
            "released ERalpha model bytes do not match the catalog before QA",
        )

        self.parity_released_model = released_model
        self.parity_released_ad = released_ad
        self.parity_model_sha256 = _sha256_path(released_model)
        self.parity_ad_sha256 = _sha256_path(released_ad)
        self.parity_erta_before = self._erta_resource_state(app)
        self.parity_model_before = self._erba_model_state(tab)
        self.parity_options_before = bool(tab.options_visible)

        app.notebook.select(tab)
        if not tab.options_visible:
            tab.parity_widgets["options_button"].invoke()
        app.update_idletasks()
        self._require(
            tab.options_visible and tab.options_frame.winfo_ismapped(),
            "ERalpha Options did not open before resource interaction",
        )
        self.parity_model_browse = self._invoke_open_file_button(
            tab.parity_widgets["model_browse_button"],
            released_model,
            purpose="ERalpha released model Browse",
        )
        self._require(
            self._same_path(tab.model_path_var.get(), released_model)
            and self._erba_model_state(tab) == self.parity_model_before,
            "browsing the current released model activated or changed it before Reload",
        )
        self.parity_model_dialog_count = len(self.dialogs)
        tab.parity_widgets["model_reload_button"].invoke()
        self._require(
            tab.single_status_var.get().startswith("Loading model"),
            "ERalpha Reload model button did not start the actual callback",
        )

    def _model_reload_finished(self) -> bool:
        tab = self.app.eralpha_tab
        status = tab.single_status_var.get()
        if status.startswith("Error:"):
            raise RuntimeError(f"approved ERalpha model reload failed: {status}")
        return bool(
            tab._model_generation > self.parity_model_before["model_generation"]
            and self._same_path(
                tab.loaded_model_path_var.get(),
                self.parity_released_model,
            )
            and status
            == f"Model loaded: {tab.selected_model_spec.model_id}"
            and self._wait_dialog(self.parity_model_dialog_count)
        )

    def _finish_model_reload_and_start_ad(self) -> None:
        app = self.app
        tab = app.eralpha_tab
        model_after = self._erba_model_state(tab)
        self._require(
            model_after["predictor_identity"]
            != self.parity_model_before["predictor_identity"]
            and model_after["model_generation"]
            > self.parity_model_before["model_generation"]
            and model_after["loaded_route_keys"]
            == ["classification/er_alpha"],
            "ERalpha Reload model did not activate a fresh preflighted predictor",
        )
        self._require(
            _sha256_path(self.parity_released_model)
            == self.parity_model_sha256,
            "ERalpha released model bytes changed during Reload",
        )
        self.parity_model_after = model_after
        self.parity_model_dialog = self.dialogs[-1]

        self.parity_ad_browse = self._invoke_open_file_button(
            tab.parity_widgets["ad_browse_button"],
            self.parity_released_ad,
            purpose="ERalpha approved AD Browse",
        )
        self._require(
            self._same_path(
                tab.ad_ref_path_var.get(),
                self.parity_released_ad,
            ),
            "ERalpha AD Browse did not retain the approved route reference",
        )
        self.parity_ad_manager_before = id(tab.erba_ad)
        self.parity_ad_dialog_count = len(self.dialogs)
        tab.parity_widgets["ad_reload_button"].invoke()
        self._require(
            tab.single_status_var.get().startswith(
                "Rebuilding AD reference cache"
            ),
            "ERalpha Reload AD button did not start the actual callback",
        )

    def _ad_reload_finished(self) -> bool:
        tab = self.app.eralpha_tab
        status = tab.single_status_var.get()
        if status.startswith("Error:"):
            raise RuntimeError(f"approved ERalpha AD reload failed: {status}")
        return bool(
            self._same_path(
                tab.loaded_ad_ref_path_var.get(),
                self.parity_released_ad,
            )
            and status.startswith("AD fitted.")
            and self._wait_dialog(self.parity_ad_dialog_count)
        )

    def _finish_resource_parity_stage(self) -> None:
        app = self.app
        tab = app.eralpha_tab
        self._require(
            _sha256_path(self.parity_released_ad) == self.parity_ad_sha256,
            "ERalpha approved AD reference bytes changed during Reload",
        )
        self._require(
            id(tab.erba_ad) != self.parity_ad_manager_before,
            "ERalpha Reload AD did not activate a fresh endpoint-local AD manager",
        )
        ad_dialog = self.dialogs[-1]
        model_candidate = self.run_root / "unapproved-erba-model.joblib"
        ad_candidate = self.run_root / "unapproved-erba-ad.xlsx"
        model_candidate.write_bytes(b"native QA unapproved model candidate\n")
        ad_candidate.write_bytes(b"native QA unapproved AD candidate\n")
        try:
            approved_state = self._erba_model_state(tab)
            model_dialog_count = len(self.dialogs)
            rejected_model_browse = self._invoke_open_file_button(
                tab.parity_widgets["model_browse_button"],
                model_candidate,
                purpose="ERalpha unapproved model Browse",
            )
            rejected_model_status = tab.single_status_var.get()
            self._require(
                self._erba_model_state(tab) == approved_state,
                "unapproved ERalpha candidate changed or loaded model state",
            )
            self._require(
                self._wait_dialog(model_dialog_count)
                and rejected_model_status.startswith("Error:")
                and (
                    "Arbitrary joblib files are not allowed"
                    in rejected_model_status
                    or "integrity" in rejected_model_status.casefold()
                ),
                "unapproved ERalpha candidate was not rejected by Browse",
            )
            rejected_model_dialog = self.dialogs[-1]

            ad_selection_before = tab.ad_ref_path_var.get()
            ad_loaded_before = tab.loaded_ad_ref_path_var.get()
            ad_calculator = tab.erba_ad.calculator_for(
                ERBATask.CLASSIFICATION,
                ERBASubtype.ER_ALPHA,
            )
            ad_calculator_identity = id(ad_calculator)
            ad_dialog_count = len(self.dialogs)
            rejected_ad_browse = self._invoke_open_file_button(
                tab.parity_widgets["ad_browse_button"],
                ad_candidate,
                purpose="ERalpha unapproved AD Browse",
            )
            rejected_ad_status = tab.single_status_var.get()
            self._require(
                tab.ad_ref_path_var.get() == ad_selection_before
                and tab.loaded_ad_ref_path_var.get() == ad_loaded_before
                and id(
                    tab.erba_ad.calculator_for(
                        ERBATask.CLASSIFICATION,
                        ERBASubtype.ER_ALPHA,
                    )
                )
                == ad_calculator_identity,
                "unapproved ERalpha AD candidate changed or fitted AD state",
            )
            self._require(
                self._wait_dialog(ad_dialog_count)
                and rejected_ad_status.startswith("Error:")
                and "approved bundled AD reference" in rejected_ad_status,
                "unapproved ERalpha AD candidate was not rejected by Browse",
            )
            rejected_ad_dialog = self.dialogs[-1]
        finally:
            model_candidate.unlink(missing_ok=True)
            ad_candidate.unlink(missing_ok=True)

        erta_after = self._erta_resource_state(app)
        self._require(
            erta_after == self.parity_erta_before,
            "ERalpha Browse/Reload interactions leaked into ERTA resource state",
        )
        self._require(
            _sha256_path(self.parity_released_model)
            == self.parity_model_sha256,
            "released ERalpha model bytes changed during candidate rejection",
        )
        self._require(
            _sha256_path(self.parity_released_ad) == self.parity_ad_sha256,
            "approved ERalpha AD bytes changed during candidate rejection",
        )
        self._record(
            "eralpha_real_resource_buttons_release_enforcement_and_isolation",
            approved_model={
                "path": str(self.parity_released_model),
                "sha256": self.parity_model_sha256,
                "browse": self.parity_model_browse,
                "before": self.parity_model_before,
                "after": self.parity_model_after,
                "dialog": self.parity_model_dialog,
            },
            approved_ad={
                "path": str(self.parity_released_ad),
                "sha256": self.parity_ad_sha256,
                "browse": self.parity_ad_browse,
                "dialog": ad_dialog,
                "manager_before": self.parity_ad_manager_before,
                "manager_after": id(tab.erba_ad),
            },
            rejected_model={
                "candidate": str(model_candidate),
                "candidate_size_bytes": len(
                    b"native QA unapproved model candidate\n"
                ),
                "browse": rejected_model_browse,
                "status": rejected_model_status,
                "dialog": rejected_model_dialog,
                "model_state_unchanged": True,
            },
            rejected_ad={
                "candidate": str(ad_candidate),
                "candidate_size_bytes": len(
                    b"native QA unapproved AD candidate\n"
                ),
                "browse": rejected_ad_browse,
                "status": rejected_ad_status,
                "dialog": rejected_ad_dialog,
                "ad_state_unchanged": True,
            },
            erta_before=self.parity_erta_before,
            erta_after=erta_after,
            endpoint_isolation=True,
            approved_resource_bytes_unchanged=True,
        )
        if bool(tab.options_visible) != self.parity_options_before:
            tab.parity_widgets["options_button"].invoke()
        app.notebook.select(app.erta_tab)
        app.parity_widgets["mode_notebook"].select(
            app.parity_widgets["single_tab"]
        )
        app.update_idletasks()

    def tick(self) -> None:
        if self._tick_active:
            return
        self._tick_active = True
        try:
            app = self.app
            tab = app.eralpha_tab
            if self.stage == -4:
                if not self._parity_startup_ready():
                    self._reschedule()
                    return
                self._start_exact_ui_parity_stage()
                self.stage = -3
                self._reschedule()
                return
            if self.stage == -3:
                if not self._model_reload_finished():
                    self._reschedule()
                    return
                self._finish_model_reload_and_start_ad()
                self.stage = -2
                self._reschedule()
                return
            if self.stage == -2:
                if not self._ad_reload_finished():
                    self._reschedule()
                    return
                self._finish_resource_parity_stage()
                self.stage = 0
                self._reschedule()
                return
            if self.stage == 0:
                self._require(app.notebook.tab(app.notebook.select(), "text") == "ERTA", "ERTA is not startup tab")
                self._require(Path(app.model_path_var.get()).is_file(), "default ERTA model resource missing")
                self._require(Path(app.ad_ref_path_var.get()).is_file(), "default ERTA AD resource missing")
                self._require(app.options_frame.winfo_ismapped() == 0, "options must start hidden")
                self._require(str(app.run_batch_button.cget("state")) == "normal", "ERTA batch control unavailable")
                self._record("erta_startup_resources_and_controls", model=app.model_path_var.get(), ad_reference=app.ad_ref_path_var.get(), options_visible=False, batch_control_state=app.run_batch_button.cget("state"))
                app.toggle_options()
                app.update_idletasks()
                self._require(app.options_visible and bool(app.options_frame.grid_info()), "options did not open")
                self._require(app.model_path_var.get() and app.ad_ref_path_var.get(), "option resources not visible")
                app.toggle_options()
                app.update_idletasks()
                self._record("erta_options_and_resource_controls", model_entry=app.model_path_var.get(), ad_entry=app.ad_ref_path_var.get(), options_hidden_after_toggle=not app.options_visible and not bool(app.options_frame.grid_info()))
                self.stage = 1
            if self.stage == 1:
                if not app.predictor.is_loaded() or not app.ad_calculator.fitted:
                    self._reschedule()
                    return
                app.cas_var.set("")
                app.smiles_var.set(self.VALID_SMILES)
                app.single_predict_clicked()
                self.stage = 2
                self._reschedule()
                return
            if self.stage == 2:
                if not self._wait_erta_result():
                    self._reschedule()
                    return
                probability = float(app.active_probability_var.get().split(":", 1)[1].strip())
                self._require(abs(probability - 0.9801110625267029) < 1e-5, "ERTA exact valid oracle probability drift")
                self._require(app.ad_domain_var.get() in {"Applicability domain: In-domain", "Applicability domain: Out-of-domain"}, "ERTA AD was not rendered")
                self._require(bool(app.structure_label.cget("image")), "input molecule structure was not rendered")
                self._require(bool(app.nearest_reference_structure_label.cget("image")), "nearest reference structure was not rendered")
                self._require(bool(app.single_ad_graph_label.cget("image")), "single AD graph was not rendered")
                self._require(len(app.prob_canvas.find_all()) > 0, "probability graph canvas is empty")
                self.erta_summary = app.prediction_summary_var.get()
                self.erta_probability = app.active_probability_var.get()
                self._record("erta_exact_valid_single", probability_positive=probability, expected_probability_positive=0.9801110625267029, summary=self.erta_summary, ad=app.ad_domain_var.get(), nearest_reference=app.nearest_reference_var.get(), probability_canvas_items=len(app.prob_canvas.find_all()), input_structure=True, reference_structure=True, ad_graph=True)
                self.erta_previous_text = app.result_text.get("1.0", "end").strip()
                app.smiles_var.set(self.LEGACY_INVALID_SMILES)
                app.single_predict_clicked()
                self.stage = 3
                self._reschedule()
                return
            if self.stage == 3:
                if not self._wait_erta_result(self.erta_previous_text):
                    self._reschedule()
                    return
                detail = app.result_text.get("1.0", "end").strip()
                self._require("Mol valid: False" in detail, "legacy invalid SMILES did not reach ERTA zero-vector path")
                self._require(app.prediction_summary_var.get() in {"Prediction: Positive", "Prediction: Negative"}, "legacy invalid prediction label missing")
                self._record("erta_legacy_invalid_direct_smiles", smiles=self.LEGACY_INVALID_SMILES, rendered_summary=app.prediction_summary_var.get(), positive_probability=app.active_probability_var.get(), details=detail)
                self._install_mock_cas()
                app.cas_var.set(self.MOCK_CAS)
                app.smiles_var.set("")
                app.pubchem_clicked()
                self.stage = 4
                self._reschedule()
                return
            if self.stage == 4:
                if app.smiles_var.get() != self.MOCK_SMILES:
                    self._reschedule()
                    return
                self._require("PubChem found CID 712" in app.status_var.get(), "mocked PubChem callback did not update status")
                self._record("erta_mocked_pubchem_callback", cas=self.MOCK_CAS, smiles=app.smiles_var.get(), status=app.status_var.get())
                self.erta_previous_text = app.result_text.get("1.0", "end").strip()
                app.single_predict_clicked()
                self.stage = 5
                self._reschedule()
                return
            if self.stage == 5:
                if not self._wait_erta_result(self.erta_previous_text):
                    self._reschedule()
                    return
                self._require(f"CAS: {self.MOCK_CAS}" in app.result_text.get("1.0", "end"), "CAS-resolved prediction details missing CAS")
                self._record("erta_mocked_pubchem_prediction", summary=app.prediction_summary_var.get(), probability=app.active_probability_var.get(), canonical_details=app.result_text.get("1.0", "end").strip())
                self.stage = 6
            if self.stage == 6:
                template_path = self.run_root / "downloaded-template.xlsx"
                self._install_picker(template_path)
                model_before = app.model_path_var.get()
                ad_before = app.ad_ref_path_var.get()
                app.browse_model()
                app.browse_ad_reference()
                self._require(app.model_path_var.get() == model_before and app.ad_ref_path_var.get() == ad_before, "model or AD browse callback changed bundled resource selection")
                self._record("erta_model_ad_browse_callbacks", model=app.model_path_var.get(), ad_reference=app.ad_ref_path_var.get())
                dialogs_before = len(self.dialogs)
                app.download_template_clicked()
                self._require(template_path.is_file(), "template callback did not create workbook")
                self._require(self._wait_dialog(dialogs_before), "template callback did not expose completion dialog")
                template = load_workbook(template_path, read_only=True, data_only=True)
                try:
                    template_headers = [
                        cell.value for cell in template.active[1]
                    ]
                    self._require(
                        template.active.max_row == 1
                        and template_headers == ["CAS"],
                        "template schema drift",
                    )
                finally:
                    template.close()
                self._record(
                    "erta_template_callback_and_dialog",
                    output=str(template_path),
                    headers=template_headers,
                    dialog=self._terminal_dialog(
                        dialogs_before,
                        kind="info",
                        title="Template downloaded",
                        context="ERTA template download",
                    ),
                )
                self.erta_batch_input_path = (
                    self._copy_shared_example_for_batch("erta")
                )
                self.erta_batch_outputs_before = {
                    path.resolve()
                    for path in self.erta_batch_input_path.parent.glob(
                        "ERTA_*.xlsx"
                    )
                }
                app.batch_input_var.set(str(self.erta_batch_input_path))
                app.batch_input_display_var.set(
                    self.erta_batch_input_path.name
                )
                app.batch_destination_var.set(str(self.erta_batch_input_path.parent))
                self.batch_dialog_count = len(self.dialogs)
                app.batch_predict_clicked()
                self.erta_success_locked_controls = (
                    self._batch_control_states(app)
                )
                self._require(
                    app._batch_active
                    and self.erta_success_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERTA shared-example batch did not lock its controls",
                )
                self.stage = 7
                self._reschedule()
                return
            if self.stage == 7:
                if (
                    app._batch_active
                    or not self._wait_dialog(self.batch_dialog_count)
                ):
                    self._reschedule()
                    return
                self._require(
                    self.erta_batch_input_path is not None,
                    "ERTA batch input path was not recorded",
                )
                input_parent = self.erta_batch_input_path.parent
                outputs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERTA_*.xlsx"
                    )
                } - self.erta_batch_outputs_before
                self._require(
                    len(outputs) == 1,
                    f"expected one fresh ERTA workbook, found {len(outputs)}",
                )
                output = outputs.pop()
                self._require(
                    output.parent == input_parent
                    and output.name.startswith("ERTA_test_")
                    and output.name.endswith("_prediction.xlsx"),
                    "ERTA workbook path or ERTA_ filename contract drift",
                )
                self.erta_batch_output_path = output
                workbook = load_workbook(output, read_only=True, data_only=True)
                try:
                    sheet = workbook.active
                    headers = [cell.value for cell in sheet[1]]
                    rows = [
                        dict(zip(headers, row))
                        for row in sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    ]
                    self._require(
                        sheet.title == "Sheet1",
                        "legacy ERTA workbook sheet name drift",
                    )
                finally:
                    workbook.close()
                required_headers = {
                    "CAS",
                    "SMILES",
                    "Canonical_SMILES",
                    "Mol_valid",
                    "Probability_Negative_0",
                    "Probability_Positive_1",
                    "Prediction",
                    "Prediction_label",
                    "Decision_rule",
                    *self.ERBA_BATCH_AD_COLUMNS,
                    "PubChem_CID",
                    "PubChem_status",
                }
                self._require(
                    required_headers <= set(headers),
                    "ERTA shared-example workbook schema is incomplete",
                )
                expected_cas = _read_shared_example_workbook(
                    self.erta_batch_input_path
                )["cas_values"]
                self._require(
                    len(rows) == 25
                    and [str(row["CAS"]) for row in rows] == expected_cas,
                    "ERTA CARSRN alias or shared-example input order drift",
                )
                self._require(
                    all(
                        row["SMILES"] == self.MOCK_SMILES
                        and row["PubChem_status"] == "Found"
                        and row["Mol_valid"] is True
                        and row["Probability_Negative_0"] is not None
                        and row["Probability_Positive_1"] is not None
                        and row["Prediction_label"] in {"Positive", "Negative"}
                        for row in rows
                    ),
                    "ERTA shared CARSRN example did not yield 25 nonblank predictions",
                )
                graph_paths = sorted(
                    path.name
                    for path in (input_parent / "graphs").glob("*.png")
                )
                self._require(graph_paths, "ERTA batch graph files were not generated")
                summary = app.batch_result.get("1.0", "end").strip()
                self._require(
                    f"Output workbook: {output}" in summary
                    and "AD In-domain:" in summary
                    and "Graph directory:" in summary,
                    "ERTA batch completion summary is incomplete",
                )
                self._require(
                    not hasattr(app, "batch_tree")
                    and not hasattr(app, "graph_combo")
                    and not hasattr(app, "graph_label"),
                    "obsolete ERTA batch table/graph preview remains in the UI",
                )
                completion_dialog = self._terminal_dialog(
                    self.batch_dialog_count,
                    kind="info",
                    title="Batch prediction done",
                    context="ERTA shared-example batch success",
                )
                self._require(
                    str(output) in completion_dialog["message"]
                    and completion_dialog[
                        "batch_controls_at_dialog"
                    ]["erta"]
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    },
                    "ERTA success dialog does not report the published workbook",
                )
                self._require(
                    self._batch_control_states(app)
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and app.batch_progress_var.get()
                    == "100% - 25/25 - Complete"
                    and int(app.batch_progress_value.get()) == 100,
                    "ERTA batch controls were not restored after success",
                )
                self.erta_success_result_identity = id(app.last_batch_result)
                self._record(
                    "erta_shared_test_batch_success_dialog_and_predictions",
                    input=str(self.erta_batch_input_path),
                    output=str(output),
                    sheet=sheet.title,
                    headers=headers,
                    input_order=expected_cas,
                    input_header="CARSRN",
                    row_count=len(rows),
                    predicted_row_count=sum(
                        row["Probability_Positive_1"] is not None
                        for row in rows
                    ),
                    mocked_pubchem_statuses=sorted(
                        {row["PubChem_status"] for row in rows}
                    ),
                    deterministic_smiles_substitution=self.MOCK_SMILES,
                    source_bytes_preserved=(
                        _sha256_path(self.qa_shared_example_path)
                        == self.shared_bundle_sha256
                    ),
                    graph_files=graph_paths,
                    batch_summary=summary,
                    progress=app.batch_progress_var.get(),
                    controls_during_run=self.erta_success_locked_controls,
                    controls=self._batch_control_states(app),
                    completion_dialog=completion_dialog,
                )
                self.stage = 72
            if self.stage == 72:
                mixed_directory = self.run_root / "mixed-batch" / "erta"
                mixed_directory.mkdir(parents=True, exist_ok=False)
                mixed_input = mixed_directory / "erta-mixed-input.xlsx"
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(["No.", "CAS", "Chemical Name", "SMILES"])
                sheet.append([2, "", "valid", self.VALID_SMILES])
                sheet.append([1, self.MOCK_CAS, "mocked-cas", ""])
                sheet.append(
                    [3, "", "legacy-invalid", self.LEGACY_INVALID_SMILES]
                )
                workbook.save(mixed_input)
                workbook.close()
                self.erta_mixed_input_path = mixed_input.resolve(
                    strict=True
                )
                self.erta_mixed_outputs_before = {
                    path.resolve()
                    for path in mixed_directory.glob("ERTA_*.xlsx")
                }
                app.batch_input_var.set(str(self.erta_mixed_input_path))
                app.batch_input_display_var.set(
                    self.erta_mixed_input_path.name
                )
                app.batch_destination_var.set(str(mixed_directory))
                self.erta_mixed_dialog_count = len(self.dialogs)
                app.batch_predict_clicked()
                self.erta_mixed_locked_controls = (
                    self._batch_control_states(app)
                )
                self._require(
                    app._batch_active
                    and self.erta_mixed_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERTA mixed batch did not start with locked controls",
                )
                self.stage = 73
                self._reschedule()
                return
            if self.stage == 73:
                if (
                    app._batch_active
                    or not self._wait_dialog(
                        self.erta_mixed_dialog_count
                    )
                ):
                    self._reschedule()
                    return
                mixed_parent = self.erta_mixed_input_path.parent
                mixed_outputs = {
                    path.resolve()
                    for path in mixed_parent.glob("ERTA_*.xlsx")
                } - self.erta_mixed_outputs_before
                self._require(
                    len(mixed_outputs) == 1,
                    f"expected one fresh mixed ERTA workbook, found {len(mixed_outputs)}",
                )
                mixed_output = mixed_outputs.pop()
                self._require(
                    mixed_output.parent == mixed_parent
                    and mixed_output.name.startswith(
                        "ERTA_erta-mixed-input_"
                    ),
                    "mixed ERTA workbook path or filename prefix drift",
                )
                workbook = load_workbook(
                    mixed_output,
                    read_only=True,
                    data_only=True,
                )
                try:
                    sheet = workbook.active
                    mixed_headers = [
                        cell.value for cell in sheet[1]
                    ]
                    mixed_rows = list(
                        sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    )
                    self._require(
                        sheet.title == "Sheet1",
                        "legacy ERTA workbook sheet name drift",
                    )
                finally:
                    workbook.close()
                expected_headers = [
                    "No.",
                    "CAS",
                    "Chemical Name",
                    "SMILES",
                    "Canonical_SMILES",
                    "Mol_valid",
                    "Probability_Negative_0",
                    "Probability_Positive_1",
                    "Prediction",
                    "Prediction_label",
                    "Decision_rule",
                    *self.ERBA_BATCH_AD_COLUMNS,
                    "PubChem_CID",
                    "PubChem_status",
                ]
                self._require(
                    mixed_headers == expected_headers,
                    "legacy ERTA mixed workbook schema/order drift",
                )
                self._require(
                    [row[0] for row in mixed_rows] == [2, 1, 3],
                    "ERTA mixed batch input order drift",
                )
                self._require(
                    mixed_rows[1][3] == self.MOCK_SMILES
                    and mixed_rows[1][-1] == "Found",
                    "mocked CAS mixed-batch row was not resolved",
                )
                self._require(
                    mixed_rows[2][5] is False,
                    "legacy invalid mixed-batch row lost its invalid marker",
                )
                mixed_graphs = sorted(
                    path.name
                    for path in (mixed_parent / "graphs").glob("*.png")
                )
                mixed_summary = app.batch_result.get(
                    "1.0", "end"
                ).strip()
                self._require(
                    mixed_graphs
                    and f"Output workbook: {mixed_output}" in mixed_summary
                    and "AD In-domain:" in mixed_summary
                    and "Graph directory:" in mixed_summary,
                    "ERTA mixed batch artifacts or completion summary are incomplete",
                )
                mixed_dialog = self._terminal_dialog(
                    self.erta_mixed_dialog_count,
                    kind="info",
                    title="Batch prediction done",
                    context="ERTA mixed batch success",
                )
                restored_controls = self._batch_control_states(app)
                self._require(
                    restored_controls
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and app.batch_progress_var.get()
                    == "100% - 3/3 - Complete"
                    and int(app.batch_progress_value.get()) == 100,
                    "ERTA mixed batch did not restore controls and progress",
                )
                self._require(
                    mixed_dialog["batch_controls_at_dialog"]["erta"]
                    == restored_controls,
                    "ERTA mixed-batch controls were not restored before its dialog",
                )
                self.erta_success_result_identity = id(
                    app.last_batch_result
                )
                self._record(
                    "erta_batch_workbook_graphs_and_order",
                    input=str(self.erta_mixed_input_path),
                    output=str(mixed_output),
                    sheet=sheet.title,
                    headers=mixed_headers,
                    input_order=[row[0] for row in mixed_rows],
                    mocked_cas_status=mixed_rows[1][-1],
                    invalid_row_mol_valid=mixed_rows[2][5],
                    graph_files=mixed_graphs,
                    batch_summary=mixed_summary,
                    progress=app.batch_progress_var.get(),
                    controls_during_run=self.erta_mixed_locked_controls,
                    controls=restored_controls,
                    completion_dialog=mixed_dialog,
                )
                self.stage = 70
            if self.stage == 70:
                failure_input = self._make_corrupt_batch_input("erta")
                self.erta_failure_input_path = failure_input
                self.erta_failure_outputs_before = {
                    path.resolve()
                    for path in failure_input.parent.glob("ERTA_*.xlsx")
                }
                self.erta_failure_graphs_before = {
                    path.resolve()
                    for path in failure_input.parent.glob("graphs")
                }
                app.batch_input_var.set(str(failure_input))
                app.batch_input_display_var.set(failure_input.name)
                app.batch_destination_var.set(str(failure_input.parent))
                self.erta_failure_dialog_count = len(self.dialogs)
                app.batch_predict_clicked()
                self.erta_failure_locked_controls = (
                    self._batch_control_states(app)
                )
                self._require(
                    app._batch_active
                    and self.erta_failure_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERTA deterministic failure did not start with locked controls",
                )
                self.stage = 71
                self._reschedule()
                return
            if self.stage == 71:
                if (
                    app._batch_active
                    or not self._wait_dialog(
                        self.erta_failure_dialog_count
                    )
                ):
                    self._reschedule()
                    return
                failure_dialog = self._terminal_dialog(
                    self.erta_failure_dialog_count,
                    kind="error",
                    title="Batch prediction failed",
                    context="ERTA deterministic batch failure",
                )
                failure_summary = app.batch_result.get(
                    "1.0", "end"
                ).strip()
                new_outputs = {
                    path.resolve()
                    for path in self.erta_failure_input_path.parent.glob(
                        "ERTA_*.xlsx"
                    )
                } - self.erta_failure_outputs_before
                restored_controls = self._batch_control_states(app)
                self._require(
                    restored_controls
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and app.batch_progress_var.get()
                    == "100% - 0/0 - Failed"
                    and int(app.batch_progress_value.get()) == 100
                    and app.status_var.get()
                    in {"Error", "Batch prediction failed"}
                    and failure_summary.startswith(
                        "Batch prediction failed."
                    )
                    and "Batch prediction completed." not in failure_summary
                    and "Output workbook:" not in failure_summary
                    and not new_outputs
                    and {
                        path.resolve()
                        for path in self.erta_failure_input_path.parent.glob(
                            "graphs"
                        )
                    }
                    == self.erta_failure_graphs_before
                    and id(app.last_batch_result)
                    == self.erta_success_result_identity
                    and failure_dialog[
                        "batch_controls_at_dialog"
                    ]["erta"]
                    == restored_controls,
                    "ERTA failure exposed false success or did not restore controls",
                )
                self._record(
                    "erta_batch_failure_dialog_and_control_recovery",
                    input=str(self.erta_failure_input_path),
                    fixture_kind="intentionally corrupt xlsx bytes",
                    dialog=failure_dialog,
                    controls_during_run=self.erta_failure_locked_controls,
                    controls=restored_controls,
                    progress=app.batch_progress_var.get(),
                    status=app.status_var.get(),
                    result=failure_summary,
                    fresh_outputs=[],
                    fresh_graph_directories=[],
                    success_dialog_emitted=False,
                    prior_success_result_retained=True,
                )
                app.notebook.select(tab)
                self.stage = 8
            if self.stage == 8:
                task, subtype = self.routes[self.route_index]
                app.notebook.select(tab)
                if self.route_index == 0:
                    tab.mode_notebook.select(tab.single_tab)
                    app.update_idletasks()
                    evidence_details = tab.single_result_text.get(
                        "1.0",
                        "end",
                    ).strip()
                    expected_caveat = (
                        "Evidence caveat: "
                        f"{HISTORICAL_EXPOSURE_CAVEAT_COMPACT}"
                    )
                    self._require(
                        tab.single_result_text.winfo_ismapped()
                        and expected_caveat in evidence_details,
                        "ERalpha details pane does not visibly present the "
                        "compact evidence caveat",
                    )
                    self._record(
                        "erba_evidence_caveat_and_dialog",
                        presentation="single_prediction_details_pane",
                        compact_caveat=HISTORICAL_EXPOSURE_CAVEAT_COMPACT,
                        rendered_details=evidence_details,
                        details_box=_widget_geometry_receipt(
                            tab.single_result_text,
                            tab,
                        ),
                    )
                tab._route_changed("single")
                tab.smiles_var.set("Oc1ccccc1")
                tab.cas_var.set("")
                tab.single_predict_clicked()
                self.stage = 9
                self._reschedule()
                return
            if self.stage == 9:
                if (
                    tab._active_single_request_id is not None
                    or tab._active_ad_request_id is not None
                    or "Evaluating" in tab.single_ad_domain_var.get()
                ):
                    self._reschedule()
                    return
                body = self._body(tab)
                task, subtype = self.routes[self.route_index]
                self._require(tab.single_status_var.get() == "ok" and body, "ERBA route did not render success")
                self._require(
                    tab.single_ad_domain_var.get()
                    in (
                        "Applicability domain: In-domain",
                        "Applicability domain: Out-of-domain",
                    ),
                    "ERBA route did not complete applicability-domain analysis",
                )
                self._require(
                    bool(tab._single_ad_graph_image)
                    and "Morgan similarity=" in tab.single_nearest_reference_var.get(),
                    "ERBA route did not render AD graph and nearest reference",
                )
                self._record("erba_route", task=task.value, subtype=subtype.value, status=tab.single_status_var.get(), rendered_result=body)
                self.route_index += 1
                if self.route_index < len(self.routes):
                    self.stage = 8
                    self._reschedule()
                    return
                self.stage = 10
            if self.stage == 10:
                tab = app.eralpha_tab
                app.notebook.select(tab)
                name, smiles, cas, _expected = self.invalid_cases[self.invalid_index]
                tab.single_task_var.set(ERBATask.CLASSIFICATION.value)
                tab.single_subtype_var.set(ERBASubtype.ER_ALPHA.value)
                tab.smiles_var.set(smiles)
                tab.cas_var.set(cas)
                tab.single_predict_clicked()
                self.stage = 11
                self._reschedule()
                return
            if self.stage == 11:
                if tab._active_single_request_id is not None:
                    self._reschedule()
                    return
                body = self._body(tab)
                self._require(tab.single_status_var.get() != "ok", f"ERBA invalid {self.invalid_cases[self.invalid_index][0]} succeeded")
                self._record(f"erba_invalid_{self.invalid_cases[self.invalid_index][0]}", rendered_result=body, rendered_status=tab.single_status_var.get())
                self.invalid_index += 1
                if self.invalid_index < len(self.invalid_cases):
                    self.stage = 10
                    self._reschedule()
                    return
                self.stage = 12
            if self.stage == 12:
                app.notebook.select(tab)
                tab.notebook.select(tab.batch_tab)
                tab.batch_task_var.set(ERBATask.CLASSIFICATION.value)
                tab.batch_subtype_var.set(ERBASubtype.ER_ALPHA.value)
                self.erba_batch_input_path = (
                    self._copy_shared_example_for_batch("eralpha")
                )
                input_parent = self.erba_batch_input_path.parent
                self.erba_batch_outputs_before = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                }
                self.erba_batch_graph_dirs_before = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                }
                tab.batch_input_var.set(str(self.erba_batch_input_path))
                tab.batch_input_display_var.set(
                    self.erba_batch_input_path.name
                )
                tab.batch_destination_var.set(str(input_parent))
                self.erba_batch_dialog_count = len(self.dialogs)
                tab.batch_predict_clicked()
                self.erba_success_locked_controls = (
                    self._batch_control_states(tab)
                )
                self._require(
                    tab._active_batch_request_id is not None
                    and self.erba_success_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERalpha shared-example batch did not start with locked "
                    f"controls: {tab.batch_status_var.get()}",
                )
                self.stage = 13
                self._reschedule()
                return
            if self.stage == 13:
                if (
                    tab._active_batch_request_id is not None
                    or not self._wait_dialog(
                        self.erba_batch_dialog_count
                    )
                ):
                    self._reschedule()
                    return
                self._require(
                    self.erba_batch_input_path is not None,
                    "ERalpha batch input path was not recorded",
                )
                input_parent = self.erba_batch_input_path.parent
                outputs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                } - self.erba_batch_outputs_before
                self._require(
                    len(outputs) == 1,
                    f"expected one new ERalpha workbook, found {len(outputs)}: "
                    f"{tab.batch_status_var.get()}",
                )
                output = outputs.pop()
                self._require(
                    output.parent == input_parent
                    and tab.batch_destination_var.get() == str(input_parent),
                    "ERalpha workbook destination is not the input workbook parent",
                )
                graph_dirs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                } - self.erba_batch_graph_dirs_before
                self._require(
                    len(graph_dirs) == 1,
                    f"expected one new input-adjacent ERalpha graph directory, found {len(graph_dirs)}",
                )
                graph_directory = graph_dirs.pop()
                graph_files = {
                    path.name for path in graph_directory.iterdir() if path.is_file()
                }
                self._require(
                    graph_directory.parent == input_parent
                    and graph_files == self.ERBA_BATCH_GRAPH_FILES,
                    "ERalpha batch graphs are incomplete or outside the input parent",
                )

                workbook = load_workbook(
                    output,
                    read_only=True,
                    data_only=True,
                )
                try:
                    sheet_names = workbook.sheetnames
                    self._require(
                        sheet_names
                        == ["Predictions", "Guide", "Input", "Metadata"]
                        and workbook.active.title == "Predictions",
                        "ERalpha batch workbook is not Predictions-first and active",
                    )
                    prediction_sheet = workbook["Predictions"]
                    prediction_headers = [
                        cell.value for cell in prediction_sheet[1]
                    ]
                    self._require(
                        all(
                            column in prediction_headers
                            for column in self.ERBA_BATCH_AD_COLUMNS
                        ),
                        "ERalpha batch workbook is missing applicability-domain columns",
                    )
                    prediction_rows = [
                        dict(zip(prediction_headers, row))
                        for row in prediction_sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    ]
                    input_sheet = workbook["Input"]
                    input_headers = [
                        cell.value for cell in input_sheet[1]
                    ]
                    input_rows = list(
                        input_sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    )
                    guide_sheet = workbook["Guide"]
                    guide_rows = list(
                        guide_sheet.iter_rows(values_only=True)
                    )
                    metadata_sheet = workbook["Metadata"]
                    metadata_headers = [
                        cell.value for cell in metadata_sheet[1]
                    ]
                    metadata_values = [
                        cell.value
                        for cell in next(
                            metadata_sheet.iter_rows(
                                min_row=2,
                                max_row=2,
                            )
                        )
                    ]
                    metadata = dict(
                        zip(metadata_headers, metadata_values)
                    )
                finally:
                    workbook.close()

                expected_cas = _read_shared_example_workbook(
                    self.erba_batch_input_path
                )["cas_values"]
                self._require(
                    input_headers == ["CARSRN"]
                    and input_rows
                    == [(cas,) for cas in expected_cas]
                    and len(input_rows) == 25,
                    "ERalpha Input sheet did not preserve the supplied CARSRN workbook",
                )
                self._require(
                    len(prediction_rows) == 25
                    and [str(row["CARSRN"]) for row in prediction_rows]
                    == expected_cas
                    and [str(row["CAS"]) for row in prediction_rows]
                    == expected_cas,
                    "ERalpha primary predictions lost CARSRN input order or CAS alias mapping",
                )
                self._require(
                    all(
                        row["Result_Status"] == "Predicted"
                        and row["binding_label"]
                        in {"binding", "non_binding"}
                        and row["non_binding_probability"] is not None
                        and row["binding_probability"] is not None
                        and row["AD"] in {"In-domain", "Out-of-domain"}
                        for row in prediction_rows
                    ),
                    "ERalpha shared CARSRN example did not yield 25 nonblank predictions",
                )
                self._require(
                    guide_rows
                    and guide_rows[0] == ("Topic", "Details")
                    and any(
                        row[0] == "ERalpha batch prediction guide"
                        for row in guide_rows[1:]
                    ),
                    "ERalpha Guide sheet content is missing",
                )

                expected_spec = tab.catalog[
                    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
                ]
                hex_digits = set("0123456789abcdef")
                self._require(
                    metadata.get("Excel_Contract_ID")
                    == ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID
                    and metadata.get("Model_ID") == expected_spec.model_id
                    and metadata.get("Model_SHA256") == expected_spec.sha256
                    and metadata.get("Performance_Evidence_Scope")
                    == "internal_historically_exposed"
                    and bool(metadata.get("Evidence_Caveat"))
                    and all(
                        isinstance(metadata.get(column), str)
                        and len(metadata[column]) == 64
                        and set(metadata[column]) <= hex_digits
                        for column in self.ERBA_PROVENANCE_COLUMNS
                    ),
                    "ERalpha batch metadata did not preserve catalog model provenance",
                )

                binding_count = sum(
                    row["binding_label"] == "binding" for row in prediction_rows
                )
                non_binding_count = sum(
                    row["binding_label"] == "non_binding" for row in prediction_rows
                )
                not_predicted_count = sum(
                    row["Result_Status"] != "Predicted" for row in prediction_rows
                )
                self._require(
                    binding_count + non_binding_count == 25
                    and not_predicted_count == 0,
                    "ERalpha shared example contains blank or not-predicted results",
                )
                summary = tab.batch_result.get("1.0", "end").strip()
                self._require(
                    tab.batch_status_var.get() == f"ERBA batch complete: {output}"
                    and f"Total rows: {len(prediction_rows)}" in summary
                    and f"Binding: {binding_count}" in summary
                    and f"Non-binding: {non_binding_count}" in summary
                    and f"Not predicted: {not_predicted_count}" in summary
                    and f"Output workbook: {output}" in summary
                    and f"Graph files: {len(graph_files)}" in summary
                    and f"Graph directory: {graph_directory}" in summary,
                    "ERalpha batch completion did not expose binding and artifact summaries",
                )
                completion_dialog = self._terminal_dialog(
                    self.erba_batch_dialog_count,
                    kind="info",
                    title="Batch prediction done",
                    context="ERalpha shared-example batch success",
                )
                self._require(
                    str(output) in completion_dialog["message"]
                    and "Not predicted: 0" in completion_dialog["message"]
                    and completion_dialog[
                        "batch_controls_at_dialog"
                    ]["eralpha"]
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    },
                    "ERalpha success dialog does not report the complete publication",
                )
                restored_controls = self._batch_control_states(tab)
                self._require(
                    restored_controls
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and tab.batch_progress_var.get()
                    == "100% - 25/25 - Completed"
                    and int(tab.batch_progress_value.get()) == 100,
                    "ERalpha batch controls were not restored after success",
                )
                self._require(
                    self.erta_batch_output_path is not None,
                    "ERTA primary workbook is unavailable for style comparison",
                )
                formatting_parity = compare_plain_primary_workbook_formatting(
                    self.erta_batch_output_path,
                    output,
                )
                self._record(
                    "eralpha_shared_test_batch_success_dialog_and_predictions",
                    input=str(self.erba_batch_input_path),
                    output=str(output),
                    sheets=sheet_names,
                    prediction_headers=prediction_headers,
                    preserved_input_rows=[list(row) for row in input_rows],
                    input_header="CARSRN",
                    metadata=metadata,
                    graph_directory=str(graph_directory),
                    graph_files=sorted(graph_files),
                    binding_count=binding_count,
                    non_binding_count=non_binding_count,
                    not_predicted_count=not_predicted_count,
                    batch_summary=summary,
                    progress=tab.batch_progress_var.get(),
                    completion_dialog=completion_dialog,
                    controls_during_run=self.erba_success_locked_controls,
                    controls=restored_controls,
                    plain_primary_formatting=formatting_parity,
                    deterministic_smiles_substitution=self.MOCK_SMILES,
                    source_bytes_preserved=(
                        _sha256_path(self.qa_shared_example_path)
                        == self.shared_bundle_sha256
                    ),
                )
                self.stage = 132
            if self.stage == 132:
                mixed_directory = (
                    self.run_root / "mixed-batch" / "eralpha"
                )
                mixed_directory.mkdir(parents=True, exist_ok=False)
                mixed_input = (
                    mixed_directory / "erba-alpha-batch-input.xlsx"
                )
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Original_Input"
                sheet.append(["Row_ID", "SMILES", "Analyst_Note"])
                sheet.append(
                    ["native-valid", "Oc1ccccc1", "preserve-valid"]
                )
                sheet.append(
                    ["native-invalid", "[*]CC", "preserve-invalid"]
                )
                workbook.save(mixed_input)
                workbook.close()
                self.erba_mixed_input_path = mixed_input.resolve(
                    strict=True
                )
                self.erba_mixed_outputs_before = {
                    path.resolve()
                    for path in mixed_directory.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                }
                self.erba_mixed_graph_dirs_before = {
                    path.resolve()
                    for path in mixed_directory.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                }
                tab.batch_input_var.set(str(self.erba_mixed_input_path))
                tab.batch_input_display_var.set(
                    self.erba_mixed_input_path.name
                )
                tab.batch_destination_var.set(str(mixed_directory))
                self.erba_mixed_dialog_count = len(self.dialogs)
                tab.batch_predict_clicked()
                self.erba_mixed_locked_controls = (
                    self._batch_control_states(tab)
                )
                self._require(
                    tab._active_batch_request_id is not None
                    and self.erba_mixed_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERalpha mixed batch did not start with locked controls",
                )
                self.stage = 133
                self._reschedule()
                return
            if self.stage == 133:
                if (
                    tab._active_batch_request_id is not None
                    or not self._wait_dialog(
                        self.erba_mixed_dialog_count
                    )
                ):
                    self._reschedule()
                    return
                mixed_parent = self.erba_mixed_input_path.parent
                mixed_outputs = {
                    path.resolve()
                    for path in mixed_parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                } - self.erba_mixed_outputs_before
                self._require(
                    len(mixed_outputs) == 1,
                    "expected one new ERalpha mixed-result workbook, found "
                    f"{len(mixed_outputs)}: {tab.batch_status_var.get()}",
                )
                mixed_output = mixed_outputs.pop()
                self._require(
                    mixed_output.parent == mixed_parent
                    and tab.batch_destination_var.get()
                    == str(mixed_parent),
                    "ERalpha mixed workbook destination is not its input parent",
                )
                mixed_graph_dirs = {
                    path.resolve()
                    for path in mixed_parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                } - self.erba_mixed_graph_dirs_before
                self._require(
                    len(mixed_graph_dirs) == 1,
                    "expected one new ERalpha mixed graph directory, found "
                    f"{len(mixed_graph_dirs)}",
                )
                mixed_graph_directory = mixed_graph_dirs.pop()
                mixed_graph_files = {
                    path.name
                    for path in mixed_graph_directory.iterdir()
                    if path.is_file()
                }
                self._require(
                    mixed_graph_directory.parent == mixed_parent
                    and mixed_graph_files
                    == self.ERBA_BATCH_GRAPH_FILES,
                    "ERalpha mixed-batch graphs are incomplete or misplaced",
                )
                workbook = load_workbook(
                    mixed_output,
                    read_only=True,
                    data_only=True,
                )
                try:
                    sheet_names = workbook.sheetnames
                    self._require(
                        sheet_names
                        == ["Predictions", "Guide", "Input", "Metadata"]
                        and workbook.active.title == "Predictions",
                        "ERalpha mixed workbook is not Predictions-first and active",
                    )
                    prediction_sheet = workbook["Predictions"]
                    prediction_headers = [
                        cell.value for cell in prediction_sheet[1]
                    ]
                    prediction_rows = [
                        dict(zip(prediction_headers, row))
                        for row in prediction_sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    ]
                    input_sheet = workbook["Input"]
                    input_headers = [
                        cell.value for cell in input_sheet[1]
                    ]
                    input_rows = list(
                        input_sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    )
                    guide_rows = list(
                        workbook["Guide"].iter_rows(values_only=True)
                    )
                    metadata_sheet = workbook["Metadata"]
                    metadata_headers = [
                        cell.value for cell in metadata_sheet[1]
                    ]
                    metadata_values = [
                        cell.value
                        for cell in next(
                            metadata_sheet.iter_rows(
                                min_row=2,
                                max_row=2,
                            )
                        )
                    ]
                    metadata = dict(
                        zip(metadata_headers, metadata_values)
                    )
                finally:
                    workbook.close()
                self._require(
                    all(
                        column in prediction_headers
                        for column in self.ERBA_BATCH_AD_COLUMNS
                    ),
                    "ERalpha mixed workbook is missing AD columns",
                )
                self._require(
                    len(prediction_rows) == 2
                    and prediction_rows[0]["Analyst_Note"]
                    == "preserve-valid"
                    and prediction_rows[1]["Analyst_Note"]
                    == "preserve-invalid",
                    "ERalpha mixed Predictions lost passthrough fields or order",
                )
                self._require(
                    prediction_rows[0]["Result_Status"] == "Predicted"
                    and prediction_rows[0]["AD"]
                    in {"In-domain", "Out-of-domain"}
                    and prediction_rows[1]["Result_Status"]
                    == "Not predicted"
                    and prediction_rows[1]["Reason_Category"]
                    == "Invalid or ambiguous SMILES"
                    and prediction_rows[1]["Reason_Description"]
                    and prediction_rows[1]["Recommended_Action"]
                    and prediction_rows[1][
                        "non_binding_probability"
                    ]
                    is None
                    and prediction_rows[1]["binding_probability"]
                    is None
                    and prediction_rows[1]["binding_label"] is None
                    and all(
                        prediction_rows[1][column] is None
                        for column in self.ERBA_BATCH_AD_COLUMNS
                    ),
                    "ERalpha mixed predicted/not-predicted row mapping is incorrect",
                )
                self._require(
                    input_headers
                    == ["Row_ID", "SMILES", "Analyst_Note"]
                    and input_rows
                    == [
                        (
                            "native-valid",
                            "Oc1ccccc1",
                            "preserve-valid",
                        ),
                        (
                            "native-invalid",
                            "[*]CC",
                            "preserve-invalid",
                        ),
                    ],
                    "ERalpha mixed Input sheet did not preserve exact source rows",
                )
                self._require(
                    any(
                        row[0] == "Not predicted rows"
                        and row[1] == "1"
                        for row in guide_rows
                    )
                    and any(
                        row[0]
                        == "Reason: Invalid or ambiguous SMILES"
                        and row[1] == "1"
                        for row in guide_rows
                    ),
                    "ERalpha mixed Guide does not explain the not-predicted row",
                )
                expected_spec = tab.catalog[
                    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
                ]
                hex_digits = set("0123456789abcdef")
                self._require(
                    metadata.get("Excel_Contract_ID")
                    == ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID
                    and metadata.get("Model_ID")
                    == expected_spec.model_id
                    and metadata.get("Model_SHA256")
                    == expected_spec.sha256
                    and metadata.get("Performance_Evidence_Scope")
                    == "internal_historically_exposed"
                    and bool(metadata.get("Evidence_Caveat"))
                    and all(
                        isinstance(metadata.get(column), str)
                        and len(metadata[column]) == 64
                        and set(metadata[column]) <= hex_digits
                        for column in self.ERBA_PROVENANCE_COLUMNS
                    ),
                    "ERalpha mixed Metadata lost catalog provenance",
                )
                binding_count = sum(
                    row["binding_label"] == "binding"
                    for row in prediction_rows
                )
                non_binding_count = sum(
                    row["binding_label"] == "non_binding"
                    for row in prediction_rows
                )
                not_predicted_count = sum(
                    row["Result_Status"] != "Predicted"
                    for row in prediction_rows
                )
                mixed_summary = tab.batch_result.get(
                    "1.0", "end"
                ).strip()
                self._require(
                    tab.batch_status_var.get()
                    == f"ERBA batch complete: {mixed_output}"
                    and f"Total rows: {len(prediction_rows)}"
                    in mixed_summary
                    and f"Binding: {binding_count}" in mixed_summary
                    and f"Non-binding: {non_binding_count}"
                    in mixed_summary
                    and "Not predicted: 1" in mixed_summary
                    and f"Output workbook: {mixed_output}"
                    in mixed_summary
                    and f"Graph files: {len(mixed_graph_files)}"
                    in mixed_summary
                    and f"Graph directory: {mixed_graph_directory}"
                    in mixed_summary,
                    "ERalpha mixed completion summary is incomplete",
                )
                warning_dialog = self._terminal_dialog(
                    self.erba_mixed_dialog_count,
                    kind="warning",
                    title="Batch prediction completed with warnings",
                    context="ERalpha mixed batch partial success",
                )
                restored_controls = self._batch_control_states(tab)
                self._require(
                    "Not predicted: 1" in warning_dialog["message"]
                    and restored_controls
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and tab.batch_progress_var.get()
                    == "100% - 2/2 - Completed"
                    and int(tab.batch_progress_value.get()) == 100
                    and warning_dialog[
                        "batch_controls_at_dialog"
                    ]["eralpha"]
                    == restored_controls,
                    "ERalpha mixed warning or control-restoration contract drift",
                )
                self._record(
                    "erba_eralpha_batch_workbook_ad_graphs_and_summary",
                    input=str(self.erba_mixed_input_path),
                    output=str(mixed_output),
                    sheets=sheet_names,
                    prediction_headers=prediction_headers,
                    preserved_input_rows=[
                        list(row) for row in input_rows
                    ],
                    metadata=metadata,
                    graph_directory=str(mixed_graph_directory),
                    graph_files=sorted(mixed_graph_files),
                    binding_count=binding_count,
                    non_binding_count=non_binding_count,
                    not_predicted_count=not_predicted_count,
                    invalid_reason_category=prediction_rows[1][
                        "Reason_Category"
                    ],
                    batch_summary=mixed_summary,
                    completion_dialog=warning_dialog,
                    progress=tab.batch_progress_var.get(),
                    controls_during_run=self.erba_mixed_locked_controls,
                    controls=restored_controls,
                )
                self.stage = 130
            if self.stage == 130:
                failure_input = self._make_corrupt_batch_input("eralpha")
                self.erba_failure_input_path = failure_input
                self.erba_failure_outputs_before = {
                    path.resolve()
                    for path in failure_input.parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                }
                self.erba_failure_graph_dirs_before = {
                    path.resolve()
                    for path in failure_input.parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                }
                tab.batch_input_var.set(str(failure_input))
                tab.batch_input_display_var.set(failure_input.name)
                tab.batch_destination_var.set(str(failure_input.parent))
                self.erba_failure_dialog_count = len(self.dialogs)
                tab.batch_predict_clicked()
                self.erba_failure_locked_controls = (
                    self._batch_control_states(tab)
                )
                self._require(
                    tab._active_batch_request_id is not None
                    and self.erba_failure_locked_controls
                    == {
                        "input": "disabled",
                        "template": "disabled",
                        "run": "disabled",
                    },
                    "ERalpha deterministic failure did not start with locked controls",
                )
                self.stage = 131
                self._reschedule()
                return
            if self.stage == 131:
                if (
                    tab._active_batch_request_id is not None
                    or not self._wait_dialog(
                        self.erba_failure_dialog_count
                    )
                ):
                    self._reschedule()
                    return
                failure_dialog = self._terminal_dialog(
                    self.erba_failure_dialog_count,
                    kind="error",
                    title="Batch prediction failed",
                    context="ERalpha deterministic batch failure",
                )
                failure_summary = tab.batch_result.get(
                    "1.0", "end"
                ).strip()
                new_outputs = {
                    path.resolve()
                    for path in self.erba_failure_input_path.parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                } - self.erba_failure_outputs_before
                new_graph_directories = {
                    path.resolve()
                    for path in self.erba_failure_input_path.parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                } - self.erba_failure_graph_dirs_before
                restored_controls = self._batch_control_states(tab)
                self._require(
                    restored_controls
                    == {
                        "input": "normal",
                        "template": "normal",
                        "run": "normal",
                    }
                    and tab.batch_progress_var.get()
                    == "100% - 0/0 - Failed"
                    and int(tab.batch_progress_value.get()) == 100
                    and tab.batch_status_var.get().startswith(
                        "ERBA batch failed:"
                    )
                    and failure_summary.startswith(
                        "Batch prediction failed."
                    )
                    and "ERalpha batch prediction completed."
                    not in failure_summary
                    and "Output workbook:" not in failure_summary
                    and not new_outputs
                    and not new_graph_directories
                    and failure_dialog[
                        "batch_controls_at_dialog"
                    ]["eralpha"]
                    == restored_controls,
                    "ERalpha failure exposed false success or did not restore controls",
                )
                self._record(
                    "eralpha_batch_failure_dialog_and_control_recovery",
                    input=str(self.erba_failure_input_path),
                    fixture_kind="intentionally corrupt xlsx bytes",
                    dialog=failure_dialog,
                    controls_during_run=self.erba_failure_locked_controls,
                    controls=restored_controls,
                    progress=tab.batch_progress_var.get(),
                    status=tab.batch_status_var.get(),
                    result=failure_summary,
                    fresh_outputs=[],
                    fresh_graph_directories=[],
                    success_dialog_emitted=False,
                )
                self.stage = 14
            if self.stage == 14:
                self._require(
                    app.eralpha_tab._selected_route("single")
                    == (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA),
                    "ERalpha page is not fixed to classification",
                )
                alpha_model = app.eralpha_tab.catalog[
                    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
                ].model_id
                self._require(
                    not hasattr(app, "erbeta_tab")
                    and app.notebook.tabs()
                    and "ERbeta"
                    not in [
                        app.notebook.tab(tab_id, "text")
                        for tab_id in app.notebook.tabs()
                    ],
                    "ERbeta remains exposed in the FDA application",
                )
                self._record(
                    "fixed_eralpha_classification_tab",
                    eralpha_model=alpha_model,
                    erbeta_present=False,
                )
                app.notebook.select(app.erta_tab)
                self._require(app.prediction_summary_var.get() != self.erta_summary or app.active_probability_var.get() != self.erta_probability, "ERTA state was not independently updated after CAS journey")
                self._require(app.last_batch_result is not None and app.notebook.tab(app.notebook.select(), "text") == "ERTA", "ERTA batch state lost after ERBA journey")
                self._record("erta_erba_erta_state_isolation", erta_summary_after_return=app.prediction_summary_var.get(), erta_probability_after_return=app.active_probability_var.get(), erba_result_retained=bool(self._body(app.eralpha_tab)), erta_batch_rows=len(app.last_batch_result))
                self._require(
                    self.shared_default_path is not None
                    and self.shared_default_path.is_file()
                    and _sha256_path(self.shared_default_path)
                    == self.shared_default_sha256,
                    "native QA changed the resolved shared default workbook",
                )
                self._require(
                    self.shared_bundle_path is not None
                    and self.shared_bundle_path.is_file()
                    and _sha256_path(self.shared_bundle_path)
                    == self.shared_bundle_sha256
                    and self.qa_shared_example_path is not None
                    and _sha256_path(self.qa_shared_example_path)
                    == self.shared_bundle_sha256
                    and self.erta_batch_input_path is not None
                    and _sha256_path(self.erta_batch_input_path)
                    == self.shared_bundle_sha256
                    and self.erba_batch_input_path is not None
                    and _sha256_path(self.erba_batch_input_path)
                    == self.shared_bundle_sha256,
                    "bundled test.xlsx or a QA input copy changed during prediction",
                )
                self._record(
                    "read_only_install_and_writable_state",
                    resource_root=str(app.project_root),
                    output_root=str(app.output_root),
                    state_root=str(app.state_root),
                    qa_run_root=str(self.run_root),
                    shared_example={
                        "resolved_default": str(self.shared_default_path),
                        "resolved_default_sha256": self.shared_default_sha256,
                        "resolved_default_bytes_unchanged": True,
                        "bundled_source": str(self.shared_bundle_path),
                        "bundled_sha256": self.shared_bundle_sha256,
                        "bundled_bytes_unchanged": True,
                        "fresh_first_run_copy": str(
                            self.qa_shared_example_path
                        ),
                        "erta_batch_copy": str(
                            self.erta_batch_input_path
                        ),
                        "eralpha_batch_copy": str(
                            self.erba_batch_input_path
                        ),
                        "all_qa_copies_byte_match_bundle": True,
                    },
                )
                check_names = [
                    receipt["check"]
                    for receipt in self.transcript["checks"]
                ]
                required_checks = (
                    self.BASELINE_CHECKS
                    | self.SHARED_EXAMPLE_BATCH_CHECKS
                    | self.BATCH_FEEDBACK_CHECKS
                )
                self._require(
                    required_checks <= set(check_names)
                    and len(self.BASELINE_CHECKS) == 22
                    and len(check_names)
                    == 26 + int(bool(self.external_driver_gate)),
                    "native QA check inventory does not retain the 22 baseline "
                    "checks plus shared-example and batch-failure contracts",
                )
                self.transcript["check_contract"] = {
                    "baseline_check_count": 22,
                    "baseline_checks": sorted(self.BASELINE_CHECKS),
                    "added_shared_example_batch_check_count": 2,
                    "added_shared_example_batch_checks": sorted(
                        self.SHARED_EXAMPLE_BATCH_CHECKS
                    ),
                    "added_batch_feedback_check_count": 2,
                    "added_batch_feedback_checks": sorted(
                        self.BATCH_FEEDBACK_CHECKS
                    ),
                    "external_driver_check_present": bool(
                        self.external_driver_gate
                    ),
                    "total_check_count": len(check_names),
                }
                app.deiconify()
                app.state("normal")
                app.update()
                self._finish(True)
        except Exception:
            self._finish(False, traceback.format_exc())
        finally:
            self._tick_active = False
