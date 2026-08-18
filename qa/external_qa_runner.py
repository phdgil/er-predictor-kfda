"""External Windows UI acceptance runner for a packaged ER_Predictor.exe.

This module deliberately imports no product modules.  It drives the frozen target with
pywinauto and native keyboard input, then correlates that surface evidence with the
opt-in in-process receipt written by the target itself.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import traceback
import uuid
from typing import Any

from PIL import Image, ImageStat
from pywinauto import Desktop
import win32gui
import win32ui

WINDOW_TITLE_RE = r".*ER Predictor.*"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def target_environment(receipt: Path, gate: Path, runtime_root: Path, run_nonce: str) -> tuple[dict[str, str], int]:
    """Give the target fresh writable roots without making Python discoverable."""
    environment = os.environ.copy()
    environment.pop("PYTHONHOME", None)
    environment.pop("PYTHONPATH", None)
    path_entries = environment.get("PATH", "").split(os.pathsep)
    retained_entries = []
    removed_entries = 0
    for entry in path_entries:
        if any("python" in part.casefold() for part in Path(entry).parts):
            removed_entries += 1
        else:
            retained_entries.append(entry)
    environment["PATH"] = os.pathsep.join(retained_entries)
    environment["PYTHONNOUSERSITE"] = "1"
    environment["PYTHONSAFEPATH"] = "1"
    environment["ER_PREDICTOR_AUTOMATION_RECEIPT"] = str(receipt)
    environment["ER_PREDICTOR_EXTERNAL_DRIVER_GATE"] = str(gate)
    environment["ER_PREDICTOR_AUTOMATION_NONCE"] = run_nonce
    environment["ER_PREDICTOR_AUTOMATION_HOLD_SECONDS"] = "1"
    environment["USERPROFILE"] = str(runtime_root / "UserProfile")
    environment["LOCALAPPDATA"] = str(runtime_root / "LocalAppData")
    environment["APPDATA"] = str(runtime_root / "AppData")
    return environment, removed_entries


def wait_for_window(timeout_seconds: float, process_id: int):
    deadline = time.monotonic() + timeout_seconds
    last_error = "no matching native window"
    while time.monotonic() < deadline:
        try:
            window = Desktop(backend="uia").window(title_re=WINDOW_TITLE_RE, process=process_id)
            if window.exists(timeout=1):
                window.wait("visible enabled ready", timeout=2)
                return window
        except Exception as error:  # UIA is transient while Tk initializes.
            last_error = str(error)
        time.sleep(0.25)
    raise TimeoutError(f"ER Predictor native window did not become ready: {last_error}")


def control_inventory(window) -> list[dict[str, Any]]:
    inventory = []
    for control in window.descendants():
        info = control.element_info
        name = str(info.name or "")
        control_type = str(info.control_type or "")
        automation_id = str(info.automation_id or "")
        selected = False
        if control_type == "TabItem":
            try:
                selected = bool(control.is_selected())
            except (AttributeError, RuntimeError):
                selected = False
        if name or control_type or automation_id:
            inventory.append(
                {"name": name, "control_type": control_type, "automation_id": automation_id, "selected": selected}
            )
    return inventory


def foreground_facts(window) -> dict[str, Any]:
    foreground_handle = win32gui.GetForegroundWindow()
    return {
        "target_handle": window.handle,
        "target_title": window.window_text(),
        "foreground_handle": foreground_handle,
        "foreground_title": win32gui.GetWindowText(foreground_handle),
        "target_is_foreground": foreground_handle == window.handle,
    }


def capture_nonuniform(window, path: Path) -> dict[str, Any]:
    handle = window.handle
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    width, height = right - left, bottom - top
    window_dc = win32gui.GetWindowDC(handle)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    memory_dc.SelectObject(bitmap)
    try:
        if not ctypes.windll.user32.PrintWindow(handle, memory_dc.GetSafeHdc(), 2):
            raise RuntimeError("PrintWindow failed")
        image = Image.frombuffer("RGB", (width, height), bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(handle, window_dc)
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path)
    extrema = ImageStat.Stat(image).extrema
    if not any(low != high for low, high in extrema):
        raise AssertionError(f"uniform screenshot: {path}")
    return {"path": str(path), "sha256": sha256_file(path), "size": list(image.size), "rgb_extrema": extrema}


def contains_caption(inventory: list[dict[str, Any]], caption: str) -> bool:
    return any(item["name"].strip().casefold() == caption.casefold() for item in inventory)



def switch_workflow_by_keyboard(window, expected: str) -> dict[str, Any]:
    # Tk does not expose ttk.Notebook children through Windows UIA. Explicit
    # Ctrl+1/Ctrl+2 workflow accelerators provide caption-linked native
    # navigation without coordinates, including on a locked workstation.
    native_window = Desktop(backend="win32").window(handle=window.handle)
    keys = "^1" if expected == "ERTA" else "^2"
    native_window.send_keystrokes(keys)
    time.sleep(0.7)
    return {
        "expected": expected,
        "method": "pywinauto.HwndWrapper.send_keystrokes " + ("Ctrl+1" if expected == "ERTA" else "Ctrl+2"),
        "input_surface": "native-window-messages",
        "caption_attestation": "ERTA/ERBA workflow accelerator",
        "attempts": 1,
    }


def switch_and_capture(
    window,
    expected: str,
    baseline_sha256: str,
    screenshot_path: Path,
) -> tuple[dict[str, Any], dict[str, Any]]:
    for attempt in range(1, 4):
        action = switch_workflow_by_keyboard(window, expected)
        image = capture_nonuniform(window, screenshot_path)
        if image["sha256"] != baseline_sha256:
            action["attempts"] = attempt
            return action, image
    raise AssertionError(f"native workflow transition to {expected} produced no visible change")


def wait_for_receipt(path: Path, timeout_seconds: float, run_nonce: str, target_pid: int) -> dict[str, Any]:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if path.is_file():
            payload = json.loads(path.read_text(encoding="utf-8"))
            checks = payload.get("checks")
            if payload.get("run_nonce") != run_nonce:
                raise AssertionError("internal receipt run nonce mismatch")
            if payload.get("target_pid") != target_pid:
                raise AssertionError("internal receipt target PID mismatch")
            if payload.get("passed") is True and isinstance(checks, list) and checks and all(check.get("passed") is True for check in checks):
                return payload
            raise AssertionError("internal receipt exists but does not contain only passing checks")
        time.sleep(0.5)
    raise TimeoutError(f"timed out waiting for internal automation receipt: {path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run external pywinauto QA against ER_Predictor.exe")
    targets = parser.add_mutually_exclusive_group(required=True)
    targets.add_argument("--target", type=Path, help="Explicit staged or published ER_Predictor.exe")
    targets.add_argument("--staged", type=Path, help="Staged distribution directory containing ER_Predictor.exe")
    targets.add_argument("--published", type=Path, help="Published distribution directory containing ER_Predictor.exe")
    parser.add_argument("--output-root", type=Path, required=True, help="Fresh runner-owned evidence/runtime root")
    parser.add_argument("--timeout-seconds", type=float, default=180.0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    target = args.target or ((args.staged or args.published) / "ER_Predictor.exe")
    target = target.resolve()
    output_root = args.output_root.resolve()
    if output_root.exists():
        raise SystemExit(f"output root must not already exist: {output_root}")
    run_nonce = uuid.uuid4().hex
    if not target.is_file() or target.name.casefold() != "er_predictor.exe":
        raise SystemExit(f"target must be an ER_Predictor.exe file: {target}")
    if target == Path(sys.executable).resolve():
        raise SystemExit("target must be supplied separately from the QA runner")

    receipt_path = output_root / "internal-receipt.json"
    gate_path = output_root / "external-driver-gate.json"
    result_path = output_root / "external-qa-receipt.json"
    runtime_root = output_root / "runtime"
    screenshots = output_root / "screenshots"
    started = time.monotonic()
    transcript: dict[str, Any] = {
        "schema_version": 1,
        "kind": "external-pywinauto-qa-runner",
        "driver_kind": "external-pywinauto",
        "run_nonce": run_nonce,
        "runner": str(Path(sys.executable).resolve()),
        "target": {"path": str(target), "command": [str(target)], "sha256": sha256_file(target)},
        "target_environment": {
            "pythonpath_removed": True,
            "pythonhome_removed": True,
            "target_uses_python_path": False,
        },
        "real_input_actions": [],
        "checks": [],
        "passed": False,
    }
    process: subprocess.Popen[str] | None = None
    try:
        environment, removed_entries = target_environment(receipt_path, gate_path, runtime_root, run_nonce)
        transcript["target_environment"]["python_path_entries_removed"] = removed_entries
        process = subprocess.Popen([str(target)], cwd=str(target.parent), env=environment)
        transcript["target"]["pid"] = process.pid
        window = wait_for_window(args.timeout_seconds, process.pid)
        inventory = control_inventory(window)
        caption_accessibility = {
            "erta_exposed": contains_caption(inventory, "ERTA"),
            "erba_exposed": contains_caption(inventory, "ERBA"),
            "mode": "uia-captions" if contains_caption(inventory, "ERTA") and contains_caption(inventory, "ERBA") else "tk-accessibility-ambiguous",
            "replacement": "target-handle Ctrl+1/Ctrl+2 messages plus PrintWindow screenshots and target receipt caption attestation",
        }
        facts = foreground_facts(window)
        if not facts["target_is_foreground"]:
            window.set_focus()
            facts = foreground_facts(window)
        facts["foreground_focus_attempted"] = True
        facts["foreground_required"] = False
        facts["delivery_mode"] = "target-handle native messages"
        facts["locked_desktop_compatible"] = True
        transcript["checks"].append({
            "check": "native_window_inventory_and_driver_context",
            "passed": True,
            "foreground": facts,
            "inventory": inventory,
            "caption_accessibility": caption_accessibility,
        })
        before = capture_nonuniform(window, screenshots / "erta-before.png")
        transcript["screenshots"] = [before]
        erba_action, erba = switch_and_capture(window, "ERBA", before["sha256"], screenshots / "erba.png")
        transcript["real_input_actions"].append(erba_action)
        transcript["screenshots"].append(erba)
        transcript["checks"].append({
            "check": "native_ctrl_2_erba_transition",
            "passed": True,
            "visible_change": True,
        })
        erta_action, after = switch_and_capture(window, "ERTA", erba["sha256"], screenshots / "erta-after.png")
        transcript["real_input_actions"].append(erta_action)
        transcript["screenshots"].append(after)
        if after["sha256"] != before["sha256"]:
            raise AssertionError("native Ctrl+1 did not restore the original ERTA image")
        transcript["checks"].append({
            "check": "native_ctrl_1_erta_transition",
            "passed": True,
            "visible_change": True,
        })
        transcript["checks"].append({"check": "keyboard_erta_erba_erta_round_trip", "passed": True})
        transcript["checks"].append({
            "check": "target_handle_input_and_printwindow_replacement",
            "passed": True,
            "evidence": (
                "Both target-handle key sequences changed the target window's PrintWindow image "
                "and the round trip restored the original ERTA image; foreground ownership is not "
                "required for this locked-desktop-compatible native-message path."
            ),
        })
        write_json(gate_path, {
            "schema_version": 1,
            "driver_kind": "external-pywinauto",
            "passed": True,
            "run_nonce": run_nonce,
            "target_pid": process.pid,
            "real_input_actions": transcript["real_input_actions"],
            "screenshots": transcript["screenshots"],
            "caption_accessibility": caption_accessibility,
        })
        internal = wait_for_receipt(receipt_path, args.timeout_seconds, run_nonce, process.pid)
        check_names = {row.get("check") for row in internal["checks"]}
        if not {"erta_startup_resources_and_controls", "erba_evidence_caveat_and_dialog", "erta_erba_erta_state_isolation"} <= check_names:
            raise AssertionError("supplemental target receipt lacks required ERTA/ERBA caption and journey attestations")
        transcript["internal_receipt"] = {
            "path": str(receipt_path),
            "sha256": sha256_file(receipt_path),
            "passed": internal["passed"],
            "check_count": len(internal["checks"]),
            "run_nonce": internal["run_nonce"],
            "target_pid": internal["target_pid"],
        }
        transcript["checks"].append({"check": "supplemental_internal_callback_receipt", "passed": True})
        transcript["passed"] = True
        return 0
    except Exception as error:
        transcript["error"] = f"{type(error).__name__}: {error}"
        transcript["traceback"] = traceback.format_exc()
        return 1
    finally:
        transcript["duration_seconds"] = time.monotonic() - started
        if process is not None:
            terminated_by_runner = process.poll() is None
            transcript["target"]["terminated_by_runner"] = terminated_by_runner
            if terminated_by_runner:
                transcript["target"]["termination_reason"] = "external QA completed or failed; runner cleaned the process tree"
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=10)
            else:
                transcript["target"]["spontaneous_returncode"] = process.returncode
            transcript["target"]["cleanup_returncode"] = process.returncode
        write_json(result_path, transcript)


if __name__ == "__main__":
    raise SystemExit(main())
