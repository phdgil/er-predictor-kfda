# -*- mode: python ; coding: utf-8 -*-
"""Standalone external QA runner; intentionally contains no ER_Predictor modules."""
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

PROJECT_ROOT = Path(SPECPATH).parent.resolve()
RUNNER = PROJECT_ROOT / "qa" / "external_qa_runner.py"

analysis = Analysis(
    [str(RUNNER)],
    pathex=[str(PROJECT_ROOT / "qa")],
    binaries=[],
    datas=[],
    hiddenimports=collect_submodules("pywinauto"),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    [],
    name="ER_Predictor_QA_Runner",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)
