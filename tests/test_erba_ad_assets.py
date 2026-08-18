from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd
import pytest

from core.ad import ADCalculator
from core.fingerprint import RDKIT_COLUMNS, detect_rdkit_columns
from prepare_erba_ad_assets import OUTPUT_ROOT, ROUTES, build_route


EXPECTED = {
    "classification/er_alpha": {"rows": 1481, "label": "label", "subtype": "ERalpha"},
}


def _reference(route) -> Path:
    return OUTPUT_ROOT / route.filename


def test_released_routes_are_exact_and_route_specific():
    assert {route.key for route in ROUTES} == set(EXPECTED)
    assert len(ROUTES) == 1
    assert len({route.filename for route in ROUTES}) == 1

    for route in ROUTES:
        expected = EXPECTED[route.key]
        reference = _reference(route)
        frame = pd.read_excel(reference)
        assert len(frame) == expected["rows"]
        assert frame["receptor_subtype"].eq(expected["subtype"]).all()
        assert expected["label"] in frame.columns
        assert frame[expected["label"]].notna().all()
        if route.key.startswith("classification/"):
            assert frame["partition"].eq("development").all()
        assert detect_rdkit_columns(frame) == RDKIT_COLUMNS
        assert len([column for column in frame.columns if str(column).startswith("RDKit_")]) == 2048
        assert Path(ADCalculator.adjacent_cache_path(str(reference))).is_file()


def test_cache_loads_only_when_bound_to_its_reference():
    for route in ROUTES:
        reference = _reference(route)
        cache = Path(ADCalculator.adjacent_cache_path(str(reference)))
        calculator = ADCalculator().load_cache(str(cache), reference_path=str(reference))
        assert calculator.fitted
        assert calculator.train_x_scaled.shape == (EXPECTED[route.key]["rows"], 2048)


def test_cache_rejects_modified_reference(tmp_path):
    route = ROUTES[0]
    reference = _reference(route)
    cache = Path(ADCalculator.adjacent_cache_path(str(reference)))
    changed_reference = tmp_path / reference.name
    shutil.copyfile(reference, changed_reference)
    frame = pd.read_excel(changed_reference)
    frame.loc[0, "label"] = 1 - int(frame.loc[0, "label"])
    frame.to_excel(changed_reference, index=False)

    with pytest.raises(ValueError, match="SHA-256"):
        ADCalculator().load_cache(str(cache), reference_path=str(changed_reference))


def test_builder_rejects_unexpected_route():
    with pytest.raises(ValueError, match="Unsupported"):
        build_route(type("Route", (), {"key": "classification/er_gamma"})())
