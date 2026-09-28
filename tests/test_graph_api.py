"""Tests for the pedagogical series-graph API."""

from __future__ import annotations

import math
from typing import Any

import pytest

from tiny_dsa.graph_api import (
    GraphApiError,
    bootstrap,
    evaluate,
    evaluate_export,
    flatten_defaults,
)
from tiny_dsa.graph_schema import EDGES, INPUT_IDS, NODES, SERIES_IDS
from tiny_dsa.model import Model

try:
    from tiny_dsa import graph_formula_evaluator as fe

    _FE_AVAILABLE = fe.is_available()
except Exception:
    _FE_AVAILABLE = False

ATOL = 1e-6


def _close(a: float, b: float, tol: float = ATOL) -> bool:
    return abs(float(a) - float(b)) <= tol


def _series_map(series: Any) -> dict[Any, float]:
    out: dict[Any, float] = {}
    for key, value in series.items():
        year = key[0] if isinstance(key, tuple) else key
        out[year] = float(value)
    return out


def test_bootstrap_export_shape() -> None:
    payload = bootstrap(backend="export")
    assert set(payload["axes"]) == {"years", "countries", "shock_params"}
    assert payload["axes"]["years"] == [1, 2, 3, 4, 5]
    assert len(payload["nodes"]) == len(NODES)
    assert len(payload["edges"]) == len(EDGES)
    assert payload["backend"] == "export"
    assert "export" in payload["backends"]
    assert set(payload["defaults"]) == set(INPUT_IDS)
    assert set(payload["values"]) == set(SERIES_IDS)


def test_evaluate_export_matches_model_defaults() -> None:
    values = evaluate_export()
    model = Model.from_defaults()
    assert _close(values["initial_debt_resolved"], float(model.initial_debt_resolved))
    assert _close(
        values["shock_magnitude_resolved"], float(model.shock_magnitude_resolved)
    )
    for series_id in (
        "shock_active",
        "shocked_growth",
        "baseline_path_internal",
        "shocked_path_internal",
        "output_baseline",
        "output_shocked",
        "output_delta",
    ):
        expected = _series_map(getattr(model, series_id))
        got = values[series_id]
        assert set(got) == set(expected)
        for key, exp in expected.items():
            assert _close(got[key], exp), f"{series_id}[{key}]"


def test_evaluate_export_shock_changes_delta() -> None:
    defaults = flatten_defaults()
    baseline = evaluate_export(defaults)
    shocked = evaluate_export(
        {
            **defaults,
            "shock_magnitudes": {
                **defaults["shock_magnitudes"],
                "Growth": -5.0,
            },
        }
    )
    assert any(
        not _close(baseline["output_delta"][y], shocked["output_delta"][y])
        for y in (2, 3, 4, 5)
    )
    # Year 1 is before shock_year=2
    assert _close(shocked["output_delta"][1], 0.0)


def test_evaluate_rejects_unknown_input() -> None:
    with pytest.raises(GraphApiError) as caught:
        evaluate({"not_a_field": 1})
    assert caught.value.status == 400
    assert "not_a_field" in caught.value.errors


def test_normalize_string_year_keys() -> None:
    values = evaluate(
        {
            "growth_baseline": {"1": 3.5, "2": 3.5, "3": 3.5, "4": 3.5, "5": 3.5},
        }
    )
    assert 1 in values["growth_baseline"]
    assert _close(values["growth_baseline"][1], 3.5)


@pytest.mark.skipif(not _FE_AVAILABLE, reason="excel-grapher / workbook fixture unavailable")
def test_formula_evaluator_defaults_match_export() -> None:
    fe.reset_driver()
    export_values = evaluate(backend="export")
    graph_values = evaluate(backend="formula_evaluator")
    for series_id in SERIES_IDS:
        left = export_values[series_id]
        right = graph_values[series_id]
        if isinstance(left, dict):
            assert set(left) == set(right), series_id
            for key, lv in left.items():
                if isinstance(lv, (int, float)) and isinstance(right[key], (int, float)):
                    assert _close(lv, right[key]), f"{series_id}[{key}]"
                else:
                    assert left[key] == right[key], f"{series_id}[{key}]"
        elif isinstance(left, (int, float)) and isinstance(right, (int, float)):
            assert _close(left, right), series_id
        else:
            assert left == right, series_id


@pytest.mark.skipif(not _FE_AVAILABLE, reason="excel-grapher / workbook fixture unavailable")
def test_formula_evaluator_input_change() -> None:
    fe.reset_driver()
    defaults = flatten_defaults()
    before = evaluate(defaults, backend="formula_evaluator")
    after = evaluate(
        {**defaults, "country_name": "Litellia"},
        backend="formula_evaluator",
    )
    assert _close(before["initial_debt_resolved"], 60.0)
    assert _close(after["initial_debt_resolved"], 80.0)
    assert not math.isclose(
        float(before["output_baseline"][5]),
        float(after["output_baseline"][5]),
        abs_tol=ATOL,
    )
