#!/usr/bin/env python3
"""Assert the docs graph evaluator formulas match tiny_dsa.Model defaults.

Mirrors ``assets/graph/app.js`` ``evaluate()`` in Python so CI does not need Node.
Keep this file in lockstep with the JS when the identity or defaults change.
"""
from __future__ import annotations

import sys
from typing import Any


YEARS = (1, 2, 3, 4, 5)
SHOCK_PARAMS = ("Growth", "Interest", "Primary balance")

DEFAULTS: dict[str, Any] = {
    "country_name": "Borvelia",
    "country_initial_debt": {"Borvelia": 60.0, "Litellia": 80.0, "Aurelium": 40.0},
    "growth_baseline": {1: 3.5, 2: 3.5, 3: 3.5, 4: 3.5, 5: 3.5},
    "interest_baseline": {1: 4.0, 2: 4.0, 3: 4.0, 4: 4.0, 5: 4.0},
    "primary_balance_baseline": {1: -1.0, 2: -0.5, 3: 0.0, 4: 0.5, 5: 1.0},
    "shock_year": 2,
    "shock_type": 1,
    "shock_magnitudes": {"Growth": -2.0, "Interest": 2.0, "Primary balance": -1.0},
}


def _choose(shock_type: int, a: float, b: float, c: float) -> float:
    if shock_type == 1:
        return a
    if shock_type == 2:
        return b
    return c


def _debt_step(prev: float, r: float, g: float, pb: float) -> float:
    return (prev * (1 + r / 100.0)) / (1 + g / 100.0) - pb


def evaluate_js_mirror(inputs: dict[str, Any]) -> dict[str, Any]:
    """Python port of assets/graph/app.js evaluate()."""
    initial_debt_resolved = float(inputs["country_initial_debt"][inputs["country_name"]])
    shock_key = SHOCK_PARAMS[inputs["shock_type"] - 1]
    shock_magnitude_resolved = float(inputs["shock_magnitudes"][shock_key])

    shock_active: dict[int, int] = {}
    shocked_growth: dict[int, float] = {}
    shocked_interest: dict[int, float] = {}
    shocked_primary_balance: dict[int, float] = {}
    baseline_path_internal: dict[int, float] = {}
    shocked_path_internal: dict[int, float] = {}

    for y in YEARS:
        active = 1 if y >= inputs["shock_year"] else 0
        shock_active[y] = active
        shocked_growth[y] = inputs["growth_baseline"][y] + _choose(
            inputs["shock_type"], shock_magnitude_resolved, 0.0, 0.0
        ) * active
        shocked_interest[y] = inputs["interest_baseline"][y] + _choose(
            inputs["shock_type"], 0.0, shock_magnitude_resolved, 0.0
        ) * active
        shocked_primary_balance[y] = inputs["primary_balance_baseline"][y] + _choose(
            inputs["shock_type"], 0.0, 0.0, shock_magnitude_resolved
        ) * active

        if y == 1:
            baseline_path_internal[y] = _debt_step(
                initial_debt_resolved,
                inputs["interest_baseline"][y],
                inputs["growth_baseline"][y],
                inputs["primary_balance_baseline"][y],
            )
            shocked_path_internal[y] = _debt_step(
                initial_debt_resolved,
                shocked_interest[y],
                shocked_growth[y],
                shocked_primary_balance[y],
            )
        else:
            baseline_path_internal[y] = _debt_step(
                baseline_path_internal[y - 1],
                inputs["interest_baseline"][y],
                inputs["growth_baseline"][y],
                inputs["primary_balance_baseline"][y],
            )
            shocked_path_internal[y] = _debt_step(
                shocked_path_internal[y - 1],
                shocked_interest[y],
                shocked_growth[y],
                shocked_primary_balance[y],
            )

    output_baseline = dict(baseline_path_internal)
    output_shocked = dict(shocked_path_internal)
    output_delta = {y: output_shocked[y] - output_baseline[y] for y in YEARS}

    return {
        "initial_debt_resolved": initial_debt_resolved,
        "shock_magnitude_resolved": shock_magnitude_resolved,
        "shock_active": shock_active,
        "shocked_growth": shocked_growth,
        "baseline_path_internal": baseline_path_internal,
        "shocked_path_internal": shocked_path_internal,
        "output_baseline": output_baseline,
        "output_shocked": output_shocked,
        "output_delta": output_delta,
    }


def _series_map(series) -> dict[int, float]:
    out: dict[int, float] = {}
    for key, value in series.items():
        year = key[0] if isinstance(key, tuple) else key
        out[int(year)] = float(value)
    return out


def python_model_goldens() -> dict[str, Any]:
    from tiny_dsa.model import Model

    m = Model.from_defaults()
    return {
        "initial_debt_resolved": float(m.initial_debt_resolved),
        "shock_magnitude_resolved": float(m.shock_magnitude_resolved),
        "shock_active": _series_map(m.shock_active),
        "shocked_growth": _series_map(m.shocked_growth),
        "baseline_path_internal": _series_map(m.baseline_path_internal),
        "shocked_path_internal": _series_map(m.shocked_path_internal),
        "output_baseline": _series_map(m.output_baseline),
        "output_shocked": _series_map(m.output_shocked),
        "output_delta": _series_map(m.output_delta),
    }


def _close(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def compare(left: dict[str, Any], right: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for key, pv in left.items():
        jv = right.get(key)
        if isinstance(pv, dict):
            if not isinstance(jv, dict):
                errors.append(f"{key}: expected map, got {jv!r}")
                continue
            for k, v in pv.items():
                if k not in jv or not _close(float(v), float(jv[k])):
                    errors.append(f"{key}[{k}]: model={v} mirror={jv.get(k)}")
        else:
            if jv is None or not _close(float(pv), float(jv)):
                errors.append(f"{key}: model={pv} mirror={jv}")
    return errors


def main() -> int:
    mirror = evaluate_js_mirror(DEFAULTS)
    model = python_model_goldens()
    errors = compare(model, mirror)
    if errors:
        print("graph evaluator mirror diverges from tiny_dsa.Model:", file=sys.stderr)
        for err in errors:
            print(f"  {err}", file=sys.stderr)
        return 1
    print("ok: assets/graph evaluate() formulas match Model.from_defaults()")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
