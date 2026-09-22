"""Pedagogical series-graph API for the docs dependency-graph viz.

Recomputes via the exported ``Model`` (default) or excel-grapher's
``FormulaEvaluator``. JSON shapes match ``assets/graph/app.js`` defaults /
evaluate output (flat key→value maps, not tuple coordinates).
"""

from __future__ import annotations

from typing import Any, Mapping

from . import data
from .graph_schema import (
    EDGES,
    INPUT_IDS,
    NODES,
    SERIES_IDS,
    BackendName,
    axes,
)
from .model import Model
from .tensor import Series

JsonValue = Any
FlatInputs = dict[str, JsonValue]
FlatValues = dict[str, JsonValue]


class GraphApiError(Exception):
    """Structured failure for HTTP / callers (validation or missing backend)."""

    def __init__(
        self,
        message: str,
        *,
        status: int = 400,
        errors: dict[str, str] | None = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.status = status
        self.errors = errors or {}

    def as_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"error": self.message}
        if self.errors:
            payload["errors"] = self.errors
        return payload


def _unwrap_key(key: object) -> object:
    if isinstance(key, tuple) and len(key) == 1:
        return key[0]
    return key


def flatten_series(value: object) -> JsonValue:
    """Series / scalar → JSON-friendly scalar or flat map."""
    if isinstance(value, Series):
        return {_unwrap_key(coord): _json_number(item) for coord, item in value.items()}
    if hasattr(value, "items") and hasattr(value, "domain"):
        return {
            _unwrap_key(coord): _json_number(item)
            for coord, item in value.items()  # type: ignore[union-attr]
        }
    return _json_number(value)


def _json_number(value: object) -> JsonValue:
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, float):
        return float(value)
    try:
        as_float = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return value
    if as_float.is_integer():
        return int(as_float)
    return as_float


def flatten_defaults() -> FlatInputs:
    """Canonical workbook defaults in viz-flat shape."""
    return {
        "country_name": data.COUNTRY_NAME_DEFAULT,
        "country_initial_debt": flatten_series(data.COUNTRY_INITIAL_DEBT_DEFAULT),
        "growth_baseline": flatten_series(data.GROWTH_BASELINE_DEFAULT),
        "interest_baseline": flatten_series(data.INTEREST_BASELINE_DEFAULT),
        "primary_balance_baseline": flatten_series(data.PRIMARY_BALANCE_BASELINE_DEFAULT),
        "shock_year": data.SHOCK_YEAR_DEFAULT,
        "shock_type": data.SHOCK_TYPE_DEFAULT,
        "shock_magnitudes": flatten_series(data.SHOCK_MAGNITUDES_DEFAULT),
    }


def _normalize_map_keys(raw: Mapping[Any, Any], *, int_keys: bool) -> dict[Any, Any]:
    out: dict[Any, Any] = {}
    for key, value in raw.items():
        if int_keys:
            out[int(key)] = float(value)
        else:
            out[str(key) if not isinstance(key, str) else key] = (
                float(value) if not isinstance(value, str) else value
            )
    return out


def normalize_inputs(raw: Mapping[str, Any] | None) -> FlatInputs:
    """Accept JSON inputs (stringified year keys ok); fill missing from defaults."""
    base = flatten_defaults()
    if not raw:
        return base
    unknown = set(raw) - set(INPUT_IDS)
    if unknown:
        raise GraphApiError(
            f"unknown inputs: {sorted(unknown)}",
            errors={name: "unknown input" for name in sorted(unknown)},
        )
    merged = dict(base)
    if "country_name" in raw:
        merged["country_name"] = str(raw["country_name"])
    if "shock_year" in raw:
        merged["shock_year"] = int(raw["shock_year"])
    if "shock_type" in raw:
        merged["shock_type"] = int(raw["shock_type"])
    if "country_initial_debt" in raw:
        merged["country_initial_debt"] = _normalize_map_keys(
            raw["country_initial_debt"], int_keys=False
        )
        # country keys stay strings; coerce values
        merged["country_initial_debt"] = {
            str(k): float(v) for k, v in merged["country_initial_debt"].items()
        }
    if "growth_baseline" in raw:
        merged["growth_baseline"] = _normalize_map_keys(raw["growth_baseline"], int_keys=True)
    if "interest_baseline" in raw:
        merged["interest_baseline"] = _normalize_map_keys(raw["interest_baseline"], int_keys=True)
    if "primary_balance_baseline" in raw:
        merged["primary_balance_baseline"] = _normalize_map_keys(
            raw["primary_balance_baseline"], int_keys=True
        )
    if "shock_magnitudes" in raw:
        merged["shock_magnitudes"] = {
            str(k): float(v) for k, v in raw["shock_magnitudes"].items()
        }
    return merged


def _ordered_values(template: Series, flat: Mapping[Any, Any]) -> tuple[Any, ...]:
    values: list[Any] = []
    for coord in template.domain:
        key = _unwrap_key(coord)
        if key not in flat:
            raise GraphApiError(
                f"missing key {key!r} for {template.schema.series_id}",
                errors={template.schema.series_id: f"missing key {key!r}"},
            )
        values.append(flat[key])
    return tuple(values)


def bind_model_inputs(flat: FlatInputs) -> dict[str, Any]:
    """Flat viz inputs → keyword args for ``Model`` / ``from_defaults``."""
    return {
        "country_name": flat["country_name"],
        "country_initial_debt": data.COUNTRY_INITIAL_DEBT_DEFAULT.with_values(
            _ordered_values(data.COUNTRY_INITIAL_DEBT_DEFAULT, flat["country_initial_debt"])
        ),
        "growth_baseline": data.GROWTH_BASELINE_DEFAULT.with_values(
            _ordered_values(data.GROWTH_BASELINE_DEFAULT, flat["growth_baseline"])
        ),
        "interest_baseline": data.INTEREST_BASELINE_DEFAULT.with_values(
            _ordered_values(data.INTEREST_BASELINE_DEFAULT, flat["interest_baseline"])
        ),
        "primary_balance_baseline": data.PRIMARY_BALANCE_BASELINE_DEFAULT.with_values(
            _ordered_values(
                data.PRIMARY_BALANCE_BASELINE_DEFAULT, flat["primary_balance_baseline"]
            )
        ),
        "shock_year": flat["shock_year"],
        "shock_type": flat["shock_type"],
        "shock_magnitudes": data.SHOCK_MAGNITUDES_DEFAULT.with_values(
            _ordered_values(data.SHOCK_MAGNITUDES_DEFAULT, flat["shock_magnitudes"])
        ),
    }


def available_backends() -> list[BackendName]:
    backends: list[BackendName] = ["export"]
    try:
        from . import graph_formula_evaluator as _fe

        if _fe.is_available():
            backends.append("formula_evaluator")
    except Exception:
        pass
    return backends


def evaluate_export(inputs: Mapping[str, Any] | None = None) -> FlatValues:
    """Recompute every series via the exported ``Model``."""
    flat = normalize_inputs(inputs)
    try:
        model = Model(**bind_model_inputs(flat))
    except (TypeError, ValueError) as exc:
        raise GraphApiError(str(exc), errors={"_model": str(exc)}) from exc

    values: FlatValues = {
        "country_name": flat["country_name"],
        "country_initial_debt": dict(flat["country_initial_debt"]),
        "growth_baseline": dict(flat["growth_baseline"]),
        "interest_baseline": dict(flat["interest_baseline"]),
        "primary_balance_baseline": dict(flat["primary_balance_baseline"]),
        "shock_year": flat["shock_year"],
        "shock_type": flat["shock_type"],
        "shock_magnitudes": dict(flat["shock_magnitudes"]),
    }
    for series_id in SERIES_IDS:
        if series_id in values:
            continue
        values[series_id] = flatten_series(getattr(model, series_id))
    return values


def evaluate_formula_evaluator(inputs: Mapping[str, Any] | None = None) -> FlatValues:
    """Recompute every series via excel-grapher ``FormulaEvaluator``."""
    from . import graph_formula_evaluator as fe

    if not fe.is_available():
        raise GraphApiError(
            "formula_evaluator backend requires excel-grapher and the Tiny DSA workbook fixture",
            status=503,
        )
    flat = normalize_inputs(inputs)
    try:
        return fe.evaluate(flat)
    except GraphApiError:
        raise
    except Exception as exc:
        raise GraphApiError(str(exc), status=503, errors={"_formula_evaluator": str(exc)}) from exc


def evaluate(
    inputs: Mapping[str, Any] | None = None,
    *,
    backend: BackendName = "export",
) -> FlatValues:
    if backend == "export":
        return evaluate_export(inputs)
    if backend == "formula_evaluator":
        return evaluate_formula_evaluator(inputs)
    raise GraphApiError(f"unknown backend: {backend!r}", status=400)


def bootstrap(*, backend: BackendName = "export") -> dict[str, Any]:
    """Schema + defaults + initial values for the viz."""
    defaults = flatten_defaults()
    return {
        "axes": axes(),
        "defaults": defaults,
        "nodes": list(NODES),
        "edges": [list(edge) for edge in EDGES],
        "values": evaluate(defaults, backend=backend),
        "backend": backend,
        "backends": available_backends(),
    }


__all__ = [
    "GraphApiError",
    "available_backends",
    "bind_model_inputs",
    "bootstrap",
    "evaluate",
    "evaluate_export",
    "evaluate_formula_evaluator",
    "flatten_defaults",
    "flatten_series",
    "normalize_inputs",
]
