"""Series-level topology for the interactive dependency-graph viz.

Nodes and cell addresses are derived from `data` where possible so the
workbook provenance stays single-sourced. Edges are authored to match the
series DAG drawn by `assets/graph/app.js`.
"""

from __future__ import annotations

from typing import Any, Literal, Mapping

from . import data

YEARS: tuple[int, ...] = tuple(data.TIME_PERIOD_AXIS.keys)
COUNTRIES: tuple[str, ...] = tuple(data.COUNTRY_AXIS.keys)
SHOCK_PARAMS: tuple[str, ...] = tuple(data.SHOCK_PARAMETER_AXIS.keys)

INPUT_IDS: tuple[str, ...] = (
    "country_name",
    "country_initial_debt",
    "growth_baseline",
    "interest_baseline",
    "primary_balance_baseline",
    "shock_year",
    "shock_type",
    "shock_magnitudes",
)

SERIES_IDS: tuple[str, ...] = INPUT_IDS + (
    "initial_debt_resolved",
    "engine_initial_debt_baseline",
    "engine_initial_debt_shocked",
    "shock_magnitude_resolved",
    "shock_active",
    "shocked_growth",
    "shocked_interest",
    "shocked_primary_balance",
    "baseline_path_internal",
    "shocked_path_internal",
    "output_baseline",
    "output_shocked",
    "output_delta",
)

# Producer → consumer, matching assets/graph/app.js SERIES_EDGES.
EDGES: tuple[tuple[str, str], ...] = (
    ("country_name", "initial_debt_resolved"),
    ("country_initial_debt", "initial_debt_resolved"),
    ("initial_debt_resolved", "engine_initial_debt_baseline"),
    ("initial_debt_resolved", "engine_initial_debt_shocked"),
    ("shock_type", "shock_magnitude_resolved"),
    ("shock_magnitudes", "shock_magnitude_resolved"),
    ("shock_year", "shock_active"),
    ("growth_baseline", "shocked_growth"),
    ("shock_type", "shocked_growth"),
    ("shock_magnitude_resolved", "shocked_growth"),
    ("shock_active", "shocked_growth"),
    ("interest_baseline", "shocked_interest"),
    ("shock_type", "shocked_interest"),
    ("shock_magnitude_resolved", "shocked_interest"),
    ("shock_active", "shocked_interest"),
    ("primary_balance_baseline", "shocked_primary_balance"),
    ("shock_type", "shocked_primary_balance"),
    ("shock_magnitude_resolved", "shocked_primary_balance"),
    ("shock_active", "shocked_primary_balance"),
    ("engine_initial_debt_baseline", "baseline_path_internal"),
    ("growth_baseline", "baseline_path_internal"),
    ("interest_baseline", "baseline_path_internal"),
    ("primary_balance_baseline", "baseline_path_internal"),
    ("engine_initial_debt_shocked", "shocked_path_internal"),
    ("shocked_growth", "shocked_path_internal"),
    ("shocked_interest", "shocked_path_internal"),
    ("shocked_primary_balance", "shocked_path_internal"),
    ("baseline_path_internal", "output_baseline"),
    ("shocked_path_internal", "output_shocked"),
    ("output_baseline", "output_delta"),
    ("output_shocked", "output_delta"),
)

BackendName = Literal["export", "formula_evaluator"]


def _flat_addresses(cells: Mapping[tuple[object, ...], str] | Mapping[Any, str]) -> dict[Any, str]:
    """Unwrap single-axis coordinate tuples to flat JSON keys."""
    out: dict[Any, str] = {}
    for coord, address in cells.items():
        if isinstance(coord, tuple) and len(coord) == 1:
            out[coord[0]] = address
        elif coord == ():
            continue
        else:
            out[coord] = address
    return out


def _scalar_address(cells: Mapping[Any, str]) -> str:
    return cells[()]


def _domain(min_value: float, max_value: float) -> dict[str, float]:
    return {"min": min_value, "max": max_value}


def _sheet_of(address: str) -> str:
    return address.split("!", 1)[0]


def _node(
    *,
    series_id: str,
    role: str,
    kind: str,
    keys: list[Any],
    address: str | None = None,
    addresses: dict[Any, str] | None = None,
    domain: dict[str, float] | None = None,
    options: list[Any] | None = None,
    option_labels: dict[Any, str] | None = None,
) -> dict[str, Any]:
    if address is not None:
        sample = address
    elif addresses:
        sample = next(iter(addresses.values()))
    else:
        raise ValueError(f"{series_id} needs address or addresses")
    node: dict[str, Any] = {
        "id": series_id,
        "role": role,
        "label": series_id,
        "sheet": _sheet_of(sample),
        "kind": kind,
        "keys": keys,
    }
    if address is not None:
        node["address"] = address
    if addresses is not None:
        node["addresses"] = addresses
    if domain is not None:
        node["domain"] = domain
    if options is not None:
        node["options"] = options
    if option_labels is not None:
        node["optionLabels"] = option_labels
    return node


NODES: tuple[dict[str, Any], ...] = (
    _node(
        series_id="country_name",
        role="input",
        kind="enum",
        keys=[None],
        address=_scalar_address(data.COUNTRY_NAME_CELLS),
        options=list(COUNTRIES),
    ),
    _node(
        series_id="country_initial_debt",
        role="input",
        kind="country_map",
        keys=list(COUNTRIES),
        addresses=_flat_addresses(data.COUNTRY_INITIAL_DEBT.cells),
        domain=_domain(0.0, 200.0),
    ),
    _node(
        series_id="growth_baseline",
        role="input",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.GROWTH_BASELINE.cells),
        domain=_domain(-10.0, 15.0),
    ),
    _node(
        series_id="interest_baseline",
        role="input",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.INTEREST_BASELINE.cells),
        domain=_domain(0.0, 20.0),
    ),
    _node(
        series_id="primary_balance_baseline",
        role="input",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.PRIMARY_BALANCE_BASELINE.cells),
        domain=_domain(-15.0, 15.0),
    ),
    _node(
        series_id="shock_year",
        role="input",
        kind="int",
        keys=[None],
        address=_scalar_address(data.SHOCK_YEAR_CELLS),
        domain=_domain(1.0, 5.0),
    ),
    _node(
        series_id="shock_type",
        role="input",
        kind="enum_int",
        keys=[None],
        address=_scalar_address(data.SHOCK_TYPE_CELLS),
        options=[1, 2, 3],
        option_labels={
            1: "1 · growth",
            2: "2 · interest",
            3: "3 · primary balance",
        },
    ),
    _node(
        series_id="shock_magnitudes",
        role="input",
        kind="shock_map",
        keys=list(SHOCK_PARAMS),
        addresses=_flat_addresses(data.SHOCK_MAGNITUDES.cells),
        domain=_domain(-30.0, 30.0),
    ),
    _node(
        series_id="initial_debt_resolved",
        role="internal",
        kind="scalar",
        keys=[None],
        address=_scalar_address(data.INITIAL_DEBT_RESOLVED_CELLS),
    ),
    _node(
        series_id="engine_initial_debt_baseline",
        role="internal",
        kind="scalar",
        keys=[None],
        address=_scalar_address(data.ENGINE_INITIAL_DEBT_BASELINE_CELLS),
    ),
    _node(
        series_id="engine_initial_debt_shocked",
        role="internal",
        kind="scalar",
        keys=[None],
        address=_scalar_address(data.ENGINE_INITIAL_DEBT_SHOCKED_CELLS),
    ),
    _node(
        series_id="shock_magnitude_resolved",
        role="internal",
        kind="scalar",
        keys=[None],
        address=_scalar_address(data.SHOCK_MAGNITUDE_RESOLVED_CELLS),
    ),
    _node(
        series_id="shock_active",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.SHOCK_ACTIVE.cells),
    ),
    _node(
        series_id="shocked_growth",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.SHOCKED_GROWTH.cells),
    ),
    _node(
        series_id="shocked_interest",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.SHOCKED_INTEREST.cells),
    ),
    _node(
        series_id="shocked_primary_balance",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.SHOCKED_PRIMARY_BALANCE.cells),
    ),
    _node(
        series_id="baseline_path_internal",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.BASELINE_PATH_INTERNAL.cells),
    ),
    _node(
        series_id="shocked_path_internal",
        role="internal",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.SHOCKED_PATH_INTERNAL.cells),
    ),
    _node(
        series_id="output_baseline",
        role="output",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.OUTPUT_BASELINE.cells),
    ),
    _node(
        series_id="output_shocked",
        role="output",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.OUTPUT_SHOCKED.cells),
    ),
    _node(
        series_id="output_delta",
        role="output",
        kind="year_map",
        keys=list(YEARS),
        addresses=_flat_addresses(data.OUTPUT_DELTA.cells),
    ),
)

NODES_BY_ID: dict[str, dict[str, Any]] = {node["id"]: node for node in NODES}


def axes() -> dict[str, list[Any]]:
    return {
        "years": list(YEARS),
        "countries": list(COUNTRIES),
        "shock_params": list(SHOCK_PARAMS),
    }


def all_cell_addresses() -> tuple[str, ...]:
    """Every series cell address — enough to seed FormulaEvaluator extraction."""
    addresses: list[str] = []
    for node in NODES:
        if "address" in node:
            addresses.append(node["address"])
        for address in node.get("addresses", {}).values():
            addresses.append(address)
    return tuple(dict.fromkeys(addresses))


def input_cell_writes(flat_inputs: Mapping[str, Any]) -> dict[str, Any]:
    """Map flat viz inputs to sheet-qualified leaf cells for FormulaEvaluator."""
    writes: dict[str, Any] = {
        NODES_BY_ID["country_name"]["address"]: flat_inputs["country_name"],
        NODES_BY_ID["shock_year"]["address"]: flat_inputs["shock_year"],
        NODES_BY_ID["shock_type"]["address"]: flat_inputs["shock_type"],
    }
    for key, address in NODES_BY_ID["country_initial_debt"]["addresses"].items():
        writes[address] = flat_inputs["country_initial_debt"][key]
    for key, address in NODES_BY_ID["growth_baseline"]["addresses"].items():
        writes[address] = flat_inputs["growth_baseline"][key]
    for key, address in NODES_BY_ID["interest_baseline"]["addresses"].items():
        writes[address] = flat_inputs["interest_baseline"][key]
    for key, address in NODES_BY_ID["primary_balance_baseline"]["addresses"].items():
        writes[address] = flat_inputs["primary_balance_baseline"][key]
    for key, address in NODES_BY_ID["shock_magnitudes"]["addresses"].items():
        writes[address] = flat_inputs["shock_magnitudes"][key]
    return writes


__all__ = [
    "YEARS",
    "COUNTRIES",
    "SHOCK_PARAMS",
    "INPUT_IDS",
    "SERIES_IDS",
    "EDGES",
    "NODES",
    "NODES_BY_ID",
    "BackendName",
    "axes",
    "all_cell_addresses",
    "input_cell_writes",
]
