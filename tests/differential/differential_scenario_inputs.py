"""Scenario input isolation for differential harnesses.

Each scenario declares only the input cells it overrides. Between scenarios the
harness restores the union of all scenario input addresses to workbook baseline
values so undeclared cells do not inherit writes from prior scenarios.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from typing import Any

from .differential_types import Axis, Scenario

InputBaselines = dict[str, Any]


def collect_scenario_input_addresses(
    axes: Iterable[Axis],
    inputs_for_excel: Callable[[Scenario], dict[str, Any]],
) -> frozenset[str]:
    """Return the union of every input cell address across all scenario points."""
    addresses: set[str] = set()
    for axis in axes:
        for point in axis.points:
            addresses.update(inputs_for_excel(point.scenario))
    return frozenset(addresses)
