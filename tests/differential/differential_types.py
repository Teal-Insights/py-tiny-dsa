"""Shared types for differential testing."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

ATOL = 1e-6


@dataclass(frozen=True)
class Scenario:
    """One identified input configuration plus a stable scenario id."""

    id: str
    inputs: Any
    expects_error_values: bool = False


@dataclass(frozen=True)
class AxisPoint:
    """One point on a scenario axis."""

    label: str
    scenario: Scenario


@dataclass(frozen=True)
class Axis:
    """Named axis grouping several scenario points for per-axis reporting."""

    name: str
    points: tuple[AxisPoint, ...]
