"""Generated functions accepting and returning named-coordinate values."""

from __future__ import annotations

from . import data, model
from .data import _CONSTANTS_0, _CONSTANTS_1
from .model import OutputBaselineInputs, OutputShockedInputs, OutputDeltaInputs
from .runtime import publish


@publish(data.OUTPUT_BASELINE.schema, constants=_CONSTANTS_0, cells=data.OUTPUT_BASELINE.cells)
def compute_output_baseline(inputs: OutputBaselineInputs) -> data.OutputBaseline:
    """Compute the baseline debt-to-GDP path for projection years 1 through 5.

    Produce the baseline debt-to-GDP trajectory for projection years 1 through 5 under the baseline growth, interest-rate, and primary-balance paths.

    Args:
        inputs: Baseline configuration supplied through the OutputBaselineInputs record, carrying the coordinate identities needed to derive the baseline debt-to-GDP path for projection years 1 through 5.

    Returns:
        The baseline debt-to-GDP path for projection years 1 through 5, expressed in percent of GDP, as an OutputBaseline value.
    """
    if not isinstance(inputs, OutputBaselineInputs):
        raise TypeError(f"compute_output_baseline() expected OutputBaselineInputs, got {type(inputs).__name__}")
    return model.Model(inputs).output_baseline


@publish(data.OUTPUT_SHOCKED.schema, constants=_CONSTANTS_1, cells=data.OUTPUT_SHOCKED.cells)
def compute_output_shocked(inputs: OutputShockedInputs) -> data.OutputShocked:
    """Compute the shocked debt-to-GDP path for projection years 1 through 5.

    Resolve the configured shock and return the shocked debt-to-GDP trajectory implied by the baseline parameters and shock configuration.

    Args:
        inputs: Shock scenario specification supplying the country profile, baseline growth, interest and primary-balance paths, and the shock year, type, and magnitude used to construct the shocked path.

    Returns:
        The shocked debt-to-GDP path for projection years 1 through 5, expressed as a percent of GDP and reported on the stable output surface.
    """
    if not isinstance(inputs, OutputShockedInputs):
        raise TypeError(f"compute_output_shocked() expected OutputShockedInputs, got {type(inputs).__name__}")
    return model.Model(inputs).output_shocked


@publish(data.OUTPUT_DELTA.schema, constants=_CONSTANTS_1, cells=data.OUTPUT_DELTA.cells)
def compute_output_delta(inputs: OutputDeltaInputs) -> data.OutputDelta:
    """Compute the output delta between the shocked and baseline debt-to-GDP paths.

    Returns the per-year difference, shocked minus baseline, expressed in percentage points of GDP.

    Args:
        inputs: Validated configuration bundle for the run, carrying the resolved shock and the shocked and baseline paths from which the delta is derived.

    Returns:
        Per-year difference between the shocked and baseline debt-to-GDP trajectories, in percentage points, for years 1 through 5.
    """
    if not isinstance(inputs, OutputDeltaInputs):
        raise TypeError(f"compute_output_delta() expected OutputDeltaInputs, got {type(inputs).__name__}")
    return model.Model(inputs).output_delta


__all__ = [
    "OutputBaselineInputs",
    "OutputShockedInputs",
    "OutputDeltaInputs",
    "compute_output_baseline",
    "compute_output_shocked",
    "compute_output_delta",
]
