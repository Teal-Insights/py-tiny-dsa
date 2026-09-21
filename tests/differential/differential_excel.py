"""Excel golden-master read helpers for differential harnesses."""

from __future__ import annotations

from typing import Any

from excel_grapher import XlError
from excel_grapher.core.address_keys import normalize_key, parse_address


def read_cell_value(sheets: Any, address: str) -> Any:
    """Read one live Excel cell, preserving error codes as strings.

    xlwings converts ``#VALUE!``, ``#N/A``, and other error cells to ``None`` by
    default. ``err_to_str=True`` returns the literal error text so the golden
    oracle can be compared against typed ``XlError`` values from the SUT.
    """
    sheet, cell = parse_address(normalize_key(address))
    return sheets[sheet].range(cell).options(err_to_str=True).value


def coerce_excel_error(value: Any) -> Any:
    """Normalize xlwings error strings to typed ``XlError`` values."""
    if isinstance(value, str):
        err = XlError.from_text(value)
        if err is not None:
            return err
    return value


def matched_error_values(golden: Any, mvp: Any) -> bool:
    """Return whether both sides are the same typed Excel error."""
    golden = coerce_excel_error(golden)
    mvp = coerce_excel_error(mvp)
    return isinstance(golden, XlError) and isinstance(mvp, XlError) and golden == mvp


def parity_exit_code(
    *,
    failed: int,
    flagged_matched_errors: int,
    allow_matched_errors: bool,
) -> int:
    """Map sweep outcomes to the harness exit code.

    A matched Excel error on a scenario without ``expects_error_values=True`` is
    not evidence of parity — both oracles errored identically, so the comparison
    exercised nothing. Such flagged comparisons fail the run unless the operator
    explicitly passes ``--allow-matched-errors`` for triage.
    """
    if failed:
        return 1
    if flagged_matched_errors and not allow_matched_errors:
        return 1
    return 0
