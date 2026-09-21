"""Shared golden-vs-oracle comparison helpers for differential harnesses."""

from __future__ import annotations

import math
from typing import Any, Literal

from excel_grapher import XlError

from .differential_excel import coerce_excel_error

ComparisonOutcome = Literal["numeric", "matched_error", "matched_blank", "mismatched"]


def classify_comparison(
    golden: Any, mvp: Any, *, atol: float, rtol: float
) -> tuple[bool, bool, ComparisonOutcome, float | None, float | None, str]:
    """Return (parity_match, healthy, outcome, abs_diff, rel_diff, note)."""
    golden = coerce_excel_error(golden)
    mvp = coerce_excel_error(mvp)

    if isinstance(golden, XlError) or isinstance(mvp, XlError):
        if isinstance(golden, XlError) and isinstance(mvp, XlError) and golden == mvp:
            return True, False, "matched_error", None, None, f"both error: {golden}"
        return (
            False,
            False,
            "mismatched",
            None,
            None,
            f"error mismatch (golden={golden!r}, mvp={mvp!r})",
        )

    if golden is None and mvp is None:
        return True, False, "matched_blank", 0.0, 0.0, "both blank"
    if golden is None or mvp is None:
        return False, False, "mismatched", None, None, "one side None"

    if isinstance(golden, bool) and isinstance(mvp, bool):
        matched = golden == mvp
        return (
            matched,
            matched,
            "numeric" if matched else "mismatched",
            None,
            None,
            "bool",
        )

    if isinstance(golden, int | float) and isinstance(mvp, int | float):
        try:
            golden_f, mvp_f = float(golden), float(mvp)
        except OverflowError:
            # Python ints can exceed double range; the diffs are unrepresentable,
            # so decide by exact equality rather than crash the ladder (§1.2).
            matched = golden == mvp
            return (
                matched,
                matched,
                "numeric" if matched else "mismatched",
                None,
                None,
                "exceeds double range; exact comparison",
            )
        if math.isnan(golden_f) and math.isnan(mvp_f):
            return True, False, "matched_error", None, None, "both NaN"
        if not (math.isfinite(golden_f) and math.isfinite(mvp_f)):
            matched = golden_f == mvp_f
            return (
                matched,
                matched,
                "numeric" if matched else "mismatched",
                None,
                None,
                "",
            )
        abs_diff = abs(golden_f - mvp_f)
        # Per technical_standard.md §1.1: rel_diff is measured against the
        # golden value alone and is infinite when golden is zero, so the
        # relative branch can never decide a zero-golden comparison. The gate
        # uses the same rel_diff that is written to the report.
        rel_diff = abs_diff / abs(golden_f) if golden_f != 0.0 else math.inf
        matched = abs_diff <= atol or rel_diff <= rtol
        return (
            matched,
            matched,
            "numeric" if matched else "mismatched",
            abs_diff,
            rel_diff,
            "",
        )

    matched = golden == mvp
    return matched, matched, "numeric" if matched else "mismatched", None, None, "exact"


def apply_health_expectation(
    *,
    expects_error_values: bool,
    parity_match: bool,
    healthy: bool,
    outcome: ComparisonOutcome,
) -> bool:
    """Treat matched errors as healthy when the scenario declares it expects them.

    ``Scenario.expects_error_values`` is the single source of truth for
    intentional error-path scenarios; scenario-id naming conventions carry no
    behavior.
    """
    if expects_error_values and parity_match and outcome == "matched_error":
        return True
    return healthy


def values_match(
    golden: Any, mvp: Any, *, atol: float, rtol: float
) -> tuple[bool, float | None, float | None, str]:
    """Compare golden vs oracle. Returns (parity_match, abs_diff, rel_diff, note)."""
    parity_match, _healthy, _outcome, abs_diff, rel_diff, note = classify_comparison(
        golden, mvp, atol=atol, rtol=rtol
    )
    return parity_match, abs_diff, rel_diff, note
