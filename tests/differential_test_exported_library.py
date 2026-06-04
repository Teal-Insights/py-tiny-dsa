"""Differential test for the exported tiny_dsa standalone library.

Drives a matrix of input configurations through two oracles and compares the
resulting output cells side-by-side. Writes a Parity Report (CSV + TXT
summary) to ``data/differential/exported_library/``.

Oracles
-------
Golden Master Oracle
    The ``tiny-dsa.xlsx`` workbook, driven by Microsoft Excel via xlwings.
    Excel runs in a dedicated, headless ``xw.App`` instance with no other
    workbooks open. Inputs are applied as raw cell writes.

MVP Oracle
    The exported standalone Python package at ``dist/`` (post commits
    8464d45 / 9831beb / 78fd00c), loaded as ``dist.tiny_dsa.api``. Inputs are
    applied via the package's Records-shaped ``set_*`` functions; outputs are
    read from the Records-shaped ``compute_*`` functions. The package does not
    access the workbook file at runtime.

Surface tested
--------------
Public API of the exported library — every function re-exported from
``dist/__init__.py``::

    set_country_name, set_growth_baseline, set_interest_baseline,
    set_primary_balance_baseline, set_shock_year, set_shock_type,
    set_shock_magnitudes, compute_output_baseline, compute_output_shocked,
    compute_output_delta

A single ``Inputs`` dataclass is the source of truth for one test point.
The runner converts it into Records for the MVP setters and into cell-
address-keyed writes for the Excel oracle; the underlying numeric inputs
applied at each cell are identical.

Coverage
--------
The sweep has three groups (118 scenarios × 15 output cells = 1,770
comparisons total):

- **Canonical** (12): country × {baseline + 3 canonical shocks}.
- **Single-axis** (61): country (3), shock_year (5), shock_type (3),
  growth-shock magnitude (5), and each continuous baseline vector at
  years {1, 3, 5} × 5 sweep values (45). Per-axis pass-rates are readable
  directly from the CSV.
- **Categorical combo** (45): full factorial of country × shock_type ×
  shock_year, covering 2- and 3-way interactions the single-axis sweeps
  cannot reach.

Pre-flight checks (fail-fast, run before any scenario)
------------------------------------------------------
**Path / package verification.** Workbook file and ``dist/__init__.py`` /
``dist/api.py`` must exist; otherwise raise with a regeneration hint.

**Staleness warning.** If the workbook's mtime is newer than
``dist/data.py``'s mtime, log a warning: the Excel oracle reads from the
current workbook state, while the MVP oracle reads from the snapshot
embedded at extraction time. A subsequent diff could be a staleness
artifact rather than a codegen bug.

**Binding-cell verification.** Each setter's ``_LEAF_INDEX_*`` table and
each compute function's ``_OUTPUT_LEAVES_*`` table is asserted to match
the cell constants declared in this script. Drift raises before any
scenario runs, isolating calculation-correctness testing from binding-
correctness testing: a wrong-cell binding bug cannot be silently masked
by a coincidence between the wrong cell's default and the test value.

Relation to ``tests/differential/differential_testing.py``
----------------------------------------------------------
The harness under ``tests/differential/`` compares Excel against
``FormulaEvaluator``. This script is its sibling for the *exported
standalone library*: it closes the M2-vs-D10 oracle gap by validating the
artifact callers actually consume against Excel.

Acceptance bar
--------------
``abs_diff <= 1e-6`` (atol) for every (scenario, cell) pair. The acceptance
threshold is 100%; any failure means the exported package's encoding of
some computation diverges from Excel within numerical tolerance.

Parity Report
-------------
``parity_report.csv``
    One row per (scenario, cell) comparison. Columns: ``scenario_id``,
    ``cell_address``, ``cell_label``, ``excel_value``, ``mvp_value``,
    ``abs_diff``, ``rel_diff``, ``passed``.
``parity_report.txt``
    Header (timestamp, paths, tolerance), aggregate pass/fail counts, first
    divergence (scenario + cell), and the full list of failing comparisons.

Usage
-----
Regenerate the exported package first if it is missing or stale::

    uv run python src/extraction_pipeline.py

From the extraction repo (default ``--layout repo``)::

    uv run python tests/differential/differential_test_exported_library.py

From the exported ``dist/`` project (``--layout exported``)::

    uv run --project dist --group validation python tests/differential_test_exported_library.py --layout exported

Microsoft Excel must be installed locally (xlwings drives Excel via COM).
"""

from __future__ import annotations

import argparse
import csv
import importlib
import logging
import math
import sys
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from types import ModuleType
from typing import Any, Iterator, Literal, cast

logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

LayoutName = Literal["repo", "exported"]

ATOL = 1e-6  # absolute tolerance against Excel; matches the principles doc


@dataclass(frozen=True)
class DifferentialConfig:
    """Runtime paths and import settings for one differential run."""

    workbook_path: Path
    package_dir: Path
    package_name: str
    import_root: Path
    report_dir: Path
    atol: float = ATOL


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse CLI flags for layout and optional path overrides."""
    parser = argparse.ArgumentParser(
        description="Differential test for the exported tiny_dsa standalone library.",
    )
    parser.add_argument(
        "--layout",
        choices=("repo", "exported"),
        default="repo",
        help="Path preset: extraction repo (default) or exported dist/tests copy.",
    )
    parser.add_argument(
        "--workbook-path",
        type=Path,
        default=None,
        help="Override workbook path (defaults depend on --layout).",
    )
    parser.add_argument(
        "--package-name",
        default=None,
        help="Override import module for the MVP oracle (defaults depend on --layout).",
    )
    parser.add_argument(
        "--import-root",
        type=Path,
        default=None,
        help="Directory added to sys.path before importing the MVP oracle.",
    )
    parser.add_argument(
        "--report-dir",
        type=Path,
        default=None,
        help="Directory for parity_report.{csv,txt} output.",
    )
    return parser.parse_args(argv)


def resolve_config(
    *,
    script_path: Path,
    layout: LayoutName,
    workbook_path: Path | None = None,
    package_name: str | None = None,
    import_root: Path | None = None,
    report_dir: Path | None = None,
) -> DifferentialConfig:
    """Resolve default paths for the extraction repo or exported dist layout."""
    script_path = script_path.resolve()
    if layout == "exported":
        dist_root = script_path.parent.parent
        tests_root = script_path.parent
        defaults = DifferentialConfig(
            workbook_path=tests_root / "fixtures" / "tiny-dsa.xlsx",
            package_dir=dist_root / "tiny_dsa",
            package_name="tiny_dsa.api",
            import_root=dist_root,
            report_dir=tests_root / "results" / "local",
        )
    else:
        project_root = script_path.parents[2]
        defaults = DifferentialConfig(
            workbook_path=project_root / "data" / "tiny-dsa.xlsx",
            package_dir=project_root / "dist" / "tiny_dsa",
            package_name="dist.tiny_dsa.api",
            import_root=project_root,
            report_dir=project_root / "data" / "differential" / "exported_library",
        )
    return DifferentialConfig(
        workbook_path=(workbook_path or defaults.workbook_path).resolve(),
        package_dir=defaults.package_dir.resolve(),
        package_name=package_name or defaults.package_name,
        import_root=(import_root or defaults.import_root).resolve(),
        report_dir=(report_dir or defaults.report_dir).resolve(),
        atol=ATOL,
    )


def config_from_args(
    script_path: Path,
    args: argparse.Namespace,
) -> DifferentialConfig:
    """Build a ``DifferentialConfig`` from parsed CLI arguments."""
    return resolve_config(
        script_path=script_path,
        layout=args.layout,
        workbook_path=args.workbook_path,
        package_name=args.package_name,
        import_root=args.import_root,
        report_dir=args.report_dir,
    )


COUNTRIES: tuple[str, ...] = ("Borvelia", "Litellia", "Aurelium")
SHOCK_TYPES: tuple[int, ...] = (1, 2, 3)
SHOCK_YEARS: tuple[int, ...] = (1, 2, 3, 4, 5)

# Record-shape constants reused across input setters.
SHOCK_PARAMETER_LABELS: tuple[str, str, str] = ("Growth", "Interest", "Primary balance")

# Single-axis perturbation values for the continuous baseline vectors.
# Each axis varies one year of one indicator vector and leaves the other
# four years (and the other two indicators) at their canonical values.
# Year endpoints 1 and 5 are covered alongside the middle (3) so the
# 5-year recursion's first and last steps are exercised explicitly.
CONTINUOUS_AXES: tuple[tuple[str, str, tuple[float, ...]], ...] = (
    ("growth", "growth_baseline", (0.0, 1.5, 3.5, 5.5, 7.0)),
    ("interest", "interest_baseline", (0.0, 2.0, 4.0, 6.0, 8.0)),
    ("primary_balance", "primary_balance_baseline", (-3.0, -1.5, 0.0, 1.5, 3.0)),
)
PERTURBATION_YEARS: tuple[int, ...] = (1, 3, 5)
SHOCK_MAGNITUDE_SWEEP: tuple[float, ...] = (-3.0, -2.0, -1.0, 0.0, 1.0)


# --------------------------------------------------------------------------
# Workbook cell mappings (mirror the named ranges defined in tiny-dsa.xlsx)
# --------------------------------------------------------------------------
#
# These constants duplicate the cell addresses declared in
# ``bindings/inputs.bindings.yaml`` and ``bindings/outputs.bindings.yaml``.
# If those sidecars change, update the constants here too.

COUNTRY_NAME_CELL = "Inputs!B5"
SHOCK_YEAR_CELL = "Inputs!B21"
SHOCK_TYPE_CELL = "Inputs!B22"
SHOCK_TABLE_CELLS: tuple[str, str, str] = (
    "Inputs!B26",  # growth-shock magnitude (pp)
    "Inputs!C26",  # interest-shock magnitude (pp)
    "Inputs!D26",  # primary-balance-shock magnitude (pp)
)
GROWTH_BASELINE_CELLS = tuple(f"Inputs!{c}16" for c in "CDEFG")
INTEREST_BASELINE_CELLS = tuple(f"Inputs!{c}17" for c in "CDEFG")
PRIMARY_BALANCE_BASELINE_CELLS = tuple(f"Inputs!{c}18" for c in "CDEFG")

OUTPUT_BASELINE_CELLS = tuple(f"Outputs!{c}12" for c in "BCDEF")
OUTPUT_SHOCKED_CELLS = tuple(f"Outputs!{c}13" for c in "BCDEF")
OUTPUT_DELTA_CELLS = tuple(f"Outputs!{c}14" for c in "BCDEF")

OUTPUT_RANGES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("output_baseline", OUTPUT_BASELINE_CELLS),
    ("output_shocked", OUTPUT_SHOCKED_CELLS),
    ("output_delta", OUTPUT_DELTA_CELLS),
)


# --------------------------------------------------------------------------
# Inputs and scenarios
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Inputs:
    """One full Tiny-DSA input configuration; matches the workbook named ranges."""

    country_name: str
    growth_baseline: tuple[float, ...]
    interest_baseline: tuple[float, ...]
    primary_balance_baseline: tuple[float, ...]
    shock_year: int
    shock_type: int
    shock_table: tuple[float, float, float]


@dataclass(frozen=True)
class Scenario:
    """One identified input configuration plus a stable scenario id."""

    id: str
    inputs: Inputs


# Canonical baseline (no-shock) configuration. Values match the project
# manifest's defaults for Borvelia; shock magnitudes are zero so the
# baseline output is shock-invariant regardless of shock_type / shock_year.
CANONICAL_BASELINE = Inputs(
    country_name="Borvelia",
    growth_baseline=(3.5, 3.5, 3.5, 3.5, 3.5),
    interest_baseline=(4.0, 4.0, 4.0, 4.0, 4.0),
    primary_balance_baseline=(-1.0, -0.5, 0.0, 0.5, 1.0),
    shock_year=2,
    shock_type=1,
    shock_table=(0.0, 0.0, 0.0),
)


def _override_year(
    vec: tuple[float, ...], year: int, value: float
) -> tuple[float, ...]:
    """Replace year-N (1-indexed) value in ``vec``, leaving other years alone."""
    return tuple(value if i == year - 1 else v for i, v in enumerate(vec))


def _canonical_scenarios() -> Iterator[Scenario]:
    """Country × {baseline + three canonical shocks} = 12 scenarios.

    Mirrors the project manifest's three canonical scenarios (growth,
    interest, primary-balance shocks at year 2) plus a no-shock baseline,
    swept across all three countries.
    """
    for country in COUNTRIES:
        base = replace(CANONICAL_BASELINE, country_name=country)
        yield Scenario(id=f"canonical:{country}:baseline", inputs=base)
        yield Scenario(
            id=f"canonical:{country}:growth_shock",
            inputs=replace(base, shock_type=1, shock_table=(-2.0, 0.0, 0.0)),
        )
        yield Scenario(
            id=f"canonical:{country}:interest_shock",
            inputs=replace(base, shock_type=2, shock_table=(0.0, 2.0, 0.0)),
        )
        yield Scenario(
            id=f"canonical:{country}:primary_balance_shock",
            inputs=replace(base, shock_type=3, shock_table=(0.0, 0.0, -1.0)),
        )


def _single_axis_perturbations() -> Iterator[Scenario]:
    """Perturb one parameter at a time around the canonical baseline.

    Country (3), shock_year (5), shock_type (3), growth-shock magnitude
    (5), and each continuous baseline vector at years {1, 3, 5} ×
    {5 sweep values} = 45. Total: 3 + 5 + 3 + 5 + 45 = 61 scenarios.
    Single-axis isolation makes the per-axis pass-rate readable from the
    parity report.
    """
    for country in COUNTRIES:
        yield Scenario(
            id=f"single_axis:country={country}",
            inputs=replace(CANONICAL_BASELINE, country_name=country),
        )

    canonical_growth_shock = replace(
        CANONICAL_BASELINE, shock_type=1, shock_table=(-2.0, 0.0, 0.0)
    )
    for year in SHOCK_YEARS:
        yield Scenario(
            id=f"single_axis:shock_year={year}",
            inputs=replace(canonical_growth_shock, shock_year=year),
        )

    full_shock_table = (-2.0, 2.0, -1.0)
    for stype in SHOCK_TYPES:
        yield Scenario(
            id=f"single_axis:shock_type={stype}",
            inputs=replace(
                CANONICAL_BASELINE, shock_type=stype, shock_table=full_shock_table
            ),
        )

    for magnitude in SHOCK_MAGNITUDE_SWEEP:
        yield Scenario(
            id=f"single_axis:growth_shock_magnitude={magnitude:+.1f}",
            inputs=replace(
                CANONICAL_BASELINE,
                shock_type=1,
                shock_table=(magnitude, 0.0, 0.0),
            ),
        )

    for indicator, attr, values in CONTINUOUS_AXES:
        base_vec: tuple[float, ...] = getattr(CANONICAL_BASELINE, attr)
        for year in PERTURBATION_YEARS:
            for value in values:
                yield Scenario(
                    id=f"single_axis:{indicator}[year={year}]={value:+.1f}",
                    inputs=replace(
                        CANONICAL_BASELINE,
                        **{attr: _override_year(base_vec, year, value)},
                    ),
                )


def _categorical_combo_scenarios() -> Iterator[Scenario]:
    """Full factorial across the three categorical axes.

    country × shock_type × shock_year = 3 × 3 × 5 = 45 scenarios. Covers
    every two- and three-way interaction the single-axis sweep cannot
    reach (e.g. country-specific behaviour of an interest-shock at year 5).
    Shock magnitudes are populated for all three types so the chosen
    shock_type drives a non-degenerate divergence from baseline.
    """
    full_shock_table = (-2.0, 2.0, -1.0)
    for country in COUNTRIES:
        for stype in SHOCK_TYPES:
            for year in SHOCK_YEARS:
                yield Scenario(
                    id=f"combo:country={country}:shock_type={stype}:shock_year={year}",
                    inputs=replace(
                        CANONICAL_BASELINE,
                        country_name=country,
                        shock_type=stype,
                        shock_year=year,
                        shock_table=full_shock_table,
                    ),
                )


def build_scenarios() -> tuple[Scenario, ...]:
    """Return the full scenario sweep.

    Three groups:

    - **Canonical** (12): country × {baseline + 3 canonical shocks}.
    - **Single-axis** (61): one parameter perturbed around the canonical
      baseline; the per-axis pass-rate is readable directly from the CSV.
    - **Categorical combo** (45): full factorial of country × shock_type
      × shock_year, covering 2- and 3-way interactions.

    Total: 118 scenarios × 15 output cells = 1,770 comparisons.
    """
    return (
        *_canonical_scenarios(),
        *_single_axis_perturbations(),
        *_categorical_combo_scenarios(),
    )


def output_cell_labels() -> tuple[tuple[str, str], ...]:
    """Return ``((cell_label, cell_address), ...)`` for the 15 comparison cells."""
    return tuple(
        (f"{name}[year={i + 1}]", cell)
        for name, cells in OUTPUT_RANGES
        for i, cell in enumerate(cells)
    )


# --------------------------------------------------------------------------
# Input adaptation (Inputs -> Excel cell writes; Inputs -> MVP Records calls)
# --------------------------------------------------------------------------


def inputs_for_excel(inputs: Inputs) -> dict[str, Any]:
    """Map ``Inputs`` to a per-A1-cell ``{address: value}`` dict for Excel.

    Excel has no Records-aware input API, so the Golden Master Oracle writes
    raw cell values. Every numeric input ends up at the same workbook cell
    the MVP setter would also target — by construction, the two oracles see
    identical inputs at the cell level.
    """
    return {
        COUNTRY_NAME_CELL: inputs.country_name,
        SHOCK_YEAR_CELL: inputs.shock_year,
        SHOCK_TYPE_CELL: inputs.shock_type,
        **dict(zip(SHOCK_TABLE_CELLS, inputs.shock_table, strict=True)),
        **dict(zip(GROWTH_BASELINE_CELLS, inputs.growth_baseline, strict=True)),
        **dict(zip(INTEREST_BASELINE_CELLS, inputs.interest_baseline, strict=True)),
        **dict(
            zip(
                PRIMARY_BALANCE_BASELINE_CELLS,
                inputs.primary_balance_baseline,
                strict=True,
            )
        ),
    }


def _time_series_records(values: tuple[float, ...]) -> list[dict[str, Any]]:
    """Build ``[{TIME_PERIOD: 1, OBS_VALUE: v1}, ...]`` for a 5-year vector."""
    return [
        {"TIME_PERIOD": i + 1, "OBS_VALUE": value} for i, value in enumerate(values)
    ]


def apply_inputs_to_mvp(api: ModuleType, ctx: Any, inputs: Inputs) -> None:
    """Apply ``Inputs`` to the MVP context via the exported ``set_*`` functions.

    Exercises the same public-API surface a consumer of the exported library
    would call. Every setter takes ``Records`` keyed by ``PARAMETER``,
    ``TIME_PERIOD``, or ``SHOCK_PARAMETER`` depending on the series shape.
    """
    api.set_country_name(
        ctx, [{"PARAMETER": "country_name", "OBS_VALUE": inputs.country_name}]
    )
    api.set_shock_year(
        ctx, [{"PARAMETER": "shock_year", "OBS_VALUE": inputs.shock_year}]
    )
    api.set_shock_type(
        ctx, [{"PARAMETER": "shock_type", "OBS_VALUE": inputs.shock_type}]
    )
    api.set_growth_baseline(ctx, _time_series_records(inputs.growth_baseline))
    api.set_interest_baseline(ctx, _time_series_records(inputs.interest_baseline))
    api.set_primary_balance_baseline(
        ctx, _time_series_records(inputs.primary_balance_baseline)
    )
    api.set_shock_magnitudes(
        ctx,
        [
            {"SHOCK_PARAMETER": label, "OBS_VALUE": magnitude}
            for label, magnitude in zip(
                SHOCK_PARAMETER_LABELS, inputs.shock_table, strict=True
            )
        ],
    )


# --------------------------------------------------------------------------
# Oracle 1: Golden Master (Excel via xlwings)
# --------------------------------------------------------------------------


def run_excel_oracle(
    workbook_path: Path,
    scenario: Scenario,
    output_addresses: tuple[str, ...],
) -> dict[str, Any]:
    """Run one scenario through Excel; return ``{cell_address: value}``.

    Opens the workbook in a dedicated, headless xlwings App instance, switches
    Excel to manual calculation so per-cell writes do not trigger intermediate
    recalcs, writes the inputs, forces one full recalculation, reads the
    output cells, and closes the workbook *without* saving.
    """
    import xlwings as xw  # imported lazily so the module loads without Excel

    logger.info("Excel oracle: %s", scenario.id)
    app = xw.App(visible=False, add_book=False)
    try:
        workbook = app.books.open(str(workbook_path))
        try:
            # Calculation must be set AFTER opening a workbook; Excel
            # rejects the property write when no workbook is loaded.
            app.calculation = "manual"
            for address, value in inputs_for_excel(scenario.inputs).items():
                sheet, cell = address.split("!", 1)
                workbook.sheets[sheet].range(cell).value = value
            workbook.app.calculate()

            def read(address: str) -> Any:
                sheet, cell = address.split("!", 1)
                return workbook.sheets[sheet].range(cell).value

            return {address: read(address) for address in output_addresses}
        finally:
            workbook.close()
    finally:
        app.quit()


# --------------------------------------------------------------------------
# Oracle 2: MVP (exported standalone library)
# --------------------------------------------------------------------------


def load_exported_library(project_root: Path, package_name: str) -> ModuleType:
    """Import the exported standalone library as a package.

    Adds the project root to ``sys.path`` if necessary, then resolves the
    package via ``importlib.import_module``. The package is not unloaded on
    return; subsequent calls will reuse it from ``sys.modules``.
    """
    logger.info("Importing exported package: %s", package_name)
    root_str = str(project_root)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)
    return importlib.import_module(package_name)


def run_mvp_oracle(
    api: ModuleType,
    scenario: Scenario,
) -> dict[str, Any]:
    """Run one scenario through the exported library; return per-cell outputs.

    Builds a fresh ``EvalContext``, applies the inputs via the Records-shaped
    ``set_*`` setters, calls each ``compute_*`` entrypoint, and maps the
    returned Records back to cell addresses for comparison against Excel.
    """
    logger.info("MVP oracle:   %s", scenario.id)
    ctx = api.make_context()
    apply_inputs_to_mvp(api, ctx, scenario.inputs)
    outputs: dict[str, Any] = {}
    for entrypoint, cells in OUTPUT_RANGES:
        compute_fn = getattr(api, f"compute_{entrypoint}")
        records = compute_fn(ctx=ctx)
        outputs.update(_records_to_cells(records, cells))
    return outputs


def _records_to_cells(
    records: list[dict[str, Any]],
    cells: tuple[str, ...],
) -> dict[str, Any]:
    """Map ``compute_*`` Records to per-cell ``{address: value}``.

    The exported library emits records in workbook order, one per cell in
    the binding's ``data_range``. The runner pairs records to cells by
    position. The specific values carried in dimension fields
    (``TIME_PERIOD``, ``SCENARIO``, …) are workbook-specific and are not
    assumed here — but when ``TIME_PERIOD`` is present on every record,
    its sequence must be strictly increasing. That cheap monotonicity
    check catches future codegen changes that reorder record emission
    without depending on the workbook's specific period labels.
    """
    if len(records) != len(cells):
        raise ValueError(f"expected {len(cells)} records, got {len(records)}")
    raw_periods = [record.get("TIME_PERIOD") for record in records]
    if all(period is not None for period in raw_periods):
        periods = cast(list[Any], raw_periods)
        if any(a >= b for a, b in zip(periods, periods[1:], strict=False)):
            raise ValueError(
                f"records' TIME_PERIOD values are not strictly increasing: "
                f"{periods!r}. compute_* may have changed its emission order."
            )
    by_cell: dict[str, Any] = {}
    for index, (record, cell) in enumerate(zip(records, cells, strict=True)):
        if "OBS_VALUE" not in record:
            raise ValueError(f"record {index}: missing OBS_VALUE: {record!r}")
        by_cell[cell] = record["OBS_VALUE"]
    return by_cell


# --------------------------------------------------------------------------
# Comparison
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Comparison:
    """One side-by-side comparison of an output cell between the two oracles.

    ``excel_value`` and ``mvp_value`` are typed ``Any`` because either oracle
    can legitimately return numbers, ``None`` (blank), or non-numeric values
    (``XlError`` enum members, error strings, …) on a given cell.
    """

    scenario_id: str
    cell_address: str
    cell_label: str
    excel_value: Any
    mvp_value: Any
    abs_diff: float | None
    rel_diff: float | None
    passed: bool


def compare_cell(
    scenario_id: str,
    cell_address: str,
    cell_label: str,
    excel: Any,
    mvp: Any,
    *,
    atol: float,
) -> Comparison:
    """Compare one cell's value across oracles; classify against ``atol``.

    Rules:

    - Both ``None`` (blank cells) match.
    - One ``None`` against a number fails.
    - Non-coercible values (e.g. ``XlError`` enum members) pass only if they
      compare equal under ``==``.
    - Non-finite values (NaN, +/-inf) pass only if equal under ``==`` (or
      both NaN).
    - Otherwise: pass iff ``abs(excel - mvp) <= atol``.
    """
    if excel is None and mvp is None:
        return Comparison(
            scenario_id, cell_address, cell_label, None, None, 0.0, 0.0, True
        )
    if excel is None or mvp is None:
        return Comparison(
            scenario_id, cell_address, cell_label, excel, mvp, None, None, False
        )
    try:
        excel_f = float(excel)  # type: ignore[arg-type]
        mvp_f = float(mvp)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return Comparison(
            scenario_id, cell_address, cell_label, excel, mvp, None, None, excel == mvp
        )
    if not (math.isfinite(excel_f) and math.isfinite(mvp_f)):
        passed = excel_f == mvp_f or (math.isnan(excel_f) and math.isnan(mvp_f))
        return Comparison(
            scenario_id, cell_address, cell_label, excel_f, mvp_f, None, None, passed
        )
    abs_diff = abs(excel_f - mvp_f)
    rel_diff = abs_diff / abs(excel_f) if excel_f != 0 else math.inf
    return Comparison(
        scenario_id,
        cell_address,
        cell_label,
        excel_f,
        mvp_f,
        abs_diff,
        rel_diff,
        abs_diff <= atol,
    )


def compare_scenario(
    scenario: Scenario,
    excel_outputs: dict[str, Any],
    mvp_outputs: dict[str, Any],
    cell_labels: tuple[tuple[str, str], ...],
    *,
    atol: float,
) -> list[Comparison]:
    """Compare every output cell for one scenario, in label order."""
    return [
        compare_cell(
            scenario.id,
            cell_address,
            cell_label,
            excel_outputs.get(cell_address),
            mvp_outputs.get(cell_address),
            atol=atol,
        )
        for cell_label, cell_address in cell_labels
    ]


def crash_comparisons(
    scenario: Scenario,
    cell_labels: tuple[tuple[str, str], ...],
    exc: BaseException,
) -> list[Comparison]:
    """Return one failing ``Comparison`` per output cell for a crashed scenario.

    Lets a per-scenario exception be recorded in the parity report instead
    of aborting the whole sweep. The exception's class and message replace
    both oracle values, the diff fields are ``None``, and ``passed`` is
    ``False`` — so the row reads as a failure with a recoverable cause.
    """
    err_repr = f"<exception: {type(exc).__name__}: {exc}>"
    return [
        Comparison(
            scenario_id=scenario.id,
            cell_address=cell_address,
            cell_label=cell_label,
            excel_value=err_repr,
            mvp_value=err_repr,
            abs_diff=None,
            rel_diff=None,
            passed=False,
        )
        for cell_label, cell_address in cell_labels
    ]


# --------------------------------------------------------------------------
# Parity report
# --------------------------------------------------------------------------

CSV_COLUMNS: tuple[str, ...] = (
    "scenario_id",
    "cell_address",
    "cell_label",
    "excel_value",
    "mvp_value",
    "abs_diff",
    "rel_diff",
    "passed",
)


def write_csv_report(comparisons: list[Comparison], path: Path) -> None:
    """Write one CSV row per comparison."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)
        for cmp in comparisons:
            writer.writerow(
                [
                    cmp.scenario_id,
                    cmp.cell_address,
                    cmp.cell_label,
                    cmp.excel_value,
                    cmp.mvp_value,
                    cmp.abs_diff,
                    cmp.rel_diff,
                    cmp.passed,
                ]
            )
    logger.info("Wrote CSV report: %s", path)


def write_txt_summary(
    comparisons: list[Comparison],
    path: Path,
    *,
    config: DifferentialConfig,
) -> None:
    """Write aggregate stats, first divergence, and any failing comparisons."""
    atol = config.atol
    path.parent.mkdir(parents=True, exist_ok=True)
    total = len(comparisons)
    failures = [c for c in comparisons if not c.passed]
    passed = total - len(failures)
    pass_rate = (100.0 * passed / total) if total else 0.0
    result = "PASS" if not failures else "FAIL"
    first = failures[0] if failures else None

    with path.open("w", encoding="utf-8") as f:
        f.write("Parity report: exported tiny_dsa standalone library vs Excel\n")
        f.write(
            f"Generated: {datetime.now(timezone.utc).isoformat(timespec='seconds')}\n"
        )
        f.write(f"Workbook:  {config.workbook_path}\n")
        f.write(
            f"Package:   {config.package_dir} (imported as {config.package_name})\n"
        )
        f.write(f"Tolerance: atol = {atol}\n\n")
        f.write(f"Total comparisons: {total}\n")
        f.write(f"Passed:            {passed}\n")
        f.write(f"Failed:            {len(failures)}\n")
        f.write(f"Pass rate:         {pass_rate:.2f}%\n")
        f.write("Acceptance bar:    100.00%\n")
        f.write(f"Result:            {result}\n")
        if first is not None:
            f.write("\nFirst divergence:\n")
            f.write(f"  scenario:  {first.scenario_id}\n")
            f.write(f"  cell:      {first.cell_address}  ({first.cell_label})\n")
            f.write(f"  excel:     {first.excel_value!r}\n")
            f.write(f"  mvp:       {first.mvp_value!r}\n")
            f.write(f"  abs_diff:  {first.abs_diff!r}\n")
            f.write(f"  rel_diff:  {first.rel_diff!r}\n")
        if failures:
            f.write(f"\nAll failing comparisons ({len(failures)}):\n")
            for c in failures:
                f.write(
                    f"  {c.scenario_id} | {c.cell_address} ({c.cell_label}) | "
                    f"excel={c.excel_value!r} mvp={c.mvp_value!r} "
                    f"abs={c.abs_diff!r} rel={c.rel_diff!r}\n"
                )
    logger.info("Wrote TXT summary: %s", path)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------


def _verify_paths(config: DifferentialConfig) -> None:
    """Fail fast with a clear message if the required inputs are missing."""
    if not config.workbook_path.is_file():
        raise FileNotFoundError(
            f"Workbook not found: {config.workbook_path}. "
            f"Confirm the tiny-dsa.xlsx workbook is at the expected path."
        )
    if (
        not (config.package_dir / "__init__.py").is_file()
        or not (config.package_dir / "api.py").is_file()
    ):
        raise FileNotFoundError(
            f"Exported package incomplete at {config.package_dir} "
            f"(expected __init__.py and api.py). "
            f"Run 'uv run python src/extraction_pipeline.py' to regenerate."
        )


def _check_staleness(config: DifferentialConfig) -> None:
    """Warn if the workbook is newer than the exported package's defaults.

    The MVP oracle reads constants (country debts, year markers, etc.) from
    the exported package's ``data.py`` snapshot, while the Excel oracle reads
    them from the workbook's current state. If the workbook has been
    modified since the last extraction, the two oracles see different
    defaults for cells the test does not explicitly overwrite, and any
    resulting parity failure would be a staleness artifact rather than a
    codegen bug. This check warns (does not raise) so the operator can
    investigate before trusting the report.
    """
    data_path = config.package_dir / "data.py"
    if not data_path.is_file():
        return  # _verify_paths handles the missing-package case
    workbook_mtime = config.workbook_path.stat().st_mtime
    data_mtime = data_path.stat().st_mtime
    if workbook_mtime > data_mtime:
        logger.warning(
            "Workbook is newer than the exported package's data.py snapshot. "
            "Workbook last modified: %s. data.py last modified: %s. "
            "Constant-cell defaults may diverge between oracles; regenerate "
            "via 'uv run python src/extraction_pipeline.py' if the workbook "
            "constants changed.",
            datetime.fromtimestamp(workbook_mtime, tz=timezone.utc).isoformat(
                timespec="seconds"
            ),
            datetime.fromtimestamp(data_mtime, tz=timezone.utc).isoformat(
                timespec="seconds"
            ),
        )


def verify_binding_cells(api: ModuleType) -> None:
    """Assert the exported library's binding tables match the script's hardcoded cells.

    Separates calculation-correctness testing from binding-correctness
    testing. If a setter's ``_LEAF_INDEX_*`` table or a compute function's
    ``_OUTPUT_LEAVES_*`` table drifts away from the cell constants declared
    at the top of this module, this raises before any scenario runs. A
    subsequent passing differential then verifies only the calculation
    engine, not the binding round-trip.

    Without this check, a wrong-cell binding bug could be masked when the
    workbook's default value at the wrong cell happens to coincide with
    the test value.
    """
    expected_inputs: dict[str, set[str]] = {
        "_LEAF_INDEX_COUNTRY_NAME": {COUNTRY_NAME_CELL},
        "_LEAF_INDEX_SHOCK_YEAR": {SHOCK_YEAR_CELL},
        "_LEAF_INDEX_SHOCK_TYPE": {SHOCK_TYPE_CELL},
        "_LEAF_INDEX_SHOCK_MAGNITUDES": set(SHOCK_TABLE_CELLS),
        "_LEAF_INDEX_GROWTH_BASELINE": set(GROWTH_BASELINE_CELLS),
        "_LEAF_INDEX_INTEREST_BASELINE": set(INTEREST_BASELINE_CELLS),
        "_LEAF_INDEX_PRIMARY_BALANCE_BASELINE": set(PRIMARY_BALANCE_BASELINE_CELLS),
    }
    for name, expected_cells in expected_inputs.items():
        table = getattr(api, name, None)
        if table is None:
            raise RuntimeError(
                f"Exported library is missing {name}; the input-binding shape "
                f"may have changed. Re-read bindings/inputs.bindings.yaml and "
                f"update this script's cell constants."
            )
        actual_cells = set(table.values())
        if actual_cells != expected_cells:
            raise RuntimeError(
                f"Binding-cell mismatch on {name}: "
                f"library targets {sorted(actual_cells)!r}, "
                f"script expects {sorted(expected_cells)!r}. "
                f"Update either the bindings or this script's constants."
            )

    expected_outputs: dict[str, tuple[str, ...]] = {
        "_OUTPUT_LEAVES_OUTPUT_BASELINE": OUTPUT_BASELINE_CELLS,
        "_OUTPUT_LEAVES_OUTPUT_SHOCKED": OUTPUT_SHOCKED_CELLS,
        "_OUTPUT_LEAVES_OUTPUT_DELTA": OUTPUT_DELTA_CELLS,
    }
    for name, expected_cells in expected_outputs.items():
        leaves = getattr(api, name, None)
        if leaves is None:
            raise RuntimeError(
                f"Exported library is missing {name}; the output-binding "
                f"shape may have changed."
            )
        actual_cells = tuple(address for address, _ in leaves)
        if actual_cells != expected_cells:
            raise RuntimeError(
                f"Binding-cell mismatch on {name}: "
                f"library targets {actual_cells!r}, "
                f"script expects {expected_cells!r}. "
                f"Update either the bindings or this script's constants."
            )
    logger.info("Binding-cell verification passed.")


def run_differential_test(config: DifferentialConfig) -> int:
    """Run every scenario through both oracles; write the parity report.

    Returns
    -------
    int
        0 if every comparison passes at ``config.atol``, 1 otherwise.
    """
    _verify_paths(config)
    _check_staleness(config)
    api = load_exported_library(config.import_root, config.package_name)
    verify_binding_cells(api)
    scenarios = build_scenarios()
    cell_labels = output_cell_labels()
    output_addresses = tuple(addr for _, addr in cell_labels)
    logger.info(
        "Sweeping %d scenarios x %d cells = %d comparisons.",
        len(scenarios),
        len(cell_labels),
        len(scenarios) * len(cell_labels),
    )

    comparisons: list[Comparison] = []
    for scenario in scenarios:
        try:
            excel_outputs = run_excel_oracle(
                config.workbook_path, scenario, output_addresses
            )
            mvp_outputs = run_mvp_oracle(api, scenario)
        except Exception as exc:
            logger.exception("Scenario %s crashed; recording as failure.", scenario.id)
            comparisons.extend(crash_comparisons(scenario, cell_labels, exc))
            continue
        comparisons.extend(
            compare_scenario(
                scenario,
                excel_outputs,
                mvp_outputs,
                cell_labels,
                atol=config.atol,
            )
        )

    config.report_dir.mkdir(parents=True, exist_ok=True)
    write_csv_report(comparisons, config.report_dir / "parity_report.csv")
    write_txt_summary(
        comparisons,
        config.report_dir / "parity_report.txt",
        config=config,
    )

    failed = sum(1 for c in comparisons if not c.passed)
    logger.info("Done. Failures: %d / %d", failed, len(comparisons))
    return 0 if failed == 0 else 1


def main(argv: list[str] | None = None) -> int:
    """Entry point: configure logging, run the differential test, return exit code."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    try:
        args = parse_args(argv)
        config = config_from_args(Path(__file__).resolve(), args)
        return run_differential_test(config)
    except Exception:
        logger.exception("Differential test failed with an unhandled exception.")
        return 2


if __name__ == "__main__":
    sys.exit(main())
