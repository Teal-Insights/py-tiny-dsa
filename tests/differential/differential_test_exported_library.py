"""Exported-library differential test harness.

Compare the generated standalone package against Microsoft Excel via xlwings.
Workbook-specific scenario definitions live in the ``Workbook-specific hooks``
section at the bottom of this module.

Run from the extraction repo after export::

    uv run python -m tests.differential.differential_test_exported_library

From the exported ``dist/`` project (Windows + Excel)::

    uv run --project dist --group validation python -m tests.differential.differential_test_exported_library --layout exported
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

from excel_grapher import XlError
from excel_grapher.core.address_keys import normalize_key, parse_address

from .differential_excel import (
    coerce_excel_error,
    matched_error_values,
    parity_exit_code,
    read_cell_value,
)
from .differential_types import ATOL, Scenario

logger = logging.getLogger(__name__)

LayoutName = Literal["repo", "exported"]


@dataclass(frozen=True)
class DifferentialConfig:
    """Runtime paths and import settings for one differential run."""

    workbook_path: Path
    package_dir: Path
    package_name: str
    import_root: Path
    report_dir: Path
    library_name: str
    atol: float = ATOL
    allow_matched_errors: bool = False


@dataclass(frozen=True)
class Comparison:
    scenario_id: str
    cell_address: str
    cell_label: str
    excel_value: Any
    mvp_value: Any
    abs_diff: float | None
    rel_diff: float | None
    passed: bool
    matched_error: bool = False
    flagged_matched_error: bool = False


CSV_COLUMNS: tuple[str, ...] = (
    "scenario_id",
    "cell_address",
    "cell_label",
    "excel_value",
    "mvp_value",
    "abs_diff",
    "rel_diff",
    "passed",
    "matched_error",
    "flagged_matched_error",
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Differential test for the exported standalone library.",
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
    parser.add_argument(
        "--allow-matched-errors",
        action="store_true",
        help=(
            "Triage escape hatch: do not fail the run when both oracles return "
            "the same Excel error on a scenario without expects_error_values=True. "
            "Flagged comparisons are still listed in the report."
        ),
    )
    return parser.parse_args(argv)


def _project_root_from_module(module_path: Path) -> Path:
    """Return repo or dist root from ``tests/differential/<module>.py``."""
    return module_path.resolve().parents[2]


def _tests_root_from_module(module_path: Path) -> Path:
    return module_path.resolve().parents[1]


def resolve_config(
    *,
    module_path: Path,
    layout: LayoutName,
    workbook_path: Path | None = None,
    package_name: str | None = None,
    import_root: Path | None = None,
    report_dir: Path | None = None,
    allow_matched_errors: bool = False,
) -> DifferentialConfig:
    """Resolve paths from ``workbook_config.py`` and the selected layout."""
    module_path = module_path.resolve()
    project_root = _project_root_from_module(module_path)

    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    from src.pipeline_config import load_pipeline_config

    pipeline = load_pipeline_config(repo_root=project_root)
    package_dir = pipeline.package_root
    package_slug = pipeline.dist_metadata.package_name
    library_name = pipeline.dist_metadata.library_name

    if layout == "exported":
        tests_root = _tests_root_from_module(module_path)
        dist_root = project_root
        defaults = DifferentialConfig(
            workbook_path=tests_root / "fixtures" / pipeline.workbook_path.name,
            package_dir=dist_root / package_slug,
            package_name=f"{package_slug}.api",
            import_root=dist_root,
            report_dir=tests_root / "results" / "local",
            library_name=library_name,
        )
    else:
        defaults = DifferentialConfig(
            workbook_path=pipeline.workbook_path,
            package_dir=package_dir,
            package_name=f"dist.{package_slug}.api",
            import_root=project_root,
            report_dir=project_root / pipeline.differential_report_dir_rel,
            library_name=library_name,
        )

    return DifferentialConfig(
        workbook_path=(workbook_path or defaults.workbook_path).resolve(),
        package_dir=defaults.package_dir.resolve(),
        package_name=package_name or defaults.package_name,
        import_root=(import_root or defaults.import_root).resolve(),
        report_dir=(report_dir or defaults.report_dir).resolve(),
        library_name=library_name,
        atol=ATOL,
        allow_matched_errors=allow_matched_errors,
    )


def config_from_args(module_path: Path, args: argparse.Namespace) -> DifferentialConfig:
    return resolve_config(
        module_path=module_path,
        layout=args.layout,
        workbook_path=args.workbook_path,
        package_name=args.package_name,
        import_root=args.import_root,
        report_dir=args.report_dir,
        allow_matched_errors=args.allow_matched_errors,
    )


def compare_cell(
    scenario_id: str,
    cell_address: str,
    cell_label: str,
    excel: Any,
    mvp: Any,
    *,
    atol: float,
    expects_error_values: bool = False,
) -> Comparison:
    raw_excel = excel
    raw_mvp = mvp
    excel = coerce_excel_error(excel)
    mvp = coerce_excel_error(mvp)

    if isinstance(excel, XlError) or isinstance(mvp, XlError):
        passed = (
            isinstance(excel, XlError) and isinstance(mvp, XlError) and excel == mvp
        )
        matched_error = passed
        return Comparison(
            scenario_id,
            cell_address,
            cell_label,
            excel,
            mvp,
            None,
            None,
            passed,
            matched_error=matched_error,
            flagged_matched_error=matched_error and not expects_error_values,
        )

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
        passed = excel == mvp
        matched_error = matched_error_values(raw_excel, raw_mvp) and passed
        return Comparison(
            scenario_id,
            cell_address,
            cell_label,
            excel,
            mvp,
            None,
            None,
            passed,
            matched_error=matched_error,
            flagged_matched_error=matched_error and not expects_error_values,
        )
    if not (math.isfinite(excel_f) and math.isfinite(mvp_f)):
        passed = excel_f == mvp_f or (math.isnan(excel_f) and math.isnan(mvp_f))
        return Comparison(
            scenario_id, cell_address, cell_label, excel_f, mvp_f, None, None, passed
        )
    abs_diff = abs(excel_f - mvp_f)
    rel_diff = abs_diff / abs(excel_f) if excel_f != 0 else math.inf
    passed = abs_diff <= atol
    return Comparison(
        scenario_id,
        cell_address,
        cell_label,
        excel_f,
        mvp_f,
        abs_diff,
        rel_diff,
        passed,
    )


def compare_scenario(
    scenario: Scenario,
    excel_outputs: dict[str, Any],
    mvp_outputs: dict[str, Any],
    cell_labels: tuple[tuple[str, str], ...],
    *,
    atol: float,
) -> list[Comparison]:
    return [
        compare_cell(
            scenario.id,
            cell_address,
            cell_label,
            excel_outputs.get(cell_address),
            mvp_outputs.get(cell_address),
            atol=atol,
            expects_error_values=scenario.expects_error_values,
        )
        for cell_label, cell_address in cell_labels
    ]


def crash_comparisons(
    scenario: Scenario,
    cell_labels: tuple[tuple[str, str], ...],
    exc: BaseException,
) -> list[Comparison]:
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


def write_csv_report(comparisons: list[Comparison], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(CSV_COLUMNS)
        for comparison in comparisons:
            writer.writerow(
                [
                    comparison.scenario_id,
                    comparison.cell_address,
                    comparison.cell_label,
                    comparison.excel_value,
                    comparison.mvp_value,
                    comparison.abs_diff,
                    comparison.rel_diff,
                    comparison.passed,
                    comparison.matched_error,
                    comparison.flagged_matched_error,
                ]
            )


def write_txt_summary(
    comparisons: list[Comparison],
    path: Path,
    *,
    config: DifferentialConfig,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    total = len(comparisons)
    failures = [comparison for comparison in comparisons if not comparison.passed]
    passed = total - len(failures)
    pass_rate = (100.0 * passed / total) if total else 0.0
    flagged = [
        comparison for comparison in comparisons if comparison.flagged_matched_error
    ]
    failing_run = bool(failures) or bool(flagged and not config.allow_matched_errors)
    result = "FAIL" if failing_run else "PASS"
    first = failures[0] if failures else None

    with path.open("w", encoding="utf-8") as handle:
        handle.write(
            f"Parity report: exported {config.library_name} standalone library vs Excel\n"
        )
        handle.write(
            f"Generated: {datetime.now(timezone.utc).isoformat(timespec='seconds')}\n"
        )
        handle.write(f"Workbook:  {config.workbook_path}\n")
        handle.write(
            f"Package:   {config.package_dir} (imported as {config.package_name})\n"
        )
        handle.write(f"Tolerance: atol = {config.atol}\n\n")
        handle.write(f"Total comparisons: {total}\n")
        handle.write(f"Passed:            {passed}\n")
        handle.write(f"Failed:            {len(failures)}\n")
        if flagged:
            allowed_note = (
                " (allowed by --allow-matched-errors)"
                if config.allow_matched_errors
                else " (fails the run)"
            )
            handle.write(f"Matched errors:    {len(flagged)} flagged{allowed_note}\n")
        handle.write(f"Pass rate:         {pass_rate:.2f}%\n")
        handle.write("Acceptance bar:    100.00%\n")
        handle.write(f"Result:            {result}\n")
        if first is not None:
            handle.write("\nFirst divergence:\n")
            handle.write(f"  scenario:  {first.scenario_id}\n")
            handle.write(f"  cell:      {first.cell_address}  ({first.cell_label})\n")
            handle.write(f"  excel:     {first.excel_value!r}\n")
            handle.write(f"  mvp:       {first.mvp_value!r}\n")
            handle.write(f"  abs_diff:  {first.abs_diff!r}\n")
            handle.write(f"  rel_diff:  {first.rel_diff!r}\n")
        if flagged:
            handle.write(
                "\nMATCHED ERROR VALUES (both oracles returned the same Excel "
                "error; fails the run unless the scenario sets "
                "expects_error_values=True or --allow-matched-errors is passed):\n"
            )
            for comparison in flagged:
                handle.write(f"  {comparison.scenario_id} :: ")
                handle.write(f"{comparison.cell_address} ({comparison.cell_label})\n")
                handle.write(f"    excel: {comparison.excel_value!r}\n")
                handle.write(f"    mvp:   {comparison.mvp_value!r}\n")


def load_exported_library(import_root: Path, package_name: str) -> ModuleType:
    root_str = str(import_root)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)
    return importlib.import_module(package_name)


def run_excel_oracle(
    workbook_path: Path,
    scenario: Scenario,
    output_addresses: tuple[str, ...],
) -> dict[str, Any]:
    import xlwings as xw

    logger.info("Excel oracle: %s", scenario.id)
    app = xw.App(visible=False, add_book=False)
    try:
        workbook = app.books.open(str(workbook_path))
        try:
            app.calculation = "manual"
            for address, value in inputs_for_excel(scenario).items():
                sheet, cell = parse_address(normalize_key(address))
                workbook.sheets[sheet].range(cell).value = value
            workbook.app.calculate()

            def read(address: str) -> Any:
                return read_cell_value(workbook.sheets, address)

            return {address: read(address) for address in output_addresses}
        finally:
            workbook.close()
    finally:
        app.quit()


def _records_to_cells(
    records: list[dict[str, Any]],
    cells: tuple[str, ...],
) -> dict[str, Any]:
    if len(records) != len(cells):
        raise ValueError(f"expected {len(cells)} records, got {len(records)}")
    raw_periods = [record.get("TIME_PERIOD") for record in records]
    if all(period is not None for period in raw_periods):
        periods = cast(list[Any], raw_periods)
        if any(a >= b for a, b in zip(periods, periods[1:], strict=False)):
            raise ValueError(
                f"records' TIME_PERIOD values are not strictly increasing: {periods!r}"
            )
    by_cell: dict[str, Any] = {}
    for index, (record, cell) in enumerate(zip(records, cells, strict=True)):
        if "OBS_VALUE" not in record:
            raise ValueError(f"record {index}: missing OBS_VALUE: {record!r}")
        by_cell[cell] = record["OBS_VALUE"]
    return by_cell


def run_mvp_oracle(api: ModuleType, scenario: Scenario) -> dict[str, Any]:
    logger.info("MVP oracle:   %s", scenario.id)
    ctx = api.make_context()
    apply_inputs_to_mvp(api, ctx, scenario)
    outputs: dict[str, Any] = {}
    for entrypoint, cells in output_ranges():
        compute_fn = getattr(api, f"compute_{entrypoint}")
        records = compute_fn(ctx=ctx)
        outputs.update(_records_to_cells(records, cells))
    return outputs


def _verify_paths(config: DifferentialConfig) -> None:
    if not config.workbook_path.is_file():
        raise FileNotFoundError(
            f"Workbook not found: {config.workbook_path}. "
            "Populate data/ and workbook_config.py before running parity tests."
        )
    if (
        not (config.package_dir / "__init__.py").is_file()
        or not (config.package_dir / "api.py").is_file()
    ):
        raise FileNotFoundError(
            f"Exported package incomplete at {config.package_dir} "
            "(expected __init__.py and api.py). "
            "Run 'uv run python -m src.extraction_pipeline' to regenerate."
        )


def _check_staleness(config: DifferentialConfig) -> None:
    data_path = config.package_dir / "data.py"
    if not data_path.is_file():
        return
    workbook_mtime = config.workbook_path.stat().st_mtime
    data_mtime = data_path.stat().st_mtime
    if workbook_mtime > data_mtime:
        logger.warning(
            "Workbook is newer than exported data.py; regenerate dist/ if constants changed."
        )


def _validate_workbook_hooks() -> None:
    scenarios = build_scenarios()
    if not scenarios:
        raise RuntimeError(
            "No differential scenarios configured. Author build_scenarios(), "
            "output_cell_labels(), output_ranges(), inputs_for_excel(), and "
            "apply_inputs_to_mvp() in "
            "tests/differential/differential_test_exported_library.py."
        )
    if not output_cell_labels():
        raise RuntimeError(
            "output_cell_labels() returned no cells. Mirror your output bindings "
            "as (label, address) pairs."
        )
    if not output_ranges():
        raise RuntimeError(
            "output_ranges() returned no compute groups. Map each compute_* "
            "entrypoint to its output cell addresses."
        )


def verify_binding_cells(api: ModuleType) -> None:
    """Assert exported binding tables match this script's cell constants."""
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
    _validate_workbook_hooks()
    _verify_paths(config)
    _check_staleness(config)

    api = load_exported_library(config.import_root, config.package_name)
    verify_binding_cells(api)
    scenarios = build_scenarios()
    cell_labels = output_cell_labels()
    output_addresses = tuple(address for _, address in cell_labels)

    comparisons: list[Comparison] = []
    for scenario in scenarios:
        try:
            excel_outputs = run_excel_oracle(
                config.workbook_path,
                scenario,
                output_addresses,
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

    failed = sum(1 for comparison in comparisons if not comparison.passed)
    flagged = sum(1 for comparison in comparisons if comparison.flagged_matched_error)
    logger.info("Done. Failures: %d / %d", failed, len(comparisons))
    if flagged:
        log = logger.warning if config.allow_matched_errors else logger.error
        log(
            "Matched error values in %d comparison(s); see report section in %s "
            "(set Scenario.expects_error_values=True when intentional)",
            flagged,
            config.report_dir / "parity_report.txt",
        )
    return parity_exit_code(
        failed=failed,
        flagged_matched_errors=flagged,
        allow_matched_errors=config.allow_matched_errors,
    )


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    try:
        args = parse_args(argv)
        config = config_from_args(Path(__file__).resolve(), args)
        return run_differential_test(config)
    except RuntimeError as exc:
        logger.error("%s", exc)
        return 2
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        return 2
    except Exception:
        logger.exception("Differential test failed with an unhandled exception.")
        return 2


# --------------------------------------------------------------------------
# Workbook-specific hooks — Tiny DSA scenario sweep and cell mappings.
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


COUNTRIES: tuple[str, ...] = ("Borvelia", "Litellia", "Aurelium")
SHOCK_TYPES: tuple[int, ...] = (1, 2, 3)
SHOCK_YEARS: tuple[int, ...] = (1, 2, 3, 4, 5)
SHOCK_PARAMETER_LABELS: tuple[str, str, str] = ("Growth", "Interest", "Primary balance")
CONTINUOUS_AXES: tuple[tuple[str, str, tuple[float, ...]], ...] = (
    ("growth", "growth_baseline", (0.0, 1.5, 3.5, 5.5, 7.0)),
    ("interest", "interest_baseline", (0.0, 2.0, 4.0, 6.0, 8.0)),
    ("primary_balance", "primary_balance_baseline", (-3.0, -1.5, 0.0, 1.5, 3.0)),
)
PERTURBATION_YEARS: tuple[int, ...] = (1, 3, 5)
SHOCK_MAGNITUDE_SWEEP: tuple[float, ...] = (-3.0, -2.0, -1.0, 0.0, 1.0)

COUNTRY_NAME_CELL = "Inputs!B5"
SHOCK_YEAR_CELL = "Inputs!B21"
SHOCK_TYPE_CELL = "Inputs!B22"
SHOCK_TABLE_CELLS: tuple[str, str, str] = (
    "Inputs!B26",
    "Inputs!C26",
    "Inputs!D26",
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
    return tuple(value if i == year - 1 else v for i, v in enumerate(vec))


def _canonical_scenarios() -> Iterator[Scenario]:
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
    return (
        *_canonical_scenarios(),
        *_single_axis_perturbations(),
        *_categorical_combo_scenarios(),
    )


def output_cell_labels() -> tuple[tuple[str, str], ...]:
    return tuple(
        (f"{name}[year={i + 1}]", cell)
        for name, cells in OUTPUT_RANGES
        for i, cell in enumerate(cells)
    )


def output_ranges() -> tuple[tuple[str, tuple[str, ...]], ...]:
    return OUTPUT_RANGES


def _inputs_for_excel(inputs: Inputs) -> dict[str, Any]:
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


def inputs_for_excel(scenario: Scenario) -> dict[str, Any]:
    return _inputs_for_excel(scenario.inputs)


def _time_series_records(values: tuple[float, ...]) -> list[dict[str, Any]]:
    return [
        {"TIME_PERIOD": i + 1, "OBS_VALUE": value} for i, value in enumerate(values)
    ]


def apply_inputs_to_mvp(api: ModuleType, ctx: Any, scenario: Scenario) -> None:
    inputs = scenario.inputs
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


if __name__ == "__main__":
    sys.exit(main())
