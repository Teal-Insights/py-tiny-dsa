"""Named calculation functions for every bound formula series."""
from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from typing import Annotated, Literal, cast
from . import data
from .tensor import Domain, Series
from .excel import XlError, as_measure, xl_add, xl_bool, xl_choose_lazy, xl_div, xl_ge, xl_index, xl_match, xl_mul, xl_sub
from .runtime import Between, CoordinateReader, axis_step, evaluate, publish, view

@publish(key=(), domain=None, cells=data.INITIAL_DEBT_RESOLVED_CELLS)
def initial_debt_resolved(*, country_profile_names: data.Series[str | None], country_name: Literal["Aurelium", "Borvelia", "Litellia"], country_initial_debt: data.CountryInitialDebt) -> float | str:
    """Resolve the initial debt-to-GDP ratio for the selected country from the country profile table.

    Reproduce the Inputs!B6 INDEX/MATCH lookup that loads a country's initial debt-to-GDP ratio from the preloaded three-row profile table.

    Args:
        country_profile_names: Country names from the first column of the country profile table, aligned to the country axis; the lookup column matched against the selected country.
        country_name: User-selected country drawn from the profile table (Aurelium, Borvelia, or Litellia); an exact match on this name selects the profile row.
        country_initial_debt: Initial debt-to-GDP ratios (percent of GDP) from the second column of the profile table, aligned to the country axis; the value returned for the matched row.

    Returns:
        The initial debt-to-GDP ratio, in percent of GDP, from the profile row whose country name matches the selection; the corresponding Excel error code as a string if the lookup fails.
    """
    data.COUNTRY_PROFILE_NAMES.schema.validate(country_profile_names)
    data.COUNTRY_INITIAL_DEBT.schema.validate(country_initial_debt)
    _initial_debt_resolved_table_0 = view(country_initial_debt, rows=data.COUNTRY_AXIS.keys)
    _initial_debt_resolved_table_1 = view(country_profile_names, rows=data.COUNTRY_AXIS.keys)
    try:
        return as_measure(xl_index(_initial_debt_resolved_table_0, xl_match(country_name, _initial_debt_resolved_table_1, 0), 1))
    except XlError as error:
        return error.code

@publish(key=(), domain=None, cells=data.ENGINE_INITIAL_DEBT_BASELINE_CELLS)
def engine_initial_debt_baseline(*, initial_debt_resolved: float | str) -> float | str:
    """Resolve the year-0 debt stock anchor for the baseline recursion on the Engine sheet.

    Establishes the initial debt-to-GDP ratio, in percent of GDP, that seeds the five-year baseline debt-dynamics recursion.

    Args:
        initial_debt_resolved: Resolved initial debt-to-GDP ratio for the selected country, in percent of GDP, as returned by the INDEX/MATCH lookup against the country profile table; supplied either as a numeric value or as a text representation that may carry an Excel error code.

    Returns:
        The initial debt-to-GDP ratio as a numeric measure in percent of GDP, or the Excel error code string when the supplied anchor cannot be converted to a measure.
    """
    try:
        return as_measure(initial_debt_resolved)
    except XlError as error:
        return error.code

@publish(key=(), domain=None, cells=data.ENGINE_INITIAL_DEBT_SHOCKED_CELLS)
def engine_initial_debt_shocked(*, initial_debt_resolved: float | str) -> float | str:
    """Year-0 debt stock anchor for the shocked recursion on the Engine sheet.

    Resolve the initial debt-to-GDP ratio from which the shocked debt-dynamics recursion is started.

    Args:
        initial_debt_resolved: Year-0 general-government debt stock as a percent of GDP, resolved from the selected country profile and used as the anchor for the shocked path.

    Returns:
        The resolved year-0 debt-to-GDP ratio, or the Excel error code if the value cannot be resolved as a measure.
    """
    try:
        return as_measure(initial_debt_resolved)
    except XlError as error:
        return error.code

@publish(key=(), domain=None, cells=data.SHOCK_MAGNITUDE_RESOLVED_CELLS)
def shock_magnitude_resolved(*, shock_type: Literal[1, 2, 3], shock_magnitudes: data.ShockMagnitudes) -> float | str:
    """Resolve the shock magnitude for the selected shock type.

    Selects the applicable shock magnitude from the shock table by offsetting from the shock table anchor to the column matching the chosen shock type, so that only the magnitude associated with the active shock is applied.

    Args:
        shock_type: Shock type code (1 for real GDP growth, 2 for the real interest rate, 3 for the primary balance) identifying which parameter the shock affects.
        shock_magnitudes: Shock magnitudes, one per shock type in shock-type order, expressed in percentage points; values not matching the selected shock type are ignored.

    Returns:
        The shock magnitude, in percentage points, associated with the selected shock type; if the underlying lookup fails, the Excel error code is returned instead.
    """
    data.SHOCK_MAGNITUDES.schema.validate(shock_magnitudes)
    try:
        return as_measure(shock_magnitudes[axis_step(data.SHOCK_PARAMETER_AXIS, 'Growth', xl_sub(shock_type, 1))])
    except XlError as error:
        return error.code

@publish(data.SHOCK_ACTIVE.schema, cells=data.SHOCK_ACTIVE.cells)
def shock_active(*, engine_year_labels: data.Series[int | str | None], shock_year: Annotated[int, Between(1, 5)]) -> data.Series[int | str | None]:
    """Return the shock-activation flag for each projection year.

    Marks which horizon years fall inside the shock period, so the shocked parameter rows apply the shock from the shock year through the end of the five-year horizon.

    Args:
        engine_year_labels: Year labels for the five-year projection horizon on the Engine sheet, one per column, tested period by period against the shock year.
        shock_year: First year in which the shock takes effect, an integer between 1 and 5; the shock persists from this year through the end of the horizon.

    Returns:
        Series of integer flags aligned to the engine year labels, holding 1 for every year at or after the shock year and 0 for every earlier year.
    """
    data.ENGINE_YEAR_LABELS.schema.validate(engine_year_labels)
    def formula(time_period: int) -> int | str | None:
        return as_measure((1 if xl_bool(xl_ge(engine_year_labels[time_period], shock_year)) else 0), 'int')

    return data.SHOCK_ACTIVE.collect(evaluate(formula, data.SHOCK_ACTIVE.required))

@publish(data.SHOCKED_GROWTH.schema, cells=data.SHOCKED_GROWTH.cells)
def shocked_growth(*, growth_baseline: data.GrowthBaseline, shock_type: Literal[1, 2, 3], shock_magnitude_resolved: float | str, shock_active: data.Series[int | str | None]) -> data.Series[float | str | None]:
    """Build the shocked real GDP growth path by applying the resolved shock magnitude to the baseline growth series from the shock year onwards.

    Produce the year-by-year real GDP growth rates used by the shocked debt recursion.

    Args:
        growth_baseline: Baseline real GDP growth rates for years 1 through 5, in percent per annum.
        shock_type: Integer code identifying the parameter the shock affects (1 for real GDP growth, 2 for the real interest rate, 3 for the primary balance); the growth shock is applied only when this equals 1.
        shock_magnitude_resolved: Shock magnitude in percentage points for the selected shock type, resolved from the shock table; added to the baseline growth rate in each active year.
        shock_active: Per-year indicator series that is 1 in years at or after the shock year and 0 otherwise, so the shock persists symmetrically through the remainder of the horizon.

    Returns:
        A series of shocked real GDP growth rates for years 1 through 5, equal to the baseline growth path where the shock is inactive and to the baseline plus the resolved growth-shock magnitude where it is active.
    """
    data.GROWTH_BASELINE.schema.validate(growth_baseline)
    def formula(time_period: int) -> float | str | None:
        return as_measure(xl_add(growth_baseline[time_period], xl_mul(xl_choose_lazy(shock_type, lambda: shock_magnitude_resolved, lambda: 0, lambda: 0), shock_active[time_period])))

    return data.SHOCKED_GROWTH.collect(evaluate(formula, data.SHOCKED_GROWTH.required))

@publish(data.SHOCKED_INTEREST.schema, cells=data.SHOCKED_INTEREST.cells)
def shocked_interest(*, interest_baseline: data.InterestBaseline, shock_type: Literal[1, 2, 3], shock_magnitude_resolved: float | str, shock_active: data.Series[int | str | None]) -> data.Series[float | str | None]:
    """Apply the selected shock magnitude to the baseline real interest rate path when the shock type is the real interest rate.

    Produce the year-by-year shocked real interest rate series used by the shocked debt-dynamics recursion.

    Args:
        interest_baseline: Baseline real interest rates, in percent per annum, for years 1 through 5; the effective real rate paid on outstanding general-government debt before any shock is applied.
        shock_type: Integer code identifying the parameter affected by the shock: 1 for real GDP growth, 2 for the real interest rate, and 3 for the primary balance. The magnitude is added only when this equals 2.
        shock_magnitude_resolved: Shock magnitude, in percentage points, resolved from the shock table for the selected shock type; added to the baseline real interest rate in each active shock year.
        shock_active: Per-year indicator returning 1 in each year from the shock year through the end of the horizon and 0 otherwise; gates where the resolved magnitude is applied.

    Returns:
        The shocked real interest rate series, in percent per annum, for years 1 through 5, equal to the baseline path where the shock is inactive and to the baseline plus the resolved magnitude in active years when the shock type is the real interest rate.
    """
    data.INTEREST_BASELINE.schema.validate(interest_baseline)
    def formula(time_period: int) -> float | str | None:
        return as_measure(xl_add(interest_baseline[time_period], xl_mul(xl_choose_lazy(shock_type, lambda: 0, lambda: shock_magnitude_resolved, lambda: 0), shock_active[time_period])))

    return data.SHOCKED_INTEREST.collect(evaluate(formula, data.SHOCKED_INTEREST.required))

@publish(data.SHOCKED_PRIMARY_BALANCE.schema, cells=data.SHOCKED_PRIMARY_BALANCE.cells)
def shocked_primary_balance(*, primary_balance_baseline: data.PrimaryBalanceBaseline, shock_type: Literal[1, 2, 3], shock_magnitude_resolved: float | str, shock_active: data.Series[int | str | None]) -> data.Series[float | str | None]:
    """Shocked primary balance path after applying the selected shock magnitude.

    Builds the year-by-year primary balance series as a percent of GDP under the configured shock, leaving the baseline path unchanged in years where the shock is inactive.

    Args:
        primary_balance_baseline: Baseline primary balance path for years 1 through 5, expressed as a percent of GDP with positive values denoting a surplus.
        shock_type: Integer between 1 and 3 selecting the parameter affected by the shock: 1 for real GDP growth, 2 for the real interest rate, and 3 for the primary balance. Only type 3 applies a magnitude to this series.
        shock_magnitude_resolved: Shock magnitude for the selected shock type, in percentage points, resolved from the shock table. It is a reduction of the primary balance when the shock type is 3 and is inert otherwise.
        shock_active: Indicator series over years 1 through 5 that is 1 in each year from the shock year onwards and 0 before it, determining where the shock applies.

    Returns:
        The shocked primary balance path as a percent of GDP for years 1 through 5, equal to the baseline path before the shock year and to the baseline path adjusted by the resolved magnitude from the shock year onwards.
    """
    data.PRIMARY_BALANCE_BASELINE.schema.validate(primary_balance_baseline)
    def formula(time_period: int) -> float | str | None:
        return as_measure(xl_add(primary_balance_baseline[time_period], xl_mul(xl_choose_lazy(shock_type, lambda: 0, lambda: 0, lambda: shock_magnitude_resolved), shock_active[time_period])))

    return data.SHOCKED_PRIMARY_BALANCE.collect(evaluate(formula, data.SHOCKED_PRIMARY_BALANCE.required))

@publish(data.BASELINE_PATH_INTERNAL.schema, cells=data.BASELINE_PATH_INTERNAL.cells)
def baseline_path_internal(*, engine_initial_debt_baseline: float | str, growth_baseline: data.GrowthBaseline, interest_baseline: data.InterestBaseline, primary_balance_baseline: data.PrimaryBalanceBaseline) -> data.Series[float | str | None]:
    """Compute the Engine sheet's internal baseline debt-to-GDP path, `baseline_path`.

    Recurse the real debt-dynamics identity over years 1 through 5 to produce the baseline trajectory that the shock overlay is later applied to.

    Args:
        engine_initial_debt_baseline: Initial debt-to-GDP ratio in percent of GDP at the start of year 1, taken from the Engine sheet's baseline anchor and used as the seed value of the recursion.
        growth_baseline: Real GDP growth rates for years 1 through 5, in percent per annum, corresponding to the `growth_baseline` input vector (Inputs!C16:G16).
        interest_baseline: Effective real interest rates paid on outstanding general-government debt for years 1 through 5, in percent per annum, corresponding to the `interest_baseline` input vector (Inputs!C17:G17).
        primary_balance_baseline: Primary fiscal balances for years 1 through 5, as a percent of GDP, with positive values denoting a surplus, corresponding to the `primary_balance_baseline` input vector (Inputs!C18:G18).

    Returns:
        The Engine-sheet `baseline_path` series of debt-to-GDP ratios in percent of GDP for years 1 through 5, each year computed as the prior year's ratio times the real snowball factor (1 + r) / (1 + g) less the year's primary balance.
    """
    data.GROWTH_BASELINE.schema.validate(growth_baseline)
    data.INTEREST_BASELINE.schema.validate(interest_baseline)
    data.PRIMARY_BALANCE_BASELINE.schema.validate(primary_balance_baseline)
    def formula(time_period: int) -> float | str | None:
        if time_period == 1:
            return as_measure(xl_sub(xl_div(xl_mul(engine_initial_debt_baseline, xl_add(1, xl_div(interest_baseline[time_period], 100))), xl_add(1, xl_div(growth_baseline[time_period], 100))), primary_balance_baseline[time_period]))
        return as_measure(xl_sub(xl_div(xl_mul(baseline_path_internal[time_period - 1], xl_add(1, xl_div(interest_baseline[time_period], 100))), xl_add(1, xl_div(growth_baseline[time_period], 100))), primary_balance_baseline[time_period]))

    baseline_path_internal = CoordinateReader('baseline_path_internal', data.BASELINE_PATH_INTERNAL.required, formula)
    return data.BASELINE_PATH_INTERNAL.collect((coord, baseline_path_internal[coord]) for coord in data.BASELINE_PATH_INTERNAL.required)

@publish(data.SHOCKED_PATH_INTERNAL.schema, cells=data.SHOCKED_PATH_INTERNAL.cells)
def shocked_path_internal(*, engine_initial_debt_shocked: float | str, shocked_growth: data.Series[float | str | None], shocked_interest: data.Series[float | str | None], shocked_primary_balance: data.Series[float | str | None]) -> data.Series[float | str | None]:
    """Compute the internal shocked debt-to-GDP path on the Engine sheet.

    Recurse the real debt-dynamics identity over the five-year horizon using shock-adjusted growth, interest, and primary-balance parameters to produce the shocked debt-to-GDP trajectory behind the `shocked_path` named range.

    Args:
        engine_initial_debt_shocked: Initial debt-to-GDP ratio (in percent of GDP) at the start of the horizon, used as the year-1 starting stock for the shocked recursion.
        shocked_growth: Real GDP growth rates (in percent per annum) by year under the shock, indexing the shocked growth row applied in the snowball factor.
        shocked_interest: Real interest rates (in percent per annum) by year under the shock, indexing the shocked interest row applied in the snowball factor.
        shocked_primary_balance: Primary balances (in percent of GDP, positive denoting surplus) by year under the shock, deducted from the snowballed debt ratio in each period.

    Returns:
        Year-by-year shocked debt-to-GDP path as a series, with each year equal to the previous year's ratio times the real snowball factor (1 + r/100) / (1 + g/100) less the year's shocked primary balance, and the first year seeded from the shocked initial debt stock.
    """
    def formula(time_period: int) -> float | str | None:
        if time_period == 1:
            return as_measure(xl_sub(xl_div(xl_mul(engine_initial_debt_shocked, xl_add(1, xl_div(shocked_interest[time_period], 100))), xl_add(1, xl_div(shocked_growth[time_period], 100))), shocked_primary_balance[time_period]))
        return as_measure(xl_sub(xl_div(xl_mul(shocked_path_internal[time_period - 1], xl_add(1, xl_div(shocked_interest[time_period], 100))), xl_add(1, xl_div(shocked_growth[time_period], 100))), shocked_primary_balance[time_period]))

    shocked_path_internal = CoordinateReader('shocked_path_internal', data.SHOCKED_PATH_INTERNAL.required, formula)
    return data.SHOCKED_PATH_INTERNAL.collect((coord, shocked_path_internal[coord]) for coord in data.SHOCKED_PATH_INTERNAL.required)

@publish(data.OUTPUT_BASELINE.schema, cells=data.OUTPUT_BASELINE.cells)
def output_baseline(*, baseline_path_internal: data.Series[float | str | None]) -> data.OutputBaseline:
    """Populate the output baseline debt-to-GDP path from the internal baseline path.

    Publish the baseline debt-to-GDP trajectory for projection years 1 through 5 on the stable Outputs surface, one year per output period.

    Args:
        baseline_path_internal: Keyword-only series holding the calculated baseline debt-to-GDP path from the Engine surface, covering projection years 1 through 5, keyed by time period.

    Returns:
        The output baseline debt-to-GDP path for projection years 1 through 5, aligned to the output baseline coordinate and suitable for downstream consumers.
    """
    def formula(time_period: int) -> float | str | None:
        return as_measure(baseline_path_internal[time_period])

    return data.OUTPUT_BASELINE.collect(evaluate(formula, data.OUTPUT_BASELINE.required))

@publish(data.OUTPUT_SHOCKED.schema, cells=data.OUTPUT_SHOCKED.cells)
def output_shocked(*, shocked_path_internal: data.Series[float | str | None]) -> data.OutputShocked:
    """Convert the Engine-sheet shocked debt-to-GDP path into the stable Outputs-sheet series.

    Expose the shocked debt-to-GDP trajectory for projection years 1 through 5 as the downstream read-only output.

    Args:
        shocked_path_internal: Engine-sheet shocked debt-to-GDP path, in percent of GDP, for projection years 1 through 5; used as the internal computation surface rather than the stable output API.

    Returns:
        The shocked debt-to-GDP path as an Outputs-sheet series covering projection years 1 through 5, in percent of GDP.
    """
    def formula(time_period: int) -> float | str | None:
        return as_measure(shocked_path_internal[time_period])

    return data.OUTPUT_SHOCKED.collect(evaluate(formula, data.OUTPUT_SHOCKED.required))

@publish(data.OUTPUT_DELTA.schema, cells=data.OUTPUT_DELTA.cells)
def output_delta(*, output_baseline: data.OutputBaseline, output_shocked: data.OutputShocked) -> data.OutputDelta:
    """Compute the difference between the shocked and baseline debt-to-GDP paths in percentage points.

    Derive the year-by-year delta of the shocked path over the baseline path for downstream comparison on the Outputs sheet.

    Args:
        output_baseline: Baseline debt-to-GDP path, in percent of GDP, for years 1 through 5.
        output_shocked: Shocked debt-to-GDP path, in percent of GDP, for years 1 through 5, computed under the active shock configuration.

    Returns:
        The delta path, shocked minus baseline, in percentage points of GDP for years 1 through 5.
    """
    def formula(time_period: int) -> float | str | None:
        return as_measure(xl_sub(output_shocked[time_period], output_baseline[time_period]))

    return data.OUTPUT_DELTA.collect(evaluate(formula, data.OUTPUT_DELTA.required))
