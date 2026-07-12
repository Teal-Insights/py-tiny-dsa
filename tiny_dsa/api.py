from __future__ import annotations

from .data import CONSTANTS, DEFAULT_INPUTS
from .internals import _resolve_formula
from ._api_helpers import EmptyMeasure, Record, Records, Scalar, SeriesInput, _apply_series_records, _coerce_records, coerce_setter_input
from .runtime import EvalContext, coerce_inputs_dict, xl_cell, xl_range_rows
import warnings


def make_context(inputs: dict[str, object] | None = None) -> EvalContext:
    """Create an EvalContext with merged inputs."""
    merged: dict[str, object] = dict(DEFAULT_INPUTS)
    merged.update(CONSTANTS)
    if inputs is not None:
        merged.update(inputs)
    return EvalContext(inputs=coerce_inputs_dict(merged), resolver=_resolve_formula, iterative_enabled=False, iterate_count=100, iterate_delta=0.001)


# --- Series binding setters (Records API) ---

_LEAF_INDEX_COUNTRY_NAME = {
    (): 'Inputs!B5',
}

def set_country_name(
    ctx: EvalContext,
    records: Records | Record | Scalar,
    *,
    strict: bool = True,
) -> None:
    """Set the country name for the debt sustainability analysis.

    Updates the selected country name in the Inputs sheet, which drives the initial debt-to-GDP ratio lookup.
    Input is normalized to a record with PARAMETER 'country_name' and OBS_VALUE; the setter checks record structure and key fields but does not validate the country name against the profile table.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: The name of the country to set.
            Optional record fields:
                - PARAMETER: The parameter being set, identifying the workbook input cell. If supplied, expected value: "country_name".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B5
        Layout: scalar
        Value type: string

    Examples:
        set_country_name(ctx, 'Borvelia')
    """
    _apply_series_records(
        ctx,
        _coerce_records(records, 'OBS_VALUE', allow_scalar=True),
        key_fields=(),
        allowed_fields=frozenset({'OBS_VALUE', 'PARAMETER'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_COUNTRY_NAME,
        strict=strict,
        fn_name='set_country_name',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_COUNTRY_INITIAL_DEBT = {
    (('COUNTRY', 'Borvelia'),): 'Inputs!B10',
    (('COUNTRY', 'Litellia'),): 'Inputs!B11',
    (('COUNTRY', 'Aurelium'),): 'Inputs!B12',
}

_KEY_ORDER_COUNTRY_INITIAL_DEBT = ('Aurelium', 'Borvelia', 'Litellia')

def set_country_initial_debt(
    ctx: EvalContext,
    records: SeriesInput,
    *,
    strict: bool = True,
    empty_measure: EmptyMeasure = "write",
) -> None:
    """Set the initial debt-to-GDP ratios for countries in the profile table.

    Updates the initial debt-to-GDP values for the selected countries in the Tiny-DSA country profile lookup table.
    Each record matches a row in the country profile table by the COUNTRY field; the OBS_VALUE is written to the corresponding initial debt cell.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
        empty_measure (EmptyMeasure): How to treat rows with missing measure values (`None` or float NaN after DataFrame coercion). "write" (default) passes values through; "skip" drops them; "error" raises. Empty key fields always raise.
            Required record fields:
                - COUNTRY: Country name as listed in the country profile table.
                - OBS_VALUE: The initial debt-to-GDP ratio, expressed as a percentage of GDP.
            Optional record fields:
                - INDICATOR: The economic indicator that the observation value represents. If supplied, expected value: "initial_debt_to_gdp".
                - UNIT_MEASURE: Unit of measurement for the debt-to-GDP ratio. If supplied, expected value: "PC_GDP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B10:B12
        Layout: series
        Value type: float

    Examples:
        set_country_initial_debt(ctx, [
            {'COUNTRY': 'Borvelia', 'OBS_VALUE': 60.0},
            {'COUNTRY': 'Litellia', 'OBS_VALUE': 80.0},
        ])

        set_country_initial_debt(ctx, [60.0, 80.0])
    """
    _apply_series_records(
        ctx,
        coerce_setter_input(
            records,
            layout='series',
            key_fields=('COUNTRY',),
            measure_field='OBS_VALUE',
            key_order=_KEY_ORDER_COUNTRY_INITIAL_DEBT,
            strict=strict,
            empty_measure=empty_measure,
            requires_address=False,
            key_dtypes={'COUNTRY': 'string'},
        ),
        key_fields=('COUNTRY',),
        allowed_fields=frozenset({'COUNTRY', 'INDICATOR', 'OBS_VALUE', 'UNIT_MEASURE'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_COUNTRY_INITIAL_DEBT,
        strict=strict,
        fn_name='set_country_initial_debt',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_GROWTH_BASELINE = {
    (('TIME_PERIOD', 1),): 'Inputs!C16',
    (('TIME_PERIOD', 2),): 'Inputs!D16',
    (('TIME_PERIOD', 3),): 'Inputs!E16',
    (('TIME_PERIOD', 4),): 'Inputs!F16',
    (('TIME_PERIOD', 5),): 'Inputs!G16',
}

_KEY_ORDER_GROWTH_BASELINE = (1, 2, 3, 4, 5)

def set_growth_baseline(
    ctx: EvalContext,
    records: SeriesInput,
    *,
    strict: bool = True,
    empty_measure: EmptyMeasure = "write",
) -> None:
    """Set baseline real GDP growth rates for projection years 1 through 5.

    Updates the growth_baseline named range (Inputs!C16:G16) with the provided growth values.
    Each record’s OBS_VALUE is assigned to the cell in the growth_baseline range corresponding to its TIME_PERIOD.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
        empty_measure (EmptyMeasure): How to treat rows with missing measure values (`None` or float NaN after DataFrame coercion). "write" (default) passes values through; "skip" drops them; "error" raises. Empty key fields always raise.
            Required record fields:
                - TIME_PERIOD: Projection year.
                - OBS_VALUE: Real GDP growth rate in percent per annum.
            Optional record fields:
                - INDICATOR: Economic indicator for the series. If supplied, expected value: "real_gdp_growth".
                - UNIT_MEASURE: Unit of measure for the observation value. If supplied, expected value: "PERCENT_PER_ANNUM".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C16:G16
        Layout: series
        Value type: float

    Examples:
        set_growth_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': 3.5},
            {'TIME_PERIOD': 2, 'OBS_VALUE': 3.5},
        ])

        set_growth_baseline(ctx, [3.5, 3.5])
    """
    _apply_series_records(
        ctx,
        coerce_setter_input(
            records,
            layout='series',
            key_fields=('TIME_PERIOD',),
            measure_field='OBS_VALUE',
            key_order=_KEY_ORDER_GROWTH_BASELINE,
            strict=strict,
            empty_measure=empty_measure,
            requires_address=False,
            key_dtypes={'TIME_PERIOD': 'int'},
        ),
        key_fields=('TIME_PERIOD',),
        allowed_fields=frozenset({'INDICATOR', 'OBS_VALUE', 'TIME_PERIOD', 'UNIT_MEASURE'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_GROWTH_BASELINE,
        strict=strict,
        fn_name='set_growth_baseline',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_INTEREST_BASELINE = {
    (('TIME_PERIOD', 1),): 'Inputs!C17',
    (('TIME_PERIOD', 2),): 'Inputs!D17',
    (('TIME_PERIOD', 3),): 'Inputs!E17',
    (('TIME_PERIOD', 4),): 'Inputs!F17',
    (('TIME_PERIOD', 5),): 'Inputs!G17',
}

_KEY_ORDER_INTEREST_BASELINE = (1, 2, 3, 4, 5)

def set_interest_baseline(
    ctx: EvalContext,
    records: SeriesInput,
    *,
    strict: bool = True,
    empty_measure: EmptyMeasure = "write",
) -> None:
    """Set the baseline real interest rates for projection years 1 through 5.

    Sets the annual real interest rate baseline values in the Inputs sheet.
    Each record's `TIME_PERIOD` maps to a projection year header, and `OBS_VALUE` sets the corresponding cell in the Inputs!C17:G17 range.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
        empty_measure (EmptyMeasure): How to treat rows with missing measure values (`None` or float NaN after DataFrame coercion). "write" (default) passes values through; "skip" drops them; "error" raises. Empty key fields always raise.
            Required record fields:
                - TIME_PERIOD: The projection year identifying the column in the baseline range.
                - OBS_VALUE: The real interest rate value for the projection year.
            Optional record fields:
                - INDICATOR: A constant identifying the series as real interest rate data. If supplied, expected value: "real_interest_rate".
                - UNIT_MEASURE: The unit of measure for the interest rate values. If supplied, expected value: "PERCENT_PER_ANNUM".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C17:G17
        Layout: series
        Value type: float

    Examples:
        set_interest_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': 4.0},
            {'TIME_PERIOD': 2, 'OBS_VALUE': 4.0},
        ])

        set_interest_baseline(ctx, [4.0, 4.0])
    """
    _apply_series_records(
        ctx,
        coerce_setter_input(
            records,
            layout='series',
            key_fields=('TIME_PERIOD',),
            measure_field='OBS_VALUE',
            key_order=_KEY_ORDER_INTEREST_BASELINE,
            strict=strict,
            empty_measure=empty_measure,
            requires_address=False,
            key_dtypes={'TIME_PERIOD': 'int'},
        ),
        key_fields=('TIME_PERIOD',),
        allowed_fields=frozenset({'INDICATOR', 'OBS_VALUE', 'TIME_PERIOD', 'UNIT_MEASURE'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_INTEREST_BASELINE,
        strict=strict,
        fn_name='set_interest_baseline',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_PRIMARY_BALANCE_BASELINE = {
    (('TIME_PERIOD', 1),): 'Inputs!C18',
    (('TIME_PERIOD', 2),): 'Inputs!D18',
    (('TIME_PERIOD', 3),): 'Inputs!E18',
    (('TIME_PERIOD', 4),): 'Inputs!F18',
    (('TIME_PERIOD', 5),): 'Inputs!G18',
}

_KEY_ORDER_PRIMARY_BALANCE_BASELINE = (1, 2, 3, 4, 5)

def set_primary_balance_baseline(
    ctx: EvalContext,
    records: SeriesInput,
    *,
    strict: bool = True,
    empty_measure: EmptyMeasure = "write",
) -> None:
    """Set the baseline primary balance path for projection years 1 through 5.

    Updates the primary balance baseline series with the provided records.
    Each record supplies a projection year (TIME_PERIOD) and its corresponding primary balance value (OBS_VALUE) for the baseline path.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
        empty_measure (EmptyMeasure): How to treat rows with missing measure values (`None` or float NaN after DataFrame coercion). "write" (default) passes values through; "skip" drops them; "error" raises. Empty key fields always raise.
            Required record fields:
                - TIME_PERIOD: Projection year index, from 1 to 5.
                - OBS_VALUE: Primary balance as a percent of GDP; positive values denote a surplus.
            Optional record fields:
                - INDICATOR: Series indicator identifier. If supplied, expected value: "primary_balance".
                - UNIT_MEASURE: Unit of measure for the observation values. If supplied, expected value: "PC_GDP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C18:G18
        Layout: series
        Value type: float

    Examples:
        set_primary_balance_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': -1.0},
            {'TIME_PERIOD': 2, 'OBS_VALUE': -0.5},
        ])

        set_primary_balance_baseline(ctx, [-1.0, -0.5])
    """
    _apply_series_records(
        ctx,
        coerce_setter_input(
            records,
            layout='series',
            key_fields=('TIME_PERIOD',),
            measure_field='OBS_VALUE',
            key_order=_KEY_ORDER_PRIMARY_BALANCE_BASELINE,
            strict=strict,
            empty_measure=empty_measure,
            requires_address=False,
            key_dtypes={'TIME_PERIOD': 'int'},
        ),
        key_fields=('TIME_PERIOD',),
        allowed_fields=frozenset({'INDICATOR', 'OBS_VALUE', 'TIME_PERIOD', 'UNIT_MEASURE'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_PRIMARY_BALANCE_BASELINE,
        strict=strict,
        fn_name='set_primary_balance_baseline',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_SHOCK_YEAR = {
    (): 'Inputs!B21',
}

def set_shock_year(
    ctx: EvalContext,
    records: Records | Record | Scalar,
    *,
    strict: bool = True,
) -> None:
    """Set the shock year in the Tiny-DSA Inputs sheet.

    Updates the value of the shock year, an integer between 1 and 5 that determines the first projection year in which the selected shock takes effect.
    Each record corresponds to the scalar shock year value; the OBS_VALUE field holds the integer year.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: The integer year, from 1 to 5, in which the shock begins.
            Optional record fields:
                - PARAMETER: Identifies the record as belonging to the shock_year parameter. If supplied, expected value: "shock_year".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B21
        Layout: scalar
        Value type: int

    Examples:
        set_shock_year(ctx, 2)
    """
    _apply_series_records(
        ctx,
        _coerce_records(records, 'OBS_VALUE', allow_scalar=True),
        key_fields=(),
        allowed_fields=frozenset({'OBS_VALUE', 'PARAMETER'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_SHOCK_YEAR,
        strict=strict,
        fn_name='set_shock_year',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_SHOCK_TYPE = {
    (): 'Inputs!B22',
}

def set_shock_type(
    ctx: EvalContext,
    records: Records | Record | Scalar,
    *,
    strict: bool = True,
) -> None:
    """Set the shock type for the scenario configuration.

    Updates the shock type in the Inputs sheet to specify which macroeconomic or fiscal parameter the shock affects.
    The OBS_VALUE field of the record is written to the scalar cell Inputs!B22.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: The shock type code: 1 for real GDP growth, 2 for real interest rate, 3 for primary balance.
            Optional record fields:
                - PARAMETER: Identifies the conceptual parameter being configured. If supplied, expected value: "shock_type".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B22
        Layout: scalar
        Value type: int

    Examples:
        set_shock_type(ctx, 1)
    """
    _apply_series_records(
        ctx,
        _coerce_records(records, 'OBS_VALUE', allow_scalar=True),
        key_fields=(),
        allowed_fields=frozenset({'OBS_VALUE', 'PARAMETER'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_SHOCK_TYPE,
        strict=strict,
        fn_name='set_shock_type',
        allow_address=False,
        requires_address=False,
    )

_LEAF_INDEX_SHOCK_MAGNITUDES = {
    (('SHOCK_PARAMETER', 'Growth'),): 'Inputs!B26',
    (('SHOCK_PARAMETER', 'Interest'),): 'Inputs!C26',
    (('SHOCK_PARAMETER', 'Primary balance'),): 'Inputs!D26',
}

_KEY_ORDER_SHOCK_MAGNITUDES = ('Growth', 'Interest', 'Primary balance')

def set_shock_magnitudes(
    ctx: EvalContext,
    records: SeriesInput,
    *,
    strict: bool = True,
    empty_measure: EmptyMeasure = "write",
) -> None:
    """Set the shock magnitudes for each shock parameter (growth, interest, primary balance).

    Defines the magnitude of the shock for each of the three possible affected parameters.
    Each record represents one shock magnitude cell in the shock table (Inputs!B26:D26); the SHOCK_PARAMETER field maps to the column header.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
        empty_measure (EmptyMeasure): How to treat rows with missing measure values (`None` or float NaN after DataFrame coercion). "write" (default) passes values through; "skip" drops them; "error" raises. Empty key fields always raise.
            Required record fields:
                - SHOCK_PARAMETER: The parameter affected by the shock (e.g., 'Growth', 'Interest', or 'Primary Balance').
                - OBS_VALUE: The shock magnitude applied to the parameter.
            Optional record fields:
                - PARAMETER: Constant field identifying the data as shock magnitude entries. If supplied, expected value: "shock_magnitude".
                - UNIT_MEASURE: Unit of measure for the shock magnitude. If supplied, expected value: "PP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B26:D26
        Layout: series
        Value type: float

    Examples:
        set_shock_magnitudes(ctx, [
            {'SHOCK_PARAMETER': 'Growth', 'OBS_VALUE': -2.0},
            {'SHOCK_PARAMETER': 'Interest', 'OBS_VALUE': 2.0},
        ])

        set_shock_magnitudes(ctx, [-2.0, 2.0])
    """
    _apply_series_records(
        ctx,
        coerce_setter_input(
            records,
            layout='series',
            key_fields=('SHOCK_PARAMETER',),
            measure_field='OBS_VALUE',
            key_order=_KEY_ORDER_SHOCK_MAGNITUDES,
            strict=strict,
            empty_measure=empty_measure,
            requires_address=False,
            key_dtypes={'SHOCK_PARAMETER': 'string'},
        ),
        key_fields=('SHOCK_PARAMETER',),
        allowed_fields=frozenset({'OBS_VALUE', 'PARAMETER', 'SHOCK_PARAMETER', 'UNIT_MEASURE'}),
        measure_field='OBS_VALUE',
        leaf_index=_LEAF_INDEX_SHOCK_MAGNITUDES,
        strict=strict,
        fn_name='set_shock_magnitudes',
        allow_address=False,
        requires_address=False,
    )

# --- Series binding output compute (Records API) ---

_OUTPUT_LEAVES_OUTPUT_BASELINE: list[tuple[str, Record]] = [
    ('Outputs!B12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!C12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!D12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!E12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!F12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PC_GDP'}),
]

def compute_output_baseline(ctx=None, *, inputs=None) -> Records:
    """Compute the baseline debt-to-GDP trajectory for projection years 1 through 5.

    Return the baseline debt-to-GDP path as a list of records.
    Each record maps to a cell in Outputs!B12:F12, with TIME_PERIOD as the year and OBS_VALUE as the debt-to-GDP ratio.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year (1 through 5).
                - OBS_VALUE: Debt-to-GDP ratio as a percentage of GDP.
            Optional record fields:
                - SCENARIO: Scenario identifier distinguishing baseline from shocked output. If supplied, expected value: "baseline".
                - UNIT_MEASURE: Unit of measure for the OBS_VALUE. If supplied, expected value: "PC_GDP".

    Source binding:
        Workbook range: Outputs!B12:F12
        Layout: series
        Value type: float

    Examples:
        compute_output_baseline(ctx=ctx)
    """
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn("inputs will be ignored because ctx was provided", UserWarning, stacklevel=2)
    measure_field = 'OBS_VALUE'
    include_address = False
    records: Records = []
    for address, static_record in _OUTPUT_LEAVES_OUTPUT_BASELINE:
        record = dict(static_record)
        record[measure_field] = xl_cell(ctx, address)
        if include_address:
            record["address"] = address
        records.append(record)
    return records

_OUTPUT_LEAVES_OUTPUT_SHOCKED: list[tuple[str, Record]] = [
    ('Outputs!B13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!C13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!D13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!E13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!F13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PC_GDP'}),
]

def compute_output_shocked(ctx=None, *, inputs=None) -> Records:
    """Returns the shocked debt-to-GDP path as a series of records.

    Provides the projected debt-to-GDP trajectory under the shocked scenario for projection years 1 through 5.
    Each record represents one projection year, with the observation value in percent of GDP.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year (1–5).
                - OBS_VALUE: Debt-to-GDP ratio under the shocked scenario.
            Optional record fields:
                - SCENARIO: Scenario identifier. If supplied, expected value: "shocked".
                - UNIT_MEASURE: Unit of measure for the observation value. If supplied, expected value: "PC_GDP".

    Source binding:
        Workbook range: Outputs!B13:F13
        Layout: series
        Value type: float

    Examples:
        compute_output_shocked(ctx=ctx)
    """
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn("inputs will be ignored because ctx was provided", UserWarning, stacklevel=2)
    measure_field = 'OBS_VALUE'
    include_address = False
    records: Records = []
    for address, static_record in _OUTPUT_LEAVES_OUTPUT_SHOCKED:
        record = dict(static_record)
        record[measure_field] = xl_cell(ctx, address)
        if include_address:
            record["address"] = address
        records.append(record)
    return records

_OUTPUT_LEAVES_OUTPUT_DELTA: list[tuple[str, Record]] = [
    ('Outputs!B14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!C14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!D14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!E14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!F14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PP'}),
]

def compute_output_delta(ctx=None, *, inputs=None) -> Records:
    """Compute the difference between the shocked and baseline debt-to-GDP paths over the projection horizon.

    Return a list of records, each containing a projection year and the corresponding difference, in percentage points, between the shocked and baseline debt-to-GDP ratios.
    Each record corresponds to one cell in the Outputs!B14:F14 range, providing the delta value for a single projection year.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year identifier, ranging from 1 to 5.
                - OBS_VALUE: Difference between the shocked and baseline debt-to-GDP ratios, expressed in percentage points of GDP.
            Optional record fields:
                - SCENARIO: Scenario classification for the delta series. If supplied, expected value: "shocked_minus_baseline".
                - UNIT_MEASURE: Unit of measurement for the debt ratio difference. If supplied, expected value: "PP".

    Source binding:
        Workbook range: Outputs!B14:F14
        Layout: series
        Value type: float

    Examples:
        compute_output_delta(ctx=ctx)
    """
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn("inputs will be ignored because ctx was provided", UserWarning, stacklevel=2)
    measure_field = 'OBS_VALUE'
    include_address = False
    records: Records = []
    for address, static_record in _OUTPUT_LEAVES_OUTPUT_DELTA:
        record = dict(static_record)
        record[measure_field] = xl_cell(ctx, address)
        if include_address:
            record["address"] = address
        records.append(record)
    return records


def list_setters() -> list[str]:
    """Return generated series-binding setter function names."""
    return ['set_country_initial_debt', 'set_country_name', 'set_growth_baseline', 'set_interest_baseline', 'set_primary_balance_baseline', 'set_shock_magnitudes', 'set_shock_type', 'set_shock_year']


def list_computes() -> list[str]:
    """Return generated series-binding compute function names."""
    return ['compute_output_baseline', 'compute_output_delta', 'compute_output_shocked']


TARGETS = {
    'Outputs!B12:Outputs!F12': xl_range_rows,
    'Outputs!B13:Outputs!F13': xl_range_rows,
    'Outputs!B14:Outputs!F14': xl_range_rows,
}


def compute_all(ctx: EvalContext | None = None, *, inputs: dict[str, object] | None = None) -> dict[str, object]:
    """Compute all target cells and return results."""
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn(
            "inputs will be ignored because ctx was provided",
            UserWarning,
            stacklevel=2,
        )
    return {target: handler(ctx, target) for target, handler in TARGETS.items()}
