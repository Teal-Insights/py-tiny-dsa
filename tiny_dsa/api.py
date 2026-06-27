from __future__ import annotations

from .data import CONSTANTS, DEFAULT_INPUTS
from .internals import _resolve_formula
from ._api_helpers import Record, Records, Scalar, SeriesInput, _apply_series_records, _coerce_records, coerce_setter_input
from .runtime import EvalContext, coerce_inputs_dict, xl_cell, xl_range
import warnings


def make_context(inputs=None):
    """Create an EvalContext with merged inputs."""
    merged = dict(DEFAULT_INPUTS)
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
    """Set the selected country name for the Tiny-DSA scenario.

    Updates the country-name input used by the workbook to derive the active country profile.
    Because this series has no dimensions, the supplied scalar record maps to the single country-name input cell.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: User-selected country name for the scenario, used by the workbook's country-profile lookup.
            Optional record fields:
                - PARAMETER: Optional context attribute identifying the workbook parameter described by the record. If supplied, expected value: "country_name".

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
) -> None:
    """Set the country profile table's initial debt-to-GDP series.

    Updates the initial debt-to-GDP values used by the country lookup table.
    Each record is matched to a country row label in the profile table, and its observation value is written to that row's initial-debt cell.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
            Required record fields:
                - COUNTRY: Country profile row to update.
                - OBS_VALUE: Initial debt-to-GDP ratio for the country profile entry.
            Optional record fields:
                - INDICATOR: Identifies the series as the initial-debt indicator. If supplied, expected value: "initial_debt_to_gdp".
                - UNIT_MEASURE: Identifies the unit used for the debt-ratio observation. If supplied, expected value: "PC_GDP".

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
) -> None:
    """Set the baseline real GDP growth path for the projection horizon.

    Updates the baseline real GDP growth assumptions used by the debt-dynamics recursion.
    Records are keyed by projection year, with each observation written to the matching year cell in the series.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
            Required record fields:
                - TIME_PERIOD: Projection year identifying the position in the baseline path.
                - OBS_VALUE: Baseline real GDP growth rate for the projection year.
            Optional record fields:
                - INDICATOR: Identifies the macroeconomic series represented by the record. If supplied, expected value: "real_gdp_growth".
                - UNIT_MEASURE: Indicates the measurement unit used for the growth observation. If supplied, expected value: "PERCENT_PER_ANNUM".

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
) -> None:
    """Set the baseline real interest-rate path for the projection horizon.

    Updates the real interest-rate assumptions used in Tiny-DSA's baseline debt recursion.
    Records are matched by projection year to the corresponding columns in the workbook series.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
            Required record fields:
                - TIME_PERIOD: Projection year identifying the position in the baseline interest-rate path.
                - OBS_VALUE: Effective real interest rate paid on outstanding general-government debt during the year, used in the real-terms snowball factor.
            Optional record fields:
                - INDICATOR: Identifies the macroeconomic indicator represented by the series. If supplied, expected value: "real_interest_rate".
                - UNIT_MEASURE: Identifies the unit convention for the observation value. If supplied, expected value: "PERCENT_PER_ANNUM".

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
) -> None:
    """Set the baseline primary-balance path for the projection horizon.

    Updates the exogenous fiscal-stance series used in the baseline debt-to-GDP recursion.
    Records are matched by projection year to the corresponding cell in the baseline primary-balance series.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
            Required record fields:
                - TIME_PERIOD: Projection year identifying the baseline horizon cell for the observation.
                - OBS_VALUE: Primary balance for the projection year; positive values denote a surplus and negative values denote a deficit.
            Optional record fields:
                - INDICATOR: Indicator identifying which baseline parameter series the record belongs to. If supplied, expected value: "primary_balance".
                - UNIT_MEASURE: Unit of measure associated with the primary-balance observation. If supplied, expected value: "PC_GDP".

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
    """Set the shock-year input for the Tiny-DSA scenario.

    Updates the first projection year in which the selected shock begins and applies through the remaining horizon.
    Because this is a scalar input, the supplied observation maps to the single shock-year workbook cell.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: Shock-year measure indicating the first projection year in which the selected shock takes effect.
            Optional record fields:
                - PARAMETER: Optional parameter attribute identifying the record as part of the shock-year input. If supplied, expected value: "shock_year".

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
    """Set the shock type used by Tiny-DSA's shock configuration.

    Updates the scalar input that determines which baseline parameter receives the configured shock.
    Because this is a scalar input, the supplied record maps directly to the single shock type cell.

    Args:
        records (Scalar | Record | Records): A bare scalar value, a single record dict, or a list of records.
            Required record fields:
                - OBS_VALUE: Integer code identifying which parameter is affected by the configured shock.
            Optional record fields:
                - PARAMETER: Optional parameter attribute identifying the shock-configuration input represented by the record. If supplied, expected value: "shock_type".

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
) -> None:
    """Set shock magnitudes by affected macro-fiscal parameter.

    Updates the shock table values used to configure the single active shock scenario.
    Records are matched by SHOCK_PARAMETER to the corresponding shock-table column.

    Args:
        records (SeriesInput): A list of records, a single record dict, a tidy pandas/polars DataFrame, or a 1-D iterable of measure values in key order.
            Required record fields:
                - SHOCK_PARAMETER: Affected parameter associated with a shock magnitude.
                - OBS_VALUE: Magnitude applied to the affected parameter when that shock is selected.
            Optional record fields:
                - PARAMETER: Identifies the parameter series represented by the record when included. If supplied, expected value: "shock_magnitude".
                - UNIT_MEASURE: Indicates the unit used for the shock magnitude when included. If supplied, expected value: "PP".

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
    """Compute the baseline debt-to-GDP output path.

    Returns the Outputs-sheet baseline trajectory computed from the baseline macroeconomic and fiscal assumptions.
    Each record corresponds to one projection-year cell in the output series, keyed by projection year.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year in the five-year horizon.
                - OBS_VALUE: Baseline general-government debt-to-GDP ratio for the projection year.
            Optional record fields:
                - SCENARIO: Identifies which trajectory the record belongs to. If supplied, expected value: "baseline".
                - UNIT_MEASURE: Indicates the unit used to express the debt ratio. If supplied, expected value: "PC_GDP".

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
    """Compute the shocked debt-to-GDP output path by projection year.

    Returns the shocked debt-to-GDP trajectory exposed on the Outputs sheet.
    Each record corresponds to one projection-year output cell in the shocked path.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year in the five-year horizon.
                - OBS_VALUE: Projected gross general-government debt-to-GDP ratio for the shocked path.
            Optional record fields:
                - SCENARIO: Scenario label identifying the shocked trajectory. If supplied, expected value: "shocked".
                - UNIT_MEASURE: Unit used to express the debt-to-GDP ratio. If supplied, expected value: "PC_GDP".

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
    """Compute the output delta series comparing the shocked and baseline debt-to-GDP paths.

    Returns the five-year difference between the shocked and baseline debt-to-GDP trajectories.
    Each record corresponds to one projection-year cell in the output delta row, with the column header identifying the year.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year in the five-year horizon.
                - OBS_VALUE: Difference between the shocked and baseline debt-to-GDP paths for the projection year.
            Optional record fields:
                - SCENARIO: Identifies the comparison scenario represented by the series. If supplied, expected value: "shocked_minus_baseline".
                - UNIT_MEASURE: Provides the unit metadata for the delta values. If supplied, expected value: "PP".

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


TARGETS = {
    'Outputs!B12:Outputs!F12': xl_range,
    'Outputs!B13:Outputs!F13': xl_range,
    'Outputs!B14:Outputs!F14': xl_range,
}


def compute_all(ctx=None, *, inputs=None):
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
