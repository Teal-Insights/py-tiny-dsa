from __future__ import annotations

from .data import CONSTANTS, DEFAULT_INPUTS
from .internals import _resolve_formula
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

Scalar = str | int | float | bool | None
Record = dict[str, object]
Records = list[Record]

def _coerce_records(records, measure_field, *, allow_scalar=False) -> Records:
    if not allow_scalar:
        return records
    if not isinstance(records, list):
        if isinstance(records, dict):
            return [records]
        return [{measure_field: records}]
    return records

def _apply_series_records(
    ctx,
    records,
    *,
    key_fields,
    allowed_fields,
    measure_field,
    leaf_index,
    strict,
    fn_name,
    allow_address=False,
    requires_address=False,
) -> None:
    updates: dict[str, object] = {}
    for index, record in enumerate(records):
        if strict:
            unknown = set(record) - allowed_fields
            if unknown:
                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")
        if measure_field not in record:
            raise ValueError(f"record[{index}]: missing required field {measure_field!r}")
        address = None
        if allow_address or requires_address:
            address = record.get("address") or record.get("cell_address")
        if requires_address and address is None:
            raise ValueError(
                f"record[{index}]: address required for {fn_name} (duplicate keys in binding)"
            )
        if address is None:
            if not requires_address:
                missing = [field for field in key_fields if field not in record]
                if missing:
                    raise ValueError(f"record[{index}]: missing key fields {missing!r}")
                key_tuple = tuple((field, record[field]) for field in key_fields)
                address = leaf_index.get(key_tuple)
                if address is None:
                    raise ValueError(
                        f"record[{index}]: no leaf matches key {dict(key_tuple)!r}"
                    )
        updates[address] = record[measure_field]
    if updates:
        ctx.set_inputs(coerce_inputs_dict(updates))

_LEAF_INDEX_COUNTRY_NAME = {
    (): 'Inputs!B5',
}

def set_country_name(
    ctx: EvalContext,
    records: Records | Record | Scalar,
    *,
    strict: bool = True,
) -> None:
    """Set the country name in the workbook.

    Updates the country name cell (Inputs!B5) to the provided value.
    Each record writes its OBS_VALUE to the country_name named range.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - OBS_VALUE: The country name to set in the workbook.
            Optional record fields:
                - PARAMETER: Identifies the parameter being set. If supplied, expected value: "country_name".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B5
        Layout: scalar
        Value type: string

    Examples:
        set_country_name(ctx, [
            {'OBS_VALUE': 'Borvelia'},
        ])
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

def set_country_initial_debt(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
    """Set the initial debt-to-GDP ratios in the country profile lookup table.

    Updates the initial debt-to-GDP ratios for each country in the profile table used for country selection and debt projection.
    Each record maps to a row in the country profile table, where COUNTRY is the row label in column A and OBS_VALUE is the cell value in column B.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - COUNTRY: The country name that identifies the row in the country profile lookup table.
                - OBS_VALUE: The initial debt-to-GDP ratio, expressed as a percent of GDP.
            Optional record fields:
                - INDICATOR: The indicator that classifies this series as the initial debt-to-GDP data. If supplied, expected value: "initial_debt_to_gdp".
                - UNIT_MEASURE: The unit of measurement for the debt-to-GDP ratio. If supplied, expected value: "PC_GDP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B10:B12
        Layout: row_series
        Value type: float

    Examples:
        set_country_initial_debt(ctx, [
            {'COUNTRY': 'Borvelia', 'OBS_VALUE': 60.0},
            {'COUNTRY': 'Litellia', 'OBS_VALUE': 80.0},
        ])
    """
    _apply_series_records(
        ctx,
        records,
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

def set_growth_baseline(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
    """Set baseline real GDP growth rates for projection years 1 through 5.

    Updates the real GDP growth assumptions used in the baseline debt projection.
    Each record corresponds to a cell in `Inputs!C16:G16`, with `TIME_PERIOD` mapping to the column header year and `OBS_VALUE` mapping to the cell value.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - TIME_PERIOD: Projection year (1 to 5) for which the growth rate applies.
                - OBS_VALUE: The baseline real GDP growth rate, expressed as a percent per annum.
            Optional record fields:
                - INDICATOR: The economic indicator that this series represents. If supplied, expected value: "real_gdp_growth".
                - UNIT_MEASURE: The unit of measurement for the growth rates. If supplied, expected value: "PERCENT_PER_ANNUM".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C16:G16
        Layout: row_series
        Value type: float

    Examples:
        set_growth_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': 3.5},
            {'TIME_PERIOD': 2, 'OBS_VALUE': 3.5},
        ])
    """
    _apply_series_records(
        ctx,
        records,
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

def set_interest_baseline(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
    """Set the baseline real interest rate path for projection years 1 through 5.

    Updates the real interest rate baseline values on the Inputs sheet.
    Each record corresponds to a projection year in the interest_baseline row (Inputs!C17:G17), keyed by TIME_PERIOD.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - TIME_PERIOD: Projection year (1 through 5).
                - OBS_VALUE: Real interest rate for the given projection year.
            Optional record fields:
                - INDICATOR: The indicator series. If supplied, expected value: "real_interest_rate".
                - UNIT_MEASURE: The unit of measure. If supplied, expected value: "PERCENT_PER_ANNUM".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C17:G17
        Layout: row_series
        Value type: float

    Examples:
        set_interest_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': 4.0},
            {'TIME_PERIOD': 2, 'OBS_VALUE': 4.0},
        ])
    """
    _apply_series_records(
        ctx,
        records,
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

def set_primary_balance_baseline(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
    """Set the baseline primary balance series as percent of GDP for projection years 1 through 5.

    Writes the primary balance baseline values to the Inputs sheet for the specified projection years.
    Each record maps TIME_PERIOD to the year column header and OBS_VALUE to the corresponding cell in the Inputs!C18:G18 range.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - TIME_PERIOD: Projection year number (1 to 5) to which the primary balance value applies.
                - OBS_VALUE: Primary balance as percent of GDP, with positive values indicating a surplus.
            Optional record fields:
                - INDICATOR: Economic indicator that identifies this series as the primary balance. If supplied, expected value: "primary_balance".
                - UNIT_MEASURE: Unit of measure, indicating the observation value is expressed as percent of GDP. If supplied, expected value: "PC_GDP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!C18:G18
        Layout: row_series
        Value type: float

    Examples:
        set_primary_balance_baseline(ctx, [
            {'TIME_PERIOD': 1, 'OBS_VALUE': -1.0},
            {'TIME_PERIOD': 2, 'OBS_VALUE': -0.5},
        ])
    """
    _apply_series_records(
        ctx,
        records,
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
    """Set the shock start year for the debt projection.

    Updates the first projection year in which the configured shock applies.
    The record's OBS_VALUE is written to the scalar cell Inputs!B21.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - OBS_VALUE: The first year, counted from the start of the projection horizon, when the shock takes effect.
            Optional record fields:
                - PARAMETER: Identifies the parameter as the shock year configuration. If supplied, expected value: "shock_year".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B21
        Layout: scalar
        Value type: int

    Examples:
        set_shock_year(ctx, [
            {'OBS_VALUE': 2},
        ])
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
    """Set the shock type for the debt sustainability scenario.

    Updates the workbook’s shock type to the supplied integer code.
    A single record provides the shock type value that is written to the workbook’s shock type cell.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - OBS_VALUE: Integer code specifying which parameter the shock affects: 1 for real GDP growth, 2 for real interest rate, 3 for primary balance.
            Optional record fields:
                - PARAMETER: Identifies the measure as the shock type parameter. If supplied, expected value: "shock_type".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B22
        Layout: scalar
        Value type: int

    Examples:
        set_shock_type(ctx, [
            {'OBS_VALUE': 1},
        ])
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

def set_shock_magnitudes(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
    """Set shock magnitudes for the debt sustainability analysis.

    Update the shock magnitudes used in the shock scenario configuration on the Inputs sheet.
    Each record corresponds to a cell in the row series Inputs!B26:D26, identified by the SHOCK_PARAMETER dimension.

    Args:
        records (Records): Records to apply to the workbook inputs.
            Required record fields:
                - SHOCK_PARAMETER: The parameter affected by the shock (growth, interest rate, or primary balance).
                - OBS_VALUE: The magnitude of the shock, expressed in percentage points.
            Optional record fields:
                - PARAMETER: The parameter this series represents. If supplied, expected value: "shock_magnitude".
                - UNIT_MEASURE: The unit of measure for the shock magnitude. If supplied, expected value: "PP".

    Returns:
        None: Applies the input updates to ctx.

    Source binding:
        Workbook range: Inputs!B26:D26
        Layout: row_series
        Value type: float

    Examples:
        set_shock_magnitudes(ctx, [
            {'SHOCK_PARAMETER': 'Growth', 'OBS_VALUE': -2.0},
            {'SHOCK_PARAMETER': 'Interest', 'OBS_VALUE': 2.0},
        ])
    """
    _apply_series_records(
        ctx,
        records,
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

_OUTPUT_LEAVES_OUTPUT_BASELINE = [
    ('Outputs!B12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!C12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!D12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!E12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!F12', {'SCENARIO': 'baseline', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PC_GDP'}),
]

def compute_output_baseline(inputs=None, *, ctx=None) -> Records:
    """Compute the baseline debt-to-GDP path from the Outputs sheet.

    Returns the baseline debt-to-GDP trajectory as a list of records, one per projection year.
    Each record represents one cell in the row series `Outputs!B12:F12`, with TIME_PERIOD from the column header in row 11 and OBS_VALUE from the cell value.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year, an integer from 1 to 5.
                - OBS_VALUE: Baseline debt-to-GDP ratio as a percentage of GDP.
            Optional record fields:
                - SCENARIO: Scenario identifier. If supplied, expected value: "baseline".
                - UNIT_MEASURE: Unit of measure for the observation values. If supplied, expected value: "PC_GDP".

    Source binding:
        Workbook range: Outputs!B12:F12
        Layout: row_series
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

_OUTPUT_LEAVES_OUTPUT_SHOCKED = [
    ('Outputs!B13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!C13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!D13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!E13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PC_GDP'}),
    ('Outputs!F13', {'SCENARIO': 'shocked', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PC_GDP'}),
]

def compute_output_shocked(inputs=None, *, ctx=None) -> Records:
    """Compute the shocked debt-to-GDP path records.

    Return the shocked debt-to-GDP path for projection years 1 through 5.
    Each record corresponds to one cell in the Outputs!B13:F13 range, with TIME_PERIOD derived from column headers and OBS_VALUE from cell values.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year.
                - OBS_VALUE: Debt-to-GDP ratio under the shock scenario, expressed as a percentage of GDP.
            Optional record fields:
                - SCENARIO: Scenario identifier for the shock path. If supplied, expected value: "shocked".
                - UNIT_MEASURE: Unit of measure for the debt ratio. If supplied, expected value: "PC_GDP".

    Source binding:
        Workbook range: Outputs!B13:F13
        Layout: row_series
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

_OUTPUT_LEAVES_OUTPUT_DELTA = [
    ('Outputs!B14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 1, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!C14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 2, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!D14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 3, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!E14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 4, 'UNIT_MEASURE': 'PP'}),
    ('Outputs!F14', {'SCENARIO': 'shocked_minus_baseline', 'TIME_PERIOD': 5, 'UNIT_MEASURE': 'PP'}),
]

def compute_output_delta(inputs=None, *, ctx=None) -> Records:
    """Computes the difference between the shocked and baseline debt-to-GDP paths over the projection horizon.

    Returns an array of records representing the year-by-year delta in debt-to-GDP ratio under the shocked-minus-baseline scenario.
    Each record maps to one cell in the Outputs!B14:F14 row, with TIME_PERIOD derived from the column header and OBS_VALUE from the cell value.

    Args:
        ctx (EvalContext | None): Existing evaluation context, if available.
        inputs (dict[str, object] | None): Optional input map when ctx is omitted.

    Returns:
        Records: Computed output records.
            Required record fields:
                - TIME_PERIOD: Projection year, an integer from 1 to 5.
                - OBS_VALUE: Debt-to-GDP difference, in percentage points, between the shocked and baseline paths.
            Optional record fields:
                - SCENARIO: Scenario identifier for the shocked-minus-baseline difference series. If supplied, expected value: "shocked_minus_baseline".
                - UNIT_MEASURE: Unit of measure for the observation values. If supplied, expected value: "PP".

    Source binding:
        Workbook range: Outputs!B14:F14
        Layout: row_series
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


def compute_all(inputs=None, *, ctx=None):
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
