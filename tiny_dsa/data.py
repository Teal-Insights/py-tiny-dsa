"""Authored domains, validated tensor types, and workbook defaults."""
from __future__ import annotations
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from .provenance import block_cells, column_cells, grid_cells, row_cells
from .tensor import Axis, Domain, Series, SeriesSpec, coordinate_runs, define_series
CODEGEN_SCHEMA_VERSION = 'named-axis-v1'
CODEGEN_FINGERPRINT = '623f46e74dbf6fd40f2d38a8f7eab8bbdd74801a163c58317aa2c92451aadd21'

COUNTRY_AXIS = Axis('COUNTRY', ('Borvelia', 'Litellia', 'Aurelium'), str)
TIME_PERIOD_AXIS = Axis('TIME_PERIOD', (1, 2, 3, 4, 5), int)
SHOCK_PARAMETER_AXIS = Axis('SHOCK_PARAMETER', ('Growth', 'Interest', 'Primary balance'), str)

FLOAT_VALUES = (int, float, bool, str, type(None))
INT_VALUES = (int, bool, str, type(None))
STR_VALUES = (str, type(None))
COUNTRY_PROFILE_NAMES: Series[str | None] = define_series(
    'country_profile_names',
    Domain.product(COUNTRY_AXIS),
    ('Borvelia', 'Litellia', 'Aurelium'),
    cells=column_cells('Inputs', 'A', 10, COUNTRY_AXIS),
    value_types=STR_VALUES,
)
ENGINE_YEAR_LABELS: Series[int | str | None] = define_series(
    'engine_year_labels',
    Domain.product(TIME_PERIOD_AXIS),
    (1, 2, 3, 4, 5),
    cells=row_cells('Engine', 5, 'C', TIME_PERIOD_AXIS),
    value_types=INT_VALUES,
)
COUNTRY_NAME_CELLS = {(): 'Inputs!B5'}
COUNTRY_NAME_DEFAULT = 'Borvelia'
COUNTRY_INITIAL_DEBT: Series[float | str | None] = define_series(
    'country_initial_debt',
    COUNTRY_PROFILE_NAMES.domain,
    (60.0, 80.0, 40.0),
    cells=column_cells('Inputs', 'B', 10, COUNTRY_AXIS),
    value_types=FLOAT_VALUES,
)
CountryInitialDebt = Series[float | str | None]
COUNTRY_INITIAL_DEBT_DEFAULT = COUNTRY_INITIAL_DEBT
GROWTH_BASELINE: Series[float | str | None] = define_series(
    'growth_baseline',
    ENGINE_YEAR_LABELS.domain,
    (3.5, 3.5, 3.5, 3.5, 3.5),
    cells=row_cells('Inputs', 16, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
GrowthBaseline = Series[float | str | None]
GROWTH_BASELINE_DEFAULT = GROWTH_BASELINE
INTEREST_BASELINE: Series[float | str | None] = define_series(
    'interest_baseline',
    ENGINE_YEAR_LABELS.domain,
    (4.0, 4.0, 4.0, 4.0, 4.0),
    cells=row_cells('Inputs', 17, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
InterestBaseline = Series[float | str | None]
INTEREST_BASELINE_DEFAULT = INTEREST_BASELINE
PRIMARY_BALANCE_BASELINE: Series[float | str | None] = define_series(
    'primary_balance_baseline',
    ENGINE_YEAR_LABELS.domain,
    (-1.0, -0.5, 0.0, 0.5, 1.0),
    cells=row_cells('Inputs', 18, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
PrimaryBalanceBaseline = Series[float | str | None]
PRIMARY_BALANCE_BASELINE_DEFAULT = PRIMARY_BALANCE_BASELINE
SHOCK_YEAR_CELLS = {(): 'Inputs!B21'}
SHOCK_YEAR_DEFAULT = 2
SHOCK_TYPE_CELLS = {(): 'Inputs!B22'}
SHOCK_TYPE_DEFAULT = 1
SHOCK_MAGNITUDES: Series[float | str | None] = define_series(
    'shock_magnitudes',
    Domain.product(SHOCK_PARAMETER_AXIS),
    (-2.0, 2.0, -1.0),
    cells=row_cells('Inputs', 26, 'B', SHOCK_PARAMETER_AXIS),
    value_types=FLOAT_VALUES,
)
ShockMagnitudes = Series[float | str | None]
SHOCK_MAGNITUDES_DEFAULT = SHOCK_MAGNITUDES
INITIAL_DEBT_RESOLVED_CELLS = {(): 'Inputs!B6'}
ENGINE_INITIAL_DEBT_BASELINE_CELLS = {(): 'Engine!B6'}
ENGINE_INITIAL_DEBT_SHOCKED_CELLS = {(): 'Engine!B20'}
SHOCK_MAGNITUDE_RESOLVED_CELLS = {(): 'Engine!B9'}
SHOCK_ACTIVE: SeriesSpec[int | str | None] = define_series(
    'shock_active',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 10, 'C', TIME_PERIOD_AXIS),
    value_types=INT_VALUES,
)
SHOCKED_GROWTH: SeriesSpec[float | str | None] = define_series(
    'shocked_growth',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 14, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
SHOCKED_INTEREST: SeriesSpec[float | str | None] = define_series(
    'shocked_interest',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 15, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
SHOCKED_PRIMARY_BALANCE: SeriesSpec[float | str | None] = define_series(
    'shocked_primary_balance',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 16, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
BASELINE_PATH_INTERNAL: SeriesSpec[float | str | None] = define_series(
    'baseline_path_internal',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 6, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
SHOCKED_PATH_INTERNAL: SeriesSpec[float | str | None] = define_series(
    'shocked_path_internal',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Engine', 20, 'C', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
OUTPUT_BASELINE: SeriesSpec[float | str | None] = define_series(
    'output_baseline',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Outputs', 12, 'B', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
OutputBaseline = Series[float | str | None]
OUTPUT_SHOCKED: SeriesSpec[float | str | None] = define_series(
    'output_shocked',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Outputs', 13, 'B', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
OutputShocked = Series[float | str | None]
OUTPUT_DELTA: SeriesSpec[float | str | None] = define_series(
    'output_delta',
    ENGINE_YEAR_LABELS.domain,
    cells=row_cells('Outputs', 14, 'B', TIME_PERIOD_AXIS),
    value_types=FLOAT_VALUES,
)
OutputDelta = Series[float | str | None]

_CONSTANT_NAMES = frozenset(('COUNTRY_PROFILE_NAMES', 'ENGINE_YEAR_LABELS'))
_CONSTANT_SCHEMAS = {'COUNTRY_PROFILE_NAMES': COUNTRY_PROFILE_NAMES.schema, 'ENGINE_YEAR_LABELS': ENGINE_YEAR_LABELS.schema}

_CONSTANTS_0 = frozenset({'country_profile_names'})
_CONSTANTS_1 = _CONSTANTS_0 | frozenset({'engine_year_labels'})


@contextmanager
def overrides(**values: object) -> Iterator[None]:
    """Temporarily replace constants after validating their named schemas."""
    unknown = values.keys() - _CONSTANT_NAMES
    if unknown:
        raise AttributeError(f'unknown constants: {sorted(unknown)}')
    namespace = globals()
    for name, value in values.items():
        if name in _CONSTANT_SCHEMAS:
            _CONSTANT_SCHEMAS[name].validate(value)
    previous = {name: namespace[name] for name in values}
    namespace.update(values)
    try:
        yield
    finally:
        namespace.update(previous)
