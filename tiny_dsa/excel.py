"""Excel value semantics shared by the generated model functions."""
from __future__ import annotations
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from enum import StrEnum
from importlib import import_module
from types import ModuleType
from typing import Any, TypeAlias, TypeGuard, TypeVar, cast, Literal, NoReturn, overload
import math
import numbers
import re
MIN_OPERATOR_FASTPATH_CELLS = 64
NormalizedAddress: TypeAlias = str
T = TypeVar('T', str, float)

class CoreXlError(StrEnum):
    VALUE = '#VALUE!'
    REF = '#REF!'
    DIV = '#DIV/0!'
    NA = '#N/A'
    NAME = '#NAME?'
    NUM = '#NUM!'
    NULL = '#NULL!'

    @classmethod
    def from_text(cls, value: str) -> CoreXlError | None:
        upper = value.strip().upper()
        for err in cls:
            if err.value == upper:
                return err
        return None
Scalar: TypeAlias = float | int | str | bool | CoreXlError | None

class XlErrorException(Exception):
    """Exception form of an Excel error code.

    The exported runtime raises Excel errors as exceptions; the evaluator keeps
    `XlError` sentinel values and never raises this type.
    """
    code: CoreXlError

    def __init__(self, code: CoreXlError) -> None:
        """Initialize the exception with an Excel error code."""
        if not isinstance(code, CoreXlError):
            raise TypeError(f'Expected XlError, got {type(code).__name__}')
        self.code = code
        super().__init__(code.value)
_EXCEL_EPOCH = datetime(1899, 12, 30)
_PLAIN_SCALAR_TYPES = frozenset({bool, int, float, str, type(None)})

def _apply_cmp(op: str, cmp: int) -> bool:
    if op == '=':
        return cmp == 0
    if op == '<>':
        return cmp != 0
    if op == '<':
        return cmp < 0
    if op == '>':
        return cmp > 0
    if op == '<=':
        return cmp <= 0
    if op == '>=':
        return cmp >= 0
    raise ValueError(f'Unknown comparison operator: {op}')

def _coerce_one(value: object, dtype: str) -> object:
    """Rewrite `int` to `float` when `dtype` is `float`; otherwise return `value`."""
    if dtype == 'float' and (not isinstance(value, bool)) and isinstance(value, int):
        return float(value)
    return value

def _criteria_compare(op: str, left: T, right: T) -> bool:
    """Compare two values of the same type."""
    if op == '=':
        return left == right
    if op == '<>':
        return left != right
    if op == '>':
        return left > right
    if op == '<':
        return left < right
    if op == '>=':
        return left >= right
    if op == '<=':
        return left <= right
    return False

def _enum_contains(value: object, allowed: object) -> bool:
    """Return whether `value` is an enum member without bool/int confusion.

    `1 in {True, False}` is true in Python because `bool` subclasses `int`.
    Membership requires the same runtime type as the declared member.
    """
    if not isinstance(allowed, (set, frozenset, list, tuple)):
        return False
    return any((type(value) is type(item) and value == item for item in allowed))

def _escape_sheet_for_formula(sheet: str) -> str:
    """Escape apostrophes for use inside quoted sheet names."""
    return sheet.replace("'", "''")

def _format_general_number(value: float | int) -> str:
    f = float(value)
    if f.is_integer():
        return str(int(f))
    return str(f)

def _format_measure_domain(domain: Mapping[str, Any]) -> str:
    """Render a measure domain for error messages."""
    if 'enum' in domain:
        values = domain['enum']
        rendered = ', '.join((repr(value) for value in sorted(values, key=repr)))
        return f'{{{rendered}}}'
    if 'between' in domain:
        bounds = domain['between']
        return f"between(min={bounds.get('min')!r}, max={bounds.get('max')!r})"
    if 'real_between' in domain:
        bounds = domain['real_between']
        return f"real_between(min={bounds.get('min')!r}, max={bounds.get('max')!r})"
    return repr(dict(domain))

def _in_closed_bounds(value: int | float, bounds: Mapping[str, Any]) -> bool:
    """Return whether `value` lies in an inclusive min/max interval."""
    lo = bounds.get('min')
    hi = bounds.get('max')
    return (lo is None or value >= lo) and (hi is None or value <= hi)

def _is_between_int(value: object) -> TypeGuard[int]:
    """Return whether `value` is a non-bool integer (`between` membership)."""
    return isinstance(value, int) and (not isinstance(value, bool))

def _is_mapping(value: object) -> TypeGuard[Mapping[str, object]]:
    return isinstance(value, Mapping) and (not isinstance(value, (str, bytes, bytearray)))

def _coerce_named_tensor(value: object, dtype: str) -> object | None:
    """Rewrite tensor members when `value` looks like a generated `Series`."""
    domain = getattr(value, 'domain', None)
    items = getattr(value, 'items', None)
    if domain is None or not callable(items) or _is_mapping(value):
        return None
    coerced = tuple((_coerce_one(member, dtype) for _coord, member in items()))
    replace = getattr(value, 'with_values', None)
    if callable(replace):
        return replace(coerced)
    return cast(Any, type(value))(domain, coerced)

def _is_measure_sequence(value: object) -> TypeGuard[Sequence[object]]:
    return isinstance(value, Sequence) and (not isinstance(value, (str, bytes, bytearray)))

def _is_real_number(value: object) -> TypeGuard[int | float]:
    """Return whether `value` is a non-bool int or float (`real_between`)."""
    return isinstance(value, (int, float)) and (not isinstance(value, bool))

def _ndarray_grid_shape(value: object) -> tuple[int, int] | None:
    """Read a 1-D/2-D ndarray-like shape as ``(nrows, ncols)`` without converting it.

    Lets `Grid` hold the array itself instead of eagerly copying every cell into
    nested lists. 1-D buffers read as single-column grids, matching
    `_as_nested_rows_from_ndarray`. Returns `None` for anything else (including
    3-D buffers) so callers keep the nested-list path.
    """
    ndim = getattr(value, 'ndim', None)
    if ndim not in (1, 2) or not callable(getattr(value, 'tolist', None)):
        return None
    shape = getattr(value, 'shape', None)
    if not isinstance(shape, tuple) or len(shape) != ndim:
        return None
    if not all((isinstance(extent, int) for extent in shape)):
        return None
    if ndim == 1:
        return (shape[0], 1)
    return (shape[0], shape[1])

def _parse_countif_criteria(criteria: str) -> tuple[str | None, str]:
    s = criteria.strip()
    for op in ('>=', '<=', '<>', '>', '<', '='):
        if s.startswith(op):
            return (op, s[len(op):].strip())
    return (None, s)

def _round_half_away_from_zero(number: float, digits: int) -> float:
    """Round `number` to `digits` places using Excel ROUND (ties away from 0)."""
    if digits > 308:
        return number
    if digits < -308:
        return 0.0
    factor = Decimal(10) ** digits
    shifted = Decimal(str(number)) * factor
    return float(shifted.to_integral_value(rounding=ROUND_HALF_UP) / factor)

def _try_import_numpy() -> ModuleType | None:
    try:
        return import_module('numpy')
    except ImportError:
        return None

def _value_in_measure_domain(value: object, domain: Mapping[str, Any]) -> bool:
    """Return whether `value` is inside a measure domain declaration."""
    if 'enum' in domain:
        return _enum_contains(value, domain['enum'])
    if 'between' in domain:
        if not _is_between_int(value):
            return False
        return _in_closed_bounds(value, domain['between'])
    if 'real_between' in domain:
        if not _is_real_number(value):
            return False
        return _in_closed_bounds(value, domain['real_between'])
    return True

def _reject_out_of_domain(value: object, domain: Mapping[str, Any], *, label: str) -> None:
    """Raise `ValueError` when a non-null `value` is outside `domain`."""
    if value is None:
        return
    if 'between' in domain and (not _is_between_int(value)):
        raise ValueError(f'{label} has type {type(value).__name__}; between requires int')
    if 'real_between' in domain and (not _is_real_number(value)):
        raise ValueError(f'{label} has type {type(value).__name__}; real_between requires int or float')
    if not _value_in_measure_domain(value, domain):
        raise ValueError(f'{label} out of domain: {value!r} not in {_format_measure_domain(domain)}')

def _vector_of(grid: Grid) -> Grid | None:
    if grid.nrows == 1 or grid.ncols == 1:
        return grid
    return None

def _wildcard_to_regex(pattern: str) -> re.Pattern[str]:
    out: list[str] = ['^']
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == '~' and i + 1 < len(pattern):
            i += 1
            out.append(re.escape(pattern[i]))
        elif ch == '*':
            out.append('.*')
        elif ch == '?':
            out.append('.')
        else:
            out.append(re.escape(ch))
        i += 1
    out.append('$')
    return re.compile(''.join(out), re.IGNORECASE)

def apply_arithmetic(op: str, ln: float, rn: float) -> float | CoreXlError:
    """Apply an Excel arithmetic operator to two coerced numbers."""
    if op == '+':
        return ln + rn
    if op == '-':
        return ln - rn
    if op == '*':
        return ln * rn
    if op == '/':
        if rn == 0:
            return CoreXlError.DIV
        return ln / rn
    if op == '^':
        try:
            value = ln ** rn
        except (ValueError, OverflowError):
            return CoreXlError.NUM
        if isinstance(value, complex):
            return CoreXlError.NUM
        return value
    raise ValueError(f'Unknown arithmetic operator: {op}')

def apply_input_value_map(value: object, mapping: Mapping[Any, Any], *, series_id: str) -> object:
    """Rewrite a public scalar input to its workbook needle.

    Args:
        value: Caller-facing measure (`input.value_map` key).
        mapping: Public value to workbook cell value.
        series_id: Binding series id used in the error message.

    Returns:
        The mapped workbook value.

    Raises:
        ValueError: When `value` is not a key of `mapping`.
    """
    try:
        return mapping[value]
    except (KeyError, TypeError):
        keys = ', '.join((repr(key) for key in sorted(mapping, key=repr)))
        raise ValueError(f'{series_id} value {value!r} is not in value_map; expected one of {{{keys}}}') from None

def coerce_input_measure(value: object, dtype: str, *, series_id: str) -> object:
    """Rewrite a public compute input using setter dtype rules.

    `int` becomes `float` when `dtype` is `float`. Sequences and tensors are
    rewritten memberwise. Other measure values (`str` error codes, bools,
    `None`) pass through. `float` is never narrowed to `int`.

    Args:
        value: One measure, a catalog-order sequence, or a named tensor.
        dtype: Binding measure dtype (`float`, `int`, `number`, ...).
        series_id: Binding series id; reserved for type-error messages.

    Returns:
        The value, possibly after a safe `int` -> `float` coercion.
    """
    tensor = _coerce_named_tensor(value, dtype)
    if tensor is not None:
        return tensor
    if _is_measure_sequence(value):
        members = [_coerce_one(member, dtype) for member in value]
        if isinstance(value, tuple):
            return tuple(members)
        if isinstance(value, list):
            return members
        return type(value)(members)
    return _coerce_one(value, dtype)

def datetime_to_excel_serial(value: datetime) -> float:
    """Convert a naive datetime to an Excel day serial (1900 date system)."""
    naive = value.replace(tzinfo=None) if value.tzinfo is not None else value
    delta = naive - _EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1000000) / 86400.0

def _try_parse_iso_date_serial(text: str) -> float | None:
    stripped = text.strip()
    if not stripped:
        return None
    try:
        if 'T' in stripped or ' ' in stripped:
            parsed = datetime.fromisoformat(stripped.replace('Z', '+00:00'))
            if parsed.tzinfo is not None:
                parsed = parsed.replace(tzinfo=None)
        else:
            parsed = datetime.combine(date.fromisoformat(stripped), datetime.min.time())
        return datetime_to_excel_serial(parsed)
    except ValueError:
        return None

def excel_casefold(value: str) -> str:
    return value.casefold()

def needs_quoting(sheet: str) -> bool:
    """Return True if a sheet name must be wrapped in single quotes in a formula."""
    return ' ' in sheet or '-' in sheet or "'" in sheet
np: ModuleType | None = _try_import_numpy()

def _materialize_grid(grid: Grid) -> Any:
    """Return the grid as a 2-D object ndarray, reusing an ndarray operand as-is.

    ndarray operands are already the buffer the array paths want, so they are
    only reshaped/cast when needed; `Range` and nested-list grids are walked
    once through ``Grid.at``.
    """
    assert np is not None
    array = grid.array
    if array is not None:
        arr = cast(Any, array)
        if arr.dtype != object:
            arr = arr.astype(object)
        if arr.ndim == 1:
            arr = arr.reshape(grid.nrows, grid.ncols)
        return arr
    return np.array([[grid.at(row0, col0) for col0 in range(grid.ncols)] for row0 in range(grid.nrows)], dtype=object)

def quote_sheet_if_needed(sheet: str) -> str:
    """Return a sheet name quoted for formulas when quoting is required."""
    if not needs_quoting(sheet):
        return sheet
    return "'" + _escape_sheet_for_formula(sheet) + "'"

def format_cell_key(sheet: str, column: str, row: int) -> NormalizedAddress:
    """Format a (sheet, column_letters, row) triple into a canonical address."""
    return f'{quote_sheet_if_needed(sheet)}!{column}{row}'

@dataclass(frozen=True, slots=True)
class Range:
    """Rectangular lazy range with consumer-driven cell access.

    Coordinates passed to `cell`, `row`, `column`, and `view` are 1-based and
    relative to this range, matching Excel function arguments.
    """
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    _resolver: Callable[[str], FormulaValue] = field(repr=False, compare=False)
    _coord_resolver: Callable[[int, int], FormulaValue] | None = field(default=None, repr=False, compare=False)

    def __post_init__(self) -> None:
        """Validate the rectangular bounds."""
        if self.start_row < 1 or self.start_col < 1:
            raise ValueError('Range coordinates must be positive')
        if self.end_row < self.start_row or self.end_col < self.start_col:
            raise ValueError('Range end must be greater than or equal to start')

    @property
    def shape(self) -> tuple[int, int]:
        """The range shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major addresses without evaluating cells."""
        for row in range(self.start_row, self.end_row + 1):
            for col in range(self.start_col, self.end_col + 1):
                yield self._address(row, col)

    def cell(self, row: int, col: int) -> FormulaValue:
        """Return a single relative cell value without evaluating siblings.

        Args:
            row: 1-based row within the range.
            col: 1-based column within the range.

        Raises:
            IndexError: If `row` or `col` is outside the range.
            XlErrorException: If the resolved cell is an Excel error.
        """
        self._validate_relative_cell(row, col)
        value = self._resolve_at(self.start_row + row - 1, self.start_col + col - 1)
        return self._raise_if_error(value)

    def row(self, row: int) -> Range:
        """Return a lazy view for one relative row."""
        nrows, _ = self.shape
        if row < 1 or row > nrows:
            raise IndexError('Range row is out of bounds')
        absolute_row = self.start_row + row - 1
        return Range(self.sheet, absolute_row, self.start_col, absolute_row, self.end_col, self._resolver, _coord_resolver=self._coord_resolver)

    def column(self, col: int) -> Range:
        """Return a lazy view for one relative column."""
        _, ncols = self.shape
        if col < 1 or col > ncols:
            raise IndexError('Range column is out of bounds')
        absolute_col = self.start_col + col - 1
        return Range(self.sheet, self.start_row, absolute_col, self.end_row, absolute_col, self._resolver, _coord_resolver=self._coord_resolver)

    def view(self, row_start: int=1, row_end: int | None=None, col_start: int=1, col_end: int | None=None) -> Range:
        """Return a lazy rectangular subrange view using relative coordinates."""
        nrows, ncols = self.shape
        row_end = nrows if row_end is None else row_end
        col_end = ncols if col_end is None else col_end
        self._validate_relative_cell(row_start, col_start)
        self._validate_relative_cell(row_end, col_end)
        if row_end < row_start or col_end < col_start:
            raise ValueError('Range view end must be greater than or equal to start')
        return Range(self.sheet, self.start_row + row_start - 1, self.start_col + col_start - 1, self.start_row + row_end - 1, self.start_col + col_end - 1, self._resolver, _coord_resolver=self._coord_resolver)

    def value_at(self, row: int, col: int) -> FormulaValue:
        """Return a single relative cell value with errors as sentinels.

        Unlike `cell`, Excel errors surface as `XlError` sentinel values (raised
        `XlErrorException`s from the resolver are caught and converted). Range
        consumers that implement Excel skip semantics (lookup scans, criteria
        matching) use this accessor; `cell`/iteration raise instead.
        """
        self._validate_relative_cell(row, col)
        try:
            return self._resolve_at(self.start_row + row - 1, self.start_col + col - 1)
        except XlErrorException as exc:
            return exc.code

    def iter_raw(self) -> Iterator[FormulaValue]:
        """Yield raw values (error sentinels included) in row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.value_at(row, col)

    def rows_raw(self) -> list[list[FormulaValue]]:
        """Materialize the range as nested row lists of raw values."""
        nrows, ncols = self.shape
        return [[self.value_at(r, c) for c in range(1, ncols + 1)] for r in range(1, nrows + 1)]

    def iter_values(self) -> Iterator[FormulaValue]:
        """Yield values in deterministic row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.cell(row, col)

    def __iter__(self) -> Iterator[FormulaValue]:
        """Yield values in deterministic row-major order."""
        return self.iter_values()

    def _resolve_at(self, row: int, col: int) -> FormulaValue:
        if self._coord_resolver is not None:
            return self._coord_resolver(row, col)
        return self._resolver(self._address(row, col))

    def _address(self, row: int, col: int) -> str:
        col_letter = _column_letter(col)
        return format_cell_key(self.sheet, col_letter, row)

    def _validate_relative_cell(self, row: int, col: int) -> None:
        nrows, ncols = self.shape
        if row < 1 or row > nrows or col < 1 or (col > ncols):
            raise IndexError('Range cell is out of bounds')

    @staticmethod
    def _raise_if_error(value: FormulaValue) -> FormulaValue:
        if isinstance(value, CoreXlError):
            raise XlErrorException(value)
        return value

def _resolve_scalar(value_fn: Callable[[], CellValue]) -> CellValue:
    """Evaluate a thunk, resolving 1x1 range views to their single cell value."""
    value = value_fn()
    if isinstance(value, Range) and value.shape == (1, 1):
        return cast('CellValue', value.cell(1, 1))
    return value

def format_key(sheet: str, cell: str) -> NormalizedAddress:
    """Format a sheet and A1 cell coordinate into a canonical address string."""
    return f'{quote_sheet_if_needed(sheet)}!{cell}'

@dataclass(frozen=True, slots=True)
class ExcelRange:
    """Rectangular worksheet reference geometry for evaluator and export."""
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        """The reference shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major sheet-qualified addresses without evaluating cells."""
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                col = _column_letter(c)
                yield format_key(self.sheet, f'{col}{r}')
CellValue: TypeAlias = float | int | str | bool | CoreXlError | ExcelRange | None
NestedGrid: TypeAlias = list[list[CellValue]]
FormulaValue: TypeAlias = CellValue | NestedGrid

def _as_nested_rows_from_ndarray(value: object) -> list[list[CellValue]] | None:
    """Convert an ndarray-like value to nested lists without importing NumPy.

    Duck-types via ``ndim`` / ``tolist`` so the grid module stays import-light
    for standalone exports that must remain NumPy-free.
    """
    ndim = getattr(value, 'ndim', None)
    tolist = getattr(value, 'tolist', None)
    if not isinstance(ndim, int) or not callable(tolist):
        return None
    if ndim == 0:
        return None
    raw = tolist()
    if ndim == 1:
        return [[cast(CellValue, cell)] for cell in raw]
    return cast('list[list[CellValue]]', raw)

class Grid:
    """Positional raw-value access over a lazy `Range`, ndarray, or nested-list array."""
    __slots__ = ('nrows', 'ncols', '_range', '_rows', '_array')

    def __init__(self, nrows: int, ncols: int, rng: Range | None, rows: list[list[CellValue]] | None, array: object=None) -> None:
        self.nrows = nrows
        self.ncols = ncols
        self._range = rng
        self._rows = rows
        self._array = array

    @staticmethod
    def wrap(value: object) -> Grid | None:
        """Wrap a range/array value; return `None` for scalar values.

        ndarray operands are kept as-is: consumers that want the buffer read
        `array`, and positional consumers pay for nested rows only on first use.
        """
        if isinstance(value, Range):
            nrows, ncols = value.shape
            return Grid(nrows, ncols, value, None)
        ndarray_shape = _ndarray_grid_shape(value)
        if ndarray_shape is not None:
            nrows, ncols = ndarray_shape
            if nrows == 0:
                return Grid(1, 1, None, [[None]])
            return Grid(nrows, ncols, None, None, value)
        ndarray_rows = _as_nested_rows_from_ndarray(value)
        if ndarray_rows is not None:
            if not ndarray_rows:
                ndarray_rows = [[None]]
            return Grid(len(ndarray_rows), len(ndarray_rows[0]), None, ndarray_rows)
        if isinstance(value, (list, tuple)):
            rows = [list(row) if isinstance(row, (list, tuple)) else [row] for row in cast('list[CellValue]', value)]
            if not rows:
                rows = [[None]]
            return Grid(len(rows), len(rows[0]), None, cast('list[list[CellValue]]', rows))
        return None

    @property
    def array(self) -> object:
        """The backing ndarray for ndarray operands, else `None`.

        Array consumers (vectorized operator fast paths) use this to skip the
        nested-list round trip; everything else goes through `at`.
        """
        return self._array

    def _nested_rows(self) -> list[list[CellValue]]:
        """Materialize (and cache) nested rows for positional access."""
        rows = self._rows
        if rows is None:
            rows = _as_nested_rows_from_ndarray(self._array)
            assert rows is not None
            self._rows = rows
        return rows

    def at(self, row0: int, col0: int) -> Scalar:
        """Return the raw value at a 0-based position (error sentinels included)."""
        if self._range is not None:
            return cast(Scalar, self._range.value_at(row0 + 1, col0 + 1))
        return cast(Scalar, self._nested_rows()[row0][col0])

    def at_flat(self, index0: int) -> Scalar:
        """Return the raw value at a 0-based row-major flat index."""
        row0, col0 = divmod(index0, self.ncols)
        return self.at(row0, col0)

    @property
    def size(self) -> int:
        """Total cell count."""
        return self.nrows * self.ncols

    def iter_raw(self) -> Iterator[Scalar]:
        """Yield raw values (error sentinels included) in row-major order."""
        for row0 in range(self.nrows):
            for col0 in range(self.ncols):
                yield self.at(row0, col0)

    def row_slice(self, row0: int) -> Range | list[list[CellValue]]:
        """Return one row as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.row(row0 + 1)
        return [list(self._nested_rows()[row0])]

    def col_slice(self, col0: int) -> Range | list[list[CellValue]]:
        """Return one column as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.column(col0 + 1)
        return [[row[col0]] for row in self._nested_rows()]

    def as_array(self) -> Range | list[list[CellValue]]:
        """Return the full grid as a lazy `Range` or nested-list copy."""
        if self._range is not None:
            return self._range
        return [list(row) for row in self._nested_rows()]

def _as_cell_result(value: object) -> FormulaValue:
    """Normalize operator results to ``FormulaValue`` (nested lists, not ndarrays)."""
    rows = _as_nested_rows_from_ndarray(value)
    if rows is not None:
        return rows
    return cast(FormulaValue, value)

def _as_grid(arg: CellValue) -> Grid:
    grid = Grid.wrap(arg)
    if grid is not None:
        return grid
    scalar = Grid.wrap([[arg]])
    assert scalar is not None
    return scalar

def _as_scalar(value: object) -> Scalar:
    if isinstance(value, (Range, list, tuple)):
        return CoreXlError.VALUE
    if Grid.wrap(value) is not None:
        return CoreXlError.VALUE
    return cast(Scalar, value)

def _broadcast_grids(left: object, right: object) -> tuple[Grid, Grid] | None | CoreXlError:
    """Wrap array operands as aligned grids; `None` when both operands are scalar.

    Named distinctly from ``operators._broadcast_pair`` so flattened export
    embeddings do not collide on the private helper name.
    """
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            return CoreXlError.VALUE
        return (left_grid, right_grid)
    if left_grid is not None:
        scalar_right = Grid.wrap([[right] * left_grid.ncols for _ in range(left_grid.nrows)])
        assert scalar_right is not None
        return (left_grid, scalar_right)
    assert right_grid is not None
    scalar_left = Grid.wrap([[left] * right_grid.ncols for _ in range(right_grid.nrows)])
    assert scalar_left is not None
    return (scalar_left, right_grid)

def _coerce_grid(value: object) -> Grid | CoreXlError | None:
    """Wrap array-like values; return `None` for scalars, errors as-is."""
    if isinstance(value, CoreXlError):
        return value
    return Grid.wrap(value)

def _is_ndarray_like(value: object) -> bool:
    """Duck-type NumPy ndarrays without importing NumPy.

    Fast-path materialization buffers expose ``ndim`` / ``flat`` / ``tolist``.
    Matches `Grid.wrap` / `_as_nested_rows_from_ndarray` so coercions stay
    import-light for NumPy-free installs and exports.
    """
    if _as_nested_rows_from_ndarray(value) is not None:
        return True
    ndim = getattr(value, 'ndim', None)
    return isinstance(ndim, int) and ndim >= 1 and hasattr(value, 'flat')

def _iter_arg_cells(value: CellValue) -> Iterator[CellValue]:
    """Yield cells from a range/array arg without raising on error sentinels.

    Prefer this over `flatten` for criteria consumers (`COUNTIF` / `AVERAGEIF`)
    so embedded exports still skip Excel error cells when `export_runtime.values.flatten`
    raises at the shared `flatten` name.
    """
    grid = Grid.wrap(value)
    if grid is not None:
        for cell in grid.iter_raw():
            yield cast(CellValue, cell)
        return
    yield value

def _iter_logical_cells(arg: CellValue) -> Iterator[CellValue]:
    """Yield scalar cells from a range/array argument in row-major order."""
    grid = Grid.wrap(arg)
    if grid is not None:
        for cell in grid.iter_raw():
            yield cast(CellValue, cell)
        return
    yield arg

def _raise_if_error(value: object) -> CellValue:
    if isinstance(value, CoreXlError):
        raise XlErrorException(value)
    return cast(CellValue, value)

def as_scalar(value: object) -> float | int | str | bool | CoreXlError | None:
    """Collapse range/array values to `#VALUE!` for scalar coercion contexts.

    Lazy `Range`, unbound `ExcelRange`, and nested lists are not valid scalar
    operands. Materialized ndarray buffers (fast-path internals) also collapse
    to `#VALUE!`. Does not evaluate cells inside a `Range`.

    Cells of an exact plain scalar type return immediately: `_is_ndarray_like`
    probes several attributes that miss on every ordinary cell, and `to_number`
    / `to_string` call this once per cell in the per-cell loops.
    """
    if type(value) in _PLAIN_SCALAR_TYPES:
        return cast('float | int | str | bool | None', value)
    if isinstance(value, (Range, ExcelRange, list, tuple)) or _is_ndarray_like(value):
        return CoreXlError.VALUE
    return cast('float | int | str | bool | XlError | None', value)

def flatten(*args: object) -> Iterator[FormulaValue]:
    """Flatten nested lists, lazy `Range` values, and ndarray buffers in row-major order.

    Full-scan reductions (`SUM`, `COUNTIF`, …) and generic-function error
    prechecks (`get_error`) use this helper to walk multi-cell args. Selective
    consumers (`INDEX`, `MATCH`, lookups) skip `get_error` so they are not
    forced to evaluate sibling cells. Ndarray inputs are supported only for
    fast-path materialization buffers, not as persisted `CellValue` results.
    """
    for arg in args:
        if _is_ndarray_like(arg):
            flat = getattr(arg, 'flat', None)
            if flat is not None:
                yield from (cast('FormulaValue', v) for v in flat)
            else:
                rows = _as_nested_rows_from_ndarray(arg)
                assert rows is not None
                yield from flatten(*rows)
            continue
        if isinstance(arg, Range):
            yield from arg.iter_raw()
            continue
        if isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
            continue
        yield cast('CellValue', arg)

def raise_if_sentinel_float(value: float | CoreXlError) -> float:
    """Return a float result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, CoreXlError):
        raise XlErrorException(value)
    return value

def raise_if_sentinel_int(value: int | CoreXlError) -> int:
    """Return an integer result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, CoreXlError):
        raise XlErrorException(value)
    return value

def raise_if_sentinel_str(value: str | CoreXlError) -> str:
    """Return a string result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, CoreXlError):
        raise XlErrorException(value)
    return value

def require_input_domain(value: object, domain: Mapping[str, Any], *, series_id: str) -> None:
    """Reject a scalar or sequence argument outside `input.domain`.

    Args:
        value: One measure, or a catalog-order sequence of measures.
        domain: Normalized `enum` / `between` / `real_between` declaration.
        series_id: Binding series id used in the error message.

    Raises:
        ValueError: When any non-`None` member is outside `domain`, or has a
            type the domain kind does not accept (`between` requires `int`).
    """
    if _is_measure_sequence(value):
        for index, member in enumerate(value):
            _reject_out_of_domain(member, domain, label=f'{series_id}[{index}]')
        return
    _reject_out_of_domain(value, domain, label=series_id)

def to_bool(value: FormulaValue) -> bool | CoreXlError:
    scalar = as_scalar(value)
    if isinstance(scalar, CoreXlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == '':
            return False
        if s == 'TRUE':
            return True
        if s == 'FALSE':
            return False
        return CoreXlError.VALUE
    return CoreXlError.VALUE

def _consume_logical_cells(args: tuple[CellValue, ...], *, short_circuit_true: bool) -> bool | CoreXlError:
    """Walk logical cells across arguments with Excel AND/OR semantics."""
    found_logical = False
    for arg in args:
        for cell in _iter_logical_cells(arg):
            if cell is None:
                continue
            if isinstance(cell, CoreXlError):
                return cell
            found_logical = True
            b = to_bool(cell)
            if isinstance(b, CoreXlError):
                return b
            if short_circuit_true:
                if b:
                    return True
            elif not b:
                return False
    if not found_logical:
        return CoreXlError.VALUE
    return not short_circuit_true

def logical_and(*args: CellValue) -> bool | CoreXlError:
    """Return logical AND across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=False)

def logical_if(cond: object, then_value: object, else_value: object=False) -> object:
    """Return `then_value` or `else_value` under Excel `IF` semantics.

    A scalar condition picks one branch. When any operand is a range or nested
    array, selection is element-wise: scalars broadcast, and mixed array shapes
    return `#VALUE!`. An omitted else is `FALSE`.
    """
    cond_grid = Grid.wrap(cond)
    then_grid = Grid.wrap(then_value)
    else_grid = Grid.wrap(else_value)
    if cond_grid is None and then_grid is None and (else_grid is None):
        flag = to_bool(cast(CellValue, cond))
        if isinstance(flag, CoreXlError):
            return flag
        return then_value if flag else else_value
    grids = [grid for grid in (cond_grid, then_grid, else_grid) if grid is not None]
    nrows, ncols = (grids[0].nrows, grids[0].ncols)
    if any((grid.nrows != nrows or grid.ncols != ncols for grid in grids[1:])):
        return CoreXlError.VALUE

    def _at(grid: Grid | None, scalar: object, row: int, col: int) -> object:
        if grid is None:
            return scalar
        return grid.at(row, col)
    result: list[list[CellValue]] = []
    for row in range(nrows):
        out_row: list[CellValue] = []
        for col in range(ncols):
            flag = to_bool(cast(CellValue, _at(cond_grid, cond, row, col)))
            if isinstance(flag, CoreXlError):
                out_row.append(flag)
            elif flag:
                out_row.append(cast(CellValue, _at(then_grid, then_value, row, col)))
            else:
                out_row.append(cast(CellValue, _at(else_grid, else_value, row, col)))
        result.append(out_row)
    return result

def logical_not(arg: CellValue) -> bool | CoreXlError:
    """Return logical NOT of an argument."""
    b = to_bool(arg)
    if isinstance(b, CoreXlError):
        return b
    return not b

def logical_or(*args: CellValue) -> bool | CoreXlError:
    """Return logical OR across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=True)

def to_string(value: FormulaValue) -> str:
    scalar = as_scalar(value)
    if isinstance(scalar, CoreXlError):
        return scalar.value
    value = cast(CellValue, scalar)
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, str):
        return value
    return str(value)

def _compare_rank_key(value: FormulaValue) -> tuple[int, float | str | bool]:
    """Return `(type_rank, key)` using Excel's number < text < logical order.

    A blank (`None`) compares as the number `0`. The empty string is text.
    Comparison never coerces across types.
    """
    if value is None:
        return (0, 0.0)
    if isinstance(value, bool):
        return (2, value)
    if isinstance(value, int | float):
        return (0, float(value))
    if isinstance(value, str):
        return (1, excel_casefold(value))
    return (1, excel_casefold(to_string(value)))

def compare_scalars(op: str, left: FormulaValue, right: FormulaValue) -> bool | CoreXlError:
    """Compare two scalar cell values using Excel type-rank rules."""
    if isinstance(left, CoreXlError):
        return left
    if isinstance(right, CoreXlError):
        return right
    left_rank, left_key = _compare_rank_key(left)
    right_rank, right_key = _compare_rank_key(right)
    if left_rank != right_rank:
        cmp = -1 if left_rank < right_rank else 1
    elif left_key == right_key:
        cmp = 0
    elif left_rank == 0:
        cmp = -1 if float(left_key) < float(right_key) else 1
    elif left_rank == 1:
        cmp = -1 if str(left_key) < str(right_key) else 1
    else:
        cmp = -1 if bool(left_key) < bool(right_key) else 1
    return _apply_cmp(op, cmp)

def _compare_scalars(op: str, left: FormulaValue, right: FormulaValue) -> bool | CoreXlError:
    return compare_scalars(op, left, right)

def concat_scalars(left: CellValue, right: CellValue) -> str:
    return to_string(left) + to_string(right)

def _concat_scalars(left: CellValue, right: CellValue) -> str:
    return concat_scalars(left, right)

def map_compare(op: str, left: object, right: object) -> object:
    """Element-wise comparison over scalar or broadcast array operands."""
    pair = _broadcast_grids(left, right)
    if isinstance(pair, CoreXlError):
        return pair
    if pair is None:
        return compare_scalars(op, cast(CellValue, left), cast(CellValue, right))
    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = compare_scalars(op, cast(CellValue, arr_left.at(row0, col0)), cast(CellValue, arr_right.at(row0, col0)))
            if isinstance(cell, CoreXlError):
                return cell
            out_row.append(cell)
        result.append(out_row)
    return result

def map_concat(left: object, right: object) -> object:
    """Element-wise string concatenation over scalar or broadcast array operands."""
    pair = _broadcast_grids(left, right)
    if isinstance(pair, CoreXlError):
        return pair
    if pair is None:
        if isinstance(left, CoreXlError):
            return left
        if isinstance(right, CoreXlError):
            return right
        return concat_scalars(cast(CellValue, left), cast(CellValue, right))
    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            lv = arr_left.at(row0, col0)
            rv = arr_right.at(row0, col0)
            out_row.append(concat_scalars(cast(CellValue, lv), cast(CellValue, rv)))
        result.append(out_row)
    return result

def reference_compare_array(op: str, arr_left: Any, arr_right: Any) -> Any | CoreXlError:
    """Element-wise comparison over broadcast object ndarrays (C-order, fail-fast)."""
    import numpy as np
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        cell = compare_scalars(op, arr_left[indices], arr_right[indices])
        if isinstance(cell, CoreXlError):
            return cell
        result[indices] = cell
    return result

def reference_concat_array(arr_left: Any, arr_right: Any) -> Any:
    """Element-wise string concatenation over broadcast object ndarrays."""
    import numpy as np
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        result[indices] = concat_scalars(arr_left[indices], arr_right[indices])
    return result

def try_coerce_string_to_float(text: str) -> float | None:
    """Parse one Excel numeric string; empty/whitespace text fails (`None`)."""
    stripped = text.strip()
    if stripped == '':
        return None
    try:
        return float(stripped)
    except ValueError:
        return _try_parse_iso_date_serial(stripped)

def to_number(value: FormulaValue) -> float | CoreXlError:
    scalar = as_scalar(value)
    if isinstance(scalar, CoreXlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        number = try_coerce_string_to_float(value)
        if number is None:
            return CoreXlError.VALUE
        return number
    return CoreXlError.VALUE

def _apply_arithmetic_cell(op: str, left: Scalar, right: Scalar) -> CellValue:
    if isinstance(left, CoreXlError):
        return left
    if isinstance(right, CoreXlError):
        return right
    ln = to_number(cast(CellValue, left))
    rn = to_number(cast(CellValue, right))
    if isinstance(ln, CoreXlError):
        return ln
    if isinstance(rn, CoreXlError):
        return rn
    return apply_arithmetic(op, ln, rn)

def _compare_values(a: object, b: object) -> int:
    a = _as_scalar(a)
    b = _as_scalar(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, CoreXlError) and (not isinstance(bn, CoreXlError)):
        return -1 if an < bn else 1 if an > bn else 0
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        return -1 if af < bf else 1 if af > bf else 0
    return 0

def _iter_aggregate_numbers(*args: CellValue) -> Iterator[float | CoreXlError]:
    """Yield numbers for SUM/AVERAGE/MIN/MAX/STDEV with Excel arg semantics.

    Range/array arguments keep only numeric cells: blanks, text, and booleans
    are skipped; `XlError` values propagate. Literal scalar arguments still
    coerce via `to_number` (so `SUM(1, "2", TRUE)` is 4).
    """
    for arg in args:
        grid = Grid.wrap(arg)
        if grid is not None:
            for cell in grid.iter_raw():
                if isinstance(cell, CoreXlError):
                    yield cell
                    return
                if cell is None or isinstance(cell, (bool, str)):
                    continue
                if isinstance(cell, numbers.Real):
                    yield float(cell)
            continue
        number = to_number(arg)
        if isinstance(number, CoreXlError):
            yield number
            return
        yield float(number)

def _value_matches_criteria(cell_value: CellValue, criteria: CellValue) -> bool:
    if isinstance(criteria, CoreXlError):
        return False
    if not isinstance(criteria, str):
        target = criteria
        if isinstance(cell_value, CoreXlError):
            return False
        if target is None:
            return cell_value is None
        if isinstance(target, bool):
            b = to_bool(cell_value)
            return not isinstance(b, CoreXlError) and b == target
        if isinstance(target, (int, float)) and (not isinstance(target, bool)):
            vn = to_number(cell_value)
            return not isinstance(vn, CoreXlError) and vn == float(target)
        return excel_casefold(to_string(cell_value)) == excel_casefold(to_string(target))
    op, rhs = _parse_countif_criteria(criteria)
    if isinstance(cell_value, CoreXlError):
        return False
    if op is None:
        if any((ch in rhs for ch in ('*', '?', '~'))):
            rx = _wildcard_to_regex(rhs)
            return rx.match(to_string(cell_value)) is not None
        return excel_casefold(to_string(cell_value)) == excel_casefold(rhs)
    try:
        rhs_num = float(rhs) if rhs != '' else 0.0
    except ValueError:
        rhs_num = None
    if rhs_num is not None:
        vn = to_number(cell_value)
        if isinstance(vn, CoreXlError):
            return False
        return _criteria_compare(op, vn, rhs_num)
    return _criteria_compare(op, excel_casefold(to_string(cell_value)), excel_casefold(rhs))

def _values_match(a: object, b: object) -> bool:
    a = _as_scalar(a)
    b = _as_scalar(b)
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, CoreXlError) and (not isinstance(bn, CoreXlError)):
        return an == bn
    return a == b

def abs_number(*args: CellValue) -> float | CoreXlError:
    """Return the absolute value of a number (Excel ``ABS``)."""
    if len(args) != 1:
        return CoreXlError.VALUE
    n = to_number(args[0])
    if isinstance(n, CoreXlError):
        return n
    return float(abs(n))

def average_cells(*args: CellValue) -> float | CoreXlError:
    """Return the average of numeric cells."""
    total = 0.0
    count = 0
    for value in _iter_aggregate_numbers(*args):
        if isinstance(value, CoreXlError):
            return value
        total += value
        count += 1
    if count == 0:
        return CoreXlError.DIV
    return float(total / count)

def countif_cells(range_values: CellValue, criteria: CellValue) -> int | CoreXlError:
    """Count cells matching criteria.

    Error cells in the range are skipped (Excel COUNTIF semantics), not
    propagated. Criteria that are themselves an `XlError` propagate.
    """
    if isinstance(criteria, CoreXlError):
        return criteria
    return sum((1 for v in _iter_arg_cells(range_values) if _value_matches_criteria(v, criteria)))

def exp_number(*args: CellValue) -> float | CoreXlError:
    """Return e raised to the power of a number (Excel ``EXP``)."""
    if len(args) != 1:
        return CoreXlError.VALUE
    n = to_number(args[0])
    if isinstance(n, CoreXlError):
        return n
    try:
        return float(math.exp(n))
    except OverflowError:
        return CoreXlError.NUM

def hlookup_cells(lookup_value: object, table_array: object, row_index_num: object, range_lookup: object=True) -> Scalar:
    """Excel HLOOKUP over a lazy grid or nested-list array."""
    rn = to_number(cast(CellValue, row_index_num))
    if isinstance(rn, CoreXlError):
        return rn
    row_index = int(rn)
    if row_index < 1:
        return CoreXlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return CoreXlError.VALUE
    if row_index > grid.nrows:
        return CoreXlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.ncols):
            if _values_match(lookup_value, grid.at(0, i)):
                return grid.at(row_index - 1, i)
        return CoreXlError.NA
    last_match_idx = None
    for i in range(grid.ncols):
        if _compare_values(grid.at(0, i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return CoreXlError.NA
    return grid.at(row_index - 1, last_match_idx)

def index_cells(array: object, row_num: object=None, col_num: object=None) -> object:
    """Excel INDEX over a lazy grid or nested-list array.

    Returns a scalar cell value, or a row/column slice (`Range` or nested list)
    when only one of `row_num` / `col_num` selects a vector.

    A `row_num` or `col_num` of `0` selects the entire column or row (Excel
    whole-vector form). Both `0` returns the full array.
    """
    grid = Grid.wrap(array)
    if grid is None:
        return CoreXlError.VALUE
    nrows, ncols = (grid.nrows, grid.ncols)
    row_omitted = row_num is None
    col_omitted = col_num is None
    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return grid.at(0, 0)
        if nrows == 1:
            return grid.at(0, ncols - 1)
        if ncols == 1:
            return grid.at(nrows - 1, 0)
        return CoreXlError.VALUE
    if row_omitted:
        col_s = as_scalar(col_num)
        if isinstance(col_s, CoreXlError):
            return col_s
        cn = to_number(cast(CellValue, col_s))
        if isinstance(cn, CoreXlError):
            return cn
        col = int(cn)
        if col == 0:
            return grid.as_array()
        if col < 1 or col > ncols:
            return CoreXlError.REF
        if nrows == 1:
            return grid.at(0, col - 1)
        return grid.col_slice(col - 1)
    row_s = as_scalar(row_num)
    if isinstance(row_s, CoreXlError):
        return row_s
    rn = to_number(cast(CellValue, row_s))
    if isinstance(rn, CoreXlError):
        return rn
    row = int(rn)
    if col_omitted:
        if row == 0:
            return grid.as_array()
        if nrows == 1:
            if row < 1 or row > ncols:
                return CoreXlError.REF
            return grid.at(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return CoreXlError.REF
            return grid.at(row - 1, 0)
        if row < 1 or row > nrows:
            return CoreXlError.REF
        return grid.row_slice(row - 1)
    col_s = as_scalar(col_num)
    if isinstance(col_s, CoreXlError):
        return col_s
    cn = to_number(cast(CellValue, col_s))
    if isinstance(cn, CoreXlError):
        return cn
    col = int(cn)
    if row == 0 and col == 0:
        return grid.as_array()
    if row == 0:
        if col < 1 or col > ncols:
            return CoreXlError.REF
        if nrows == 1:
            return grid.at(0, col - 1)
        return grid.col_slice(col - 1)
    if col == 0:
        if row < 1 or row > nrows:
            return CoreXlError.REF
        if ncols == 1:
            return grid.at(row - 1, 0)
        return grid.row_slice(row - 1)
    if nrows == 1:
        if row < 1 or row > ncols:
            return CoreXlError.REF
        return grid.at(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return CoreXlError.REF
        return grid.at(row - 1, 0)
    if row < 1 or row > nrows:
        return CoreXlError.REF
    if col < 1 or col > ncols:
        return CoreXlError.REF
    return grid.at(row - 1, col - 1)

def large_kth(array: CellValue, k: CellValue) -> float | CoreXlError:
    """Return the k-th largest numeric value."""
    kk = to_number(k)
    if isinstance(kk, CoreXlError):
        return kk
    kth = int(kk)
    if kth < 1:
        return CoreXlError.NUM
    nums: list[float] = []
    for v in flatten(array):
        if isinstance(v, CoreXlError):
            return v
        if v is None or isinstance(v, bool):
            continue
        if isinstance(v, numbers.Real) and (not isinstance(v, bool)):
            nums.append(float(v))
    if kth > len(nums):
        return CoreXlError.NUM
    nums.sort(reverse=True)
    return float(nums[kth - 1])

def left_chars(text: CellValue, num_chars: CellValue=1) -> str | CoreXlError:
    """Return the leftmost characters of text."""
    scalar = as_scalar(text)
    if isinstance(scalar, CoreXlError):
        return scalar
    s = to_string(cast(CellValue, scalar))
    n = to_number(num_chars)
    if isinstance(n, CoreXlError):
        return n
    chars = int(n)
    if chars < 0:
        return CoreXlError.VALUE
    return s[:chars]

def lookup_cells(lookup_value: object, lookup_vector_or_array: object, result_vector: object=None) -> Scalar:
    """Excel LOOKUP over a lazy grid or nested-list array.

    Error values in the lookup vector are skipped (Excel), so idioms such as
    `LOOKUP(2, 1/(rng<>0), rng)` return the last non-error match.
    """
    grid = _coerce_grid(lookup_vector_or_array)
    if grid is None:
        return CoreXlError.VALUE
    if isinstance(grid, CoreXlError):
        return grid
    result_grid = _coerce_grid(result_vector) if result_vector is not None else None
    if result_vector is not None and result_grid is None:
        return CoreXlError.VALUE
    if isinstance(result_grid, CoreXlError):
        return result_grid
    if result_grid is None:
        vector = _vector_of(grid)
        if vector is not None:
            lookup_flat = vector
            result_flat = vector
        elif grid.nrows >= grid.ncols:
            lookup_flat = Grid.wrap(grid.col_slice(0))
            result_flat = Grid.wrap(grid.col_slice(grid.ncols - 1))
            assert lookup_flat is not None and result_flat is not None
        else:
            lookup_flat = Grid.wrap(grid.row_slice(0))
            result_flat = Grid.wrap(grid.row_slice(grid.nrows - 1))
            assert lookup_flat is not None and result_flat is not None
    else:
        if _vector_of(grid) is None or _vector_of(result_grid) is None:
            return CoreXlError.NA
        if grid.size != result_grid.size:
            return CoreXlError.NA
        lookup_flat = grid
        result_flat = result_grid
    last_match_idx = None
    for i in range(lookup_flat.size):
        cell = lookup_flat.at_flat(i)
        if isinstance(cell, CoreXlError):
            continue
        if _compare_values(cell, lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return CoreXlError.NA
    return result_flat.at_flat(last_match_idx)

def map_arithmetic(op: str, left: object, right: object) -> object:
    """Element-wise arithmetic over scalar or broadcast array operands.

    Array results embed per-element ``XlError`` values rather than collapsing
    the whole operation to a scalar error.
    """
    pair = _broadcast_grids(left, right)
    if isinstance(pair, CoreXlError):
        return pair
    if pair is None:
        return _apply_arithmetic_cell(op, cast(Scalar, left), cast(Scalar, right))
    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            out_row.append(_apply_arithmetic_cell(op, arr_left.at(row0, col0), arr_right.at(row0, col0)))
        result.append(out_row)
    return result

def map_unary(op: str, value: object) -> object:
    """Apply a unary Excel operator over scalar or array operands."""
    grid = Grid.wrap(value)
    if grid is None:
        if isinstance(value, CoreXlError):
            return value
        number = to_number(cast(CellValue, value))
        if isinstance(number, CoreXlError):
            return number
        if op == '-':
            return -number
        if op == '+':
            return +number
        if op == '%':
            return number / 100.0
        raise ValueError(f'Unknown unary operator: {op}')
    result: list[list[CellValue]] = []
    for row0 in range(grid.nrows):
        out_row: list[CellValue] = []
        for col0 in range(grid.ncols):
            cell = grid.at(row0, col0)
            if isinstance(cell, CoreXlError):
                return cell
            number = to_number(cast(CellValue, cell))
            if isinstance(number, CoreXlError):
                return number
            if op == '-':
                out_row.append(-number)
            elif op == '+':
                out_row.append(+number)
            elif op == '%':
                out_row.append(number / 100.0)
            else:
                raise ValueError(f'Unknown unary operator: {op}')
        result.append(out_row)
    return result

def match_cells(lookup_value: object, lookup_array: object, match_type: object=1) -> int | CoreXlError:
    """Excel MATCH over a lazy grid or nested-list array."""
    mt = to_number(cast(CellValue, match_type))
    if isinstance(mt, CoreXlError):
        return mt
    match_type_int = int(mt)
    if isinstance(lookup_array, CoreXlError):
        return lookup_array
    grid = Grid.wrap(lookup_array)
    if grid is None:
        grid_wrapped = Grid.wrap([[lookup_array]])
        assert grid_wrapped is not None
        grid = grid_wrapped
    if match_type_int == 0:
        for i in range(grid.size):
            if _values_match(lookup_value, grid.at_flat(i)):
                return i + 1
        return CoreXlError.NA
    if match_type_int == 1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        return CoreXlError.NA if last_match is None else last_match
    if match_type_int == -1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        return CoreXlError.NA if last_match is None else last_match
    return CoreXlError.VALUE

def max_cells(*args: CellValue) -> float | CoreXlError:
    """Return the maximum of numeric cells."""
    found = False
    current = 0.0
    for value in _iter_aggregate_numbers(*args):
        if isinstance(value, CoreXlError):
            return value
        if not found or value > current:
            current = value
            found = True
    if not found:
        return 0.0
    return current

def min_cells(*args: CellValue) -> float | CoreXlError:
    """Return the minimum of numeric cells."""
    found = False
    current = 0.0
    for value in _iter_aggregate_numbers(*args):
        if isinstance(value, CoreXlError):
            return value
        if not found or value < current:
            current = value
            found = True
    if not found:
        return 0.0
    return current

def npv_cells(rate: CellValue, *values: CellValue) -> float | CoreXlError:
    """Return net present value for a rate and cash flows."""
    r = to_number(rate)
    if isinstance(r, CoreXlError):
        return r
    result = 0.0
    count = 0
    for v in flatten(*values):
        n = to_number(v)
        if isinstance(n, CoreXlError):
            return n
        count += 1
        result += float(n) / (1 + r) ** count
    if count == 0:
        return CoreXlError.VALUE
    return result

def numbervalue_parse(text: CellValue, decimal_separator: CellValue='.', group_separator: CellValue=',') -> float | CoreXlError:
    """Convert text to a number with explicit decimal and group separators."""
    if isinstance(text, CoreXlError):
        return text
    if isinstance(decimal_separator, CoreXlError):
        return decimal_separator
    if isinstance(group_separator, CoreXlError):
        return group_separator
    if not isinstance(text, str):
        return to_number(text)
    dec_sep = to_string(decimal_separator)
    grp_sep = to_string(group_separator)
    if dec_sep == '' or dec_sep == grp_sep:
        return CoreXlError.VALUE
    s = text.replace('\xa0', ' ').strip()
    if s == '':
        return 0.0
    currency_symbols = '$€£¥'
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == '':
            return CoreXlError.VALUE
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1].strip()
        if s == '':
            return CoreXlError.VALUE
    percent = False
    if s.endswith('%'):
        percent = True
        s = s[:-1].strip()
        if s == '':
            return CoreXlError.VALUE
    sign = 1.0
    if s.startswith(('+', '-')):
        if s[0] == '-':
            sign = -1.0
        s = s[1:].strip()
        if s == '':
            return CoreXlError.VALUE
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == '':
            return CoreXlError.VALUE
    if grp_sep:
        s = s.replace(grp_sep, '')
    if dec_sep != '.':
        s = s.replace(dec_sep, '.')
    try:
        value = float(s)
    except ValueError:
        return CoreXlError.VALUE
    if percent:
        value /= 100.0
    if negative:
        value = -abs(value)
    return value * sign

def rank_number(number: CellValue, ref: CellValue, order: CellValue=0) -> int | CoreXlError:
    """Return the rank of a number within a reference range."""
    nn = to_number(number)
    if isinstance(nn, CoreXlError):
        return nn
    oo = to_number(order)
    if isinstance(oo, CoreXlError):
        return oo
    ascending = int(oo) != 0
    nums: list[float] = []
    for v in flatten(ref):
        if isinstance(v, CoreXlError):
            return v
        if v is None or isinstance(v, bool):
            continue
        if isinstance(v, numbers.Real) and (not isinstance(v, bool)):
            nums.append(float(v))
    if ascending:
        return 1 + sum((1 for v in nums if v < nn))
    return 1 + sum((1 for v in nums if v > nn))

def reference_arithmetic_array(op: str, arr_left: Any, arr_right: Any) -> Any | CoreXlError:
    """Element-wise arithmetic over broadcast object ndarrays.

    Per-element errors (operand sentinels, coercion failures, ``#DIV/0!``,
    ``#NUM!``) are embedded in the result array rather than collapsing the
    whole operation to a scalar error (Excel array arithmetic).
    """
    import numpy as np
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        left_cell = arr_left[indices]
        right_cell = arr_right[indices]
        if isinstance(left_cell, CoreXlError):
            result[indices] = left_cell
            continue
        if isinstance(right_cell, CoreXlError):
            result[indices] = right_cell
            continue
        ln = to_number(left_cell)
        rn = to_number(right_cell)
        if isinstance(ln, CoreXlError):
            result[indices] = ln
            continue
        if isinstance(rn, CoreXlError):
            result[indices] = rn
            continue
        result[indices] = apply_arithmetic(op, ln, rn)
    return result

def round_number(number: CellValue, num_digits: CellValue) -> float | CoreXlError:
    """Round a number to the given number of digits (Excel `ROUND`).

    Halfway cases round away from zero (`ROUND(2.5, 0)` is `3`), matching Excel
    rather than Python's ties-to-even `round`.
    """
    n = to_number(number)
    if isinstance(n, CoreXlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, CoreXlError):
        return d
    return _round_half_away_from_zero(n, int(d))

def rounddown_number(number: CellValue, num_digits: CellValue) -> float | CoreXlError:
    """Round a number down to the given number of digits."""
    n = to_number(number)
    if isinstance(n, CoreXlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, CoreXlError):
        return d
    digits = int(d)
    factor = 10 ** digits
    if n >= 0:
        return float(math.floor(n * factor) / factor)
    return float(math.ceil(n * factor) / factor)

def stdev_cells(*args: CellValue) -> float | CoreXlError:
    """Return sample standard deviation of numeric cells."""
    nums: list[float] = []
    for value in _iter_aggregate_numbers(*args):
        if isinstance(value, CoreXlError):
            return value
        nums.append(value)
    if len(nums) < 2:
        return CoreXlError.DIV
    mean = sum(nums) / len(nums)
    variance = sum(((x - mean) ** 2 for x in nums)) / (len(nums) - 1)
    return float(variance ** 0.5)

def sum_cells(*args: CellValue) -> float | CoreXlError:
    """Return the sum of numeric cells."""
    total = 0.0
    for value in _iter_aggregate_numbers(*args):
        if isinstance(value, CoreXlError):
            return value
        total += value
    return total

def sumproduct_cells(*args: CellValue) -> float | CoreXlError:
    """Return the sum of element-wise products across aligned array arguments."""
    if len(args) == 0:
        return 0.0
    grids = [_as_grid(arg) for arg in args]
    shape = (grids[0].nrows, grids[0].ncols)
    for grid in grids[1:]:
        if (grid.nrows, grid.ncols) != shape:
            return CoreXlError.VALUE
    result = 0.0
    for index0 in range(grids[0].size):
        product = 1.0
        for grid in grids:
            cell = cast(CellValue, grid.at_flat(index0))
            if isinstance(cell, CoreXlError):
                return cell
            if isinstance(cell, str):
                number = 0.0
            else:
                number = to_number(cell)
                if isinstance(number, CoreXlError):
                    return number
            product *= number
        result += product
    return result

def text_format(value: CellValue, format_text: CellValue) -> str | CoreXlError:
    """Format a value as text using a format string."""
    scalar = as_scalar(value)
    if isinstance(scalar, CoreXlError):
        return scalar
    if isinstance(format_text, CoreXlError):
        return format_text
    fmt = to_string(format_text)
    n = to_number(cast(CellValue, scalar))
    if isinstance(n, CoreXlError):
        return to_string(cast(CellValue, scalar))
    if fmt == '0':
        return str(int(round(n)))
    if fmt == '0.0':
        return f'{n:.1f}'
    if fmt == '0.00':
        return f'{n:.2f}'
    if fmt == '0.000':
        return f'{n:.3f}'
    if fmt == '#,##0':
        return f'{int(round(n)):,}'
    if fmt == '#,##0.00':
        return f'{n:,.2f}'
    if fmt == '0%':
        return f'{int(round(n * 100))}%'
    if fmt == '0.0%':
        return f'{n * 100:.1f}%'
    if fmt == '0.00%':
        return f'{n * 100:.2f}%'
    if n == int(n):
        return str(int(n))
    return str(n)

def try_fastpath_arithmetic_array(op: str, arr_left: Any, arr_right: Any) -> Any | CoreXlError | None:
    return None

def try_fastpath_compare_array(op: str, arr_left: Any, arr_right: Any) -> Any | CoreXlError | None:
    return None

def try_fastpath_concat_array(arr_left: Any, arr_right: Any) -> Any | None:
    return None

def _apply_large_grid_binary(op: str | None, left: object, right: object, left_grid: Grid | None, right_grid: Grid | None, *, kind: str) -> FormulaValue | None:
    """Materialize large grids once, then run fastpath or reference loops.

    Returns `None` when operands are below ``MIN_OPERATOR_FASTPATH_CELLS`` (or
    cannot form a grid pair) so the caller can use shared ``map_*`` loops.
    When this path materializes, a fastpath miss reuses the same arrays via
    ``reference_*_array`` — it does not fall through to a second Range walk.

    Without NumPy (`fast` extra), always returns `None` so callers use
    ``operator_maps``.
    """
    if np is None:
        return None
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            return None
        if left_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_left = _materialize_grid(left_grid)
        arr_right = _materialize_grid(right_grid)
    elif left_grid is not None:
        if left_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_left = _materialize_grid(left_grid)
        arr_right = np.full(arr_left.shape, right, dtype=object)
    else:
        assert right_grid is not None
        if right_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_right = _materialize_grid(right_grid)
        arr_left = np.full(arr_right.shape, left, dtype=object)
    if kind == 'compare':
        assert op is not None
        fast = try_fastpath_compare_array(op, arr_left, arr_right)
        if fast is not None:
            return _as_cell_result(fast)
        return _as_cell_result(reference_compare_array(op, arr_left, arr_right))
    if kind == 'concat':
        fast = try_fastpath_concat_array(arr_left, arr_right)
        if fast is not None:
            return _as_cell_result(fast)
        return _as_cell_result(reference_concat_array(arr_left, arr_right))
    assert op is not None
    fast = try_fastpath_arithmetic_array(op, arr_left, arr_right)
    if fast is not None:
        return _as_cell_result(fast)
    return _as_cell_result(reference_arithmetic_array(op, arr_left, arr_right))

def _xl_arithmetic(op: str, left: FormulaValue, right: FormulaValue) -> FormulaValue:
    if isinstance(left, CoreXlError):
        return left
    if isinstance(right, CoreXlError):
        return right
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is not None or right_grid is not None:
        large = _apply_large_grid_binary(op, left, right, left_grid, right_grid, kind='arithmetic')
        if large is not None:
            return large
        return cast(FormulaValue, map_arithmetic(op, left, right))
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, CoreXlError):
        return ln
    if isinstance(rn, CoreXlError):
        return rn
    return apply_arithmetic(op, ln, rn)

def _xl_compare(op: str, left: FormulaValue, right: FormulaValue) -> FormulaValue:
    if isinstance(left, CoreXlError):
        return left
    if isinstance(right, CoreXlError):
        return right
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is not None or right_grid is not None:
        large = _apply_large_grid_binary(op, left, right, left_grid, right_grid, kind='compare')
        if large is not None:
            return large
        return cast(FormulaValue, map_compare(op, left, right))
    return _compare_scalars(op, left, right)

def _xl_concat(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    if isinstance(left, CoreXlError):
        return left
    if isinstance(right, CoreXlError):
        return right
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is not None or right_grid is not None:
        large = _apply_large_grid_binary(None, left, right, left_grid, right_grid, kind='concat')
        if large is not None:
            return large
        return cast(FormulaValue, map_concat(left, right))
    return _concat_scalars(cast(CellValue, left), cast(CellValue, right))

def value_from_text(text: CellValue) -> float | CoreXlError:
    """Convert locale-formatted text to a number (Excel VALUE)."""
    parsed = numbervalue_parse(text)
    if parsed is CoreXlError.VALUE and isinstance(text, str):
        return to_number(text)
    return parsed

def vlookup_cells(lookup_value: object, table_array: object, col_index_num: object, range_lookup: object=True) -> Scalar:
    """Excel VLOOKUP over a lazy grid or nested-list array."""
    cn = to_number(cast(CellValue, col_index_num))
    if isinstance(cn, CoreXlError):
        return cn
    col_index = int(cn)
    if col_index < 1:
        return CoreXlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return CoreXlError.VALUE
    if col_index > grid.ncols:
        return CoreXlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.nrows):
            if _values_match(lookup_value, grid.at(i, 0)):
                return grid.at(i, col_index - 1)
        return CoreXlError.NA
    last_match_idx = None
    for i in range(grid.nrows):
        if _compare_values(grid.at(i, 0), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return CoreXlError.NA
    return grid.at(last_match_idx, col_index - 1)

def _core_add(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic('+', left, right)

def _core_concat(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_concat(left, right)

def _shared_countif(range_values: CellValue, criteria: CellValue) -> int:
    """Count cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_int(countif_cells(range_values, criteria))

def _core_div(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic('/', left, right)

def _core_eq(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('=', left, right)

def _core_ge(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('>=', left, right)

def _core_gt(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('>', left, right)

def _shared_hlookup(lookup_value: CellValue, table_array: CellValue, row_index_num: CellValue, range_lookup: CellValue=True) -> CellValue:
    return _raise_if_error(hlookup_cells(lookup_value, table_array, row_index_num, range_lookup))

def _shared_iferror(value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]) -> CellValue:
    """Excel IFERROR over lazily-evaluated value and fallback thunks."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return fallback_fn()
    if isinstance(value, CoreXlError):
        return fallback_fn()
    return value

def _shared_ifna(value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]) -> CellValue:
    """Excel IFNA: catch `#N/A` only; other Excel errors propagate."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        if exc.code == CoreXlError.NA:
            return fallback_fn()
        raise
    if value == CoreXlError.NA:
        return fallback_fn()
    return value

def _shared_isblank(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISBLANK: IS functions do not propagate errors."""
    try:
        value = value_fn()
    except XlErrorException:
        return False
    if isinstance(value, Range):
        if value.shape != (1, 1):
            return False
        value = value.value_at(1, 1)
    return value is None

def _shared_iserror(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISERROR: True when evaluating the argument produces any Excel error."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return True
    return isinstance(value, CoreXlError)

def _shared_isna(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISNA: True when evaluating the argument produces `#N/A`."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        return exc.code == CoreXlError.NA
    return value == CoreXlError.NA

def _shared_isnumber(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISNUMBER: False for errors; True only for non-bool numbers."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return False
    if isinstance(value, CoreXlError):
        return False
    return not isinstance(value, bool) and isinstance(value, (int, float))

def _shared_istext(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISTEXT: False for errors; True only for non-error strings."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return False
    return isinstance(value, str) and (not isinstance(value, CoreXlError))

def _shared_large(array: CellValue, k: CellValue) -> float:
    """Return the k-th largest value, raising on Excel errors."""
    return raise_if_sentinel_float(large_kth(array, k))

def _core_le(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('<=', left, right)

def _shared_left(text: CellValue, num_chars: CellValue=1) -> str:
    """Return the leftmost characters of text, raising on Excel errors."""
    return raise_if_sentinel_str(left_chars(text, num_chars))

def _shared_lookup(lookup_value: CellValue, lookup_vector_or_array: CellValue, result_vector: CellValue=None) -> CellValue:
    return _raise_if_error(lookup_cells(lookup_value, lookup_vector_or_array, result_vector))

def _core_lt(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('<', left, right)

def _core_mul(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic('*', left, right)

def _core_ne(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare('<>', left, right)

def _core_neg(value: FormulaValue) -> FormulaValue:
    return cast(FormulaValue, map_unary('-', value))

def _shared_npv(rate: CellValue, *values: CellValue) -> float:
    """Return net present value, raising on Excel errors."""
    return raise_if_sentinel_float(npv_cells(rate, *values))

def _shared_numbervalue(text: CellValue, decimal_separator: CellValue='.', group_separator: CellValue=',') -> float:
    """Convert text to a number, raising on Excel errors."""
    return raise_if_sentinel_float(numbervalue_parse(text, decimal_separator, group_separator))

def _core_pos(value: FormulaValue) -> FormulaValue:
    return cast(FormulaValue, map_unary('+', value))

def _core_pow(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic('^', left, right)

def _shared_rank(number: CellValue, ref: CellValue, order: CellValue=0) -> int:
    """Return the rank of a number in a list, raising on Excel errors."""
    return raise_if_sentinel_int(rank_number(number, ref, order))

def _shared_round(number: CellValue, num_digits: CellValue) -> float:
    """Round a number, raising on Excel coercion errors."""
    return raise_if_sentinel_float(round_number(number, num_digits))

def _shared_rounddown(number: CellValue, num_digits: CellValue) -> float:
    """Round a number down, raising on Excel coercion errors."""
    return raise_if_sentinel_float(rounddown_number(number, num_digits))

def _shared_stdev(*args: CellValue) -> float:
    """Return sample standard deviation, raising on Excel errors."""
    return raise_if_sentinel_float(stdev_cells(*args))

def _core_sub(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic('-', left, right)

def _shared_text(value: CellValue, format_text: CellValue) -> str:
    """Format a value as text, raising on Excel errors."""
    return raise_if_sentinel_str(text_format(value, format_text))

def xlookup_cells(lookup_value: object, lookup_array: object, return_array: object, if_not_found: object=None, match_mode: object=0, search_mode: object=1) -> object:
    """Excel XLOOKUP (exact match; search forward or backward)."""
    mm = to_number(cast(CellValue, match_mode))
    if isinstance(mm, CoreXlError):
        return mm
    sm = to_number(cast(CellValue, search_mode))
    if isinstance(sm, CoreXlError):
        return sm
    mm_i = int(mm)
    sm_i = int(sm)
    if mm_i != 0:
        return CoreXlError.VALUE
    if sm_i not in (1, -1):
        return CoreXlError.VALUE
    keys = Grid.wrap(lookup_array)
    vals = Grid.wrap(return_array)
    if keys is None or vals is None:
        return CoreXlError.VALUE
    if keys.size != vals.size:
        return CoreXlError.VALUE
    idxs = range(keys.size) if sm_i == 1 else range(keys.size - 1, -1, -1)
    for i in idxs:
        if _values_match(lookup_value, keys.at_flat(i)):
            return vals.at_flat(i)
    return CoreXlError.NA if if_not_found is None else if_not_found

def _shared_xlookup(lookup_value: CellValue, lookup_array: CellValue, return_array: CellValue, if_not_found: CellValue=None, match_mode: CellValue=0, search_mode: CellValue=1) -> CellValue:
    return _raise_if_error(xlookup_cells(lookup_value, lookup_array, return_array, if_not_found, match_mode, search_mode))

def _column_letter(index: int) -> str:
    """Return worksheet column letters for a 1-based column index."""
    letters = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord('A') + remainder) + letters
    return letters

SharedXlError = XlErrorException

XL_ERROR_CODES = frozenset(
    {
        "#VALUE!",
        "#REF!",
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NUM!",
        "#NULL!",
    }
)


class XlError(Exception):
    """Excel error value raised as a Python exception."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


def is_error(value: object) -> TypeGuard[str]:
    """True when `value` is an Excel error code string."""
    return isinstance(value, str) and value in XL_ERROR_CODES


@overload
def as_measure(value: object, dtype: Literal["float"] = "float") -> float | str: ...


@overload
def as_measure(value: object, dtype: Literal["int"]) -> int | str: ...


@overload
def as_measure(value: object, dtype: Literal["str"]) -> str | int | float | bool: ...


@overload
def as_measure(value: object, dtype: Literal["bool"]) -> bool | str: ...


@overload
def as_measure(value: object, dtype: Literal["datetime"]) -> datetime | str: ...


def as_measure(value: object, dtype: str = "float") -> int | float | str | bool | datetime | None:
    """Coerce a helper result to a measure: number or cached text.

    Operators still raise `XlError`. Series-member boundaries catch that and
    store `err.code` here so a `#REF!` cell does not abort the rest of a series.
    Non-numeric cached strings (`n/a`, `..`) pass through as measures.
    Blank cells (`None`) stay `None`.

    Overloads narrow the return by `dtype`: the default `float` path is
    `float | str` so generated `list[float | str]` accumulators type-check.
    A `str` measure keeps Excel numbers and bools so `INDEX(...)=1` on a
    string-dtyped computed 0/1 series stays numeric, matching the evaluator.
    """
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, XlError):
        return value.code
    if dtype == "int":
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value)
        raise TypeError(f"cannot coerce {type(value).__name__} to int measure")
    if dtype == "str":
        if isinstance(value, bool | int | float):
            return value
        return str(value)
    if dtype == "bool":
        return bool(value)
    if dtype == "datetime":
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        raise TypeError(f"cannot coerce {type(value).__name__} to datetime measure")
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, int | float):
        return float(value)
    raise TypeError(f"cannot coerce {type(value).__name__} to float measure")


def _raise_stored_error(value: object) -> None:
    """Re-raise a cached Excel error-code measure."""
    if isinstance(value, str) and is_error(value):
        raise XlError(value)


def _raise_stored_errors_in(value: object) -> None:
    """Re-raise stored error-code measures in scalars, sequences, and views."""
    if isinstance(value, Range):
        for item in value.iter_values():
            _raise_stored_error(item)
        return
    if isinstance(value, str) or not isinstance(value, Sequence):
        _raise_stored_error(value)
        return
    for item in value:
        _raise_stored_errors_in(item)


def _as_core_cells(value: object) -> CellValue:
    """Convert stored error-code measures to core sentinels for shared helpers."""
    if isinstance(value, str):
        converted = CoreXlError.from_text(value)
        return converted if converted is not None else value
    if isinstance(value, Sequence):
        return cast(CellValue, [_as_core_cells(item) for item in value])
    return cast(CellValue, value)


def _adapt_core(value: object) -> object:
    """Raise `XlError` when `core` returned a sentinel."""
    if isinstance(value, CoreXlError):
        raise XlError(value.value)
    return value


def _as_formula(value: object) -> FormulaValue:
    """Narrow a generated-code operand to a `core` formula value."""
    return cast(FormulaValue, value)


def _arith_operand(value: object) -> FormulaValue:
    """Prepare an arithmetic operand for `core`.

    Blank cells (`None`) stay `None` so `to_number` coerces them to `0`. Empty
    text (`""`) is left as text so arithmetic raises `#VALUE!` (Excel / #420).
    """
    _raise_stored_error(value)
    return _as_formula(value)


def _as_number(value: object) -> float:
    """Coerce `value` via core `to_number`, re-raising stored error codes."""
    number = to_number(_arith_operand(value))
    if isinstance(number, CoreXlError):
        raise XlError(number.value)
    return float(number)


def xl_add(left: object, right: object) -> object:
    """Excel `+` via `core.operators.xl_add`."""
    return _adapt_core(_core_add(_arith_operand(left), _arith_operand(right)))


def xl_concat(left: object, right: object) -> object:
    """Excel `&` with shared blank, boolean, number, and error semantics."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_concat(_as_formula(left), _as_formula(right)))


def xl_sub(left: object, right: object) -> object:
    """Excel `-` via `core.operators.xl_sub`."""
    return _adapt_core(_core_sub(_arith_operand(left), _arith_operand(right)))


def xl_mul(left: object, right: object) -> object:
    """Excel `*` via `core.operators.xl_mul`."""
    return _adapt_core(_core_mul(_arith_operand(left), _arith_operand(right)))


def xl_div(numerator: object, denominator: object) -> object:
    """Excel `/` via `core.operators.xl_div`."""
    return _adapt_core(_core_div(_arith_operand(numerator), _arith_operand(denominator)))


def xl_pow(left: object, right: object) -> object:
    """Excel `^` via `core.operators.xl_pow`."""
    return _adapt_core(_core_pow(_arith_operand(left), _arith_operand(right)))


def xl_neg(value: object) -> object:
    """Excel unary `-` via `core.operators.xl_neg`."""
    return _adapt_core(_core_neg(_arith_operand(value)))


def xl_pos(value: object) -> object:
    """Excel unary `+` via `core.operators.xl_pos`."""
    return _adapt_core(_core_pos(_arith_operand(value)))


def xl_eq(left: object, right: object) -> object:
    """Excel `=` via `core.operators.xl_eq`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_eq(_as_formula(left), _as_formula(right)))


def xl_ne(left: object, right: object) -> object:
    """Excel `<>` via `core.operators.xl_ne`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_ne(_as_formula(left), _as_formula(right)))


def xl_lt(left: object, right: object) -> object:
    """Excel `<` via `core.operators.xl_lt`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_lt(_as_formula(left), _as_formula(right)))


def xl_gt(left: object, right: object) -> object:
    """Excel `>` via `core.operators.xl_gt`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_gt(_as_formula(left), _as_formula(right)))


def xl_le(left: object, right: object) -> object:
    """Excel `<=` via `core.operators.xl_le`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_le(_as_formula(left), _as_formula(right)))


def xl_ge(left: object, right: object) -> object:
    """Excel `>=` via `core.operators.xl_ge`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_ge(_as_formula(left), _as_formula(right)))


OPERATOR_TABLE = {
    "+": xl_add,
    "-": xl_sub,
    "*": xl_mul,
    "/": xl_div,
    "^": xl_pow,
    "=": xl_eq,
    "<>": xl_ne,
    "<": xl_lt,
    ">": xl_gt,
    "<=": xl_le,
    ">=": xl_ge,
    "-u": xl_neg,
    "+u": xl_pos,
}


def xl_bool(value: object) -> bool:
    """Coerce an `IF` condition with Excel boolean rules.

    `to_bool("")` is `False` (evaluator / `to_bool`), not Excel's `#VALUE!`.
    Non-boolean text such as `"nope"` raises `#VALUE!`.
    """
    _raise_stored_error(value)
    result = _adapt_core(to_bool(_as_formula(value)))
    assert isinstance(result, bool), f"IF condition returned {type(result).__name__}"
    return result


def xl_exp(*args: object) -> object:
    """Excel `EXP` via `core.math_funcs.exp_number`."""
    for arg in args:
        _raise_stored_error(arg)
    return _adapt_core(exp_number(*cast(tuple[CellValue, ...], args)))


def xl_abs(*args: object) -> object:
    """Excel `ABS` via `core.math_funcs.abs_number`."""
    for arg in args:
        _raise_stored_error(arg)
    return _adapt_core(abs_number(*cast(tuple[CellValue, ...], args)))


def xl_value(*args: object) -> object:
    """Excel `VALUE` via `core.text_funcs.value_from_text`."""
    for arg in args:
        _raise_stored_error(arg)
    if len(args) != 1:
        raise XlError("#VALUE!")
    return _adapt_core(value_from_text(cast(CellValue, args[0])))


def xl_sum(*args: object) -> object:
    """Excel `SUM` via `core.math_funcs.sum_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(sum_cells(*(_as_core_cells(arg) for arg in args)))


def xl_average(*args: object) -> object:
    """Excel `AVERAGE` via `core.math_funcs.average_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(average_cells(*(_as_core_cells(arg) for arg in args)))


def xl_min(*args: object) -> object:
    """Excel `MIN` via `core.math_funcs.min_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(min_cells(*(_as_core_cells(arg) for arg in args)))


def xl_max(*args: object) -> object:
    """Excel `MAX` via `core.math_funcs.max_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(max_cells(*(_as_core_cells(arg) for arg in args)))


def xl_if(cond: object, then_value: object, else_value: object = False) -> object:
    """Excel `IF` via `core.logic_funcs.logical_if` (scalar or element-wise)."""
    return _adapt_core(logical_if(cond, then_value, else_value))


def xl_and(*args: object) -> object:
    """Excel `AND` via `core.logic_funcs.logical_and`."""
    return _adapt_core(logical_and(*(_as_core_cells(arg) for arg in args)))


def xl_or(*args: object) -> object:
    """Excel `OR` via `core.logic_funcs.logical_or`."""
    return _adapt_core(logical_or(*(_as_core_cells(arg) for arg in args)))


def xl_not(arg: object) -> object:
    """Excel `NOT` via `core.logic_funcs.logical_not`."""
    return _adapt_core(logical_not(_as_core_cells(arg)))


def xl_sumproduct(*args: object) -> object:
    """Excel `SUMPRODUCT` via `core.sumproduct.sumproduct_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(sumproduct_cells(*(_as_core_cells(arg) for arg in args)))


def xl_choose(index: object, *choices: float) -> float:
    """Excel `CHOOSE`: 1-based selection over already-evaluated arguments."""
    position = int(_as_number(index))
    if position < 1 or position > len(choices):
        raise XlError("#VALUE!")
    return choices[position - 1]


def xl_choose_lazy(index: object, *choices: Callable[[], object]) -> object:
    """Select one CHOOSE branch before evaluating its workbook expression."""
    selected = int(xl_choose(index, *range(len(choices))))
    return choices[selected]()


def xl_choose_range(index: object, cells: Range) -> object:
    """Select the `index`-th cell of a one-row or one-column view, as `CHOOSE` lists it."""
    rows, cols = cells.shape
    selected = int(xl_choose(index, *range(rows * cols)))
    return cells.cell(selected // cols + 1, selected % cols + 1)


def xl_lookup_cell(measure: object, workbook: object) -> object:
    """Return `measure`, restoring `workbook`'s Excel type after dtype stringify.

    INDEX/MATCH tables still read the bound series so `overrides` apply. A
    string measure that is only `str(workbook)` is the series dtype hiding a
    number or bool; Excel `INDEX` returns the worksheet type.
    """
    if measure == workbook:
        return measure
    if isinstance(measure, str) and measure == str(workbook):
        return workbook
    return measure


def _as_native_grid(natives: object, height: int, width: int) -> tuple[tuple[object, ...], ...]:
    """Interpret `natives` as a `height` by `width` row-major grid."""
    if isinstance(natives, str) or not isinstance(natives, Sequence):
        raise TypeError("natives must be a nested sequence")
    rows = list(natives)
    if height == 1 and width == len(rows) and (not rows or not isinstance(rows[0], Sequence)):
        return (tuple(rows),)
    grid: list[tuple[object, ...]] = []
    for row in rows:
        if isinstance(row, str) or not isinstance(row, Sequence):
            grid.append((row,))
        else:
            grid.append(tuple(row))
    if len(grid) != height or any(len(row) != width for row in grid):
        got = f"{len(grid)}x{len(grid[0]) if grid else 0}"
        raise ValueError(f"natives shape {got} != {height}x{width}")
    return tuple(grid)


def xl_typed_range(values: Range, natives: object) -> Range:
    """Apply `xl_lookup_cell` to each cell of `values` using `natives`.

    `natives` is a row-major nested sequence matching `values.shape`.
    """
    height, width = values.shape
    grid = _as_native_grid(natives, height, width)

    def resolve(row: int, column: int) -> FormulaValue:
        return cast(
            FormulaValue,
            xl_lookup_cell(values.cell(row, column), grid[row - 1][column - 1]),
        )

    return Range("", 1, 1, height, width, lambda address: None, _coord_resolver=resolve)


def xl_index(array: object, row_num: object = None, col_num: object = None) -> object:
    """Excel `INDEX` via `core.lookup_funcs.index_cells`."""
    return _adapt_core(index_cells(array, row_num, col_num))


def xl_match(lookup: object, lookup_array: object, match_type: int = 0) -> int:
    """Excel `MATCH` via `core.lookup_funcs.match_cells`."""
    _raise_stored_error(lookup)
    result = match_cells(lookup, _as_core_cells(lookup_array), match_type)
    adapted = _adapt_core(result)
    if not isinstance(adapted, int | float):
        raise TypeError(f"MATCH returned {type(adapted).__name__}")
    return int(adapted)


def xl_vlookup(
    lookup: object,
    table_array: object,
    col_index_num: object,
    range_lookup: object = True,
) -> object:
    """Excel `VLOOKUP` via `core.lookup_funcs.vlookup_cells`."""
    _raise_stored_error(lookup)
    _raise_stored_error(col_index_num)
    _raise_stored_error(range_lookup)
    return _adapt_core(vlookup_cells(lookup, table_array, col_index_num, range_lookup))


def _call_shared(function: Callable[..., object], *args: object) -> object:
    """Translate the shared runtime's exception channel at the export boundary."""
    try:
        return function(*args)
    except SharedXlError as exc:
        raise XlError(exc.code.value) from exc


def _shared_value(function: Callable[..., object], *args: object) -> object:
    """Pass scalar and range values to a shared worksheet function."""
    return _call_shared(function, *(_shared_operand(arg) for arg in args))


def _shared_operand(value: object) -> object:
    """Preserve error cells for each shared function to consume or propagate."""
    if isinstance(value, str):
        return CoreXlError(value) if is_error(value) else value
    if isinstance(value, Sequence):
        return tuple(_shared_operand(item) for item in value)
    return value


def _shared_thunk(value: Callable[[], object]) -> Callable[[], object]:
    """Expose stored or raised inverted-tree errors to a shared lazy consumer."""

    def evaluate() -> object:
        try:
            result = value()
            _raise_stored_error(result)
            return result
        except XlError as exc:
            raise SharedXlError(CoreXlError(exc.code)) from exc

    return evaluate


def xl_iferror(value: Callable[[], object], fallback: Callable[[], object]) -> object:
    """Evaluate IFERROR lazily with shared error-consumer semantics."""
    return _call_shared(_shared_iferror, _shared_thunk(value), _shared_thunk(fallback))


def xl_ifna(value: Callable[[], object], fallback: Callable[[], object]) -> object:
    """Evaluate IFNA lazily, catching only NA errors."""
    return _call_shared(_shared_ifna, _shared_thunk(value), _shared_thunk(fallback))


def xl_iserror(value: Callable[[], object]) -> object:
    """Inspect an expression for errors without propagating them."""
    return _call_shared(_shared_iserror, _shared_thunk(value))


def xl_isna(value: Callable[[], object]) -> object:
    """Inspect an expression for an NA error."""
    return _call_shared(_shared_isna, _shared_thunk(value))


def xl_isblank(value: Callable[[], object]) -> object:
    """Inspect an expression for a blank using shared semantics."""
    return _call_shared(_shared_isblank, _shared_thunk(value))


def xl_isnumber(value: Callable[[], object]) -> object:
    """Inspect a possibly failing expression for a numeric value."""
    return _call_shared(_shared_isnumber, _shared_thunk(value))


def xl_istext(value: Callable[[], object]) -> object:
    """Inspect a possibly failing expression for a text value."""
    return _call_shared(_shared_istext, _shared_thunk(value))


def xl_npv(*args: object) -> object:
    """Compute NPV with the shared financial implementation."""
    return _shared_value(_shared_npv, *args)


def xl_rank(*args: object) -> object:
    """Compute RANK with the shared statistical implementation."""
    return _shared_value(_shared_rank, *args)


def xl_large(*args: object) -> object:
    """Compute LARGE with the shared statistical implementation."""
    return _shared_value(_shared_large, *args)


def xl_stdev(*args: object) -> object:
    """Compute STDEV with the shared statistical implementation."""
    return _shared_value(_shared_stdev, *args)


def xl_countif(*args: object) -> object:
    """Compute COUNTIF with the shared criteria implementation."""
    return _shared_value(_shared_countif, *args)


def xl_round(*args: object) -> object:
    """Round numbers with the shared Excel implementation."""
    return _shared_value(_shared_round, *args)


def xl_rounddown(*args: object) -> object:
    """Round toward zero with the shared Excel implementation."""
    return _shared_value(_shared_rounddown, *args)


def xl_numbervalue(*args: object) -> object:
    """Parse numeric text with the shared Excel implementation."""
    return _shared_value(_shared_numbervalue, *args)


def xl_text(*args: object) -> object:
    """Format a value as text with the shared Excel implementation."""
    return _shared_value(_shared_text, *args)


def xl_left(*args: object) -> object:
    """Extract leading text with the shared Excel implementation."""
    return _shared_value(_shared_left, *args)


def xl_hlookup(*args: object) -> object:
    """Look up a horizontal table with shared Excel semantics."""
    return _shared_value(_shared_hlookup, *args)


def xl_lookup(*args: object) -> object:
    """Look up a vector or table with shared Excel semantics."""
    return _shared_value(_shared_lookup, *args)


def xl_xlookup(*args: object) -> object:
    """Look up corresponding arrays with shared Excel semantics."""
    return _shared_value(_shared_xlookup, *args)


def xl_at(values: Sequence[T], index: object) -> T:
    """Return `values[index]` (0-based), raising `#VALUE!` when out of range.

    `index` is coerced with core `to_number` and truncated toward zero.
    """
    position = int(_as_number(index))
    if position < 0 or position >= len(values):
        raise XlError("#VALUE!")
    return values[position]


def xl_raise(code: str) -> NoReturn:
    """Raise `XlError(code)` from generated expression position."""
    raise XlError(code)
