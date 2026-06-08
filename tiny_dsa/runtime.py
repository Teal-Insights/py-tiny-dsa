"""Standalone runtime for generated Excel formula code."""

from __future__ import annotations

import warnings
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from enum import StrEnum
from typing import TypeAlias, cast

import fastpyxl.utils.cell
import numpy as np

class CircularReferenceWarning(RuntimeWarning):
    """Warning emitted when a circular reference is encountered (default Excel mode)."""

@dataclass(slots=True)
class EvalContext:
    """Per-run evaluation state for generated spreadsheets.

    The exported-code path needs a mutable inputs mapping and a cache that is scoped
    to a single compute call, so callers can run many scenarios without global state.
    """

    inputs: dict[str, CellValue]
    resolver: Callable[[str], Callable[[EvalContext], CellValue] | None]
    cache: dict[str, CellValue] = field(default_factory=dict)
    computing: set[str] = field(default_factory=set)
    deps: dict[str, set[str]] = field(default_factory=dict)
    reverse_deps: dict[str, set[str]] = field(default_factory=dict)
    stack: list[str] = field(default_factory=list)
    iterative_enabled: bool = False
    iterate_count: int = 100
    iterate_delta: float = 0.001
    iteration_values: dict[str, CellValue] = field(default_factory=dict)

    def _record_dependency(self, parent: str, child: str) -> None:
        if parent == child:
            return
        self.deps.setdefault(parent, set()).add(child)
        self.reverse_deps.setdefault(child, set()).add(parent)

    def invalidate(self, addresses: Iterable[str]) -> None:
        """Invalidate cached values for the given addresses and their dependents."""
        to_visit = list(addresses)
        seen: set[str] = set()
        while to_visit:
            addr = to_visit.pop()
            if addr in seen:
                continue
            seen.add(addr)

            self.cache.pop(addr, None)
            self.computing.discard(addr)

            dependents = list(self.reverse_deps.get(addr, set()))
            to_visit.extend(dependents)

            for dep in self.deps.get(addr, set()):
                parents = self.reverse_deps.get(dep)
                if parents is not None:
                    parents.discard(addr)
                    if not parents:
                        self.reverse_deps.pop(dep, None)

            self.deps.pop(addr, None)
            self.reverse_deps.pop(addr, None)

    def set_inputs(self, inputs: dict[str, CellValue]) -> None:
        """Update input values and invalidate dependent cached results."""
        changed = [k for k, v in inputs.items() if self.inputs.get(k) != v]
        self.inputs.update(inputs)
        if changed:
            self.invalidate(changed)

NormalizedAddress: TypeAlias = str

class XlError(StrEnum):
    VALUE = "#VALUE!"
    REF = "#REF!"
    DIV = "#DIV/0!"
    NA = "#N/A"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NULL = "#NULL!"

    @classmethod
    def from_text(cls, value: str) -> XlError | None:
        upper = value.strip().upper()
        for err in cls:
            if err.value == upper:
                return err
        return None

def _escape_sheet_for_formula(sheet: str) -> str:
    """Escape apostrophes for use inside quoted sheet names."""
    return sheet.replace("'", "''")

def _format_general_number(value: float | int) -> str:
    f = float(value)
    if f.is_integer():
        return str(int(f))
    return str(f)

def _quote_sheet_if_needed(sheet: str) -> str:
    if " " in sheet or "-" in sheet or "'" in sheet:
        return f"'{sheet}'"
    return sheet

def _format_address(sheet: str, row: int, col: int) -> str:
    sheet_name = _quote_sheet_if_needed(sheet)
    col_letter = fastpyxl.utils.cell.get_column_letter(col)
    return f"{sheet_name}!{col_letter}{row}"

def excel_casefold(value: str) -> str:
    return value.casefold()

def needs_quoting(sheet: str) -> bool:
    """Return True if a sheet name must be wrapped in single quotes in a formula."""
    return " " in sheet or "-" in sheet or "'" in sheet

def quote_sheet_if_needed(sheet: str) -> str:
    """Return a sheet name quoted for formulas when quoting is required."""
    if not needs_quoting(sheet):
        return sheet
    return "'" + _escape_sheet_for_formula(sheet) + "'"

def format_key(sheet: str, cell: str) -> NormalizedAddress:
    """Format a sheet and A1 cell coordinate into a canonical address string."""
    return f"{quote_sheet_if_needed(sheet)}!{cell}"

@dataclass(frozen=True, slots=True)
class ExcelRange:
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                col = fastpyxl.utils.cell.get_column_letter(c)
                yield format_key(self.sheet, f"{col}{r}")

    def resolve(self, evaluate_fn: Callable[[str], CellValue]) -> np.ndarray:
        values: list[CellValue] = [evaluate_fn(addr) for addr in self.cell_addresses()]
        rows, cols = self.shape
        return np.array(values, dtype=object).reshape((rows, cols))

CellValue: TypeAlias = float | int | str | bool | XlError | ExcelRange | np.ndarray | None

def coerce_inputs_dict(values: Mapping[str, object]) -> dict[str, CellValue]:
    """Widen inferred default-input dicts to `dict[str, CellValue]` for `EvalContext`."""
    return cast(dict[str, CellValue], dict(values))

def split_sheet_qualified_address(address: str) -> tuple[str, str] | None:
    """Split `sheet!coord` into `(sheet_name, coord)`.

    Handles quoted sheet names, including Excel's doubled-single-quote escape
    (`'O''Neil'!A1` -> sheet `O'Neil`).

    Returns `None` when *address* has no sheet qualifier (plain `A1`).
    """
    if address.startswith("'"):
        i = 1
        while i < len(address):
            if address[i] == "'":
                if i + 1 < len(address) and address[i + 1] == "'":
                    i += 2
                    continue
                break
            i += 1
        if i >= len(address):
            return None
        sheet = address[1:i].replace("''", "'")
        rest = address[i + 1 :]
        if not rest.startswith("!"):
            return None
        return sheet, rest[1:]

    if "!" not in address:
        return None
    sheet, cell = address.rsplit("!", 1)
    return sheet, cell

def _parse_sheet_address(address: str) -> tuple[str, str] | None:
    return split_sheet_qualified_address(address)

def _parse_range_address(address: str) -> tuple[str, str, str] | XlError:
    if ":" not in address:
        return XlError.VALUE
    start_text, end_text = address.split(":", 1)
    start = _parse_sheet_address(start_text)
    if start is None:
        return XlError.VALUE
    sheet, start_cell = start
    if "!" in end_text:
        end = _parse_sheet_address(end_text)
        if end is None:
            return XlError.VALUE
        end_sheet, end_cell = end
        if end_sheet != sheet:
            return XlError.VALUE
    else:
        end_cell = end_text
    return sheet, start_cell, end_cell

def to_bool(value: CellValue) -> bool | XlError:
    if value is None:
        return False
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == "":
            return False
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE

def to_number(value: CellValue) -> float | XlError:
    if value is None:
        return 0.0
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return 0.0
        try:
            return float(s)
        except ValueError:
            return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE

def _compare_values(a: CellValue, b: CellValue) -> int:
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return -1 if an < bn else 1 if an > bn else 0
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        return -1 if af < bf else 1 if af > bf else 0
    return 0

def _values_match(a: CellValue, b: CellValue) -> bool:
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    return a == b

def index_excel_range(
    base: ExcelRange,
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> ExcelRange | XlError:
    """Map INDEX(row,col) over *base* to an absolute range (single cell or slice).

    Mirrors `excel_grapher.runtime.lookup.xl_index` geometry
    so OFFSET(INDEX(...), ...) receives a true cell reference.
    """
    nrows = base.end_row - base.start_row + 1
    ncols = base.end_col - base.start_col + 1
    row_omitted = row_num is None
    col_omitted = col_num is None

    def abs_cell(r0: int, c0: int) -> ExcelRange:
        r = base.start_row + r0
        c = base.start_col + c0
        return ExcelRange(base.sheet, r, c, r, c)

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return abs_cell(0, 0)
        if nrows == 1:
            return abs_cell(0, ncols - 1)
        if ncols == 1:
            return abs_cell(nrows - 1, 0)
        return XlError.VALUE

    if row_omitted:
        cn = to_number(col_num)
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return abs_cell(0, col - 1)
        c0 = base.start_col + col - 1
        return ExcelRange(base.sheet, base.start_row, c0, base.end_row, c0)

    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return abs_cell(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return abs_cell(row - 1, 0)
        if row < 1 or row > nrows:
            return XlError.REF
        r0 = base.start_row + row - 1
        return ExcelRange(base.sheet, r0, base.start_col, r0, base.end_col)

    cn = to_number(col_num)
    if isinstance(cn, XlError):
        return cn
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return abs_cell(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return abs_cell(row - 1, 0)
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return abs_cell(row - 1, col - 1)

def to_int(value: CellValue) -> int | XlError:
    """Coerce a CellValue to an integer using Excel-style numeric coercion.

    For functions that operate on integer indices (e.g. CHOOSE/INDEX/MATCH)
    while propagating Excel errors.
    """
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return int(n)

def to_string(value: CellValue) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, XlError):
        return value.value
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, ExcelRange):
        return XlError.VALUE.value
    return str(value)

def _xl_compare(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    def _cmp_str(a: str, b: str) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    def _cmp_float(a: float, b: float) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    if isinstance(left, str) and isinstance(right, str):
        return _cmp_str(excel_casefold(left), excel_casefold(right))

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError) or isinstance(rn, XlError):
        return _cmp_str(excel_casefold(to_string(left)), excel_casefold(to_string(right)))

    return _cmp_float(float(ln), float(rn))

def xl_add(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln + rn

def xl_circular_reference() -> CellValue:
    """Excel default behavior for circular references (non-iterative calculation)."""
    warnings.warn(
        "Circular reference detected; returning 0 (iterative calculation is disabled).",
        CircularReferenceWarning,
        stacklevel=2,
    )
    return 0

def xl_cell(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a single cell address under the given context.

    Resolution order:
    - cached value (per ctx)
    - user-provided inputs
    - exported formula implementation (via resolver)
    - missing cell raises KeyError
    """
    if ctx.stack:
        ctx._record_dependency(ctx.stack[-1], address)

    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    if address in ctx.inputs:
        v = ctx.inputs[address]
        ctx.cache[address] = v
        return v

    fn = ctx.resolver(address)
    if fn is None:
        raise KeyError(f"Cell {address} not found in graph")

    ctx.computing.add(address)
    ctx.stack.append(address)
    try:
        v = fn(ctx)
        # Excel treats "empty" formula results as 0 in most numeric contexts; the evaluator
        # normalizes those Nones to 0. Structural blank-range cells intentionally stay None
        # so INDEX/MATCH (and similar) see true empty cells in object arrays.
        if v is None and not getattr(fn, "__structural_blank__", False):
            v = 0
        ctx.cache[address] = v
        return v
    finally:
        ctx.computing.discard(address)
        if ctx.stack and ctx.stack[-1] == address:
            ctx.stack.pop()

def xl_div(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    if rn == 0:
        return XlError.DIV
    return ln / rn

def xl_eval(
    ctx: EvalContext,
    address: str,
    fn: Callable[[EvalContext], CellValue],
) -> CellValue:
    """Evaluate a known formula implementation under the given context."""
    if ctx.stack:
        ctx._record_dependency(ctx.stack[-1], address)

    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    if address in ctx.inputs:
        v = ctx.inputs[address]
        ctx.cache[address] = v
        return v

    ctx.computing.add(address)
    ctx.stack.append(address)
    try:
        v = fn(ctx)
        if v is None:
            v = 0
        ctx.cache[address] = v
        return v
    finally:
        ctx.computing.discard(address)
        if ctx.stack and ctx.stack[-1] == address:
            ctx.stack.pop()

def xl_ge(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare(">=", left, right)

def xl_index_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int] | XlError:
    """INDEX semantics that return a reference suitable for OFFSET."""
    if isinstance(ref, ExcelRange):
        base = ref
    else:
        match ref:
            case (sheet, r1, c1):
                base = ExcelRange(sheet=sheet, start_row=r1, start_col=c1, end_row=r1, end_col=c1)
            case (sheet, r1, c1, r2, c2):
                base = ExcelRange(sheet=sheet, start_row=r1, start_col=c1, end_row=r2, end_col=c2)
            case _:
                return XlError.VALUE

    out = index_excel_range(base, row_num, col_num)
    if isinstance(out, XlError):
        return out
    if out.start_row == out.end_row and out.start_col == out.end_col:
        return (out.sheet, out.start_row, out.start_col)
    return (out.sheet, out.start_row, out.start_col, out.end_row, out.end_col)

def xl_match(
    lookup_value: CellValue, lookup_array: CellValue, match_type: CellValue = 1
) -> int | XlError:
    mt = to_number(match_type)
    if isinstance(mt, XlError):
        return mt
    match_type_int = int(mt)
    if isinstance(lookup_array, XlError):
        return lookup_array
    if isinstance(lookup_array, np.ndarray):
        flat = np.ravel(lookup_array)
    elif isinstance(lookup_array, (list, tuple)):
        flat = np.ravel(np.array(lookup_array, dtype=object))
    else:
        flat = np.array([lookup_array], dtype=object)
    if match_type_int == 0:
        for i, val in enumerate(flat):
            if _values_match(lookup_value, val):
                return i + 1
        return XlError.NA
    if match_type_int == 1:
        last_match = None
        for i, val in enumerate(flat):
            if _compare_values(val, lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    if match_type_int == -1:
        last_match = None
        for i, val in enumerate(flat):
            if _compare_values(val, lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    return XlError.VALUE

def xl_mul(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln * rn

def xl_offset(
    ctx: EvalContext,
    ref_info: tuple[str, int, int] | tuple[str, int, int, int, int] | XlError,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> CellValue:
    rr = to_number(rows)
    if isinstance(rr, XlError):
        return rr
    cc = to_number(cols)
    if isinstance(cc, XlError):
        return cc

    if isinstance(ref_info, XlError):
        return ref_info

    match ref_info:
        case (sheet, base_row, base_col):
            base_end_row, base_end_col = base_row, base_col
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            pass
        case _:
            return XlError.VALUE

    base_h = int(base_end_row - base_row + 1)
    base_w = int(base_end_col - base_col + 1)

    if height is None:
        h = base_h
    else:
        hh = to_number(height)
        if isinstance(hh, XlError):
            return hh
        h = int(hh)

    if width is None:
        w = base_w
    else:
        ww = to_number(width)
        if isinstance(ww, XlError):
            return ww
        w = int(ww)

    target_row = int(base_row + int(rr))
    target_col = int(base_col + int(cc))

    if target_row < 1 or target_col < 1:
        return XlError.REF
    if h <= 0 or w <= 0:
        return XlError.VALUE

    if h == 1 and w == 1:
        addr = _format_address(sheet, target_row, target_col)
        return xl_cell(ctx, addr)

    result: list[list[CellValue]] = []
    for r in range(target_row, target_row + h):
        row_values: list[CellValue] = []
        for c in range(target_col, target_col + w):
            addr = _format_address(sheet, r, c)
            row_values.append(xl_cell(ctx, addr))
        result.append(row_values)
    return np.array(result, dtype=object)

def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range and return a 2D numpy array of values."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        return parsed
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        return XlError.VALUE

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    rng = ExcelRange(sheet, start_row, start_col_idx, end_row, end_col_idx)
    return rng.resolve(lambda addr: xl_cell(ctx, addr))

def xl_sub(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln - rn
