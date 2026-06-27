from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from datetime import date, datetime, timedelta
from typing import TYPE_CHECKING, Any, Literal, TypeGuard, cast

from .runtime import coerce_inputs_dict

Layout = Literal["scalar", "series", "matrix"]
SetterInput = object
Scalar = str | int | float | bool | datetime | None
Record = dict[str, object]
Records = list[Record]
SeriesInput = Records | Record | Sequence[Scalar]

if TYPE_CHECKING:
    import pandas as pd
    import polars as pl

    SeriesInput = Records | Record | Sequence[Scalar] | pd.DataFrame | pl.DataFrame
else:
    SeriesInput = Records | Record | Sequence[Scalar] | object

"""Coerce workbook and manifest values into series-binding scalars."""

_EXCEL_EPOCH = datetime(1899, 12, 30)

_BOOL_TRUE = frozenset({"true", "1", "yes"})

_BOOL_FALSE = frozenset({"false", "0", "no"})

_DATETIME_CLS = datetime

_DATE_CLS = date

def _ensure_naive_datetime(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value
    raise ValueError(f"Timezone-aware datetime values are not supported: {value!r}")

def _normalize_date(value: date) -> datetime:
    return datetime.combine(value, datetime.min.time())

def _parse_iso_datetime(text: str) -> datetime:
    stripped = text.strip()
    if not stripped:
        raise ValueError(f"Cannot coerce {text!r} to datetime")
    if "T" in stripped or " " in stripped:
        parsed = datetime.fromisoformat(stripped.replace("Z", "+00:00"))
        return _ensure_naive_datetime(parsed)
    return _normalize_date(date.fromisoformat(stripped))

def _excel_serial_to_datetime(serial: float) -> datetime:
    """Convert an Excel day-fraction serial to a naive datetime."""
    return _EXCEL_EPOCH + timedelta(days=serial)

def _coerce_datetime(raw: Any, *, excel_serial: bool) -> datetime:
    if isinstance(raw, _DATETIME_CLS):
        return _ensure_naive_datetime(raw)
    if isinstance(raw, _DATE_CLS):
        return _normalize_date(raw)
    if isinstance(raw, str):
        return _parse_iso_datetime(raw)
    if excel_serial and isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return _excel_serial_to_datetime(float(raw))
    raise ValueError(f"Cannot coerce {raw!r} to datetime")

def _coerce_bool(raw: Any) -> bool:
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, int) and not isinstance(raw, bool):
        if raw == 0:
            return False
        if raw == 1:
            return True
        raise ValueError(f"Cannot coerce {raw!r} to bool")
    if isinstance(raw, float):
        if raw == 0.0:
            return False
        if raw == 1.0:
            return True
        raise ValueError(f"Cannot coerce {raw!r} to bool")
    text = str(raw).strip().lower()
    if text in _BOOL_TRUE:
        return True
    if text in _BOOL_FALSE:
        return False
    raise ValueError(f"Cannot coerce {raw!r} to bool")

def coerce_scalar(raw: Any, read_as: str) -> Scalar:
    """Coerce a workbook or manifest value using an explicit or automatic read mode."""
    if raw is None:
        return None
    if read_as == "auto":
        if isinstance(raw, bool):
            return raw
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        if isinstance(raw, float):
            return raw
        if isinstance(raw, str):
            return raw
        if isinstance(raw, _DATETIME_CLS):
            return _ensure_naive_datetime(raw)
        if isinstance(raw, _DATE_CLS):
            return _normalize_date(raw)
        return str(raw)
    if read_as == "string":
        return str(raw)
    if read_as == "int":
        return int(raw)
    if read_as == "float":
        return float(raw)
    if read_as == "number":
        if isinstance(raw, int) and not isinstance(raw, bool):
            return raw
        return float(raw)
    if read_as == "bool":
        return _coerce_bool(raw)
    if read_as == "datetime":
        return _coerce_datetime(raw, excel_serial=True)
    raise ValueError(f"Unknown read mode: {read_as!r}")

def coerce_constant(value: Any, *, read_as: str) -> Scalar:
    """Coerce a manifest constant using the effective read mode."""
    return coerce_scalar(value, read_as)

"""Coerce setter caller input into canonical record lists for series bindings."""

def _is_mapping(value: object) -> TypeGuard[Mapping[str, object]]:
    return isinstance(value, Mapping) and not isinstance(value, (str, bytes, bytearray))

def _is_pandas_dataframe(data: object) -> bool:
    cls = type(data)
    module = cls.__module__
    return cls.__name__ == "DataFrame" and (module == "pandas" or module.startswith("pandas."))

def _is_polars_dataframe(data: object) -> bool:
    cls = type(data)
    return cls.__module__.startswith("polars.") and cls.__name__ == "DataFrame"

def _is_tabular_dataframe(data: object) -> bool:
    return _is_pandas_dataframe(data) or _is_polars_dataframe(data)

def _import_pandas() -> Any:
    try:
        import pandas as pd
    except ImportError as exc:
        raise ImportError(
            "DataFrame input requires pandas; install it or pass records / a 1D iterable"
        ) from exc
    return pd

def _import_polars() -> Any:
    try:
        import polars as pl
    except ImportError as exc:
        raise ImportError(
            "DataFrame input requires polars; install it or pass records / a 1D iterable"
        ) from exc
    return pl

def _coerce_scalar_records(
    data: object,
    measure_field: str,
) -> Records:
    """Normalize scalar-layout setter input to a record list."""
    if isinstance(data, list):
        return cast(Records, data)
    if _is_mapping(data):
        return [dict(data)]
    return [{measure_field: data}]

def _is_records_list(data: Sequence[object]) -> TypeGuard[Records]:
    if not data:
        return True
    return all(_is_mapping(item) for item in data)

def _coerce_key_value(
    field: str,
    raw: object,
    key_dtypes: Mapping[str, str] | None,
) -> object:
    if key_dtypes is None:
        return raw
    read_as = key_dtypes.get(field)
    if read_as is None:
        return raw
    return coerce_scalar(raw, read_as)

def _dataframe_column_names(data: object) -> list[str]:
    columns = getattr(data, "columns", None)
    if columns is None:
        raise TypeError(f"unsupported DataFrame-like input: {type(data)!r}")
    return [str(column) for column in columns]

def _validate_dataframe_columns(
    column_names: list[str],
    *,
    key_fields: tuple[str, ...],
    measure_field: str,
    strict: bool,
) -> None:
    required = set(key_fields) | {measure_field}
    present = set(column_names)
    missing = sorted(required - present)
    if missing:
        raise ValueError(f"missing required column(s): {missing!r}")
    if strict:
        unknown = sorted(present - required)
        if unknown:
            raise ValueError(f"unknown columns {unknown!r}")

def _row_dicts_from_dataframe(data: object) -> list[Record]:
    if _is_pandas_dataframe(data):
        pd = _import_pandas()
        if not isinstance(data, pd.DataFrame):
            raise ImportError(
                "DataFrame input requires pandas; install it or pass records / a 1D iterable"
            )
        return data.to_dict(orient="records")
    if _is_polars_dataframe(data):
        pl = _import_polars()
        if not isinstance(data, pl.DataFrame):
            raise ImportError(
                "DataFrame input requires polars; install it or pass records / a 1D iterable"
            )
        return data.to_dicts()
    raise TypeError(f"unsupported DataFrame-like input: {type(data)!r}")

def _apply_key_dtypes(
    records: Records,
    *,
    key_fields: tuple[str, ...],
    key_dtypes: Mapping[str, str] | None,
) -> Records:
    """Coerce key field values on each record using binding read modes."""
    if not key_dtypes:
        return records
    coerced: list[dict[str, object]] = []
    for record in records:
        updated = dict(record)
        for field in key_fields:
            if field in updated:
                updated[field] = _coerce_key_value(field, updated[field], key_dtypes)
        coerced.append(updated)
    return coerced

def _coerce_dataframe_records(
    data: object,
    *,
    key_fields: tuple[str, ...],
    measure_field: str,
    strict: bool,
) -> Records:
    column_names = _dataframe_column_names(data)
    _validate_dataframe_columns(
        column_names,
        key_fields=key_fields,
        measure_field=measure_field,
        strict=strict,
    )
    records: list[dict[str, object]] = []
    for row in _row_dicts_from_dataframe(data):
        record: dict[str, object] = {field: row[field] for field in key_fields}
        record[measure_field] = row[measure_field]
        records.append(record)
    return records

def _non_scalar_input_hint(
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
) -> str:
    if layout == "matrix":
        columns = ", ".join([*key_fields, measure_field])
        return f"pass records or a tidy DataFrame with columns {columns!r}"
    return "pass records, a 1D iterable of measure values, or a tidy DataFrame"

def _unsupported_non_scalar_input_type_error(
    data: object,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
) -> TypeError:
    hint = _non_scalar_input_hint(
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
    )
    return TypeError(f"unsupported {layout} setter input type {type(data)!r}; {hint}")

def _coerce_positional_records(
    data: Iterable[object],
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...],
) -> Records:
    if len(key_fields) != 1:
        if layout == "matrix":
            raise ValueError(
                "positional measure values are not supported for matrix setters; "
                "pass records or a tidy DataFrame"
            )
        raise ValueError(
            "positional measure values require a single-key series binding; "
            f"got key_fields={list(key_fields)!r}"
        )
    values = list(data)
    if len(values) != len(key_order):
        raise ValueError(
            f"expected {len(key_order)} values for positional input, got {len(values)}"
        )
    key_field = key_fields[0]
    return [
        {key_field: key, measure_field: value} for key, value in zip(key_order, values, strict=True)
    ]

def _coerce_non_scalar_records(
    data: SetterInput,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...] | None,
    strict: bool,
) -> Records:
    """Normalize series/matrix setter input to records before key coercion."""
    if _is_tabular_dataframe(data):
        return _coerce_dataframe_records(
            data,
            key_fields=key_fields,
            measure_field=measure_field,
            strict=strict,
        )

    if _is_mapping(data):
        return [dict(data)]

    if isinstance(data, list):
        if _is_records_list(data):
            return data
        if key_order is None:
            raise ValueError("positional input requires key_order")
        return _coerce_positional_records(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
            key_order=key_order,
        )

    if isinstance(data, (str, bytes, bytearray)):
        raise _unsupported_non_scalar_input_type_error(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
        )

    if isinstance(data, Iterable):
        if key_order is None:
            raise ValueError("positional input requires key_order")
        return _coerce_positional_records(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
            key_order=key_order,
        )

    raise _unsupported_non_scalar_input_type_error(
        data,
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
    )

def coerce_setter_input(
    data: SetterInput,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...] | None,
    strict: bool,
    key_dtypes: Mapping[str, str] | None = None,
) -> Records:
    """Normalize caller input into records for ``_apply_series_records``.

    Args:
        data: Scalar value, record(s), 1D measure values, or tidy DataFrame.
        layout: Binding layout (`scalar`, `series`, or `matrix`).
        key_fields: Key column names from the binding manifest.
        measure_field: Measure concept name (e.g. `OBS_VALUE`).
        key_order: Canonical key values for positional measure iterables.
        strict: When true, reject unknown DataFrame columns.
        key_dtypes: Optional read modes per key field applied to all input shapes.

    Returns:
        List of record dicts ready for leaf resolution.

    Raises:
        ImportError: When a DataFrame-like value is passed but pandas/polars is missing.
        TypeError: When the input shape is unsupported for the layout.
        ValueError: When columns, keys, or positional lengths are invalid.
    """
    if layout == "scalar":
        if _is_tabular_dataframe(data):
            raise TypeError("scalar setters do not accept DataFrame input")
        return _coerce_scalar_records(data, measure_field)

    records = _coerce_non_scalar_records(
        data,
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
        key_order=key_order,
        strict=strict,
    )
    return _apply_key_dtypes(
        records,
        key_fields=key_fields,
        key_dtypes=key_dtypes,
    )

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
    first_record_by_address: dict[str, int] = {}
    for index, record in enumerate(records):
        if strict:
            unknown = set(record) - allowed_fields
            if unknown:
                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")
        if measure_field not in record:
            raise ValueError(f"record[{index}]: missing required field {measure_field!r}")
        address = None
        key_tuple = None
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
        elif not requires_address and all(field in record for field in key_fields):
            key_tuple = tuple((field, record[field]) for field in key_fields)
        assert address is not None
        if address in updates:
            prior = first_record_by_address[address]
            if key_tuple is not None:
                detail = f"duplicate key {dict(key_tuple)!r} matches record[{prior}]"
            else:
                detail = f"duplicate cell {address!r} matches record[{prior}]"
            raise ValueError(f"record[{index}]: {detail}")
        first_record_by_address[address] = index
        updates[address] = record[measure_field]
    if updates:
        ctx.set_inputs(coerce_inputs_dict(updates))
