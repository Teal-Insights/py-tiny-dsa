"""Memoized evaluator and per-output input bundles."""

from __future__ import annotations

from dataclasses import Field, dataclass, fields
from functools import cached_property
from pathlib import Path
from typing import Annotated, Any, ClassVar, Literal, Self

from . import data, internals, validation
from .runtime import Between
from .workbook import read_bound_inputs


def _bind_inputs(target: object, inputs: dict[str, Any], *, validate: bool = True) -> None:
    if not validate:
        for name, value in inputs.items():
            setattr(target, name, value)
        return
    for name, value in inputs.items():
        check = validation.CHECKS.get(name)
        setattr(target, name, value if check is None else check(value))


class _BoundInputs:
    """Shared workbook bind and CHECKS validation for Model and input bundles."""

    __dataclass_fields__: ClassVar[dict[str, Field[Any]]]
    _INPUT_IDS: tuple[str, ...] = ()

    @classmethod
    def from_workbook(cls, workbook: Path | str, **overrides: object) -> Self:
        """Bind input leaves from a populated workbook of this vintage."""
        declared = getattr(cls, "__dataclass_fields__", None)
        names = tuple(declared) if declared else cls._INPUT_IDS
        unknown = overrides.keys() - set(names)
        if unknown:
            raise TypeError(f"unknown inputs: {sorted(unknown)}")
        values = read_bound_inputs(Path(workbook), names, data)
        values.update(overrides)
        return cls(**values)

    def __post_init__(self) -> None:
        self._validate()

    def _validate(self) -> None:
        holder = Model.__new__(Model)
        names = {field.name for field in fields(self)}
        values = {name: getattr(self, name) for name in names}
        _bind_inputs(holder, values)
        for name in names:
            object.__setattr__(self, name, getattr(holder, name))


class _SnapshotInputs(_BoundInputs):
    """Per-output bundle factory over `data.*_DEFAULT` leaves."""

    @classmethod
    def from_defaults(cls, **overrides: object) -> Self:
        names = {field.name for field in fields(cls)}
        unexpected = overrides.keys() - names
        if unexpected:
            listed = ", ".join(sorted(unexpected))
            raise TypeError(f"{cls.__name__}.from_defaults() got unknown argument(s): {listed}")
        values = {
            name: (
                overrides[name] if name in overrides else getattr(data, f"{name.upper()}_DEFAULT")
            )
            for name in names
        }
        return cls(**values)


class Model(_BoundInputs):
    """Formula series of the workbook, evaluated on demand from bound inputs.

    Each attribute evaluates its named formula once per model. Only the
    inputs bound at construction are available, so a public function
    supplies exactly the leaves of its output. Unknown constructor
    names fail closed. `from_defaults` binds every input from
    `data.*_DEFAULT`. `from_workbook` reads those input cells from a
    populated workbook of this vintage.
    """

    country_name: Literal["Aurelium", "Borvelia", "Litellia"]
    country_initial_debt: data.CountryInitialDebt
    growth_baseline: data.GrowthBaseline
    interest_baseline: data.InterestBaseline
    primary_balance_baseline: data.PrimaryBalanceBaseline
    shock_year: Annotated[int, Between(1, 5)]
    shock_type: Literal[1, 2, 3]
    shock_magnitudes: data.ShockMagnitudes
    _INPUT_IDS: tuple[str, ...] = ("country_name", "country_initial_debt", "growth_baseline", "interest_baseline", "primary_balance_baseline", "shock_year", "shock_type", "shock_magnitudes")

    def __init__(self, bundle: _BoundInputs | None = None, /, **inputs: Any) -> None:
        if bundle is not None:
            if not isinstance(bundle, _BoundInputs):
                raise TypeError(
                    f"Model() bundle must be a bound inputs instance, not {type(bundle).__name__}"
                )
            if inputs:
                raise TypeError("Model() does not accept keyword inputs with a bound bundle")
            values = {field.name: getattr(bundle, field.name) for field in fields(bundle)}
            _bind_inputs(self, values, validate=False)
            return
        unknown = inputs.keys() - {"country_name", "country_initial_debt", "growth_baseline", "interest_baseline", "primary_balance_baseline", "shock_year", "shock_type", "shock_magnitudes"}
        if unknown:
            raise TypeError(f"unknown inputs: {sorted(unknown)}")
        _bind_inputs(self, inputs)

    @classmethod
    def from_defaults(
        cls,
        *,
        country_name: Literal["Aurelium", "Borvelia", "Litellia"] = data.COUNTRY_NAME_DEFAULT,
        country_initial_debt: data.CountryInitialDebt = data.COUNTRY_INITIAL_DEBT_DEFAULT,
        growth_baseline: data.GrowthBaseline = data.GROWTH_BASELINE_DEFAULT,
        interest_baseline: data.InterestBaseline = data.INTEREST_BASELINE_DEFAULT,
        primary_balance_baseline: data.PrimaryBalanceBaseline = data.PRIMARY_BALANCE_BASELINE_DEFAULT,
        shock_year: Annotated[int, Between(1, 5)] = data.SHOCK_YEAR_DEFAULT,
        shock_type: Literal[1, 2, 3] = data.SHOCK_TYPE_DEFAULT,
        shock_magnitudes: data.ShockMagnitudes = data.SHOCK_MAGNITUDES_DEFAULT,
    ) -> Model:
        """Bind every input from `data.*_DEFAULT`, then apply overrides."""
        return cls(
            country_name=country_name,
            country_initial_debt=country_initial_debt,
            growth_baseline=growth_baseline,
            interest_baseline=interest_baseline,
            primary_balance_baseline=primary_balance_baseline,
            shock_year=shock_year,
            shock_type=shock_type,
            shock_magnitudes=shock_magnitudes,
        )

    @cached_property
    def initial_debt_resolved(self) -> float | str:
        return internals.initial_debt_resolved(country_profile_names=data.COUNTRY_PROFILE_NAMES, country_name=self.country_name, country_initial_debt=self.country_initial_debt)

    @cached_property
    def engine_initial_debt_baseline(self) -> float | str:
        return internals.engine_initial_debt_baseline(initial_debt_resolved=self.initial_debt_resolved)

    @cached_property
    def engine_initial_debt_shocked(self) -> float | str:
        return internals.engine_initial_debt_shocked(initial_debt_resolved=self.initial_debt_resolved)

    @cached_property
    def shock_magnitude_resolved(self) -> float | str:
        return internals.shock_magnitude_resolved(shock_type=self.shock_type, shock_magnitudes=self.shock_magnitudes)

    @cached_property
    def shock_active(self) -> data.Series[int | str | None]:
        return internals.shock_active(engine_year_labels=data.ENGINE_YEAR_LABELS, shock_year=self.shock_year)

    @cached_property
    def shocked_growth(self) -> data.Series[float | str | None]:
        return internals.shocked_growth(growth_baseline=self.growth_baseline, shock_type=self.shock_type, shock_magnitude_resolved=self.shock_magnitude_resolved, shock_active=self.shock_active)

    @cached_property
    def shocked_interest(self) -> data.Series[float | str | None]:
        return internals.shocked_interest(interest_baseline=self.interest_baseline, shock_type=self.shock_type, shock_magnitude_resolved=self.shock_magnitude_resolved, shock_active=self.shock_active)

    @cached_property
    def shocked_primary_balance(self) -> data.Series[float | str | None]:
        return internals.shocked_primary_balance(primary_balance_baseline=self.primary_balance_baseline, shock_type=self.shock_type, shock_magnitude_resolved=self.shock_magnitude_resolved, shock_active=self.shock_active)

    @cached_property
    def baseline_path_internal(self) -> data.Series[float | str | None]:
        return internals.baseline_path_internal(engine_initial_debt_baseline=self.engine_initial_debt_baseline, growth_baseline=self.growth_baseline, interest_baseline=self.interest_baseline, primary_balance_baseline=self.primary_balance_baseline)

    @cached_property
    def shocked_path_internal(self) -> data.Series[float | str | None]:
        return internals.shocked_path_internal(engine_initial_debt_shocked=self.engine_initial_debt_shocked, shocked_growth=self.shocked_growth, shocked_interest=self.shocked_interest, shocked_primary_balance=self.shocked_primary_balance)

    @cached_property
    def output_baseline(self) -> data.OutputBaseline:
        return internals.output_baseline(baseline_path_internal=self.baseline_path_internal)

    @cached_property
    def output_shocked(self) -> data.OutputShocked:
        return internals.output_shocked(shocked_path_internal=self.shocked_path_internal)

    @cached_property
    def output_delta(self) -> data.OutputDelta:
        return internals.output_delta(output_baseline=self.output_baseline, output_shocked=self.output_shocked)


@dataclass(frozen=True, kw_only=True)
class OutputBaselineInputs(_SnapshotInputs):
    """Bound input leaves for `compute_output_baseline`."""

    country_name: Literal["Aurelium", "Borvelia", "Litellia"]
    country_initial_debt: data.CountryInitialDebt
    growth_baseline: data.GrowthBaseline
    interest_baseline: data.InterestBaseline
    primary_balance_baseline: data.PrimaryBalanceBaseline


@dataclass(frozen=True, kw_only=True)
class OutputShockedInputs(_SnapshotInputs):
    """Bound input leaves for `compute_output_shocked`."""

    country_name: Literal["Aurelium", "Borvelia", "Litellia"]
    country_initial_debt: data.CountryInitialDebt
    growth_baseline: data.GrowthBaseline
    interest_baseline: data.InterestBaseline
    primary_balance_baseline: data.PrimaryBalanceBaseline
    shock_year: Annotated[int, Between(1, 5)]
    shock_type: Literal[1, 2, 3]
    shock_magnitudes: data.ShockMagnitudes


@dataclass(frozen=True, kw_only=True)
class OutputDeltaInputs(_SnapshotInputs):
    """Bound input leaves for `compute_output_delta`."""

    country_name: Literal["Aurelium", "Borvelia", "Litellia"]
    country_initial_debt: data.CountryInitialDebt
    growth_baseline: data.GrowthBaseline
    interest_baseline: data.InterestBaseline
    primary_balance_baseline: data.PrimaryBalanceBaseline
    shock_year: Annotated[int, Between(1, 5)]
    shock_type: Literal[1, 2, 3]
    shock_magnitudes: data.ShockMagnitudes


__all__ = [
    "Model",
    "OutputBaselineInputs",
    "OutputShockedInputs",
    "OutputDeltaInputs",
]
