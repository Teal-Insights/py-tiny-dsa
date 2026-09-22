"""Input schema, domain, and value-map checks for bound Model arguments."""

from __future__ import annotations

from typing import Annotated, Literal

from . import data
from .excel import coerce_input_measure
from .runtime import Between, RealBetween, require_annotated_domain


def _check_country_name(
    country_name: Literal["Aurelium", "Borvelia", "Litellia"],
) -> Literal["Aurelium", "Borvelia", "Litellia"]:
    """Validate `country_name` before the model reads it."""
    country_name = coerce_input_measure(country_name, dtype="string", series_id="country_name")
    require_annotated_domain(
        country_name,
        Literal["Aurelium", "Borvelia", "Litellia"],
        series_id="country_name",
    )
    return country_name


def _check_country_initial_debt(
    country_initial_debt: data.CountryInitialDebt,
) -> data.CountryInitialDebt:
    """Validate `country_initial_debt` before the model reads it."""
    data.COUNTRY_INITIAL_DEBT.schema.validate(country_initial_debt)
    country_initial_debt = coerce_input_measure(country_initial_debt, dtype="float", series_id="country_initial_debt")
    for coordinate in data.COUNTRY_INITIAL_DEBT.required:
        require_annotated_domain(
            country_initial_debt[coordinate],
            Annotated[float, RealBetween(0.0, 200.0)],
            series_id="country_initial_debt" + repr(coordinate),
        )
    return country_initial_debt


def _check_growth_baseline(growth_baseline: data.GrowthBaseline) -> data.GrowthBaseline:
    """Validate `growth_baseline` before the model reads it."""
    data.GROWTH_BASELINE.schema.validate(growth_baseline)
    growth_baseline = coerce_input_measure(growth_baseline, dtype="float", series_id="growth_baseline")
    for coordinate in data.GROWTH_BASELINE.required:
        require_annotated_domain(
            growth_baseline[coordinate],
            Annotated[float, RealBetween(-10.0, 15.0)],
            series_id="growth_baseline" + repr(coordinate),
        )
    return growth_baseline


def _check_interest_baseline(interest_baseline: data.InterestBaseline) -> data.InterestBaseline:
    """Validate `interest_baseline` before the model reads it."""
    data.INTEREST_BASELINE.schema.validate(interest_baseline)
    interest_baseline = coerce_input_measure(interest_baseline, dtype="float", series_id="interest_baseline")
    for coordinate in data.INTEREST_BASELINE.required:
        require_annotated_domain(
            interest_baseline[coordinate],
            Annotated[float, RealBetween(0.0, 20.0)],
            series_id="interest_baseline" + repr(coordinate),
        )
    return interest_baseline


def _check_primary_balance_baseline(
    primary_balance_baseline: data.PrimaryBalanceBaseline,
) -> data.PrimaryBalanceBaseline:
    """Validate `primary_balance_baseline` before the model reads it."""
    data.PRIMARY_BALANCE_BASELINE.schema.validate(primary_balance_baseline)
    primary_balance_baseline = coerce_input_measure(primary_balance_baseline, dtype="float", series_id="primary_balance_baseline")
    for coordinate in data.PRIMARY_BALANCE_BASELINE.required:
        require_annotated_domain(
            primary_balance_baseline[coordinate],
            Annotated[float, RealBetween(-15.0, 15.0)],
            series_id="primary_balance_baseline" + repr(coordinate),
        )
    return primary_balance_baseline


def _check_shock_year(shock_year: Annotated[int, Between(1, 5)]) -> Annotated[int, Between(1, 5)]:
    """Validate `shock_year` before the model reads it."""
    shock_year = coerce_input_measure(shock_year, dtype="int", series_id="shock_year")
    require_annotated_domain(shock_year, Annotated[int, Between(1, 5)], series_id="shock_year")
    return shock_year


def _check_shock_type(shock_type: Literal[1, 2, 3]) -> Literal[1, 2, 3]:
    """Validate `shock_type` before the model reads it."""
    shock_type = coerce_input_measure(shock_type, dtype="int", series_id="shock_type")
    require_annotated_domain(shock_type, Literal[1, 2, 3], series_id="shock_type")
    return shock_type


def _check_shock_magnitudes(shock_magnitudes: data.ShockMagnitudes) -> data.ShockMagnitudes:
    """Validate `shock_magnitudes` before the model reads it."""
    data.SHOCK_MAGNITUDES.schema.validate(shock_magnitudes)
    shock_magnitudes = coerce_input_measure(shock_magnitudes, dtype="float", series_id="shock_magnitudes")
    for coordinate in data.SHOCK_MAGNITUDES.required:
        require_annotated_domain(
            shock_magnitudes[coordinate],
            Annotated[float, RealBetween(-30.0, 30.0)],
            series_id="shock_magnitudes" + repr(coordinate),
        )
    return shock_magnitudes


CHECKS = {
    "country_name": _check_country_name,
    "country_initial_debt": _check_country_initial_debt,
    "growth_baseline": _check_growth_baseline,
    "interest_baseline": _check_interest_baseline,
    "primary_balance_baseline": _check_primary_balance_baseline,
    "shock_year": _check_shock_year,
    "shock_type": _check_shock_type,
    "shock_magnitudes": _check_shock_magnitudes,
}
