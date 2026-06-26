from __future__ import annotations

from .runtime import (
    XlError,
    np,
    to_bool,
    to_int,
    xl_add,
    xl_cell,
    xl_div,
    xl_eval,
    xl_ge,
    xl_index_ref,
    xl_match,
    xl_mul,
    xl_offset,
    xl_sub,
)

# --- Formula cell functions ---

def debt_to_gdp_shock_impact(ctx, time_period):
    """Return the shock impact on Debt-to-GDP as the difference between shocked and baseline ratios.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection year index (1-based, mapping to engine column letter).

    Returns:
        The difference between shocked and baseline debt-to-GDP ratios.

    Note:
        Covers Outputs!B14:F14. Excel: =Engine!{col}20-Engine!{col}6 where col corresponds to time_period.
"""
    shocked = shocked_debt_to_gdp(ctx, time_period=time_period)
    baseline = baseline_debt_to_gdp(ctx, time_period=time_period)
    return xl_sub(shocked, baseline)

def baseline_debt_to_gdp(ctx, time_period):
    """Return the baseline debt-to-GDP ratio for the given projection time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Integer projection year (1 to 5).

Returns:
    Debt-to-GDP percentage for the given time period.

Note:
    Covers Engine!C6:G6.
    Excel: For column C: =Inputs!B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18.
    For columns D-G: =Engine!{prev_col}6*(1+Inputs!{col}17/100)/(1+Inputs!{col}16/100)-Inputs!{col}18.
"""
    if time_period == 1:
        prior_debt = initial_debt_to_gdp(ctx)
    else:
        prior_debt = baseline_debt_to_gdp(ctx, time_period - 1)
    column_letter = chr(ord('C') + time_period - 1)
    growth_rate_div100 = xl_div(xl_cell(ctx, f'Inputs!{column_letter}17'), 100.0)
    interest_rate_div100 = xl_div(xl_cell(ctx, f'Inputs!{column_letter}16'), 100.0)
    primary_balance = xl_cell(ctx, f'Inputs!{column_letter}18')
    debt_to_gdp = xl_sub(xl_mul(prior_debt, xl_div(xl_add(1.0, growth_rate_div100), xl_add(1.0, interest_rate_div100))), primary_balance)
    return debt_to_gdp

def shocked_debt_to_gdp(ctx, time_period):
    """Compute debt-to-GDP under the shocked path for a given projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Integer projection period (1=first year, column C).

    Returns:
        Debt-to-GDP percentage for the period, accounting for shock choices.

    Note:
        Covers Engine!C20:G20. Excel template: For first period, =Inputs!B6*(1+(Inputs!{col}17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!{col}10)/100)/(1+(Inputs!{col}16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!{col}10)/100)-Engine!{col}16; for subsequent periods, prior period's debt replaces Inputs!B6.
"""
    if time_period == 1:
        prior_debt = initial_debt_to_gdp(ctx)
    else:
        prior_debt = shocked_debt_to_gdp(ctx, time_period - 1)
    shock_param_raw = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_param_raw, XlError):
        return shock_param_raw
    shock_param = to_int(shock_param_raw)
    if isinstance(shock_param, XlError):
        return shock_param
    if shock_param < 1 or shock_param > 3:
        return XlError.VALUE
    if shock_param == 1:
        numerator_shock = 0.0
        denominator_shock = shock_magnitude(ctx)
    elif shock_param == 2:
        numerator_shock = shock_magnitude(ctx)
        denominator_shock = 0.0
    else:
        numerator_shock = 0.0
        denominator_shock = 0.0
    shock_active_value = shock_active(ctx, time_period=time_period)
    col_letter = chr(ord('C') + time_period - 1)
    interest_rate = xl_cell(ctx, f'Inputs!{col_letter}17')
    growth_rate = xl_cell(ctx, f'Inputs!{col_letter}16')
    numerator = xl_add(1.0, xl_div(xl_add(interest_rate, xl_mul(numerator_shock, shock_active_value)), 100.0))
    denominator = xl_add(1.0, xl_div(xl_add(growth_rate, xl_mul(denominator_shock, shock_active_value)), 100.0))
    debt = xl_mul(prior_debt, xl_div(numerator, denominator))
    return xl_sub(debt, primary_balance_shocked(ctx, time_period=time_period))

def primary_balance_shocked(ctx, time_period):
    """Return the shocked primary balance as a percentage of GDP for a given projection year.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection time period (1-5).

    Returns:
        Shocked primary balance (% of GDP).

    Note:
        Covers Engine!C16:G16. Excel: =Inputs!{col}18+CHOOSE(Inputs!$B$22,0,0,$B$9)*{col}10
"""
    col = chr(ord('C') + time_period - 1)
    base_balance = xl_cell(ctx, f'Inputs!{col}18')
    shock_active_val = shock_active(ctx, time_period=time_period)
    shock_choice_raw = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_choice_raw, XlError):
        shock_increment = shock_choice_raw
    else:
        shock_choice = to_int(shock_choice_raw)
        if isinstance(shock_choice, XlError):
            shock_increment = shock_choice
        elif shock_choice < 1 or shock_choice > 3:
            shock_increment = XlError.VALUE
        elif shock_choice == 3:
            shock_increment = shock_magnitude(ctx)
        else:
            shock_increment = 0.0
    shock_term = xl_mul(shock_increment, shock_active_val)
    if isinstance(shock_term, XlError):
        return shock_term
    return xl_add(base_balance, shock_term)

def shock_active(ctx, time_period):
    """Return 1.0 if the shock is active for the given time period, else 0.0.

    Args:
        ctx: Workbook evaluation context.
        time_period: Time period index (1-based) corresponding to projection columns.

    Returns:
        1.0 if the year for the column is >= the shock year, else 0.0.

    Note:
        Covers Engine!C10:G10. Excel: =IF(Engine!{col}5>=Inputs!$B$21,1,0).
"""
    col = chr(ord('C') + time_period - 1)
    year_value = xl_cell(ctx, f'Engine!{col}5')
    shock_year = xl_cell(ctx, 'Inputs!B21')
    is_active_bool = to_bool(xl_ge(year_value, shock_year))
    if isinstance(is_active_bool, XlError):
        return is_active_bool
    elif is_active_bool:
        return 1.0
    else:
        return 0.0

def initial_debt_to_gdp(ctx):
    """Look up the initial debt-to-GDP ratio for the selected country.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Initial debt-to-GDP ratio from the country profile table.

    Note:
        Covers Inputs!B6. Excel: =INDEX($A$10:$C$12,MATCH($B$5,$A$10:$A$12,0),2).
"""
    country_selector = xl_cell(ctx, 'Inputs!B5')
    country_labels = np.array([[xl_cell(ctx, 'Inputs!A10')], [xl_cell(ctx, 'Inputs!A11')], [xl_cell(ctx, 'Inputs!A12')]], dtype=object)
    lookup_column = np.array(country_labels, dtype=object)
    matched_row = xl_match(country_selector, lookup_column, 0.0)
    profile_table = ('Inputs', 10, 1, 12, 3)
    initial_value = xl_index_ref(profile_table, matched_row, 2.0)
    return xl_offset(ctx, initial_value, 0.0, 0.0)

def shock_magnitude(ctx):
    """Retrieve the shock magnitude (in percentage points) for the active shock type.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Shock magnitude value from the corresponding column in the shock magnitude table.

    Note:
        Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    shock_type = xl_cell(ctx, 'Inputs!B22')
    column_offset = xl_sub(shock_type, 1.0)
    return xl_offset(ctx, ('Inputs', 26, 2), 0.0, column_offset, None, None)

# --- Projection public address aliases ---

# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B12': ('baseline_debt_to_gdp', {'time_period': 1}),
    'Outputs!B13': ('shocked_debt_to_gdp', {'time_period': 1}),
    'Outputs!B14': ('debt_to_gdp_shock_impact', {'time_period': 1}),
    'Outputs!C12': ('baseline_debt_to_gdp', {'time_period': 2}),
    'Outputs!C13': ('shocked_debt_to_gdp', {'time_period': 2}),
    'Outputs!C14': ('debt_to_gdp_shock_impact', {'time_period': 2}),
    'Outputs!D12': ('baseline_debt_to_gdp', {'time_period': 3}),
    'Outputs!D13': ('shocked_debt_to_gdp', {'time_period': 3}),
    'Outputs!D14': ('debt_to_gdp_shock_impact', {'time_period': 3}),
    'Outputs!E12': ('baseline_debt_to_gdp', {'time_period': 4}),
    'Outputs!E13': ('shocked_debt_to_gdp', {'time_period': 4}),
    'Outputs!E14': ('debt_to_gdp_shock_impact', {'time_period': 4}),
    'Outputs!F12': ('baseline_debt_to_gdp', {'time_period': 5}),
    'Outputs!F13': ('shocked_debt_to_gdp', {'time_period': 5}),
    'Outputs!F14': ('debt_to_gdp_shock_impact', {'time_period': 5}),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'shock_magnitude',
    'Inputs!B6': 'initial_debt_to_gdp',
}

def _address_to_func_name(address):
    name = []
    prev_underscore = False
    for ch in address.lower():
        if ch == "'":
            continue
        if "a" <= ch <= "z" or "0" <= ch <= "9":
            name.append(ch)
            prev_underscore = False
        else:
            if not prev_underscore:
                name.append("_")
                prev_underscore = True
    base = "".join(name).strip("_")
    return f"cell_{base}"

def _resolve_formula(address):
    fn = _RESOLVED_FORMULAS.get(address)
    if fn is not None:
        return fn
    dispatch = _ADDRESS_DISPATCH.get(address)
    if dispatch is not None:
        helper_name, key_kwargs = dispatch
        helper = globals()[helper_name]

        def _bound(ctx, _helper=helper, _key_kwargs=key_kwargs):
            return _helper(ctx, **_key_kwargs)

        _RESOLVED_FORMULAS[address] = _bound
        return _bound
    symbol_name = _SYMBOL_DISPATCH.get(address)
    if symbol_name is not None:
        helper = globals()[symbol_name]

        def _bound(ctx, _helper=helper):
            return _helper(ctx)

        _RESOLVED_FORMULAS[address] = _bound
        return _bound
    name = _address_to_func_name(address)
    fn = globals().get(name)
    if fn is not None:
        _RESOLVED_FORMULAS[address] = fn
    return fn
