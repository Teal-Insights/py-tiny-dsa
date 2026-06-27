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

def debt_to_gdp_pct_shock_delta(ctx, time_period):
    """Return the shocked debt-to-GDP percentage less the baseline debt-to-GDP percentage.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection time period, where 1 corresponds to Engine column C and 5 corresponds to Engine column G.

    Returns:
        Difference between shocked and baseline debt-to-GDP percentages for the time period.

    Note:
        Covers Outputs!B14:F14. Excel: =Engine!C20-Engine!C6 through =Engine!G20-Engine!G6.
"""
    shocked_debt_to_gdp_percentage = shocked_debt_to_gdp_pct(ctx, time_period=time_period)
    baseline_debt_to_gdp_percentage = baseline_debt_to_gdp_pct(ctx, time_period=time_period)
    return xl_sub(shocked_debt_to_gdp_percentage, baseline_debt_to_gdp_percentage)

def baseline_debt_to_gdp_pct(ctx, time_period):
    """Return the baseline debt-to-GDP percentage for a projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection period (1 through 5).

    Returns:
        Baseline debt-to-GDP percentage for the requested projection period, or XlError.VALUE for an unsupported period.

    Note:
        Covers Engine!C6:G6. Excel: Engine!C6: =Inputs!B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18; Engine!D6: =Engine!C6*(1+Inputs!D17/100)/(1+Inputs!D16/100)-Inputs!D18; Engine!E6: =Engine!D6*(1+Inputs!E17/100)/(1+Inputs!E16/100)-Inputs!E18; Engine!F6: =Engine!E6*(1+Inputs!F17/100)/(1+Inputs!F16/100)-Inputs!F18; Engine!G6: =Engine!F6*(1+Inputs!G17/100)/(1+Inputs!G16/100)-Inputs!G18.
"""
    period_input_addresses = {1: ('Inputs!C16', 'Inputs!C17', 'Inputs!C18'), 2: ('Inputs!D16', 'Inputs!D17', 'Inputs!D18'), 3: ('Inputs!E16', 'Inputs!E17', 'Inputs!E18'), 4: ('Inputs!F16', 'Inputs!F17', 'Inputs!F18'), 5: ('Inputs!G16', 'Inputs!G17', 'Inputs!G18')}.get(time_period)
    if period_input_addresses is None:
        return XlError.VALUE
    growth_rate_pct_address, interest_rate_pct_address, primary_balance_pct_address = period_input_addresses
    if time_period == 1:
        prior_debt_to_gdp_pct = initial_debt_to_gdp(ctx)
    else:
        prior_debt_to_gdp_pct = baseline_debt_to_gdp_pct(ctx, time_period=time_period - 1)
    growth_rate_pct = xl_cell(ctx, growth_rate_pct_address)
    interest_rate_pct = xl_cell(ctx, interest_rate_pct_address)
    primary_balance_pct = xl_cell(ctx, primary_balance_pct_address)
    return xl_sub(xl_div(xl_mul(prior_debt_to_gdp_pct, xl_add(1.0, xl_div(interest_rate_pct, 100.0))), xl_add(1.0, xl_div(growth_rate_pct, 100.0))), primary_balance_pct)

def shocked_debt_to_gdp_pct(ctx, time_period):
    """Return shocked-path debt-to-GDP for a projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection period number, 1 through 5.

    Returns:
        Shocked-path debt-to-GDP percentage for the requested projection period, or XlError.VALUE when time_period is outside the covered range.

    Note:
        Covers Engine!C20:G20. Excel: Engine!C20 = Inputs!B6*(1+(Inputs!C17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!C10)/100)/(1+(Inputs!C16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!C10)/100)-Engine!C16; Engine!D20:G20 = Engine!{prior_col}20*(1+(Inputs!{col}17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!{col}10)/100)/(1+(Inputs!{col}16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!{col}10)/100)-Engine!{col}16.
"""
    period_rate_addresses = {1: {'interest_rate': 'Inputs!C17', 'growth_rate': 'Inputs!C16'}, 2: {'interest_rate': 'Inputs!D17', 'growth_rate': 'Inputs!D16'}, 3: {'interest_rate': 'Inputs!E17', 'growth_rate': 'Inputs!E16'}, 4: {'interest_rate': 'Inputs!F17', 'growth_rate': 'Inputs!F16'}, 5: {'interest_rate': 'Inputs!G17', 'growth_rate': 'Inputs!G16'}}.get(time_period)
    if period_rate_addresses is None:
        return XlError.VALUE
    if time_period == 1:
        prior_debt_to_gdp = initial_debt_to_gdp(ctx)
    else:
        prior_debt_to_gdp = shocked_debt_to_gdp_pct(ctx, time_period=time_period - 1)
    baseline_interest_rate_pct = xl_cell(ctx, period_rate_addresses['interest_rate'])
    interest_shock_parameter_raw = xl_cell(ctx, 'Inputs!B22')
    interest_shock_parameter_index = interest_shock_parameter_raw if isinstance(interest_shock_parameter_raw, XlError) else to_int(interest_shock_parameter_raw)
    interest_shock_selection = interest_shock_parameter_index if isinstance(interest_shock_parameter_index, XlError) else XlError.VALUE if interest_shock_parameter_index < 1 or interest_shock_parameter_index > 3 else {1: 'no_shock', 2: 'selected_shock', 3: 'no_shock'}.get(interest_shock_parameter_index, XlError.VALUE)
    interest_rate_shock_magnitude_pp = interest_shock_selection if isinstance(interest_shock_selection, XlError) else selected_shock_magnitude_pp(ctx) if interest_shock_selection == 'selected_shock' else 0.0 if interest_shock_selection == 'no_shock' else XlError.VALUE
    interest_rate_factor = xl_add(1.0, xl_div(xl_add(baseline_interest_rate_pct, xl_mul(interest_rate_shock_magnitude_pp, shock_active(ctx, time_period=time_period))), 100.0))
    debt_after_interest_factor = xl_mul(prior_debt_to_gdp, interest_rate_factor)
    baseline_growth_rate_pct = xl_cell(ctx, period_rate_addresses['growth_rate'])
    growth_shock_parameter_raw = xl_cell(ctx, 'Inputs!B22')
    growth_shock_parameter_index = growth_shock_parameter_raw if isinstance(growth_shock_parameter_raw, XlError) else to_int(growth_shock_parameter_raw)
    growth_shock_selection = growth_shock_parameter_index if isinstance(growth_shock_parameter_index, XlError) else XlError.VALUE if growth_shock_parameter_index < 1 or growth_shock_parameter_index > 3 else {1: 'selected_shock', 2: 'no_shock', 3: 'no_shock'}.get(growth_shock_parameter_index, XlError.VALUE)
    growth_rate_shock_magnitude_pp = growth_shock_selection if isinstance(growth_shock_selection, XlError) else selected_shock_magnitude_pp(ctx) if growth_shock_selection == 'selected_shock' else 0.0 if growth_shock_selection == 'no_shock' else XlError.VALUE
    growth_rate_factor = xl_add(1.0, xl_div(xl_add(baseline_growth_rate_pct, xl_mul(growth_rate_shock_magnitude_pp, shock_active(ctx, time_period=time_period))), 100.0))
    debt_to_gdp_before_primary_balance = xl_div(debt_after_interest_factor, growth_rate_factor)
    primary_balance_shocked_pct = primary_balance_shocked_pct_gdp(ctx, time_period=time_period)
    return xl_sub(debt_to_gdp_before_primary_balance, primary_balance_shocked_pct)

def primary_balance_shocked_pct_gdp(ctx, time_period):
    """Return the shocked primary balance percentage of GDP for a projection time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Projection time period (1 through 5).

Returns:
    The result of adding the baseline primary balance to the CHOOSE-selected shock magnitude multiplied by the shock-active indicator.

Note:
    Covers Engine!C16:G16. Excel: =Inputs!{col}18+CHOOSE(Inputs!B22,0,0,Engine!B9)*Engine!{col}10.
"""
    primary_balance_input_address = {1: 'Inputs!C18', 2: 'Inputs!D18', 3: 'Inputs!E18', 4: 'Inputs!F18', 5: 'Inputs!G18'}.get(time_period)
    if primary_balance_input_address is None:
        return XlError.VALUE
    baseline_primary_balance_pct_gdp = xl_cell(ctx, primary_balance_input_address)
    shock_selection_raw = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_selection_raw, XlError):
        selected_primary_balance_shock_pp = shock_selection_raw
    else:
        shock_selection_index = to_int(shock_selection_raw)
        if isinstance(shock_selection_index, XlError):
            selected_primary_balance_shock_pp = shock_selection_index
        else:
            selected_primary_balance_shock_pp = XlError.VALUE if shock_selection_index < 1 or shock_selection_index > 3 else 0.0 if shock_selection_index == 1 else 0.0 if shock_selection_index == 2 else selected_shock_magnitude_pp(ctx) if shock_selection_index == 3 else XlError.VALUE
    active_shock_indicator = shock_active(ctx, time_period=time_period)
    return xl_add(baseline_primary_balance_pct_gdp, xl_mul(selected_primary_balance_shock_pp, active_shock_indicator))

def shock_active(ctx, time_period):
    """Return 1.0 when the shock is active for the given time period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection time period, 1 through 5.

    Returns:
        1.0 if the projection year is at or after the shock year, 0.0 if it is not, or an XlError if the time period is unsupported or Excel comparison/coercion returns an error.

    Note:
        Covers Engine!C10:G10. Excel: =IF(Engine!{column}5>=Inputs!B21,1,0).
"""
    projection_year_address_by_time_period = {1: 'Engine!C5', 2: 'Engine!D5', 3: 'Engine!E5', 4: 'Engine!F5', 5: 'Engine!G5'}
    projection_year_address = projection_year_address_by_time_period.get(time_period)
    if projection_year_address is None:
        return XlError.VALUE
    projection_year = xl_cell(ctx, projection_year_address)
    shock_year = xl_cell(ctx, 'Inputs!B21')
    shock_activation_comparison = xl_ge(projection_year, shock_year)
    shock_activation_condition = to_bool(shock_activation_comparison)
    return shock_activation_condition if isinstance(shock_activation_condition, XlError) else 1.0 if shock_activation_condition else 0.0

def initial_debt_to_gdp(ctx):
    """Look up the initial debt-to-GDP ratio for the selected country.

Args:
    ctx: Workbook evaluation context.

Returns:
    Initial debt-to-GDP ratio from the country selector table.

Note:
    Covers Inputs!B6. Excel: =INDEX(Inputs!A10:Inputs!C12,MATCH(Inputs!B5,Inputs!A10:Inputs!A12,0),2).
"""
    selected_country = xl_cell(ctx, 'Inputs!B5')
    country_selector_labels = np.array(np.array([[xl_cell(ctx, 'Inputs!A10')], [xl_cell(ctx, 'Inputs!A11')], [xl_cell(ctx, 'Inputs!A12')]], dtype=object), dtype=object)
    selected_country_row_index = xl_match(selected_country, country_selector_labels, 0.0)
    initial_debt_to_gdp_reference = xl_index_ref(('Inputs', 10, 1, 12, 3), selected_country_row_index, 2.0)
    return xl_offset(ctx, initial_debt_to_gdp_reference, 0.0, 0.0)

def selected_shock_magnitude_pp(ctx):
    """Look up the shock magnitude in percentage points for the selected shock type.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Shock magnitude in percentage points from the shock activation inputs.

    Note:
        Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    shock_type_selection = xl_cell(ctx, 'Inputs!B22')
    shock_magnitude_column_offset = xl_sub(shock_type_selection, 1.0)
    return xl_offset(ctx, ('Inputs', 26, 2), 0.0, shock_magnitude_column_offset, None, None)

# --- Projection public address aliases ---

# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B12': ('baseline_debt_to_gdp_pct', {'time_period': 1}),
    'Outputs!B13': ('shocked_debt_to_gdp_pct', {'time_period': 1}),
    'Outputs!B14': ('debt_to_gdp_pct_shock_delta', {'time_period': 1}),
    'Outputs!C12': ('baseline_debt_to_gdp_pct', {'time_period': 2}),
    'Outputs!C13': ('shocked_debt_to_gdp_pct', {'time_period': 2}),
    'Outputs!C14': ('debt_to_gdp_pct_shock_delta', {'time_period': 2}),
    'Outputs!D12': ('baseline_debt_to_gdp_pct', {'time_period': 3}),
    'Outputs!D13': ('shocked_debt_to_gdp_pct', {'time_period': 3}),
    'Outputs!D14': ('debt_to_gdp_pct_shock_delta', {'time_period': 3}),
    'Outputs!E12': ('baseline_debt_to_gdp_pct', {'time_period': 4}),
    'Outputs!E13': ('shocked_debt_to_gdp_pct', {'time_period': 4}),
    'Outputs!E14': ('debt_to_gdp_pct_shock_delta', {'time_period': 4}),
    'Outputs!F12': ('baseline_debt_to_gdp_pct', {'time_period': 5}),
    'Outputs!F13': ('shocked_debt_to_gdp_pct', {'time_period': 5}),
    'Outputs!F14': ('debt_to_gdp_pct_shock_delta', {'time_period': 5}),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'selected_shock_magnitude_pp',
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
