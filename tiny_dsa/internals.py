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

def debt_to_gdp_shock_difference_percent(ctx, time_period):
    """Return the shocked debt-to-GDP percentage less the baseline debt-to-GDP percentage.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection time period corresponding to Outputs columns B through F.

    Returns:
        Difference between shocked and baseline debt-to-GDP percentages for the time period.

    Note:
        Covers Outputs!B14:F14. Excel: =Engine!{col}20-Engine!{col}6.
"""
    shocked_debt_percent = shocked_debt_to_gdp_percent(ctx, time_period=time_period)
    baseline_debt_percent = baseline_debt_to_gdp_percent(ctx, time_period=time_period)
    return xl_sub(shocked_debt_percent, baseline_debt_percent)

def baseline_debt_to_gdp_percent(ctx, time_period):
    """Return the baseline debt-to-GDP percentage for a projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection period number, where 1 through 5 map to Engine!C6 through Engine!G6.

    Returns:
        The baseline debt-to-GDP percentage for the requested projection period.

    Note:
        Covers Engine!C6:G6. Excel: Engine!C6 =Inputs!B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18; Engine!D6:G6 =Engine!{previous_col}6*(1+Inputs!{col}17/100)/(1+Inputs!{col}16/100)-Inputs!{col}18.
"""
    if time_period == 1:
        projection_inputs_column = 'C'
    elif time_period == 2:
        projection_inputs_column = 'D'
    elif time_period == 3:
        projection_inputs_column = 'E'
    elif time_period == 4:
        projection_inputs_column = 'F'
    elif time_period == 5:
        projection_inputs_column = 'G'
    else:
        raise XlError('Unsupported time_period for Engine!C6:G6')
    if time_period == 1:
        prior_debt_to_gdp_percent = selected_country_initial_debt_to_gdp(ctx)
    else:
        previous_time_period = time_period - 1
        prior_debt_to_gdp_percent = baseline_debt_to_gdp_percent(ctx, time_period=previous_time_period)
    debt_accumulation_rate_percent = xl_cell(ctx, f'Inputs!{projection_inputs_column}17')
    shocked_debt_to_gdp_percent_value = shocked_debt_to_gdp_percent(ctx, time_period=time_period)
    debt_to_gdp_adjustment_percent = xl_cell(ctx, f'Inputs!{projection_inputs_column}18')
    return xl_sub(xl_div(xl_mul(prior_debt_to_gdp_percent, xl_add(1.0, xl_div(debt_accumulation_rate_percent, 100.0))), xl_add(1.0, xl_div(shocked_debt_to_gdp_percent_value, 100.0))), debt_to_gdp_adjustment_percent)

def shocked_debt_to_gdp_percent(ctx, time_period):
    """Return the shocked debt-to-GDP percentage for the projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection period number (1 through 5).

    Returns:
        Shocked debt-to-GDP percentage for the requested projection period.

    Note:
        Covers Engine!C20:G20. Excel: first period =Inputs!B6*(1+(Inputs!C17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!C10)/100)/(1+(Inputs!C16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!C10)/100)-Engine!C16; later periods =Engine!{previous_col}20*(1+(Inputs!{col}17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!{col}10)/100)/(1+(Inputs!{col}16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!{col}10)/100)-Engine!{col}16.
"""
    projection_column_by_time_period = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    projection_column = projection_column_by_time_period[time_period]
    if time_period == 1:
        prior_debt_to_gdp_percent = selected_country_initial_debt_to_gdp(ctx)
    else:
        previous_time_period = time_period - 1
        prior_debt_to_gdp_percent = shocked_debt_to_gdp_percent(ctx, time_period=previous_time_period)
    nominal_interest_rate_percent = xl_cell(ctx, f'Inputs!{projection_column}17')
    nominal_shock_parameter_selector = xl_cell(ctx, 'Inputs!B22')
    if isinstance(nominal_shock_parameter_selector, XlError):
        nominal_interest_shock_adjustment = nominal_shock_parameter_selector
    else:
        nominal_shock_parameter_index = to_int(nominal_shock_parameter_selector)
        if isinstance(nominal_shock_parameter_index, XlError):
            nominal_interest_shock_adjustment = nominal_shock_parameter_index
        elif nominal_shock_parameter_index < 1 or nominal_shock_parameter_index > 3:
            nominal_interest_shock_adjustment = XlError.VALUE
        elif nominal_shock_parameter_index == 1:
            nominal_interest_shock_adjustment = 0.0
        elif nominal_shock_parameter_index == 2:
            nominal_interest_shock_adjustment = selected_shock_magnitude_pp(ctx)
        elif nominal_shock_parameter_index == 3:
            nominal_interest_shock_adjustment = 0.0
        else:
            nominal_interest_shock_adjustment = XlError.VALUE
    nominal_shock_active_indicator = shock_active(ctx, time_period=time_period)
    interest_rate_shocked_factor = xl_add(1.0, xl_div(xl_add(nominal_interest_rate_percent, xl_mul(nominal_interest_shock_adjustment, nominal_shock_active_indicator)), 100.0))
    debt_with_interest_shock = xl_mul(prior_debt_to_gdp_percent, interest_rate_shocked_factor)
    real_gdp_growth_percent = xl_cell(ctx, f'Inputs!{projection_column}16')
    growth_shock_parameter_selector = xl_cell(ctx, 'Inputs!B22')
    if isinstance(growth_shock_parameter_selector, XlError):
        real_gdp_growth_shock_adjustment = growth_shock_parameter_selector
    else:
        growth_shock_parameter_index = to_int(growth_shock_parameter_selector)
        if isinstance(growth_shock_parameter_index, XlError):
            real_gdp_growth_shock_adjustment = growth_shock_parameter_index
        elif growth_shock_parameter_index < 1 or growth_shock_parameter_index > 3:
            real_gdp_growth_shock_adjustment = XlError.VALUE
        elif growth_shock_parameter_index == 1:
            real_gdp_growth_shock_adjustment = selected_shock_magnitude_pp(ctx)
        elif growth_shock_parameter_index == 2:
            real_gdp_growth_shock_adjustment = 0.0
        elif growth_shock_parameter_index == 3:
            real_gdp_growth_shock_adjustment = 0.0
        else:
            real_gdp_growth_shock_adjustment = XlError.VALUE
    growth_shock_active_indicator = shock_active(ctx, time_period=time_period)
    growth_rate_shocked_factor = xl_add(1.0, xl_div(xl_add(real_gdp_growth_percent, xl_mul(real_gdp_growth_shock_adjustment, growth_shock_active_indicator)), 100.0))
    debt_after_growth_shock = xl_div(debt_with_interest_shock, growth_rate_shocked_factor)
    primary_balance_shocked = primary_balance_shocked_percent_gdp(ctx, time_period=time_period)
    return xl_sub(debt_after_growth_shock, primary_balance_shocked)

def primary_balance_shocked_percent_gdp(ctx, time_period):
    """Return the shocked primary balance as a percent of GDP for a projection period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection period number from 1 through 5.

    Returns:
        Workbook value for the primary balance, shocked (% of GDP).

    Note:
        Covers Engine!C16:G16. Excel: =Inputs!{col}18+CHOOSE(Inputs!B22,0,0,Engine!B9)*Engine!{col}10.
"""
    if time_period == 1:
        projection_column = 'C'
    elif time_period == 2:
        projection_column = 'D'
    elif time_period == 3:
        projection_column = 'E'
    elif time_period == 4:
        projection_column = 'F'
    elif time_period == 5:
        projection_column = 'G'
    else:
        return XlError.VALUE
    unshocked_primary_balance_percent_gdp = xl_cell(ctx, 'Inputs!' + projection_column + '18')
    shock_parameter_selector = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_parameter_selector, XlError):
        primary_balance_shock_magnitude_pp = shock_parameter_selector
    else:
        selected_shock_parameter = to_int(shock_parameter_selector)
        if isinstance(selected_shock_parameter, XlError):
            primary_balance_shock_magnitude_pp = selected_shock_parameter
        elif selected_shock_parameter < 1 or selected_shock_parameter > 3:
            primary_balance_shock_magnitude_pp = XlError.VALUE
        elif selected_shock_parameter == 1:
            primary_balance_shock_magnitude_pp = 0.0
        elif selected_shock_parameter == 2:
            primary_balance_shock_magnitude_pp = 0.0
        elif selected_shock_parameter == 3:
            primary_balance_shock_magnitude_pp = selected_shock_magnitude_pp(ctx)
        else:
            primary_balance_shock_magnitude_pp = XlError.VALUE
    shock_active_indicator = shock_active(ctx, time_period=time_period)
    primary_balance_shock_adjustment = xl_mul(primary_balance_shock_magnitude_pp, shock_active_indicator)
    return xl_add(unshocked_primary_balance_percent_gdp, primary_balance_shock_adjustment)

def shock_active(ctx, time_period):
    """Return 1.0 when the projection year is at or after the shock year.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection time period (1 through 5).

    Returns:
        1.0 if the projection year is at or after the shock year, else 0.0, or an XlError if the comparison cannot be evaluated.

    Note:
        Covers Engine!C10, Engine!D10, Engine!E10, Engine!F10, and Engine!G10. Excel: =IF(Engine!{col}5>=Inputs!B21,1,0).
"""
    projection_year_workbook_column_by_time_period = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    projection_year_workbook_column = projection_year_workbook_column_by_time_period[time_period]
    projection_year = xl_cell(ctx, f'Engine!{projection_year_workbook_column}5')
    shock_year = xl_cell(ctx, 'Inputs!B21')
    projection_year_at_or_after_shock_year = xl_ge(projection_year, shock_year)
    shock_active_condition = to_bool(projection_year_at_or_after_shock_year)
    if isinstance(shock_active_condition, XlError):
        return shock_active_condition
    if shock_active_condition:
        return 1.0
    else:
        return 0.0

def selected_country_initial_debt_to_gdp(ctx):
    """Look up the initial debt-to-GDP ratio for the selected country.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Initial debt-to-GDP percentage from the selected country's profile row.

    Note:
        Covers Inputs!B6. Excel: =INDEX(Inputs!A10:Inputs!C12,MATCH(Inputs!B5,Inputs!A10:Inputs!A12,0),2).
"""
    selected_country_name = xl_cell(ctx, 'Inputs!B5')
    country_selector_names = np.array(np.array([[xl_cell(ctx, 'Inputs!A10')], [xl_cell(ctx, 'Inputs!A11')], [xl_cell(ctx, 'Inputs!A12')]], dtype=object), dtype=object)
    selected_country_row = xl_match(selected_country_name, country_selector_names, 0.0)
    initial_debt_to_gdp_reference = xl_index_ref(('Inputs', 10, 1, 12, 3), selected_country_row, 2.0)
    return xl_offset(ctx, initial_debt_to_gdp_reference, 0.0, 0.0)

def selected_shock_magnitude_pp(ctx):
    """Return the shock magnitude in percentage points for the selected shock type.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Shock magnitude in percentage points selected from the shock activation inputs.

    Note:
        Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    shock_type = xl_cell(ctx, 'Inputs!B22')
    shock_type_column_offset = xl_sub(shock_type, 1.0)
    return xl_offset(ctx, ('Inputs', 26, 2), 0.0, shock_type_column_offset, None, None)

# --- Projection public address aliases ---

# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B12': ('baseline_debt_to_gdp_percent', {'time_period': 1}),
    'Outputs!B13': ('shocked_debt_to_gdp_percent', {'time_period': 1}),
    'Outputs!B14': ('debt_to_gdp_shock_difference_percent', {'time_period': 1}),
    'Outputs!C12': ('baseline_debt_to_gdp_percent', {'time_period': 2}),
    'Outputs!C13': ('shocked_debt_to_gdp_percent', {'time_period': 2}),
    'Outputs!C14': ('debt_to_gdp_shock_difference_percent', {'time_period': 2}),
    'Outputs!D12': ('baseline_debt_to_gdp_percent', {'time_period': 3}),
    'Outputs!D13': ('shocked_debt_to_gdp_percent', {'time_period': 3}),
    'Outputs!D14': ('debt_to_gdp_shock_difference_percent', {'time_period': 3}),
    'Outputs!E12': ('baseline_debt_to_gdp_percent', {'time_period': 4}),
    'Outputs!E13': ('shocked_debt_to_gdp_percent', {'time_period': 4}),
    'Outputs!E14': ('debt_to_gdp_shock_difference_percent', {'time_period': 4}),
    'Outputs!F12': ('baseline_debt_to_gdp_percent', {'time_period': 5}),
    'Outputs!F13': ('shocked_debt_to_gdp_percent', {'time_period': 5}),
    'Outputs!F14': ('debt_to_gdp_shock_difference_percent', {'time_period': 5}),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'selected_shock_magnitude_pp',
    'Inputs!B6': 'selected_country_initial_debt_to_gdp',
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
