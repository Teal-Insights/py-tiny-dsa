from __future__ import annotations

from .runtime import EvalContext, XlError, xl_bool, xl_cell, xl_compare, xl_eval, xl_index_ref, xl_int, xl_match, xl_number, xl_offset, xl_raise, xl_range

# --- Formula cell functions ---

def shocked_vs_projected_debt_to_gdp(ctx: EvalContext, time_period: int) -> float:
    """Return the difference between the shocked and projected debt-to-GDP ratio for a given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    Shocked debt-to-GDP minus projected debt-to-GDP, as a float.

Note:
    Covers Outputs!B14:F14. Excel: =Engine!C20-Engine!C6.
"""
    if time_period == 1:
        shocked = xl_number(shocked_debt_to_gdp(ctx))
        projected = xl_number(projected_debt_to_gdp(ctx))
    else:
        shocked = xl_number(shocked_debt_ratio_step(ctx, time_period=time_period))
        projected = xl_number(debt_to_gdp_forecast(ctx, time_period=time_period))
    return shocked - projected

def debt_to_gdp_forecast(ctx: EvalContext, time_period: int) -> float:
    """Compute the projected debt-to-GDP ratio for a given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Index of the time period (2 through 5).

Returns:
    Projected debt-to-GDP ratio for the period.

Note:
    Covers Engine!D6:G6. Excel: =Engine!C6*(1+Inputs!D17/100)/(1+Inputs!D16/100)-Inputs!D18.
"""
    column_map = {2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    col = column_map[time_period]
    if time_period == 2:
        prev = xl_number(projected_debt_to_gdp(ctx))
    else:
        prev = xl_number(debt_to_gdp_forecast(ctx, time_period - 1))
    growth = xl_number(xl_cell(ctx, f'Inputs!{col}17')) / 100.0
    interest = xl_number(xl_cell(ctx, f'Inputs!{col}16')) / 100.0
    primary = xl_number(xl_cell(ctx, f'Inputs!{col}18'))
    num = prev * (1.0 + growth)
    den = 1.0 + interest
    div = num / den if den != 0 else xl_raise(XlError.DIV)
    return xl_number(div) - primary

def shocked_debt_ratio_step(ctx: EvalContext, time_period: int) -> float:
    """Compute the debt-to-GDP ratio for a given time period under a shock scenario.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (2 through 5).

Returns:
    The shocked debt-to-GDP ratio as a percentage of GDP.

Note:
    Covers Engine!D20:G20. Excel: =Engine!C20*(1+(Inputs!D17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!D10)/100)/(1+(Inputs!D16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!D10)/100)-Engine!D16.
"""
    column_by_period = {2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    col = column_by_period[time_period]
    if time_period == 2:
        prev_value = shocked_debt_to_gdp(ctx)
    else:
        prev_value = shocked_debt_ratio_step(ctx, time_period - 1)
    scenario_idx = xl_int(xl_cell(ctx, 'Inputs!B22'))
    if scenario_idx < 1 or scenario_idx > 3:
        xl_raise(XlError.VALUE)
    shock_mag = shock_magnitude_resolved(ctx)
    num_coeff = (0.0, shock_mag, 0.0)[scenario_idx - 1]
    den_coeff = (shock_mag, 0.0, 0.0)[scenario_idx - 1]
    shock_mult = shock_active(ctx, time_period=time_period)
    num_rate = xl_number(xl_cell(ctx, f'Inputs!{col}17'))
    den_rate = xl_number(xl_cell(ctx, f'Inputs!{col}16'))
    num_term = (num_rate + num_coeff * shock_mult) / 100.0
    den_term = (den_rate + den_coeff * shock_mult) / 100.0
    denominator = 1.0 + den_term
    if denominator == 0.0:
        xl_raise(XlError.DIV)
    intermediate = prev_value * (1.0 + num_term) / denominator
    return xl_number(intermediate - shocked_primary_balance(ctx, time_period=time_period))

def shocked_primary_balance(ctx: EvalContext, time_period: int) -> float:
    """Return the shocked primary balance as percentage of GDP for a given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    The calculated shocked primary balance value.

Note:
    Covers Engine!C16:G16. Excel: =Inputs!C18+CHOOSE(Inputs!B22,0,0,Engine!B9)*Engine!C10.
"""
    column_by_period = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    col = column_by_period[time_period]
    base_value = xl_number(xl_cell(ctx, f'Inputs!{col}18'))
    shock_option = xl_int(xl_cell(ctx, 'Inputs!B22'))
    shock_amount_map = {1: 0.0, 2: 0.0, 3: shock_magnitude_resolved(ctx)}
    shock_amount = shock_amount_map.get(shock_option)
    if shock_amount is None:
        xl_raise(XlError.VALUE)
    shock_active_value = shock_active(ctx, time_period=time_period)
    return base_value + shock_amount * shock_active_value

def shock_active(ctx: EvalContext, time_period: int) -> float:
    """Return 1.0 when the engine value for the given time period meets or exceeds the shock threshold.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    1.0 if the engine value is at or above the threshold, else 0.0.

Note:
    Covers Engine!C10:G10. Excel: =IF(Engine!C5>=Inputs!B21,1,0).
"""
    column_by_period = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    column = column_by_period[time_period]
    engine_value = xl_cell(ctx, f'Engine!{column}5')
    threshold = xl_cell(ctx, 'Inputs!B21')
    met_threshold = xl_bool(xl_compare('>=', engine_value, threshold))
    return 1.0 if met_threshold else 0.0

def initial_debt_to_gdp(ctx: EvalContext) -> float:
    """Initial debt-to-GDP ratio, looked up from a table based on a matched value.

Args:
    ctx: Workbook evaluation context.

Returns:
    The initial debt-to-GDP ratio.

Note:
    Covers Inputs!B6. Excel: =INDEX(Inputs!A10:Inputs!C12,MATCH(Inputs!B5,Inputs!A10:Inputs!A12,0),2).
"""
    lookup_value = xl_cell(ctx, 'Inputs!B5')
    lookup_array = xl_range(ctx, 'Inputs!A10:Inputs!A12')
    matched_row = xl_match(lookup_value, lookup_array, 0.0)
    index_ref = xl_index_ref(('Inputs', 10, 1, 12, 3), matched_row, 2.0)
    return xl_offset(ctx, index_ref, 0.0, 0.0)

def projected_debt_to_gdp(ctx: EvalContext) -> float:
    """Compute projected debt-to-GDP ratio based on the debt dynamics equation: initial debt * (1 + growth rate) / (1 + nominal interest rate) minus primary balance.

Args:
    ctx: Workbook evaluation context.

Returns:
    Projected debt-to-GDP ratio for the next period.

Note:
    Covers Engine!C6. Excel: =Inputs!B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18.
"""
    initial_debt = xl_number(initial_debt_to_gdp(ctx))
    growth_rate_pct = xl_number(xl_cell(ctx, 'Inputs!C17'))
    interest_rate_pct = xl_number(xl_cell(ctx, 'Inputs!C16'))
    primary_balance_pct_gdp = xl_number(xl_cell(ctx, 'Inputs!C18'))
    growth_factor = 1.0 + growth_rate_pct / 100.0
    interest_factor = 1.0 + interest_rate_pct / 100.0
    if interest_factor == 0.0:
        xl_raise(XlError.DIV)
    return initial_debt * growth_factor / interest_factor - primary_balance_pct_gdp

def shock_magnitude_resolved(ctx: EvalContext) -> float:
    """Resolve the shock magnitude for the current scenario period by offsetting from the base shock magnitude range.

Args:
    ctx: Workbook evaluation context.

Returns:
    Shock magnitude in percentage points (PP).

Note:
    Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    period_index = xl_number(xl_cell(ctx, 'Inputs!B22')) - 1.0
    magnitude_cell = xl_offset(ctx, ('Inputs', 26, 2), 0.0, period_index, None, None)
    return xl_number(magnitude_cell)

def shocked_debt_to_gdp(ctx: EvalContext) -> float:
    """Compute the debt-to-GDP ratio for time period 1 under a shock scenario.

Args:
    ctx: Workbook evaluation context.

Returns:
    The shocked debt-to-GDP ratio as a percentage of GDP.

Note:
    Covers Engine!C20. Excel: =Inputs!B6*(1+(Inputs!C17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!C10)/100)/(1+(Inputs!C16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!C10)/100)-Engine!C16.
"""
    initial_debt = xl_number(initial_debt_to_gdp(ctx))
    base_interest = xl_number(xl_cell(ctx, 'Inputs!C17'))
    base_growth = xl_number(xl_cell(ctx, 'Inputs!C16'))
    shock_index = xl_int(xl_cell(ctx, 'Inputs!B22'))
    if shock_index < 1 or shock_index > 3:
        xl_raise(XlError.VALUE)
    if shock_index == 1:
        interest_shock = 0.0
        growth_shock = shock_magnitude_resolved(ctx)
    elif shock_index == 2:
        interest_shock = shock_magnitude_resolved(ctx)
        growth_shock = 0.0
    else:
        interest_shock = 0.0
        growth_shock = 0.0
    shock_active_status = xl_number(shock_active(ctx, time_period=1))
    adjusted_interest = base_interest + interest_shock * shock_active_status
    adjusted_growth = base_growth + growth_shock * shock_active_status
    numerator = initial_debt * (1.0 + adjusted_interest / 100.0)
    denominator = 1.0 + adjusted_growth / 100.0
    if denominator == 0.0:
        xl_raise(XlError.DIV)
    ratio = numerator / denominator
    primary_balance = xl_number(shocked_primary_balance(ctx, time_period=1))
    return ratio - primary_balance

# --- Projection public address aliases ---

def cell_engine_b20(ctx):
    return initial_debt_to_gdp(ctx)

def cell_engine_b6(ctx):
    return initial_debt_to_gdp(ctx)

def cell_outputs_b12(ctx):
    return projected_debt_to_gdp(ctx)

def cell_outputs_b13(ctx):
    return shocked_debt_to_gdp(ctx)

# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B14': ('shocked_vs_projected_debt_to_gdp', {'time_period': 1}),
    'Outputs!C12': ('debt_to_gdp_forecast', {'time_period': 2}),
    'Outputs!C13': ('shocked_debt_ratio_step', {'time_period': 2}),
    'Outputs!C14': ('shocked_vs_projected_debt_to_gdp', {'time_period': 2}),
    'Outputs!D12': ('debt_to_gdp_forecast', {'time_period': 3}),
    'Outputs!D13': ('shocked_debt_ratio_step', {'time_period': 3}),
    'Outputs!D14': ('shocked_vs_projected_debt_to_gdp', {'time_period': 3}),
    'Outputs!E12': ('debt_to_gdp_forecast', {'time_period': 4}),
    'Outputs!E13': ('shocked_debt_ratio_step', {'time_period': 4}),
    'Outputs!E14': ('shocked_vs_projected_debt_to_gdp', {'time_period': 4}),
    'Outputs!F12': ('debt_to_gdp_forecast', {'time_period': 5}),
    'Outputs!F13': ('shocked_debt_ratio_step', {'time_period': 5}),
    'Outputs!F14': ('shocked_vs_projected_debt_to_gdp', {'time_period': 5}),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'shock_magnitude_resolved',
    'Engine!C20': 'shocked_debt_to_gdp',
    'Engine!C6': 'projected_debt_to_gdp',
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
