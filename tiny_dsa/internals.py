from __future__ import annotations

from .runtime import EvalContext, XlError, xl_bool, xl_cell, xl_compare, xl_eval, xl_index_ref, xl_int, xl_match, xl_number, xl_offset, xl_raise, xl_range

# --- Formula cell functions ---

def shock_impact_on_debt_to_gdp(ctx: EvalContext, time_period: int) -> float:
    """Compute the impact of the shock on the debt-to-GDP ratio for a given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    Difference between the shocked and baseline/projected debt-to-GDP ratio, in percentage points.

Note:
    Covers Outputs!B14:F14. Excel: =Engine!C20-Engine!C6.
"""
    if time_period == 1:
        shocked = shocked_debt_to_gdp_period_1(ctx)
        baseline = baseline_debt_to_gdp(ctx)
    else:
        shocked = compute_shocked_debt_to_gdp(ctx, time_period=time_period)
        baseline = project_debt_to_gdp(ctx, time_period=time_period)
    return xl_number(shocked) - xl_number(baseline)

def project_debt_to_gdp(ctx: EvalContext, time_period: int) -> float:
    """Compute the projected debt-to-GDP ratio for a given time period using standard debt dynamics.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (2 to 5).

Returns:
    Projected debt-to-GDP ratio as a float.

Note:
    Covers Engine!D6:G6. Excel: =Engine!C6*(1+Inputs!D17/100)/(1+Inputs!D16/100)-Inputs!D18.
"""
    column_by_period = {2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    col = column_by_period[time_period]
    if time_period == 2:
        previous = baseline_debt_to_gdp(ctx)
    else:
        previous = project_debt_to_gdp(ctx, time_period - 1)
    growth_rate = xl_number(xl_cell(ctx, f'Inputs!{col}17')) / 100.0
    interest_rate = xl_number(xl_cell(ctx, f'Inputs!{col}16')) / 100.0
    primary_balance = xl_number(xl_cell(ctx, f'Inputs!{col}18'))
    numerator = previous * (1.0 + growth_rate)
    denominator = 1.0 + interest_rate
    if denominator != 0.0:
        return numerator / denominator - primary_balance
    else:
        return xl_raise(XlError.DIV)

def compute_shocked_debt_to_gdp(ctx: EvalContext, time_period: int) -> float:
    """Compute the debt-to-GDP ratio for the shocked scenario at a given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (2 through 5).

Returns:
    Debt-to-GDP ratio as a percentage of GDP.

Note:
    Covers Engine!D20:G20. Excel: =Engine!C20*(1+(Inputs!D17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!D10)/100)/(1+(Inputs!D16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!D10)/100)-Engine!D16.
"""
    col_map = {2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    col = col_map[time_period]
    if time_period == 2:
        prev_debt = shocked_debt_to_gdp_period_1(ctx)
    else:
        prev_debt = compute_shocked_debt_to_gdp(ctx, time_period - 1)
    numerator_rate = xl_number(xl_cell(ctx, f'Inputs!{col}17'))
    denominator_rate = xl_number(xl_cell(ctx, f'Inputs!{col}16'))
    scenario_type = xl_int(xl_cell(ctx, 'Inputs!B22'))
    shock_mag = shock_magnitude_resolved(ctx)
    active = shock_active(ctx, time_period=time_period)
    shock_numerator = shock_mag * active if scenario_type == 2 else 0.0
    shock_denominator = shock_mag * active if scenario_type == 1 else 0.0
    numerator = prev_debt * (1.0 + (numerator_rate + shock_numerator) / 100.0)
    denom = 1.0 + (denominator_rate + shock_denominator) / 100.0
    quotient = numerator / denom if denom != 0.0 else xl_raise(XlError.DIV)
    balance = shocked_primary_balance(ctx, time_period=time_period)
    return quotient - balance

def shocked_primary_balance(ctx: EvalContext, time_period: int) -> float:
    """Return the shocked primary balance (as percentage of GDP) for the given time period.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    The shocked primary balance incorporating the shock adjustment when the scenario is shocked.

Note:
    Covers Engine!C16:G16. Excel: =Inputs!C18+CHOOSE(Inputs!B22,0,0,Engine!B9)*Engine!C10.
"""
    col_letters = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    base = xl_number(xl_cell(ctx, f'Inputs!{col_letters[time_period]}18'))
    scenario_index = xl_int(xl_cell(ctx, 'Inputs!B22'))
    if 1 <= scenario_index <= 3:
        if scenario_index == 3:
            shock_magnitude = shock_magnitude_resolved(ctx)
        else:
            shock_magnitude = 0.0
    else:
        xl_raise(XlError.VALUE)
    shock_active_flag = xl_number(shock_active(ctx, time_period=time_period))
    return base + shock_magnitude * shock_active_flag

def shock_active(ctx: EvalContext, time_period: int) -> float:
    """Return 1.0 if the observed value for the given time period meets or exceeds the shock threshold.

Args:
    ctx: Workbook evaluation context.
    time_period: Time period index (1 through 5).

Returns:
    1.0 if the observed value is at or above the threshold, else 0.0.

Note:
    Covers Engine!C10:G10. Excel: =IF(Engine!C5>=Inputs!B21,1,0).
"""
    column_by_period = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}
    current_value = xl_cell(ctx, f'Engine!{column_by_period[time_period]}5')
    threshold = xl_cell(ctx, 'Inputs!B21')
    meets_threshold = xl_bool(xl_compare('>=', current_value, threshold))
    return 1.0 if meets_threshold else 0.0

def initial_debt_to_gdp(ctx: EvalContext) -> float | str:
    """Return the initial debt-to-GDP ratio by looking up a value in a table.

Args:
    ctx: Workbook evaluation context.

Returns:
    The initial debt-to-GDP ratio. The return type depends on cell contents.

Note:
    Covers Inputs!B6. Excel: =INDEX(Inputs!A10:Inputs!C12,MATCH(Inputs!B5,Inputs!A10:Inputs!A12,0),2).
"""
    lookup_key = xl_cell(ctx, 'Inputs!B5')
    lookup_array = xl_range(ctx, 'Inputs!A10:Inputs!A12')
    match_row = xl_match(lookup_key, lookup_array, 0.0)
    table_range = ('Inputs', 10, 1, 12, 3)
    target_ref = xl_index_ref(table_range, match_row, 2.0)
    return xl_offset(ctx, target_ref, 0.0, 0.0)

def baseline_debt_to_gdp(ctx: EvalContext) -> float:
    """Calculate the baseline debt-to-GDP ratio for the next period using standard debt dynamics.

Args:
    ctx: Workbook evaluation context.

Returns:
    Projected debt-to-GDP ratio as a percentage of GDP.

Note:
    Covers Engine!C6. Excel: =Inputs!B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18.
"""
    initial_debt = xl_number(initial_debt_to_gdp(ctx))
    interest_rate_pct = xl_number(xl_cell(ctx, 'Inputs!C17'))
    gdp_growth_pct = xl_number(xl_cell(ctx, 'Inputs!C16'))
    primary_balance = xl_number(xl_cell(ctx, 'Inputs!C18'))
    interest_factor = 1.0 + interest_rate_pct / 100.0
    growth_factor = 1.0 + gdp_growth_pct / 100.0
    numerator = initial_debt * interest_factor
    denominator = growth_factor
    if denominator == 0:
        xl_raise(XlError.DIV)
    projected_debt_ratio = numerator / denominator
    return projected_debt_ratio - primary_balance

def shock_magnitude_resolved(ctx: EvalContext) -> float:
    """Retrieve the resolved shock magnitude (in percentage points) based on the selected scenario index.

Args:
    ctx: Workbook evaluation context.

Returns:
    Magnitude of the shock for the selected scenario, in percentage points.

Note:
    Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    scenario_index = xl_number(xl_cell(ctx, 'Inputs!B22'))
    column_offset = scenario_index - 1.0
    row_offset = 0.0
    return xl_offset(ctx, ('Inputs', 26, 2), row_offset, column_offset, None, None)

def shocked_debt_to_gdp_period_1(ctx: EvalContext) -> float:
    """Compute the debt-to-GDP ratio for the shocked scenario at time period 1, incorporating shock adjustments to interest and growth rates.

Args:
    ctx: Workbook evaluation context.

Returns:
    Debt-to-GDP ratio as a percentage of GDP.

Note:
    Covers Engine!C20. Excel: =Inputs!B6*(1+(Inputs!C17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!C10)/100)/(1+(Inputs!C16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!C10)/100)-Engine!C16.
"""
    initial_debt = xl_number(initial_debt_to_gdp(ctx))
    c17 = xl_number(xl_cell(ctx, 'Inputs!C17'))
    c16 = xl_number(xl_cell(ctx, 'Inputs!C16'))
    scenario_index = xl_int(xl_cell(ctx, 'Inputs!B22'))
    if not 1 <= scenario_index <= 3:
        xl_raise(XlError.VALUE)
    shock_active_val = xl_number(shock_active(ctx, time_period=1))
    shock_magnitude = xl_number(shock_magnitude_resolved(ctx))
    if scenario_index == 1:
        numerator_shock_addend = 0.0
        denominator_shock_addend = shock_magnitude
    elif scenario_index == 2:
        numerator_shock_addend = shock_magnitude
        denominator_shock_addend = 0.0
    else:
        numerator_shock_addend = 0.0
        denominator_shock_addend = 0.0
    numerator_adjustment = (c17 + numerator_shock_addend * shock_active_val) / 100.0
    denominator_adjustment = (c16 + denominator_shock_addend * shock_active_val) / 100.0
    compound_numerator = initial_debt * (1.0 + numerator_adjustment)
    compound_denominator = 1.0 + denominator_adjustment
    debt_before_pb = compound_numerator / compound_denominator if compound_denominator != 0.0 else xl_raise(XlError.DIV)
    shocked_pb = xl_number(shocked_primary_balance(ctx, time_period=1))
    return debt_before_pb - shocked_pb

# --- Projection public address aliases ---

def cell_engine_b20(ctx):
    return initial_debt_to_gdp(ctx)

def cell_engine_b6(ctx):
    return initial_debt_to_gdp(ctx)

def cell_outputs_b12(ctx):
    return baseline_debt_to_gdp(ctx)

def cell_outputs_b13(ctx):
    return shocked_debt_to_gdp_period_1(ctx)

# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B14': ('shock_impact_on_debt_to_gdp', {'time_period': 1}),
    'Outputs!C12': ('project_debt_to_gdp', {'time_period': 2}),
    'Outputs!C13': ('compute_shocked_debt_to_gdp', {'time_period': 2}),
    'Outputs!C14': ('shock_impact_on_debt_to_gdp', {'time_period': 2}),
    'Outputs!D12': ('project_debt_to_gdp', {'time_period': 3}),
    'Outputs!D13': ('compute_shocked_debt_to_gdp', {'time_period': 3}),
    'Outputs!D14': ('shock_impact_on_debt_to_gdp', {'time_period': 3}),
    'Outputs!E12': ('project_debt_to_gdp', {'time_period': 4}),
    'Outputs!E13': ('compute_shocked_debt_to_gdp', {'time_period': 4}),
    'Outputs!E14': ('shock_impact_on_debt_to_gdp', {'time_period': 4}),
    'Outputs!F12': ('project_debt_to_gdp', {'time_period': 5}),
    'Outputs!F13': ('compute_shocked_debt_to_gdp', {'time_period': 5}),
    'Outputs!F14': ('shock_impact_on_debt_to_gdp', {'time_period': 5}),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'shock_magnitude_resolved',
    'Engine!C20': 'shocked_debt_to_gdp_period_1',
    'Engine!C6': 'baseline_debt_to_gdp',
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
