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

def output_delta(ctx, col):
    """Return the difference between shocked and baseline debt-to-GDP.

    Args:
        ctx: Workbook evaluation context.
        col: Engine column letter (C through G).

    Returns:
        The shocked debt-to-GDP minus the baseline debt-to-GDP.

    Note:
        Covers Outputs!B14:F14 (columns C..G). Excel formula: =Engine!{col}20-Engine!{col}6.
"""
    return xl_sub(debt_to_gdp(ctx, col), baseline_debt(ctx, col))

def debt_to_gdp(ctx, col):
    """Compute shocked debt-to-GDP ratio for a given projection column.

    Args:
        ctx: Workbook evaluation context.
        col: Engine column letter (C through G).

    Returns:
        The debt-to-GDP ratio for the column after shocks.

    Note:
        Covers Engine!C20:G20. Excel: ={PRIOR_DEBT}*(1+(Inputs!{COL}17+CHOOSE(Inputs!B22,0,Engine!B9,0)*Engine!{COL}10)/100)/(1+(Inputs!{COL}16+CHOOSE(Inputs!B22,Engine!B9,0,0)*Engine!{COL}10)/100)-Engine!{COL}16
        where {PRIOR_DEBT} is Inputs!B6 for C20, and the previous column's result for D20:G20.
"""
    if col == 'C':
        prior_debt = xl_eval(ctx, 'Inputs!B6', initial_debt_to_gdp)
    else:
        prev_col = chr(ord(col) - 1)
        prior_debt = debt_to_gdp(ctx, prev_col)
    shock_type_raw = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_type_raw, XlError):
        return shock_type_raw
    shock_type = to_int(shock_type_raw)
    if isinstance(shock_type, XlError):
        return shock_type
    if shock_type < 1 or shock_type > 3:
        return XlError.VALUE
    shock_magnitude = xl_eval(ctx, 'Engine!B9', shock_magnitude_resolved)
    if shock_type == 1:
        choose_num = 0.0
        choose_den = shock_magnitude
    elif shock_type == 2:
        choose_num = shock_magnitude
        choose_den = 0.0
    else:
        choose_num = 0.0
        choose_den = 0.0
    interest_rate = xl_cell(ctx, f'Inputs!{col}17')
    growth_rate = xl_cell(ctx, f'Inputs!{col}16')
    shock_active_col = shock_active(ctx, col)
    primary_balance = primary_balance_shocked(ctx, col)
    numerator = xl_add(1.0, xl_div(xl_add(interest_rate, xl_mul(choose_num, shock_active_col)), 100.0))
    denominator = xl_add(1.0, xl_div(xl_add(growth_rate, xl_mul(choose_den, shock_active_col)), 100.0))
    result = xl_sub(xl_div(xl_mul(prior_debt, numerator), denominator), primary_balance)
    return result

def baseline_debt(ctx, col):
    """Return the baseline debt-to-GDP ratio for a given projection column.

    Args:
        ctx: Workbook evaluation context.
        col: Engine column letter (C through G).

    Returns:
        Baseline debt-to-GDP ratio computed as (prior_debt * (1 + interest_baseline/100) / (1 + growth_baseline/100)) - primary_balance_baseline.

    Note:
        Covers Engine!C6:G6. Excel: ={PRIOR_DEBT}*(1+Inputs!{col}17/100)/(1+Inputs!{col}16/100)-Inputs!{col}18 where PRIOR_DEBT is Inputs!B6 for column C, else Engine!{prev_col}6.
"""
    if col == 'C':
        prior_debt = xl_eval(ctx, 'Inputs!B6', initial_debt_to_gdp)
    else:
        prev_col = chr(ord(col) - 1)
        prior_debt = baseline_debt(ctx, prev_col)
    growth_baseline = xl_cell(ctx, f'Inputs!{col}16')
    interest_baseline = xl_cell(ctx, f'Inputs!{col}17')
    primary_balance_baseline = xl_cell(ctx, f'Inputs!{col}18')
    return xl_sub(xl_div(xl_mul(prior_debt, xl_add(1.0, xl_div(interest_baseline, 100.0))), xl_add(1.0, xl_div(growth_baseline, 100.0))), primary_balance_baseline)

def primary_balance_shocked(ctx, col):
    """Return the primary balance including shock effects for a given projection column.

Args:
    ctx: Workbook evaluation context.
    col: Engine column letter (C through G).

Returns:
    The shocked primary balance, as a float or XlError.

Note:
    Covers Engine!C16:G16. Excel: =Inputs!{col}18+CHOOSE(Inputs!$B$22,0,0,Engine!$B$9)*Engine!{col}10.
"""
    shock_type_raw = xl_cell(ctx, 'Inputs!B22')
    if isinstance(shock_type_raw, XlError):
        return shock_type_raw
    shock_type_int = to_int(shock_type_raw)
    if isinstance(shock_type_int, XlError):
        return shock_type_int
    if shock_type_int < 1 or shock_type_int > 3:
        return XlError.VALUE
    if shock_type_int == 1 or shock_type_int == 2:
        shock_multiplier = 0.0
    elif shock_type_int == 3:
        shock_multiplier = xl_eval(ctx, 'Engine!B9', shock_magnitude_resolved)
    else:
        shock_multiplier = XlError.VALUE
    baseline = xl_cell(ctx, f'Inputs!{col}18')
    shock_active_val = shock_active(ctx, col)
    return xl_add(baseline, xl_mul(shock_multiplier, shock_active_val))

def shock_active(ctx, col):
    """Return 1.0 when the shock is active for the given projection column.

    Args:
        ctx: Workbook evaluation context.
        col: Engine column letter (C through G).

    Returns:
        1.0 if the projection year is at or after the shock year, else 0.0.

    Note:
        Covers Engine!C10:G10. Excel: =IF(Engine!{col}5>=Inputs!$B$21,1,0).
"""
    projection_year = xl_cell(ctx, f'Engine!{col}5')
    shock_year = xl_cell(ctx, 'Inputs!B21')
    condition = xl_ge(projection_year, shock_year)
    bool_val = to_bool(condition)
    if isinstance(bool_val, XlError):
        return bool_val
    return 1.0 if bool_val else 0.0

def initial_debt_to_gdp(ctx):
    """Look up the initial debt-to-GDP ratio for the selected country.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Initial debt-to-GDP ratio from the country profile table.

    Note:
        Covers Inputs!B6. Excel: =INDEX($A$10:$C$12,MATCH($B$5,$A$10:$A$12,0),2).
"""
    selected_country = xl_cell(ctx, 'Inputs!B5')
    country_a10 = xl_cell(ctx, 'Inputs!A10')
    country_a11 = xl_cell(ctx, 'Inputs!A11')
    country_a12 = xl_cell(ctx, 'Inputs!A12')
    country_lookup_array = np.array([[country_a10], [country_a11], [country_a12]], dtype=object)
    matched_position = xl_match(selected_country, country_lookup_array, 0.0)
    initial_debt_to_gdp_ref = xl_index_ref(('Inputs', 10, 1, 12, 3), matched_position, 2.0)
    return xl_offset(ctx, initial_debt_to_gdp_ref, 0.0, 0.0)

def shock_magnitude_resolved(ctx):
    """Resolved shock magnitude from the shock magnitude table based on shock type.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        The shock magnitude value from Inputs!B26:D26 corresponding to the shock type in Inputs!B22.

    Note:
        Covers Engine!B9. Excel: =OFFSET(Inputs!$B$26,0,Inputs!$B$22-1).
"""
    shock_type = xl_cell(ctx, 'Inputs!B22')
    column_offset = xl_sub(shock_type, 1.0)
    magnitude = xl_offset(ctx, ('Inputs', 26, 2), 0.0, column_offset, None, None)
    return magnitude


# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B12': ('baseline_debt', 'C'),
    'Outputs!B13': ('debt_to_gdp', 'C'),
    'Outputs!B14': ('output_delta', 'C'),
    'Outputs!C12': ('baseline_debt', 'D'),
    'Outputs!C13': ('debt_to_gdp', 'D'),
    'Outputs!C14': ('output_delta', 'D'),
    'Outputs!D12': ('baseline_debt', 'E'),
    'Outputs!D13': ('debt_to_gdp', 'E'),
    'Outputs!D14': ('output_delta', 'E'),
    'Outputs!E12': ('baseline_debt', 'F'),
    'Outputs!E13': ('debt_to_gdp', 'F'),
    'Outputs!E14': ('output_delta', 'F'),
    'Outputs!F12': ('baseline_debt', 'G'),
    'Outputs!F13': ('debt_to_gdp', 'G'),
    'Outputs!F14': ('output_delta', 'G'),
}
_SYMBOL_DISPATCH = {
    'Engine!B9': 'shock_magnitude_resolved',
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
        helper_name, column = dispatch
        helper = globals()[helper_name]

        def _bound(ctx, _helper=helper, _column=column):
            return _helper(ctx, _column)

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
