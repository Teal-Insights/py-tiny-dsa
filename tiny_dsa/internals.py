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

def output_delta(ctx, time_period):
    """Return the difference between shocked and baseline debt-to-GDP for a given time period.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection year (1-5).

    Returns:
        The difference (shocked minus baseline) as a float.

    Note:
        Covers Outputs!B14:F14. Excel: =Engine!{col}20-Engine!{col}6.
"""
    return xl_sub(debt_to_gdp(ctx, time_period), baseline_debt(ctx, time_period))

def baseline_debt(ctx, time_period):
    """Compute the baseline debt-to-GDP ratio for a given projection year.

Args:
    ctx: Workbook evaluation context.
    time_period: Projection year index (1 to 5).

Returns:
    Baseline debt-to-GDP ratio.

Note:
    Covers Engine!C6:G6. Excel template: ={PRIOR_DEBT}*(1+Inputs!{COL}17/100)/(1+Inputs!{COL}16/100)-Inputs!{COL}18,
    where {PRIOR_DEBT} is Inputs!B6 for time_period=1, or the previous year's baseline debt otherwise.
    The growth rate Inputs!{COL}16 is replaced by primary_balance_shocked(ctx, time_period).
"""
    col = {1: 'C', 2: 'D', 3: 'E', 4: 'F', 5: 'G'}[time_period]
    if time_period == 1:
        prior_debt = xl_cell(ctx, 'Inputs!B6')
    else:
        prior_debt = baseline_debt(ctx, time_period - 1)
    interest_rate = xl_cell(ctx, f'Inputs!{col}17')
    growth_rate = primary_balance_shocked(ctx, time_period=time_period)
    primary_balance = xl_cell(ctx, f'Inputs!{col}18')
    return xl_sub(xl_div(xl_mul(prior_debt, xl_add(1.0, xl_div(interest_rate, 100.0))), xl_add(1.0, xl_div(growth_rate, 100.0))), primary_balance)

def debt_to_gdp(ctx, time_period):
    """Compute the shocked debt-to-GDP ratio for a given projection year.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection year index (1..5).

    Returns:
        Shocked debt-to-GDP ratio as a float or XlError on error.

    Note:
        Covers Engine!C20:G20. Excel formula:
        =PRIOR_DEBT*(1+(Inputs!col17+CHOOSE(Inputs!$B$22,0,$B$9,0)*col10)/100)
        /(1+(Inputs!col16+CHOOSE(Inputs!$B$22,$B$9,0,0)*col10)/100)-col16,
        where PRIOR_DEBT is Inputs!B6 for the first column (C) and the previous
        column's result otherwise.
"""
    col_letter = chr(ord('C') + time_period - 1)
    if time_period == 1:
        prior_debt = xl_cell(ctx, 'Inputs!B6')
    else:
        prior_debt = debt_to_gdp(ctx, time_period - 1)
    growth_baseline = xl_cell(ctx, f'Inputs!{col_letter}16')
    interest_baseline = xl_cell(ctx, f'Inputs!{col_letter}17')
    shock_type_raw = xl_cell(ctx, 'Inputs!B22')
    shock_magnitude = xl_cell(ctx, 'Engine!B9')
    shock_active_val = shock_active(ctx, time_period=time_period)
    primary_balance = primary_balance_shocked(ctx, time_period=time_period)
    shock_type_int = to_int(shock_type_raw)
    if isinstance(shock_type_int, XlError):
        shock_type = shock_type_int
    elif shock_type_int < 1 or shock_type_int > 3:
        shock_type = XlError.VALUE
    else:
        shock_type = shock_type_int
    if isinstance(shock_type, XlError):
        growth_adjustment = shock_type
        interest_adjustment = shock_type
    elif shock_type == 1:
        growth_adjustment = shock_magnitude
        interest_adjustment = 0
    elif shock_type == 2:
        growth_adjustment = 0
        interest_adjustment = shock_magnitude
    else:
        growth_adjustment = 0
        interest_adjustment = 0
    effective_interest = xl_add(interest_baseline, xl_mul(interest_adjustment, shock_active_val))
    effective_growth = xl_add(growth_baseline, xl_mul(growth_adjustment, shock_active_val))
    numerator = xl_mul(prior_debt, xl_add(1.0, xl_div(effective_interest, 100.0)))
    fraction = xl_div(numerator, xl_add(1.0, xl_div(effective_growth, 100.0)))
    shocked_debt = xl_sub(fraction, primary_balance)
    return shocked_debt

def primary_balance_shocked(ctx, time_period):
    """Return the primary balance including the shock for the given projection year.

    Args:
        ctx: Workbook evaluation context.
        time_period: Projection year as an integer (1 for column C, 2 for D, ..., 5 for G).

    Returns:
        The shocked primary balance as a float or an XlError.

    Note:
        Covers Engine!C16:G16. Excel formula: =Inputs!{col}18+CHOOSE(Inputs!B22,0,0,Engine!B9)*Engine!{col}10.
"""
    raw_shock_type = xl_cell(ctx, 'Inputs!B22')
    if isinstance(raw_shock_type, XlError):
        shock_factor = raw_shock_type
    else:
        shock_type_int = to_int(raw_shock_type)
        if isinstance(shock_type_int, XlError):
            shock_factor = shock_type_int
        elif shock_type_int < 1 or shock_type_int > 3:
            shock_factor = XlError.VALUE
        elif shock_type_int == 1 or shock_type_int == 2:
            shock_factor = 0.0
        elif shock_type_int == 3:
            shock_factor = xl_eval(ctx, 'Engine!B9', shock_magnitude_resolved)
        else:
            shock_factor = XlError.VALUE
    col_letter = chr(ord('C') + time_period - 1)
    baseline = xl_cell(ctx, f'Inputs!{col_letter}18')
    active = shock_active(ctx, time_period=time_period)
    return xl_add(baseline, xl_mul(shock_factor, active))

def shock_active(ctx, time_period):
    """Return 1.0 when the shock is active for the given projection column.

    Args:
        ctx: Workbook evaluation context.
        time_period: Integer projection year index (1 = first year, column C).

    Returns:
        1.0 if the projection year is at or after the shock year, else 0.0.

    Note:
        Covers Engine!C10:G10. Excel: =IF(Engine!{col}5>=Inputs!$B$21,1,0).
"""
    col_letter = chr(ord('C') + time_period - 1)
    year_cell = f'Engine!{col_letter}5'
    shock_year_cell = 'Inputs!B21'
    ge_result = xl_ge(xl_cell(ctx, year_cell), xl_cell(ctx, shock_year_cell))
    bool_result = to_bool(ge_result)
    if isinstance(bool_result, XlError):
        return bool_result
    return 1.0 if bool_result else 0.0

def initial_debt_to_gdp(ctx):
    """Look up the initial debt-to-GDP ratio for the selected country.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        Initial debt-to-GDP ratio from the country profile table.

    Note:
        Covers Inputs!B6. Excel: =INDEX(Inputs!A10:Inputs!C12, MATCH(Inputs!B5, Inputs!A10:Inputs!A12, 0), 2).
"""
    selected_country = xl_cell(ctx, 'Inputs!B5')
    country_list = np.array([[xl_cell(ctx, 'Inputs!A10')], [xl_cell(ctx, 'Inputs!A11')], [xl_cell(ctx, 'Inputs!A12')]], dtype=object)
    match_row = xl_match(selected_country, country_list, 0.0)
    data_range = ('Inputs', 10, 1, 12, 3)
    value = xl_offset(ctx, xl_index_ref(data_range, match_row, 2.0), 0.0, 0.0)
    return value

def shock_magnitude_resolved(ctx):
    """Resolved shock magnitude for the selected shock type.

    Args:
        ctx: Workbook evaluation context.

    Returns:
        The shock magnitude value from Inputs!B26:D26 corresponding to the shock type (1-3).

    Note:
        Covers Engine!B9. Excel: =OFFSET(Inputs!B26,0,Inputs!B22-1).
"""
    shock_type = xl_cell(ctx, 'Inputs!B22')
    column_offset = xl_sub(shock_type, 1.0)
    return xl_offset(ctx, ('Inputs', 26, 2), 0.0, column_offset, None, None)


# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
_ADDRESS_DISPATCH = {
    'Outputs!B12': ('baseline_debt', {'time_period': 1}),
    'Outputs!B13': ('debt_to_gdp', {'time_period': 1}),
    'Outputs!B14': ('output_delta', {'time_period': 1}),
    'Outputs!C12': ('baseline_debt', {'time_period': 2}),
    'Outputs!C13': ('debt_to_gdp', {'time_period': 2}),
    'Outputs!C14': ('output_delta', {'time_period': 2}),
    'Outputs!D12': ('baseline_debt', {'time_period': 3}),
    'Outputs!D13': ('debt_to_gdp', {'time_period': 3}),
    'Outputs!D14': ('output_delta', {'time_period': 3}),
    'Outputs!E12': ('baseline_debt', {'time_period': 4}),
    'Outputs!E13': ('debt_to_gdp', {'time_period': 4}),
    'Outputs!E14': ('output_delta', {'time_period': 4}),
    'Outputs!F12': ('baseline_debt', {'time_period': 5}),
    'Outputs!F13': ('debt_to_gdp', {'time_period': 5}),
    'Outputs!F14': ('output_delta', {'time_period': 5}),
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
