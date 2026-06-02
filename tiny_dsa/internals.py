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

def cell_inputs_b6(ctx):
    '''Formula: =INDEX($A$10:$C$12,MATCH($B$5,$A$10:$A$12,0),2)'''
    return xl_offset(ctx, xl_index_ref(('Inputs', 10, 1, 12, 3), xl_match(xl_cell(ctx, 'Inputs!B5'), np.array(np.array([[xl_cell(ctx, 'Inputs!A10')], [xl_cell(ctx, 'Inputs!A11')], [xl_cell(ctx, 'Inputs!A12')]], dtype=object), dtype=object), 0.0), 2.0), 0.0, 0.0)


def cell_engine_b6(ctx):
    '''Formula: =Inputs!B6'''
    return xl_eval(ctx, 'Inputs!B6', cell_inputs_b6)


def cell_engine_c6(ctx):
    '''Formula: =B6*(1+Inputs!C17/100)/(1+Inputs!C16/100)-Inputs!C18'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!B6', cell_engine_b6), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!C17'), 100.0))), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!C16'), 100.0))), xl_cell(ctx, 'Inputs!C18'))


def cell_outputs_b12(ctx):
    '''Formula: =Engine!C6'''
    return xl_eval(ctx, 'Engine!C6', cell_engine_c6)


def cell_engine_d6(ctx):
    '''Formula: =C6*(1+Inputs!D17/100)/(1+Inputs!D16/100)-Inputs!D18'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!C6', cell_engine_c6), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!D17'), 100.0))), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!D16'), 100.0))), xl_cell(ctx, 'Inputs!D18'))


def cell_outputs_c12(ctx):
    '''Formula: =Engine!D6'''
    return xl_eval(ctx, 'Engine!D6', cell_engine_d6)


def cell_engine_e6(ctx):
    '''Formula: =D6*(1+Inputs!E17/100)/(1+Inputs!E16/100)-Inputs!E18'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!D6', cell_engine_d6), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!E17'), 100.0))), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!E16'), 100.0))), xl_cell(ctx, 'Inputs!E18'))


def cell_outputs_d12(ctx):
    '''Formula: =Engine!E6'''
    return xl_eval(ctx, 'Engine!E6', cell_engine_e6)


def cell_engine_f6(ctx):
    '''Formula: =E6*(1+Inputs!F17/100)/(1+Inputs!F16/100)-Inputs!F18'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!E6', cell_engine_e6), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!F17'), 100.0))), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!F16'), 100.0))), xl_cell(ctx, 'Inputs!F18'))


def cell_outputs_e12(ctx):
    '''Formula: =Engine!F6'''
    return xl_eval(ctx, 'Engine!F6', cell_engine_f6)


def cell_engine_g6(ctx):
    '''Formula: =F6*(1+Inputs!G17/100)/(1+Inputs!G16/100)-Inputs!G18'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!F6', cell_engine_f6), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!G17'), 100.0))), xl_add(1.0, xl_div(xl_cell(ctx, 'Inputs!G16'), 100.0))), xl_cell(ctx, 'Inputs!G18'))


def cell_outputs_f12(ctx):
    '''Formula: =Engine!G6'''
    return xl_eval(ctx, 'Engine!G6', cell_engine_g6)


def cell_engine_b9(ctx):
    '''Formula: =OFFSET(Inputs!$B$26,0,Inputs!$B$22-1)'''
    return xl_offset(ctx, ('Inputs', 26, 2), 0.0, xl_sub(xl_cell(ctx, 'Inputs!B22'), 1.0), None, None)


def cell_engine_c10(ctx):
    '''Formula: =IF(C5>=Inputs!$B$21,1,0)'''
    return (_t2 if isinstance((_t2 := to_bool((_t1 := xl_ge(xl_cell(ctx, 'Engine!C5'), xl_cell(ctx, 'Inputs!B21'))))), XlError) else ((1.0) if _t2 else (0.0)))


def cell_engine_c16(ctx):
    '''Formula: =Inputs!C18+CHOOSE(Inputs!$B$22,0,0,$B$9)*C10'''
    return xl_add(xl_cell(ctx, 'Inputs!C18'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((0.0) if _t2 == 2 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!C10', cell_engine_c10)))


def cell_engine_b20(ctx):
    '''Formula: =Inputs!B6'''
    return xl_eval(ctx, 'Inputs!B6', cell_inputs_b6)


def cell_engine_c15(ctx):
    '''Formula: =Inputs!C17+CHOOSE(Inputs!$B$22,0,$B$9,0)*C10'''
    return xl_add(xl_cell(ctx, 'Inputs!C17'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!C10', cell_engine_c10)))


def cell_engine_c14(ctx):
    '''Formula: =Inputs!C16+CHOOSE(Inputs!$B$22,$B$9,0,0)*C10'''
    return xl_add(xl_cell(ctx, 'Inputs!C16'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 1 else (((0.0) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!C10', cell_engine_c10)))


def cell_engine_c20(ctx):
    '''Formula: =B20*(1+C15/100)/(1+C14/100)-C16'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!B20', cell_engine_b20), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!C15', cell_engine_c15), 100.0))), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!C14', cell_engine_c14), 100.0))), xl_eval(ctx, 'Engine!C16', cell_engine_c16))


def cell_outputs_b13(ctx):
    '''Formula: =Engine!C20'''
    return xl_eval(ctx, 'Engine!C20', cell_engine_c20)


def cell_engine_d10(ctx):
    '''Formula: =IF(D5>=Inputs!$B$21,1,0)'''
    return (_t2 if isinstance((_t2 := to_bool((_t1 := xl_ge(xl_cell(ctx, 'Engine!D5'), xl_cell(ctx, 'Inputs!B21'))))), XlError) else ((1.0) if _t2 else (0.0)))


def cell_engine_d15(ctx):
    '''Formula: =Inputs!D17+CHOOSE(Inputs!$B$22,0,$B$9,0)*D10'''
    return xl_add(xl_cell(ctx, 'Inputs!D17'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!D10', cell_engine_d10)))


def cell_engine_d14(ctx):
    '''Formula: =Inputs!D16+CHOOSE(Inputs!$B$22,$B$9,0,0)*D10'''
    return xl_add(xl_cell(ctx, 'Inputs!D16'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 1 else (((0.0) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!D10', cell_engine_d10)))


def cell_engine_d16(ctx):
    '''Formula: =Inputs!D18+CHOOSE(Inputs!$B$22,0,0,$B$9)*D10'''
    return xl_add(xl_cell(ctx, 'Inputs!D18'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((0.0) if _t2 == 2 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!D10', cell_engine_d10)))


def cell_engine_d20(ctx):
    '''Formula: =C20*(1+D15/100)/(1+D14/100)-D16'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!C20', cell_engine_c20), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!D15', cell_engine_d15), 100.0))), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!D14', cell_engine_d14), 100.0))), xl_eval(ctx, 'Engine!D16', cell_engine_d16))


def cell_outputs_c13(ctx):
    '''Formula: =Engine!D20'''
    return xl_eval(ctx, 'Engine!D20', cell_engine_d20)


def cell_engine_e10(ctx):
    '''Formula: =IF(E5>=Inputs!$B$21,1,0)'''
    return (_t2 if isinstance((_t2 := to_bool((_t1 := xl_ge(xl_cell(ctx, 'Engine!E5'), xl_cell(ctx, 'Inputs!B21'))))), XlError) else ((1.0) if _t2 else (0.0)))


def cell_engine_e15(ctx):
    '''Formula: =Inputs!E17+CHOOSE(Inputs!$B$22,0,$B$9,0)*E10'''
    return xl_add(xl_cell(ctx, 'Inputs!E17'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!E10', cell_engine_e10)))


def cell_engine_e16(ctx):
    '''Formula: =Inputs!E18+CHOOSE(Inputs!$B$22,0,0,$B$9)*E10'''
    return xl_add(xl_cell(ctx, 'Inputs!E18'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((0.0) if _t2 == 2 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!E10', cell_engine_e10)))


def cell_engine_e14(ctx):
    '''Formula: =Inputs!E16+CHOOSE(Inputs!$B$22,$B$9,0,0)*E10'''
    return xl_add(xl_cell(ctx, 'Inputs!E16'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 1 else (((0.0) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!E10', cell_engine_e10)))


def cell_engine_e20(ctx):
    '''Formula: =D20*(1+E15/100)/(1+E14/100)-E16'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!D20', cell_engine_d20), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!E15', cell_engine_e15), 100.0))), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!E14', cell_engine_e14), 100.0))), xl_eval(ctx, 'Engine!E16', cell_engine_e16))


def cell_outputs_d13(ctx):
    '''Formula: =Engine!E20'''
    return xl_eval(ctx, 'Engine!E20', cell_engine_e20)


def cell_engine_f10(ctx):
    '''Formula: =IF(F5>=Inputs!$B$21,1,0)'''
    return (_t2 if isinstance((_t2 := to_bool((_t1 := xl_ge(xl_cell(ctx, 'Engine!F5'), xl_cell(ctx, 'Inputs!B21'))))), XlError) else ((1.0) if _t2 else (0.0)))


def cell_engine_f15(ctx):
    '''Formula: =Inputs!F17+CHOOSE(Inputs!$B$22,0,$B$9,0)*F10'''
    return xl_add(xl_cell(ctx, 'Inputs!F17'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!F10', cell_engine_f10)))


def cell_engine_f16(ctx):
    '''Formula: =Inputs!F18+CHOOSE(Inputs!$B$22,0,0,$B$9)*F10'''
    return xl_add(xl_cell(ctx, 'Inputs!F18'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((0.0) if _t2 == 2 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!F10', cell_engine_f10)))


def cell_engine_f14(ctx):
    '''Formula: =Inputs!F16+CHOOSE(Inputs!$B$22,$B$9,0,0)*F10'''
    return xl_add(xl_cell(ctx, 'Inputs!F16'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 1 else (((0.0) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!F10', cell_engine_f10)))


def cell_engine_f20(ctx):
    '''Formula: =E20*(1+F15/100)/(1+F14/100)-F16'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!E20', cell_engine_e20), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!F15', cell_engine_f15), 100.0))), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!F14', cell_engine_f14), 100.0))), xl_eval(ctx, 'Engine!F16', cell_engine_f16))


def cell_outputs_e13(ctx):
    '''Formula: =Engine!F20'''
    return xl_eval(ctx, 'Engine!F20', cell_engine_f20)


def cell_engine_g10(ctx):
    '''Formula: =IF(G5>=Inputs!$B$21,1,0)'''
    return (_t2 if isinstance((_t2 := to_bool((_t1 := xl_ge(xl_cell(ctx, 'Engine!G5'), xl_cell(ctx, 'Inputs!B21'))))), XlError) else ((1.0) if _t2 else (0.0)))


def cell_engine_g14(ctx):
    '''Formula: =Inputs!G16+CHOOSE(Inputs!$B$22,$B$9,0,0)*G10'''
    return xl_add(xl_cell(ctx, 'Inputs!G16'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 1 else (((0.0) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!G10', cell_engine_g10)))


def cell_engine_g16(ctx):
    '''Formula: =Inputs!G18+CHOOSE(Inputs!$B$22,0,0,$B$9)*G10'''
    return xl_add(xl_cell(ctx, 'Inputs!G18'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((0.0) if _t2 == 2 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!G10', cell_engine_g10)))


def cell_engine_g15(ctx):
    '''Formula: =Inputs!G17+CHOOSE(Inputs!$B$22,0,$B$9,0)*G10'''
    return xl_add(xl_cell(ctx, 'Inputs!G17'), xl_mul((_t1 if isinstance((_t1 := xl_cell(ctx, 'Inputs!B22')), XlError) else (_t2 if isinstance((_t2 := to_int(_t1)), XlError) else XlError.VALUE if _t2 < 1 or _t2 > 3 else ((0.0) if _t2 == 1 else (((xl_eval(ctx, 'Engine!B9', cell_engine_b9)) if _t2 == 2 else (((0.0) if _t2 == 3 else (XlError.VALUE)))))))), xl_eval(ctx, 'Engine!G10', cell_engine_g10)))


def cell_engine_g20(ctx):
    '''Formula: =F20*(1+G15/100)/(1+G14/100)-G16'''
    return xl_sub(xl_div(xl_mul(xl_eval(ctx, 'Engine!F20', cell_engine_f20), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!G15', cell_engine_g15), 100.0))), xl_add(1.0, xl_div(xl_eval(ctx, 'Engine!G14', cell_engine_g14), 100.0))), xl_eval(ctx, 'Engine!G16', cell_engine_g16))


def cell_outputs_f13(ctx):
    '''Formula: =Engine!G20'''
    return xl_eval(ctx, 'Engine!G20', cell_engine_g20)


def cell_outputs_b14(ctx):
    '''Formula: =B13-B12'''
    return xl_sub(xl_eval(ctx, 'Outputs!B13', cell_outputs_b13), xl_eval(ctx, 'Outputs!B12', cell_outputs_b12))


def cell_outputs_c14(ctx):
    '''Formula: =C13-C12'''
    return xl_sub(xl_eval(ctx, 'Outputs!C13', cell_outputs_c13), xl_eval(ctx, 'Outputs!C12', cell_outputs_c12))


def cell_outputs_d14(ctx):
    '''Formula: =D13-D12'''
    return xl_sub(xl_eval(ctx, 'Outputs!D13', cell_outputs_d13), xl_eval(ctx, 'Outputs!D12', cell_outputs_d12))


def cell_outputs_e14(ctx):
    '''Formula: =E13-E12'''
    return xl_sub(xl_eval(ctx, 'Outputs!E13', cell_outputs_e13), xl_eval(ctx, 'Outputs!E12', cell_outputs_e12))


def cell_outputs_f14(ctx):
    '''Formula: =F13-F12'''
    return xl_sub(xl_eval(ctx, 'Outputs!F13', cell_outputs_f13), xl_eval(ctx, 'Outputs!F12', cell_outputs_f12))


# --- Formula resolver ---
_RESOLVED_FORMULAS = {}
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
    name = _address_to_func_name(address)
    fn = globals().get(name)
    if fn is not None:
        _RESOLVED_FORMULAS[address] = fn
    return fn
