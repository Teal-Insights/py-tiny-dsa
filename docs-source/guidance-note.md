# User Guide for the Tiny Debt Sustainability Analysis Tool (Tiny-DSA)

**Prepared by Teal Insights**

**Version 0.2.0 | May 2026**

---

## CONTENTS

```
ACRONYMS
I.   Introduction
II.  Functional Overview
     A. Spreadsheet management rules
     B. Selecting a country
     C. Setting the baseline parameters
     D. Configuring the shock
     E. Reading the outputs
III. Illustrative Example
IV.  Detailed Discussion of the Methodology
     A. The debt-dynamics identity
     B. The shock specification
     C. A realism check on the outputs
     D. Limitations of the modeling approach
References
```

---

## ACRONYMS

| | |
|---|---|
| DSA | Debt Sustainability Analysis |
| GDP | Gross Domestic Product |
| IMF | International Monetary Fund |
| LIC-DSF | Low-Income Country Debt Sustainability Framework |
| pp | Percentage points |
| PB | Primary balance |
| Q-CRAFT | Quantitative Climate Risk Assessment Fiscal Tool |
| WB | World Bank |

---

## I. Introduction[^1]

[^1]: Tiny-DSA is a methodological artifact prepared by Teal Insights. It is not intended for policy use, is not validated against country-specific data, and is not endorsed by the IMF, the World Bank, or any other institution. The modeling choices follow the standard public-debt accounting literature; the documentation conventions follow the IMF/World Bank LIC-DSF and IMF FAD Q-CRAFT guidance documents.

Tiny-DSA is a stylized debt-sustainability tool that projects a country's general-government debt-to-GDP ratio over a five-year horizon under a baseline macroeconomic and fiscal path and one configurable shock. It is intentionally minimal. The tool implements a single accounting identity, recursed annually for five periods, across three Excel worksheets totaling fewer than sixty formula cells. It does not reproduce the indicator-and-threshold structure of operational debt-sustainability frameworks.

The intended reader is a domain economist familiar with the standard public-debt accounting identity, who needs to understand what the tool computes, what its inputs and outputs mean, and where its simplifications depart from operational practice. The debt ratio and the primary balance are expressed as percentages of GDP. Real GDP growth and the real interest rate are expressed in percent per annum. The tool projects in real terms throughout.

Tiny-DSA omits, deliberately, the contribution of inflation to the debt ratio (treating the recursion in real terms throughout); the exchange-rate dynamics and currency decomposition of debt that drive much of the variation in low-income-country debt trajectories; the residual financing terms that account for stock-flow adjustments, contingent-liability calls, and other below-the-line operations; and any endogenous response of the primary balance to the macroeconomic environment.[^2] The fiscal stance is taken as exogenous, the snowball factor is computed in real terms, and shocks are step changes to a single parameter from a configurable shock-year onwards. These omissions hold the tool small enough to be hand-traceable end-to-end.

[^2]: For the corresponding treatment in operational frameworks, see International Monetary Fund and World Bank (2018), Sections II.B–II.D and Appendix III, and the discussion in Escolano (2010).

This User Guide explains how to set up and use the Tiny-DSA worksheets and how to interpret the outputs. The remainder of the guide is structured as follows. Section II provides a functional overview of the workbook, including its three-sheet architecture, the color conventions, the named ranges that constitute its API, and the operational steps required to select a country, set baseline parameters, and configure a shock. Section III walks through three illustrative scenarios using a fictional country, Borvelia. Section IV provides a detailed discussion of the methodology, including the debt-dynamics identity, the shock specification, a realism check on the outputs, and the limitations of the modeling approach.

---

## II. Functional Overview

Tiny-DSA is organized across three worksheets: **Inputs**, **Engine**, and **Outputs**. The Inputs sheet contains all user-editable parameters and the country profile lookup table. The Engine sheet contains the recursive calculations: the baseline path, the shock-activation logic, the shocked parameter rows, and the shocked path. The Outputs sheet exposes a stable comparison surface for downstream consumers. Every formula on the Engine sheet reads from the Inputs sheet; every formula on the Outputs sheet reads from the Engine sheet.

### A. Spreadsheet management rules

Tiny-DSA has been developed in accordance with strict formatting rules so that the user-editable surface and the calculated surface are visually distinct.

The following coloring and formatting conventions are used:

- Cells with **cornsilk** (pale yellow) fill and a thin border are user-editable inputs. The user is expected to modify these cells to specify a scenario.
- Cells with **pale-violet** fill are lookup-table data. They are not user-editable in normal operation; they are exposed for inspection.
- Cells with **pale-green** fill are calculated outputs that the user should not modify directly. Their values are derived from the inputs through the formulas described in Section IV.
- Cells without fill are either labels (in column A) or intermediate calculations.

Numerical formats follow the convention `0.00` for nonnegative ratios, `0.00;-0.00` for ratios that may be negative, and `+0.00;-0.00` for differences (delta values and shock magnitudes), with the leading sign displayed for clarity.

Excel-level data validation enforces three input constraints. The country name (Inputs!B5) is restricted to values from the country profile table. The shock year (Inputs!B21) is restricted to integers between 1 and 5. The shock type (Inputs!B22) is restricted to integers between 1 and 3. Entries outside these ranges produce an error dialog and are rejected.

Fifteen named ranges expose the workbook's input/output API. These named ranges are stable across versions and constitute the contract that any downstream Python rebuild must honor:

| Named range | Cells | Description |
|---|---|---|
| `country_name` | Inputs!B5 | User-selected country, drawn from the profile table |
| `initial_debt_to_gdp` | Inputs!B6 | Initial debt-to-GDP ratio, looked up by country via INDEX/MATCH |
| `country_profile_table` | Inputs!A10:C12 | Three-row lookup table of country profiles |
| `growth_baseline` | Inputs!C16:G16 | Real GDP growth rates for years 1 through 5 |
| `interest_baseline` | Inputs!C17:G17 | Real interest rates for years 1 through 5 |
| `primary_balance_baseline` | Inputs!C18:G18 | Primary balance for years 1 through 5 |
| `shock_year` | Inputs!B21 | The year in which the shock begins |
| `shock_type` | Inputs!B22 | The parameter affected by the shock (1, 2, or 3) |
| `shock_table` | Inputs!B26:D26 | Shock magnitudes, one per shock type |
| `shock_magnitude_resolved` | Engine!B9 | Shock magnitude for the selected shock type, looked up via OFFSET |
| `baseline_path` | Engine!C6:G6 | Calculated baseline debt-to-GDP path, internal to the Engine sheet |
| `shocked_path` | Engine!C20:G20 | Calculated shocked debt-to-GDP path, internal to the Engine sheet |
| `output_baseline` | Outputs!B12:F12 | Baseline debt-to-GDP path on the Outputs sheet |
| `output_shocked` | Outputs!B13:F13 | Shocked debt-to-GDP path on the Outputs sheet |
| `output_delta` | Outputs!B14:F14 | Difference, shocked minus baseline, in percentage points |

The distinction between the Engine-sheet ranges (`baseline_path`, `shocked_path`) and the Outputs-sheet ranges (`output_baseline`, `output_shocked`) is intentional. The Engine ranges are the computation surface and may be inspected during parity testing or formula-graph extraction. The Outputs ranges are the stable read-only API for downstream consumers and should be the only ranges referenced by external code.

### B. Selecting a country

**The first step in using Tiny-DSA is to choose a country in the Inputs sheet.** The country name is entered in cell `country_name` (Inputs!B5), a cornsilk-filled input cell. Three countries are pre-loaded in the country profile table at Inputs!A10:C12: Borvelia (initial debt 60.0% of GDP, stylized emerging market), Litellia (initial debt 80.0% of GDP, stylized HIPC), and Aurelium (initial debt 40.0% of GDP, stylized advanced economy). Other country names are rejected by the data-validation rule.

The initial debt-to-GDP ratio is then automatically derived from the profile table. Cell `initial_debt_to_gdp` (Inputs!B6) evaluates `INDEX($A$10:$C$12, MATCH($B$5, $A$10:$A$12, 0), 2)`, which returns the second column of the row whose first column matches the selected country name. This is the same INDEX/MATCH lookup pattern that operational tools such as Q-CRAFT and the LIC-DSF use to load country-specific data from a preloaded master table.

The country profile table is intentionally small. It is exposed for inspection and may be extended by adding rows, but the data-validation list on `country_name` would need to be widened correspondingly. In practice, the profile table is read-only.

### C. Setting the baseline parameters

**The baseline scenario is defined by three time-series inputs in the BASELINE PARAMETERS section.** All three are entered as five-cell vectors covering years 1 through 5.

The user enters real GDP growth rates in cells `growth_baseline` (Inputs!C16:G16). Real growth is expressed in percent per annum. The default trajectory is flat at 3.5 percent for each year, consistent with a country in steady-state convergence at moderate pace. Time-varying growth paths are supported by entering different values in each cell.

The user enters real interest rates in cells `interest_baseline` (Inputs!C17:G17). The interest rate represents the effective real rate paid on outstanding general-government debt during the year, in percent per annum. The default trajectory is flat at 4.0 percent, generating a small positive interest-growth differential of 0.5 percentage points.

The user enters the primary balance in cells `primary_balance_baseline` (Inputs!C18:G18). The primary balance is expressed as a percent of GDP, with positive values denoting a surplus. The default trajectory consolidates from a deficit of 1.0 percent of GDP in year 1 to a surplus of 1.0 percent of GDP by year 5, in increments of 0.5 percentage points. This represents a stylized fiscal-consolidation path of the kind frequently assumed in DSA exercises.

Once the country is selected and the baseline parameters are entered, the Engine sheet automatically computes the year-by-year baseline debt-to-GDP trajectory in cells `baseline_path` (Engine!C6:G6).

### D. Configuring the shock

**The shock is defined by three inputs in the SHOCK CONFIGURATION and SHOCK TABLE sections.**

The user enters the shock year in cell `shock_year` (Inputs!B21). The shock year is an integer between 1 and 5 indicating the first year in which the shock takes effect. The shock applies from the shock year onwards through the end of the horizon. The default is 2.

The user enters the shock type in cell `shock_type` (Inputs!B22). The shock type is an integer between 1 and 3 indicating which parameter the shock affects: 1 for real GDP growth, 2 for the real interest rate, and 3 for the primary balance. The default is 1, denoting a growth shock.

The user enters the shock magnitudes in cells `shock_table` (Inputs!B26:D26). The three cells correspond to the three shock types in order. The default magnitudes are -2.0, +2.0, and -1.0 percentage points respectively, denoting a 2-percentage-point reduction in growth, a 2-percentage-point increase in the interest rate, and a 1-percentage-point reduction in the primary balance. Only the magnitude associated with the selected shock type is applied in any given run; the other two values are ignored.

Once the shock is configured, the Engine sheet computes the year-by-year shocked debt-to-GDP trajectory in cells `shocked_path` (Engine!C20:G20). The mechanics of how the magnitude is selected and applied are discussed in Section IV.B.

### E. Reading the outputs

The Outputs sheet reports three time series for years 1 through 5: the baseline debt-to-GDP path, the shocked debt-to-GDP path, and the difference (shocked minus baseline) in percentage points of GDP. The output cells are formatted with pale-green fill to distinguish them from the workbook's intermediate calculations and from the user-editable inputs. Downstream consumers of the workbook should read from the Outputs sheet rather than from the internal `baseline_path` and `shocked_path` ranges on the Engine sheet, which are exposed only for inspection during parity testing.

A summary block at the top of the Outputs sheet reports the country name, the initial debt-to-GDP ratio, the shock year, and the shock type for the active configuration, providing context for the trajectory tables.

---

## III. Illustrative Example

This section walks through three illustrative scenarios for a fictional country, Borvelia, using the canonical defaults. Borvelia is a stylized small open economy with general-government debt of 60 percent of GDP at end of year 0. Borvelia's authorities project flat real GDP growth of 3.5 percent per year over the medium-term, flat real interest rates of 4.0 percent, and a fiscal-consolidation path that swings the primary balance from a deficit of 1.0 percent of GDP in year 1 to a surplus of 1.0 percent of GDP by year 5. Borvelia's macroeconomic and fiscal team is interested in assessing the debt-trajectory consequences of three alternative shocks, each beginning in year 2 and persisting through the end of the five-year horizon: a 2-percentage-point reduction in real growth, a 2-percentage-point increase in the real interest rate, and a 1-percentage-point reduction in the primary balance.

### III.A. The baseline trajectory

Under baseline assumptions, the snowball factor (1 + 0.04) / (1 + 0.035) takes the value 1.00483. Applied recursively to the initial debt stock and net of the consolidating primary balance, the trajectory rises modestly during the deficit years before consolidating as the primary balance turns positive. Year-1 debt is 60.00 × 1.00483 + 1.00 = 61.29 percent of GDP. Year-5 debt is 61.49 percent, only 1.49 percentage points above the starting level despite the cumulative deficits of years 1 and 2. The fiscal consolidation roughly offsets the modestly positive snowball.

### III.B. The growth-shock scenario

A two-percentage-point reduction in real growth from year 2 onwards raises the snowball factor from 1.00483 to (1 + 0.04) / (1 + 0.015) = 1.02463 in years 2 through 5. Year-1 debt is unchanged at 61.29 percent of GDP, since the shock has not yet taken effect. From year 2 onwards the higher snowball factor produces materially higher debt accumulation. Year-5 debt is 66.58 percent of GDP, 5.09 percentage points above the baseline. The fiscal consolidation visible in the baseline is preserved under the shock but is not sufficient to reverse the upward trajectory within the five-year horizon, illustrating the standard finding that permanent reductions in real growth are difficult to offset through primary-balance adjustment alone over short horizons.

### III.C. The interest-rate-shock scenario

A two-percentage-point increase in the real interest rate from year 2 onwards raises the snowball factor from 1.00483 to (1 + 0.06) / (1 + 0.035) = 1.02415 in years 2 through 5. The trajectory under this shock closely tracks the growth-shock trajectory because the snowball factor is nearly identical. Year-5 debt is 66.45 percent of GDP, 4.97 percentage points above the baseline. The proximity of the two shock trajectories reflects the symmetry of the snowball factor in real growth and real interest: a 2-percentage-point movement in either parameter produces a roughly comparable effect on the debt path, with the small difference reflecting the convexity of the (1+r)/(1+g) ratio.

### III.D. The primary-balance-shock scenario

A one-percentage-point reduction in the primary balance from year 2 onwards leaves the snowball factor unchanged at the baseline value of 1.00483 but adds one percentage point of GDP to the deficit (or subtracts it from the surplus) in each of years 2 through 5. Year-5 debt is 65.52 percent of GDP, 4.03 percentage points above the baseline. The shock magnitude is half that of the growth and interest shocks (1pp instead of 2pp), so the smaller debt impact is expected; the channel is also direct rather than amplifying through the snowball, so the cumulative effect is a roughly linear function of the cumulative shock dose.

The full trajectories are reported in Table 1.

**Table 1. Borvelia: Debt-to-GDP Trajectory under Baseline and Three Shock Scenarios (% of GDP)**

| Year | 0 | 1 | 2 | 3 | 4 | 5 |
|---|---|---|---|---|---|---|
| Baseline | 60.00 | 61.29 | 62.09 | 62.39 | 62.19 | 61.49 |
| Growth shock | 60.00 | 61.29 | 63.30 | 64.86 | 65.96 | 66.58 |
| Interest shock | 60.00 | 61.29 | 63.27 | 64.79 | 65.83 | 66.45 |
| Primary balance shock | 60.00 | 61.29 | 63.09 | 64.39 | 65.21 | 65.52 |

A user replicating any one scenario in Excel can verify each cell by hand against the formulas in Section IV.A.

---

## IV. Detailed Discussion of the Methodology

Tiny-DSA is grounded in the standard public-debt accounting identity, recursed annually over a five-year horizon. The shock module is a step-change overlay on a single user-selected parameter, applied symmetrically through the remainder of the horizon from a user-specified shock year. This section describes the identity, the shock specification, a realism check on the outputs, and the limitations of the modeling approach.

### A. The debt-dynamics identity

The evolution of the debt-to-GDP ratio in real terms follows the standard accounting decomposition:[^3]

[^3]: This formulation is the real-rate, real-growth analogue of the textbook nominal debt-dynamics identity. See Escolano (2010) for a derivation from the budget constraint and a discussion of when the real-rate, real-growth representation is preferred. The IMF/World Bank LIC-DSF (International Monetary Fund and World Bank, 2018, Appendix III) treats the equivalent decomposition with currency, inflation, and stock-flow terms made explicit; Tiny-DSA collapses those terms by construction.

```
debt(t) = debt(t-1) × (1 + r(t)) / (1 + g(t))  −  primary_balance(t)
```

where `debt(t)` denotes gross general-government debt as a percent of GDP at end of year *t*, `r(t)` is the average real interest rate paid on outstanding debt during year *t* in percent per annum, `g(t)` is the real growth rate of GDP during year *t* in percent per annum, and `primary_balance(t)` is the primary fiscal balance during year *t* as a percent of GDP, with positive values denoting a surplus.

The factor `(1 + r) / (1 + g)` is the real-terms snowball factor. It captures the automatic dynamics of the debt ratio in the absence of discretionary fiscal adjustment. When the real interest rate exceeds real growth, the snowball factor exceeds unity and the debt ratio rises mechanically over time; when real growth exceeds the real interest rate, the snowball factor is below unity and the debt ratio declines mechanically. The primary-balance term modifies this trajectory by the discretionary fiscal stance: a primary surplus reduces the debt ratio in addition to the snowball, and a primary deficit raises it.

The recursion is implemented in cells `baseline_path` (Engine!C6:G6) of the worksheet for the baseline path. Each cell reads the previous year's debt ratio from the cell immediately to its left on the Engine sheet, and the current year's growth, interest rate, and primary balance from the corresponding columns of `growth_baseline`, `interest_baseline`, and `primary_balance_baseline` on the Inputs sheet via cross-sheet reference. The shocked path in `shocked_path` (Engine!C20:G20) follows the same recursion using the shocked parameter values constructed in rows 14–16 of the Engine sheet. The recursion terminates after year 5; longer horizons are not supported in this version of the tool.

### B. The shock specification

The shock specification proceeds in three steps.

**First, the shock magnitude is selected from the shock table by `OFFSET` against a dynamic column index.** Cell `shock_magnitude_resolved` (Engine!B9) evaluates `OFFSET(Inputs!$B$26, 0, Inputs!$B$22 - 1)`, which walks zero rows down and `shock_type - 1` columns to the right of the anchor at the first cell of the shock table, returning the magnitude associated with the selected shock type. This use of `OFFSET` with a dynamic column argument is methodologically essential to the toy: it exercises the path that any Python rebuild must take through the formula graph in a form that is small enough to verify cell-by-cell but representative of the constraint-resolution problem at scale in tools such as the LIC-DSF.

**Second, an indicator function in row 10 of the Engine sheet determines whether the shock is active in each year.** For each year *t* in cells C10:G10, the cell evaluates `IF(year(t) >= shock_year, 1, 0)`, returning 1 if the year falls within the shock period and 0 otherwise.

**Third, the shock magnitude is applied to the selected parameter via `CHOOSE` against the shock type.** The shocked growth rate in row 14 is computed as `Inputs!{col}16 + CHOOSE(Inputs!$B$22, $B$9, 0, 0) × shock_active(t)`. The `CHOOSE` function returns the shock magnitude only when shock_type equals 1 (the growth case) and zero otherwise; analogous formulas in rows 15 and 16 apply the magnitude to the interest rate and primary balance respectively when shock_type equals 2 or 3. The shock therefore enters one and only one of the three parameter rows in any given run.

The shock is permanent in the sense that, once active, it applies symmetrically through the remainder of the horizon. The magnitude is constant across the shock period, and there is no decay or partial reversion. There is no mechanism for simultaneous multi-parameter shocks, no correlation across parameters, and no propagation of the shock through endogenous channels. A user wishing to model multi-parameter shocks in the current version of the tool would need to enter the corresponding adjustments directly in the baseline parameter rows; the SHOCK CONFIGURATION block supports only one parameter per run.

### C. A realism check on the outputs

A useful sanity check on the baseline trajectory is the comparison of the average primary balance to the debt-stabilizing primary balance.[^4] The debt-stabilizing primary balance is the level of primary balance that, given the snowball factor, holds the debt ratio constant over time:

[^4]: The IMF Q-CRAFT user guide (Tim and Rahman, 2024, p. 13) provides an analogous realism check for productivity assumptions, comparing implied long-run productivity growth to OECD historical averages.

```
pb* = D × (r − g) / (1 + g)
```

where `D` is the debt-to-GDP ratio in steady state, `r` is the real interest rate, and `g` is the real growth rate. At Borvelia's canonical defaults (D = 60, r = 4.0, g = 3.5), pb* equals 60 × 0.005 / 1.035 ≈ 0.29 percent of GDP. The baseline primary-balance path averages 0.0 percent of GDP, somewhat below pb*; the recursion therefore implies a modestly rising debt path in the baseline, consistent with the trajectory reported in Section III. If the team finds that the baseline trajectory either falls or rises substantially without a corresponding gap between the average primary balance and pb*, the discrepancy is a signal that the recursion is mis-specified or that one of the parameter rows has an unintended sign convention.

This check is a useful template for the verification practice that the rebuild team should apply at every scale: outputs verify against an economic-logic check, not only against a formula-level golden master. A formula-level rebuild can produce numerically correct outputs while drifting from economic logic — for example by inverting the sign of the primary balance or by applying the snowball factor in nominal terms. Realism checks of this kind are a low-cost defense against that class of error.

### D. Limitations of the modeling approach

Tiny-DSA is a methodological exercise, not a debt-sustainability framework. Its limitations are first-order for any operational use.

The tool projects in real terms throughout, treating inflation as residual to the real-rate, real-growth specification. Operational frameworks decompose the debt ratio into its real-rate, real-growth, and inflation components separately, in part because the inflation channel is a substantial driver of debt-ratio variation in countries with high or volatile inflation, and in part because nominal magnitudes are required for translation into local-currency debt-stock projections.

The tool does not separate debt by currency, holder, or instrument category. The exchange-rate channel, which in low-income-country debt sustainability analysis is frequently first-order through the revaluation of foreign-currency-denominated debt stocks, is absent. The IMF/World Bank LIC-DSF treats the external and domestic debt universes separately and applies a depreciation stress test as one of its standard scenarios; Tiny-DSA collapses these distinctions to a single ratio.

The primary balance is exogenous in Tiny-DSA. It does not respond to the debt path, the macroeconomic environment, or the shock. In operational frameworks, fiscal rules and policy reaction functions allow the primary balance to adjust endogenously, and this adjustment can substantially modify the projected debt trajectory under stress.

The tool does not represent residual financing terms, contingent liabilities, or stock-flow adjustments. In operational analyses these are nontrivial: the IMF/World Bank LIC-DSF Guidance Note (Appendix V) documents the residual-financing framework, and the 2024 supplement extends it. Tiny-DSA omits these terms by construction.

The country profile lookup table contains three stylized profiles, each differing only in initial debt-to-GDP. Operational frameworks load 100+ country-specific datasets covering historical macroeconomic and fiscal series, demographic projections, and country-specific calibrations. The lookup pattern in Tiny-DSA exercises the structural mechanism (INDEX/MATCH against a master table) but at trivial data scale.

Finally, Tiny-DSA does not produce indicator-and-threshold output, does not classify the country into a debt-distress risk category, and does not generate any of the policy-relevant signals that operational DSA frameworks produce. The output is a single ratio path under a baseline and a single shocked scenario. For operational analysis of low-income-country debt sustainability, users should refer to the IMF/World Bank LIC-DSF and the IMF FAD Q-CRAFT tool.

---

## References

Batini, N., Di Serio, M., Fragetta, M., and Gomes, S. (2024). *Q-CRAFT: A Quantitative Climate Risk Assessment Fiscal Tool*. IMF Working Paper. International Monetary Fund Fiscal Affairs Department.

Escolano, J. (2010). *A Practical Guide to Public Debt Dynamics, Fiscal Sustainability, and Cyclical Adjustment of Budgetary Aggregates*. IMF Technical Notes and Manuals 10/02. International Monetary Fund.

International Monetary Fund and World Bank (2018). *Guidance Note on the Bank-Fund Debt Sustainability Framework for Low-Income Countries*. Policy Paper. February 2018.

International Monetary Fund and World Bank (2024). *Supplement to 2018 Guidance Note on the Bank-Fund Debt Sustainability Framework for Low-Income Countries*. Policy Paper. July 2024.

Tim, T., and Rahman, J. (2024). *User Guide for the Quantitative Climate Risk Assessment Fiscal Tool (Q-CRAFT)*. International Monetary Fund Fiscal Affairs Department. October 2024.

---

*Reuse of this User Guide and the Tiny-DSA tool does not imply any endorsement of any resulting research and/or product.*
