# DCF Model Guide (Non-Finance Background)

This document explains what the model does, in plain language.

## 1) What problem this solves
A DCF (Discounted Cash Flow) model estimates what a business is worth **today** based on how much cash it can generate in the future.

Think of it as:
1. Forecast future cash the business can produce.
2. Discount that cash back to today using a risk-adjusted rate.
3. Convert enterprise value to equity value (after debt/cash adjustments).
4. Compare implied value vs current/ask price.

## 2) Main inputs you control
- Revenue and EBITDA (or upload historical financial statements).
- Growth rate assumptions.
- WACC / discount rate.
- Terminal growth rate.
- Capital structure items (cash, debt).

If you leave fields blank in the portal, scenario defaults are used.

## 3) What the model computes
- 5-year operating forecast (revenue, costs, EBITDA, EBIT, NOPAT).
- Reinvestment needs (Capex + working capital).
- Unlevered free cash flow each year.
- Present value of forecast cash flows.
- Terminal value using:
  - Gordon Growth method
  - Exit Multiple method
  - Blended terminal value (weighted mix of both)

## 4) Bank-grade safeguards now included
- **Terminal growth cap:** capped vs GDP-style long-run ceiling.
- **Terminal spread floor:** enforces minimum spread between WACC and terminal growth.
- **Blended valuation output:** reduces single-method bias.
- **Diagnostic outputs:** implied exit multiple from Gordon and implied perpetuity growth from Exit.
- **Audit checks:** terminal FCF positivity, spread sanity, and balance integrity checks.

## 5) How to read deliverables
- **Dashboard tab:** fast decision summary for management.
- **Inputs tab:** all key assumptions (editable cells).
- **Forecast tab:** projected operating and cash-flow lines.
- **Valuation tab:** detailed DCF mechanics and implied prices.
- **Sensitivity tab:** valuation change across WACC / terminal growth combinations.
- **Checks tab:** PASS/ALERT integrity status.
- **Word report:** narrative and visuals for client/investor communication.

## 6) Recommended workflow
1. Start with Base scenario.
2. Validate assumptions in Inputs.
3. Review Dashboard and Checks first.
4. Test sensitivity (WACC and terminal growth).
5. Compare implied blended price to ask/current level.
6. Document assumptions used in final report.

## 7) Important note
This is an analytical model, not investment advice. Output quality depends on input quality and assumption realism.
