# Automated DCF Model Generator (Premium Delivery Guide)

This project generates a **sell-ready, formula-linked Excel DCF model** from raw financial data.

The output workbook is designed to look and behave like an investment banking / transaction model:

- inputs in blue font,
- linked formulas in black font,
- accounting number formats,
- sensitivity heatmap,
- executive dashboard,
- print areas/scaling,
- audit checks.

### Premium workbook design (new)
- Executive-first workbook order: `Dashboard -> Inputs -> Forecast -> Valuation -> Sensitivity -> Checks`.
- Cohesive investment-banking style theme (title bars, tab colors, section cards, alternating row bands).
- KPI cards for WACC, EV, and implied price to improve first-look readability.
- Expanded 5x5 sensitivity matrix with a cleaner heatmap for investment committee use.
- Enhanced checks tab with conditional pass/fail highlighting for quick QA sign-off.

---

## 1) What this model does

### Data ingestion & setup
- Accepts `.csv`, `.xlsx`.
- Includes API connector scaffolding for QuickBooks / Xero / NetSuite in `src/dcf_generator/api_connectors.py`.
- Maps chart of accounts to standardized model lines.
- Detects fiscal/calendar basis and stub-period signals.
- Normalizes non-recurring items into EBITDA add-backs.

### Forecasting engine
- Revenue driver via CAGR / YoY / manual inputs (config-driven).
- COGS and Opex via `% of Revenue`, fixed+inflation, or historical average.
- Working capital logic (DSO/DPO/DIO).
- Capex and depreciation schedules.

### WACC module
- CAPM-based cost of equity.
- Synthetic rating based cost of debt from interest coverage.
- Debt/equity weighting.
- Optional API beta fetching hook.

### DCF valuation core
- Mid-year or end-period discounting.
- Gordon Growth and Exit Multiple terminal valuation.
- Enterprise-to-equity bridge (cash, debt, minority, preferred).
- Implied share price output.

### Scenario + sensitivity + checks
- Base / Bull / Bear scenarios.
- 2D sensitivity table (WACC vs terminal growth).
- Balance check and terminal growth sanity check.

---

## 2) Project structure

- `src/dcf_generator/main.py` — CLI entrypoint.
- `src/dcf_generator/pipeline.py` — orchestration.
- `src/dcf_generator/excel_export.py` — formula-linked Excel writer.
- `data/sample_financials.csv` — sample input.
- `config.example.json` — example configuration override.

---

## 3) Install and run

## Requirements
- Python 3.11+ (tested with local venv).

## Install
```bash
pip install -r requirements.txt
```

## Client-facing web portal
Launch a browser-based interface to collect inputs and generate files for clients:

```bash
python webapp.py
```

Then open:

```text
http://127.0.0.1:5000
```

Windows shortcut launcher:

```bash
run_portal.bat
```

Important:
- Do **not** open `templates/index.html` directly from File Explorer (`file:///...`).
- Always run through Flask (`http://127.0.0.1:5000`), otherwise `/generate` will fail.

## Deploy to Vercel (use anytime)

Yes, this app can be hosted on Vercel.

### Files already added
- `api/index.py` (Vercel Python serverless entrypoint)
- `vercel.json` (routing/build config)

### Deploy steps
1. Push this project to a GitHub repo.
2. In Vercel, click **Add New Project** and import the repo.
3. Framework preset: **Other**.
4. Build/Output commands: leave default (Vercel uses `vercel.json`).
5. Deploy.

### Local test before deploy
```bash
vercel dev
```

### Practical notes for Vercel serverless
- Large uploads and very large generated files may hit serverless request/response limits.
- Long-running model generation can hit execution timeout on lower plans.
- If you expect bigger files or heavy concurrent use, move generation artifacts to object storage (S3/Vercel Blob) and return a download link.

What it does:
- Accepts uploaded `.csv/.xlsx` financials OR quick manual inputs.
- Generates a premium Excel model and Word valuation memo.
- Returns a downloadable ZIP file containing both deliverables.

Word report quality (premium mode):
- 1-minute executive page with football-field scenario chart and KPI table.
- Dynamic narrative text based on WACC, terminal growth, and upside.
- WACC sanity-check breakdown and growth-vs-history commentary.
- Financial health visuals (revenue bridge and margin trend charts) + FCF build table.
- Sensitivity/scenario section with risk-warning narrative.
- Peer comparison table with premium/discount context.
- Optional logo in report header (upload in portal).

## Run with default assumptions
```bash
python -m src.dcf_generator.main --input data/sample_financials.csv --output output/dcf_model.xlsx --scenario Base
```

## Run with custom config
```bash
python -m src.dcf_generator.main --input data/sample_financials.csv --output output/dcf_model_custom.xlsx --scenario Bull --config config.example.json
```

---

## 4) Input file contract

Required columns:
- `period`
- `account`
- `amount`

Optional but recommended:
- `statement` (`IS`/`BS`)
- `is_non_recurring` (`true`/`false`)

Example:
```csv
period,account,statement,amount,is_non_recurring
2023-12-31,Revenue,IS,1500000,false
2023-12-31,Restructuring Charge,IS,40000,true
```

---

## 5) Excel output (what the client receives)

The generated workbook includes:

1. `Inputs` — all model levers (blue input cells).
2. `Forecast` — formula-linked annual projections.
3. `Valuation` — fully linked DCF bridge and implied price.
4. `Sensitivity` — formula-driven valuation matrix with heatmap.
5. `Dashboard` — executive summary and football-field style comparison chart.
6. `Checks` — balance and sanity controls.

### Formula-linking standard
- Forecast formulas reference only `Inputs` and prior forecast rows.
- Valuation formulas reference `Forecast` + `Inputs` (no hardcoded calculation constants).
- Sensitivity cells are formula-linked to valuation drivers.
- Dashboard cells reference valuation output cells.
- Checks tab references forecast and input cells.

### Visual QA standard for premium delivery
- No blank-looking sheets: every tab has title/header styling and bounded tables.
- Every numeric section uses explicit formats (percent, accounting, share price).
- Critical statuses are color-signaled (`PASS` green, `FAIL/ALERT` red).
- Default print setup fits key tabs into client-review page ranges.

---

## 6) Professional formatting standard

- Inputs: blue font.
- Formulas: black font.
- Numbers: accounting format
  `_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)`.
- Header styling and borders for auditability.
- Freeze panes enabled on key tabs.
- Print area + fit-to-page settings for presentation exports.

---

## 7) How to sell this as a $4,000 model

You can’t guarantee a fixed market price from code alone, but this package now supports a **premium positioning** if you deliver it correctly.

Use this delivery bundle for each client:

1. **Customized assumptions file** (`config.clientname.json`).
2. **Client-branded workbook output** (`dcf_clientname_v1.xlsx`).
3. **Assumption memo** (1–2 pages: data source, drivers, WACC logic).
4. **QA proof pack**:
   - screenshots of formula tracing,
   - checks tab all PASS,
   - scenario and sensitivity outputs.
5. **Change log** (`v1`, `v2`) to demonstrate iterative value.

### Suggested commercial framing
- Base model delivery: $2,000–$2,500.
- Data integration + mapping customization: $750–$1,000.
- Scenario and IC-ready sensitivity pack + walkthrough: $750–$1,000.

That naturally supports a ~$4,000 engagement when combined.

---

## 8) QA checklist before sending to a buyer

- Run all three scenarios (`Base`, `Bull`, `Bear`).
- Confirm `Checks` tab has no FAIL / ALERT unless intentionally flagged.
- Open Excel and use trace precedents on random model cells.
- Verify print layout on `Dashboard`, `Valuation`, `Sensitivity`.
- Confirm no broken references (`#REF!`, `#DIV/0!`, `#N/A`).

---

## 9) Extension roadmap (if you want higher-ticket deals)

- Live authenticated API ingestion and tenant mapping templates.
- Monthly/quarterly forecast support with true stub-period math.
- Debt schedule with revolver, mandatory amortization, and cash sweep.
- Monte Carlo simulation and probabilistic valuation bands.
- Auto-generated investment committee PDF report pack.
