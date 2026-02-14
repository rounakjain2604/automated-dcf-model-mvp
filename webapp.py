from __future__ import annotations

import io
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt
from flask import Flask, render_template, request, send_file

from src.dcf_generator.config import DCFConfig
from src.dcf_generator.pipeline import run_dcf_pipeline


app = Flask(__name__)


def _to_float(value: str, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _build_cfg_from_form(form) -> DCFConfig:
    cfg = DCFConfig()

    growth = _to_float(form.get("growth_rate", "3"), 3.0) / 100
    wacc = _to_float(form.get("wacc", "24"), 24.0) / 100
    terminal_growth = _to_float(form.get("terminal_growth", "2"), 2.0) / 100
    cash = _to_float(form.get("cash", "50000"), 50000)
    debt = _to_float(form.get("debt", "0"), 0)
    ask_price = _to_float(form.get("ask_price", "4000000"), 4000000)

    cfg.forecast.revenue_method = "cagr"
    cfg.forecast.revenue_cagr = growth
    cfg.forecast.cogs_method = "pct_revenue"
    cfg.forecast.opex_method = "pct_revenue"

    cfg.wacc.target_debt_weight = 0.0
    cfg.wacc.target_equity_weight = 1.0
    cfg.wacc.risk_free_rate = 0.04
    cfg.wacc.beta = 1.0
    cfg.wacc.size_premium = 0.05
    cfg.wacc.country_risk_premium = 0.0
    cfg.wacc.market_risk_premium = max(wacc - (cfg.wacc.risk_free_rate + cfg.wacc.size_premium), 0.01)

    cfg.valuation.terminal_growth_rate = terminal_growth
    cfg.valuation.cash = cash
    cfg.valuation.debt = debt
    cfg.valuation.minority_interest = 0.0
    cfg.valuation.preferred_stock = 0.0
    cfg.valuation.fully_diluted_shares = 1.0
    cfg.valuation.gdp_growth_cap = 0.035
    cfg.valuation.exit_ev_ebitda_multiple = 4.5 if wacc >= 0.24 else 6.0

    cfg._web_meta = {"ask_price": ask_price}  # type: ignore[attr-defined]
    return cfg


def _generate_synthetic_input(form, destination: Path) -> Path:
    revenue = _to_float(form.get("revenue", "2058911"), 2058911)
    ebitda = _to_float(form.get("ebitda", "1072911"), 1072911)

    cogs = revenue * 0.17
    opex = max(revenue - cogs - ebitda, 0)

    historical_revenue = revenue / 1.03
    historical_cogs = historical_revenue * 0.17
    historical_opex = max(historical_revenue - historical_cogs - (ebitda / 1.02), 0)

    df = pd.DataFrame(
        [
            ["2023-12-31", "Revenue", "IS", historical_revenue, False],
            ["2023-12-31", "COGS", "IS", historical_cogs, False],
            ["2023-12-31", "Operating Expenses", "IS", historical_opex, False],
            ["2023-12-31", "Depreciation", "IS", 25000, False],
            ["2023-12-31", "Accounts Receivable", "BS", historical_revenue * 0.11, False],
            ["2023-12-31", "Inventory", "BS", historical_revenue * 0.008, False],
            ["2023-12-31", "Accounts Payable", "BS", historical_revenue * 0.03, False],
            ["2023-12-31", "Cash", "BS", _to_float(form.get("cash", "50000"), 50000), False],
            ["2023-12-31", "Debt", "BS", _to_float(form.get("debt", "0"), 0), False],
            ["2024-12-31", "Revenue", "IS", revenue, False],
            ["2024-12-31", "COGS", "IS", cogs, False],
            ["2024-12-31", "Operating Expenses", "IS", opex, False],
            ["2024-12-31", "Depreciation", "IS", 25000, False],
            ["2024-12-31", "Accounts Receivable", "BS", revenue * 0.11, False],
            ["2024-12-31", "Inventory", "BS", revenue * 0.008, False],
            ["2024-12-31", "Accounts Payable", "BS", revenue * 0.03, False],
            ["2024-12-31", "Cash", "BS", _to_float(form.get("cash", "50000"), 50000), False],
            ["2024-12-31", "Debt", "BS", _to_float(form.get("debt", "0"), 0), False],
        ],
        columns=["period", "account", "statement", "amount", "is_non_recurring"],
    )
    df.to_csv(destination, index=False)
    return destination


def _build_word_report(output_path: Path, company_name: str, result: dict, ask_price: float) -> None:
    valuation = result["valuation_summary"]
    ev_g = float(valuation["Enterprise Value (Gordon)"])
    ev_e = float(valuation["Enterprise Value (Exit)"])

    low = min(ev_g, ev_e)
    high = max(ev_g, ev_e)
    midpoint = (low + high) / 2
    uplift = midpoint - ask_price

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Segoe UI"
    style.font.size = Pt(10.5)

    doc.add_heading(f"PROJECT {company_name.upper()} | VALUATION REPORT", level=1)
    doc.add_paragraph(f"Date: {date.today().isoformat()}")
    doc.add_paragraph("Strictly Private & Confidential | Prepared by Rounak Jain, CFA L2 Candidate")

    doc.add_heading("Executive Summary", level=2)
    doc.add_paragraph(f"Implied Enterprise Value Range: ${low/1_000_000:,.2f}M - ${high/1_000_000:,.2f}M")
    doc.add_paragraph(f"Intrinsic Midpoint: ${midpoint/1_000_000:,.2f}M")
    doc.add_paragraph(f"Ask Price: ${ask_price/1_000_000:,.2f}M")
    doc.add_paragraph(f"Implied Instant Equity Cushion: ${uplift:,.0f}")

    doc.add_heading("Narrative", level=2)
    doc.add_paragraph(
        f"At conservative assumptions, intrinsic value is approximately ${midpoint/1_000_000:,.2f}M. "
        f"Relative to the ask price of ${ask_price/1_000_000:,.2f}M, buyers gain an estimated "
        f"${uplift:,.0f} of immediate equity value."
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/generate")
def generate():
    company_name = request.form.get("company_name", "Company")
    scenario = request.form.get("scenario", "Base")
    ask_price = _to_float(request.form.get("ask_price", "4000000"), 4000000)

    cfg = _build_cfg_from_form(request.form)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        uploaded = request.files.get("financial_file")
        input_path = tmp_path / "input.csv"

        if uploaded and uploaded.filename:
            suffix = Path(uploaded.filename).suffix.lower() or ".csv"
            input_path = tmp_path / f"input{suffix}"
            uploaded.save(input_path)
        else:
            input_path = _generate_synthetic_input(request.form, input_path)

        excel_path = tmp_path / f"{company_name.replace(' ', '_').lower()}_valuation.xlsx"
        word_path = tmp_path / f"{company_name.replace(' ', '_').lower()}_valuation_report.docx"

        result = run_dcf_pipeline(input_path, excel_path, cfg, scenario_name=scenario)
        _build_word_report(word_path, company_name, result, ask_price)

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(excel_path, arcname=excel_path.name)
            zf.write(word_path, arcname=word_path.name)

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"{company_name.replace(' ', '_').lower()}_valuation_pack.zip",
        )


if __name__ == "__main__":
    app.run(debug=True)
