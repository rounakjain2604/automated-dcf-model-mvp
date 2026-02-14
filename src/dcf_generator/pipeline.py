from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

import pandas as pd

from .config import DCFConfig
from .excel_export import export_workbook
from .forecast import build_forecast
from .ingestion import ingest_financials
from .mapping import map_chart_of_accounts
from .normalization import normalize_non_recurring
from .valuation import run_dcf
from .wacc import compute_wacc, fetch_comparable_beta


def run_dcf_pipeline(input_path: str | Path, output_path: str | Path, cfg: DCFConfig, scenario_name: str = "Base") -> dict:
    ingested = ingest_financials(input_path)
    mapped = map_chart_of_accounts(ingested.raw_data)
    normalized = normalize_non_recurring(mapped.mapped_data)

    scenario = cfg.scenarios.get(scenario_name)
    if scenario is None:
        raise ValueError(f"Unknown scenario '{scenario_name}'. Available: {list(cfg.scenarios.keys())}")

    forecast_result = build_forecast(normalized, cfg.forecast, scenario)

    beta = None
    try:
        beta = fetch_comparable_beta()
    except Exception:
        beta = None

    wacc = compute_wacc(cfg.wacc, beta_override=beta)
    valuation = run_dcf(forecast_result.forecast, wacc, cfg.valuation)

    audit = _build_checks(forecast_result.forecast, cfg)

    valuation_summary = {
        "WACC": wacc.wacc,
        "Enterprise Value (Gordon)": valuation.enterprise_value_gordon,
        "Enterprise Value (Exit)": valuation.enterprise_value_exit,
        "Equity Value (Gordon)": valuation.equity_value_gordon,
        "Equity Value (Exit)": valuation.equity_value_exit,
        "Implied Price (Gordon)": valuation.implied_share_price_gordon,
        "Implied Price (Exit)": valuation.implied_share_price_exit,
    }

    period_meta = {
        "period_basis": ingested.period_basis,
        "has_stub_period": ingested.has_stub_period,
    }

    historical_revenue = (
        normalized.loc[normalized["standard_account"] == "Revenue", ["period", "amount"]]
        .groupby("period", as_index=False)["amount"]
        .sum()
        .sort_values("period")
    )
    historical_growth_3y_avg = 0.0
    if len(historical_revenue) >= 2:
        growth = historical_revenue["amount"].pct_change().dropna()
        if not growth.empty:
            historical_growth_3y_avg = float(growth.tail(3).mean())

    export_workbook(
        output_path,
        cfg=cfg,
        scenario_name=scenario_name,
        scenario=scenario,
        period_meta=period_meta,
        forecast_df=forecast_result.forecast,
        wacc_result=wacc,
        valuation_summary=valuation_summary,
        historical_growth_3y_avg=historical_growth_3y_avg,
    )

    return {
        "period_meta": period_meta,
        "unmapped_accounts": mapped.unmapped_accounts,
        "valuation_summary": valuation_summary,
        "audit": audit,
        "scenario": scenario_name,
        "config": asdict(cfg),
        "forecast_rows": forecast_result.forecast.to_dict(orient="records"),
        "historical_growth_3y_avg": historical_growth_3y_avg,
    }


def _build_checks(forecast_df: pd.DataFrame, cfg: DCFConfig) -> pd.DataFrame:
    rows = []

    for _, row in forecast_df.iterrows():
        assets = row["AR"] + row["Inventory"]
        liabilities_plus_equity = row["AP"] + max(assets - row["AP"], 0)
        balance_gap = assets - liabilities_plus_equity

        rows.append(
            {
                "period": row["period"],
                "check": "Balance Sheet Check",
                "value": balance_gap,
                "status": "PASS" if abs(balance_gap) < 1e-6 else "FAIL",
            }
        )

    rows.append(
        {
            "period": forecast_df.iloc[-1]["period"],
            "check": "Terminal Growth <= GDP Growth Cap",
            "value": cfg.valuation.terminal_growth_rate,
            "status": "PASS" if cfg.valuation.terminal_growth_rate <= cfg.valuation.gdp_growth_cap else "ALERT",
        }
    )

    if cfg.valuation.terminal_growth_rate > cfg.valuation.gdp_growth_cap:
        rows.append(
            {
                "period": forecast_df.iloc[-1]["period"],
                "check": "Sanity Alert",
                "value": "Terminal Growth exceeds GDP cap",
                "status": "ALERT",
            }
        )

    return pd.DataFrame(rows)
