from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from .config import ForecastConfig, ScenarioSet


@dataclass
class ForecastResult:
    forecast: pd.DataFrame
    schedules: dict[str, pd.DataFrame]


def build_forecast(mapped_data: pd.DataFrame, cfg: ForecastConfig, scenario: ScenarioSet) -> ForecastResult:
    hist = _build_historical_summary(mapped_data)
    if hist.empty:
        raise ValueError("No usable historical data after mapping.")

    last_period = hist["period"].max()
    base_row = hist.sort_values("period").iloc[-1]

    periods = [pd.Timestamp(last_period) + pd.DateOffset(years=i) for i in range(1, cfg.years + 1)]
    rows: list[dict] = []

    revenue = float(base_row["Revenue"])
    gross_margin_shift = scenario.margin_delta_bps / 10_000

    ppe_existing = max(float(base_row.get("PPE", revenue * 0.2)), 0.0)
    ppe_new = 0.0

    for idx, period in enumerate(periods, start=1):
        growth = _revenue_growth(cfg, idx)
        revenue = revenue * (1 + growth) * scenario.revenue_multiplier

        cogs = _cost_value(cfg.cogs_method, cfg.cogs_pct_revenue, cfg.cogs_fixed, cfg.cogs_inflation, revenue, idx, hist, "COGS")
        opex = _cost_value(cfg.opex_method, cfg.opex_pct_revenue, cfg.opex_fixed, cfg.opex_inflation, revenue, idx, hist, "Operating Expenses")

        cogs = cogs * (1 - gross_margin_shift)
        ebitda = revenue - cogs - opex

        capex = revenue * cfg.capex_pct_revenue if cfg.capex_method == "pct_revenue" else cfg.capex_fixed
        capex *= scenario.capex_multiplier

        dep_existing = ppe_existing * cfg.depreciation_rate
        dep_new = ppe_new * cfg.depreciation_rate
        depreciation = dep_existing + dep_new

        ebit = ebitda - depreciation
        nopat = ebit * (1 - cfg.tax_rate)

        adj_dso = cfg.dso + scenario.working_capital_days_delta
        adj_dpo = cfg.dpo + scenario.working_capital_days_delta
        adj_dio = cfg.dio + scenario.working_capital_days_delta

        ar = revenue * adj_dso / 365
        inv = cogs * adj_dio / 365
        ap = cogs * adj_dpo / 365
        nwc = ar + inv - ap

        prev_nwc = rows[-1]["NWC"] if rows else float(base_row.get("NWC", nwc))
        delta_nwc = nwc - prev_nwc

        fcf = nopat + depreciation - capex - delta_nwc

        rows.append(
            {
                "period": period,
                "Revenue": revenue,
                "COGS": cogs,
                "Operating Expenses": opex,
                "EBITDA": ebitda,
                "Depreciation": depreciation,
                "EBIT": ebit,
                "NOPAT": nopat,
                "Capex": capex,
                "AR": ar,
                "Inventory": inv,
                "AP": ap,
                "NWC": nwc,
                "Delta NWC": delta_nwc,
                "FCF": fcf,
                "PPE Existing": ppe_existing,
                "PPE New": ppe_new,
                "Dep Existing": dep_existing,
                "Dep New": dep_new,
            }
        )

        ppe_existing = max(ppe_existing - dep_existing, 0)
        ppe_new = max(ppe_new + capex - dep_new, 0)

    fc = pd.DataFrame(rows)
    schedules = {
        "capex_dep": fc[["period", "PPE Existing", "PPE New", "Capex", "Dep Existing", "Dep New", "Depreciation"]].copy(),
        "working_capital": fc[["period", "AR", "Inventory", "AP", "NWC", "Delta NWC"]].copy(),
    }
    return ForecastResult(forecast=fc, schedules=schedules)


def _build_historical_summary(mapped_data: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        mapped_data.groupby(["period", "standard_account"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .pivot(index="period", columns="standard_account", values="amount")
        .fillna(0.0)
        .reset_index()
    )

    for col in ["Revenue", "COGS", "Operating Expenses", "Depreciation", "Accounts Receivable", "Inventory", "Accounts Payable"]:
        if col not in grouped.columns:
            grouped[col] = 0.0

    grouped["NWC"] = grouped["Accounts Receivable"] + grouped["Inventory"] - grouped["Accounts Payable"]
    grouped["PPE"] = grouped.get("PPE", pd.Series(np.zeros(len(grouped))))
    return grouped


def _revenue_growth(cfg: ForecastConfig, year_idx: int) -> float:
    if cfg.revenue_method == "cagr":
        return cfg.revenue_cagr
    if cfg.revenue_method == "yoy":
        return cfg.revenue_yoy
    return cfg.revenue_manual.get(year_idx, 0.0)


def _cost_value(method: str, pct: float, fixed: float, inflation: float, revenue: float, year_idx: int, hist: pd.DataFrame, hist_col: str) -> float:
    if method == "pct_revenue":
        return revenue * pct
    if method == "fixed_inflation":
        return fixed * ((1 + inflation) ** max(year_idx - 1, 0))
    hist_series = hist[hist_col] if hist_col in hist.columns else pd.Series([0.0])
    return float(hist_series.tail(3).mean()) if not hist_series.empty else 0.0
