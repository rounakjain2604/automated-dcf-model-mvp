from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Literal


DiscountConvention = Literal["mid_year", "end_period"]
TerminalMethod = Literal["gordon_growth", "exit_multiple", "both"]


@dataclass
class ScenarioSet:
    revenue_multiplier: float = 1.0
    margin_delta_bps: float = 0.0
    working_capital_days_delta: float = 0.0
    capex_multiplier: float = 1.0


@dataclass
class ForecastConfig:
    years: int = 5
    revenue_method: Literal["cagr", "yoy", "manual"] = "cagr"
    revenue_cagr: float = 0.08
    revenue_yoy: float = 0.07
    revenue_manual: Dict[int, float] = field(default_factory=dict)

    cogs_method: Literal["pct_revenue", "fixed_inflation", "historical_avg"] = "pct_revenue"
    cogs_pct_revenue: float = 0.45
    cogs_fixed: float = 0.0
    cogs_inflation: float = 0.03

    opex_method: Literal["pct_revenue", "fixed_inflation", "historical_avg"] = "pct_revenue"
    opex_pct_revenue: float = 0.25
    opex_fixed: float = 0.0
    opex_inflation: float = 0.03

    dso: float = 45.0
    dpo: float = 40.0
    dio: float = 50.0

    capex_method: Literal["pct_revenue", "fixed"] = "pct_revenue"
    capex_pct_revenue: float = 0.04
    capex_fixed: float = 0.0
    depreciation_rate: float = 0.12
    tax_rate: float = 0.25


@dataclass
class WACCConfig:
    risk_free_rate: float = 0.042
    market_risk_premium: float = 0.055
    beta: float = 1.1
    size_premium: float = 0.0
    country_risk_premium: float = 0.0
    target_debt_weight: float = 0.3
    target_equity_weight: float = 0.7
    interest_coverage_ratio: float = 5.0
    tax_rate: float = 0.25


@dataclass
class ValuationConfig:
    terminal_method: TerminalMethod = "both"
    terminal_growth_rate: float = 0.025
    exit_ev_ebitda_multiple: float = 10.0
    discount_convention: DiscountConvention = "mid_year"
    cash: float = 0.0
    debt: float = 0.0
    minority_interest: float = 0.0
    preferred_stock: float = 0.0
    fully_diluted_shares: float = 1_000_000.0
    gdp_growth_cap: float = 0.035
    terminal_value_blend_weight_gordon: float = 0.5
    terminal_spread_floor_bps: float = 50.0


@dataclass
class DCFConfig:
    forecast: ForecastConfig = field(default_factory=ForecastConfig)
    wacc: WACCConfig = field(default_factory=WACCConfig)
    valuation: ValuationConfig = field(default_factory=ValuationConfig)
    scenarios: Dict[str, ScenarioSet] = field(
        default_factory=lambda: {
            "Base": ScenarioSet(),
            "Bull": ScenarioSet(revenue_multiplier=1.1, margin_delta_bps=150, working_capital_days_delta=-3),
            "Bear": ScenarioSet(revenue_multiplier=0.9, margin_delta_bps=-150, working_capital_days_delta=4),
        }
    )
