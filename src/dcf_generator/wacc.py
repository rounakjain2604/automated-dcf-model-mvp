from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import requests

from .config import WACCConfig


@dataclass
class WACCResult:
    cost_of_equity: float
    cost_of_debt_pre_tax: float
    cost_of_debt_after_tax: float
    wacc: float
    synthetic_rating: str


def fetch_comparable_beta(api_url: str | None = None, api_key: str | None = None) -> float | None:
    if not api_url:
        return None
    headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
    resp = requests.get(api_url, headers=headers, timeout=15)
    resp.raise_for_status()
    payload = resp.json()

    if isinstance(payload, dict) and "beta" in payload:
        return float(payload["beta"])
    if isinstance(payload, list) and payload:
        betas = [float(item.get("beta", 0.0)) for item in payload if item.get("beta") is not None]
        if betas:
            return float(pd.Series(betas).mean())
    return None


def synthetic_credit_spread(interest_coverage_ratio: float) -> tuple[str, float]:
    grid = [
        (8.5, "AAA", 0.007),
        (6.5, "AA", 0.010),
        (5.0, "A", 0.012),
        (4.0, "BBB", 0.015),
        (3.0, "BB", 0.020),
        (2.0, "B", 0.030),
        (1.0, "CCC", 0.045),
        (0.0, "CC", 0.060),
    ]
    for threshold, rating, spread in grid:
        if interest_coverage_ratio >= threshold:
            return rating, spread
    return "D", 0.08


def compute_wacc(cfg: WACCConfig, beta_override: float | None = None) -> WACCResult:
    beta = beta_override if beta_override is not None else cfg.beta
    cost_of_equity = cfg.risk_free_rate + beta * cfg.market_risk_premium + cfg.size_premium + cfg.country_risk_premium

    rating, spread = synthetic_credit_spread(cfg.interest_coverage_ratio)
    cost_of_debt_pre_tax = cfg.risk_free_rate + spread
    cost_of_debt_after_tax = cost_of_debt_pre_tax * (1 - cfg.tax_rate)

    wacc = cfg.target_equity_weight * cost_of_equity + cfg.target_debt_weight * cost_of_debt_after_tax
    return WACCResult(
        cost_of_equity=cost_of_equity,
        cost_of_debt_pre_tax=cost_of_debt_pre_tax,
        cost_of_debt_after_tax=cost_of_debt_after_tax,
        wacc=wacc,
        synthetic_rating=rating,
    )
