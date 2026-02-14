from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from .config import ValuationConfig
from .wacc import WACCResult


@dataclass
class ValuationResult:
    valuation_table: pd.DataFrame
    enterprise_value_gordon: float
    enterprise_value_exit: float
    equity_value_gordon: float
    equity_value_exit: float
    implied_share_price_gordon: float
    implied_share_price_exit: float


def run_dcf(forecast: pd.DataFrame, wacc: WACCResult, cfg: ValuationConfig) -> ValuationResult:
    df = forecast.copy().reset_index(drop=True)
    periods = np.arange(1, len(df) + 1)

    if cfg.discount_convention == "mid_year":
        discount_factors = 1 / ((1 + wacc.wacc) ** (periods - 0.5))
    else:
        discount_factors = 1 / ((1 + wacc.wacc) ** periods)

    df["Discount Factor"] = discount_factors
    df["PV of FCF"] = df["FCF"] * df["Discount Factor"]

    terminal_fcf = float(df.iloc[-1]["FCF"])
    terminal_ebitda = float(df.iloc[-1]["EBITDA"])

    g = cfg.terminal_growth_rate
    gordon_tv = terminal_fcf * (1 + g) / max(wacc.wacc - g, 1e-6)
    exit_tv = terminal_ebitda * cfg.exit_ev_ebitda_multiple

    terminal_discount = float(df.iloc[-1]["Discount Factor"])
    pv_gordon_tv = gordon_tv * terminal_discount
    pv_exit_tv = exit_tv * terminal_discount

    pv_sum = float(df["PV of FCF"].sum())
    ev_gordon = pv_sum + pv_gordon_tv
    ev_exit = pv_sum + pv_exit_tv

    equity_gordon = _enterprise_to_equity(ev_gordon, cfg)
    equity_exit = _enterprise_to_equity(ev_exit, cfg)

    price_gordon = equity_gordon / max(cfg.fully_diluted_shares, 1e-9)
    price_exit = equity_exit / max(cfg.fully_diluted_shares, 1e-9)

    valuation_table = df[["period", "FCF", "Discount Factor", "PV of FCF"]].copy()
    return ValuationResult(
        valuation_table=valuation_table,
        enterprise_value_gordon=ev_gordon,
        enterprise_value_exit=ev_exit,
        equity_value_gordon=equity_gordon,
        equity_value_exit=equity_exit,
        implied_share_price_gordon=price_gordon,
        implied_share_price_exit=price_exit,
    )


def _enterprise_to_equity(enterprise_value: float, cfg: ValuationConfig) -> float:
    return enterprise_value - cfg.debt - cfg.minority_interest - cfg.preferred_stock + cfg.cash
