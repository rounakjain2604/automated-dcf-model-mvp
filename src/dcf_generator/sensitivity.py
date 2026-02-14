from __future__ import annotations

import pandas as pd

from .config import ValuationConfig
from .valuation import run_dcf
from .wacc import WACCResult


def build_sensitivity_table(
    forecast: pd.DataFrame,
    base_wacc: WACCResult,
    valuation_cfg: ValuationConfig,
    wacc_values: list[float],
    growth_values: list[float],
) -> pd.DataFrame:
    table = pd.DataFrame(index=[f"g={g:.2%}" for g in growth_values], columns=[f"wacc={w:.2%}" for w in wacc_values])

    for g in growth_values:
        for w in wacc_values:
            updated_wacc = WACCResult(
                cost_of_equity=base_wacc.cost_of_equity,
                cost_of_debt_pre_tax=base_wacc.cost_of_debt_pre_tax,
                cost_of_debt_after_tax=base_wacc.cost_of_debt_after_tax,
                wacc=w,
                synthetic_rating=base_wacc.synthetic_rating,
            )
            cfg = ValuationConfig(**{**valuation_cfg.__dict__, "terminal_growth_rate": g})
            result = run_dcf(forecast, updated_wacc, cfg)
            table.loc[f"g={g:.2%}", f"wacc={w:.2%}"] = result.implied_share_price_gordon

    return table
