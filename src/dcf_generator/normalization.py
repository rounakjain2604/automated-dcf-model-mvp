from __future__ import annotations

import re

import pandas as pd


NON_RECURRING_PATTERNS = [
    r"one[- ]time",
    r"restructur",
    r"settlement",
    r"impairment",
    r"gain/loss",
]


def normalize_non_recurring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    pattern = re.compile("|".join(NON_RECURRING_PATTERNS), re.IGNORECASE)

    inferred_non_recurring = out["account"].astype(str).str.contains(pattern, na=False)
    non_recurring_flag = out["is_non_recurring"] | inferred_non_recurring

    out["ebitda_add_back"] = 0.0
    mask = non_recurring_flag & out["statement"].eq("IS")
    out.loc[mask, "ebitda_add_back"] = out.loc[mask, "amount"]
    return out
