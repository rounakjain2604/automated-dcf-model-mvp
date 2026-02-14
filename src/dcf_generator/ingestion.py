from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import pandas as pd


PeriodBasis = Literal["fiscal", "calendar"]


@dataclass
class IngestionResult:
    raw_data: pd.DataFrame
    period_basis: PeriodBasis
    has_stub_period: bool


def load_source(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    elif path.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
    else:
        raise ValueError("Unsupported file format. Use .csv, .xlsx, or .xls")

    required = {"period", "account", "amount"}
    missing = required.difference(df.columns.str.lower())
    if missing:
        normalized = {col.lower(): col for col in df.columns}
        if any(col not in normalized for col in required):
            raise ValueError(f"Missing required columns: {sorted(missing)}")

    return _normalize_columns(df)


def detect_period_characteristics(df: pd.DataFrame) -> tuple[PeriodBasis, bool]:
    periods = pd.to_datetime(df["period"], errors="coerce")
    month_counts = periods.dt.month.value_counts(dropna=True)
    fiscal = "fiscal" if (not month_counts.empty and month_counts.index[0] != 12) else "calendar"

    if periods.dropna().empty:
        return fiscal, False

    ordered = periods.dropna().sort_values().unique()
    if len(ordered) < 2:
        return fiscal, False

    diffs_days = pd.Series(ordered[1:]) - pd.Series(ordered[:-1])
    median_gap = diffs_days.dt.days.median()
    has_stub = bool(median_gap and abs(median_gap - 365) > 40)
    return fiscal, has_stub


def ingest_financials(path: str | Path) -> IngestionResult:
    df = load_source(path)
    period_basis, has_stub = detect_period_characteristics(df)
    return IngestionResult(raw_data=df, period_basis=period_basis, has_stub_period=has_stub)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {col: col.strip().lower() for col in df.columns}
    out = df.rename(columns=renamed).copy()
    out["period"] = pd.to_datetime(out["period"], errors="coerce")
    out["amount"] = pd.to_numeric(out["amount"], errors="coerce").fillna(0.0)
    if "statement" not in out.columns:
        out["statement"] = ""
    if "is_non_recurring" not in out.columns:
        out["is_non_recurring"] = False
    out["is_non_recurring"] = out["is_non_recurring"].astype(bool)
    out["account"] = out["account"].astype(str)
    return out
