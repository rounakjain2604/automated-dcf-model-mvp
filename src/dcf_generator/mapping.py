from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

import pandas as pd


STANDARD_ACCOUNT_MAP: Dict[str, str] = {
    "revenue": "Revenue",
    "sales": "Revenue",
    "cogs": "COGS",
    "cost of goods": "COGS",
    "opex": "Operating Expenses",
    "sg&a": "Operating Expenses",
    "depreciation": "Depreciation",
    "amortization": "Depreciation",
    "cash": "Cash",
    "accounts receivable": "Accounts Receivable",
    "inventory": "Inventory",
    "accounts payable": "Accounts Payable",
    "debt": "Debt",
    "equity": "Equity",
}


@dataclass
class MappingResult:
    mapped_data: pd.DataFrame
    unmapped_accounts: list[str]


def map_chart_of_accounts(df: pd.DataFrame) -> MappingResult:
    mapped = df.copy()
    mapped["standard_account"] = mapped["account"].apply(_map_account)
    unmapped = sorted(mapped.loc[mapped["standard_account"] == "Other", "account"].unique().tolist())

    mapped["statement"] = mapped.apply(_infer_statement, axis=1)
    return MappingResult(mapped_data=mapped, unmapped_accounts=unmapped)


def _map_account(account: str) -> str:
    key = account.lower().strip()
    for phrase, standard in STANDARD_ACCOUNT_MAP.items():
        if phrase in key:
            return standard
    return "Other"


def _infer_statement(row: pd.Series) -> str:
    if row.get("statement"):
        return str(row["statement"]).upper()

    account = str(row["standard_account"])
    if account in {"Revenue", "COGS", "Operating Expenses", "Depreciation"}:
        return "IS"
    if account in {"Cash", "Accounts Receivable", "Inventory", "Accounts Payable", "Debt", "Equity"}:
        return "BS"
    return "IS"
