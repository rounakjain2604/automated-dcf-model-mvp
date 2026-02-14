from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd
import requests


@dataclass
class APIConfig:
    base_url: str
    token: str
    company_id: str | None = None


class BaseConnector:
    def __init__(self, config: APIConfig):
        self.config = config

    def _get(self, endpoint: str, params: dict[str, Any] | None = None) -> dict | list:
        headers = {"Authorization": f"Bearer {self.config.token}", "Accept": "application/json"}
        url = f"{self.config.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        resp = requests.get(url, headers=headers, params=params, timeout=20)
        resp.raise_for_status()
        return resp.json()


class QuickBooksConnector(BaseConnector):
    def fetch_trial_balance(self) -> pd.DataFrame:
        payload = self._get("reports/TrialBalance")
        return pd.DataFrame(payload if isinstance(payload, list) else payload.get("rows", []))


class XeroConnector(BaseConnector):
    def fetch_trial_balance(self) -> pd.DataFrame:
        payload = self._get("Reports/TrialBalance")
        return pd.DataFrame(payload if isinstance(payload, list) else payload.get("Reports", []))


class NetSuiteConnector(BaseConnector):
    def fetch_trial_balance(self) -> pd.DataFrame:
        payload = self._get("query/v1/suiteql", params={"q": "SELECT * FROM transaction"})
        return pd.DataFrame(payload if isinstance(payload, list) else payload.get("items", []))
