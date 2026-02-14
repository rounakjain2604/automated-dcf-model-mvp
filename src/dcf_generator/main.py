from __future__ import annotations

import argparse
import json
from pathlib import Path

from .config import DCFConfig
from .pipeline import run_dcf_pipeline


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automated DCF Model Generator")
    parser.add_argument("--input", required=True, help="Path to input CSV/XLSX financial file")
    parser.add_argument("--output", required=True, help="Path to output Excel model")
    parser.add_argument("--scenario", default="Base", choices=["Base", "Bull", "Bear"], help="Scenario toggle")
    parser.add_argument("--config", required=False, help="Optional JSON config override file")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    cfg = DCFConfig()

    if args.config:
        cfg = _load_config_override(args.config, cfg)

    result = run_dcf_pipeline(args.input, args.output, cfg, scenario_name=args.scenario)

    print("DCF model generated successfully")
    print(f"Output: {args.output}")
    print(json.dumps(result["valuation_summary"], indent=2, default=str))


def _load_config_override(config_path: str, current_cfg: DCFConfig) -> DCFConfig:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    payload = json.loads(path.read_text(encoding="utf-8"))

    if "forecast" in payload:
        for key, value in payload["forecast"].items():
            setattr(current_cfg.forecast, key, value)

    if "wacc" in payload:
        for key, value in payload["wacc"].items():
            setattr(current_cfg.wacc, key, value)

    if "valuation" in payload:
        for key, value in payload["valuation"].items():
            setattr(current_cfg.valuation, key, value)

    return current_cfg


if __name__ == "__main__":
    main()
