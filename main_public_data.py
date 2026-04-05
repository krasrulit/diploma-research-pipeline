from __future__ import annotations

import argparse
from pathlib import Path

from src.merge_public_data import run_pipeline
from src.utils import load_env_file

PROJECT_ROOT = Path(__file__).resolve().parent


def main() -> None:
    load_env_file(PROJECT_ROOT / ".env")

    parser = argparse.ArgumentParser(
        description="Build additional public-data workbook for the diploma company shortlist.",
    )
    parser.add_argument(
        "--shortlist",
        type=Path,
        required=True,
        help="Path to the shortlist Excel file.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "additional_public_data.xlsx",
        help="Path to the final Excel workbook.",
    )
    parser.add_argument(
        "--processed-dir",
        type=Path,
        default=PROJECT_ROOT / "data_processed",
        help="Directory for processed intermediate outputs.",
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=PROJECT_ROOT / "data_raw",
        help="Directory for downloaded raw files.",
    )
    parser.add_argument(
        "--max-companies",
        type=int,
        default=None,
        help="Optional cap for testing on a small subset of companies.",
    )
    parser.add_argument(
        "--moex-history",
        action="store_true",
        help="Load optional MOEX bond history and save it to data_processed/moex_bond_history.csv.",
    )
    parser.add_argument(
        "--moex-history-from",
        default="2019-01-01",
        help="Start date for optional MOEX bond history.",
    )
    parser.add_argument(
        "--moex-history-to",
        default=None,
        help="End date for optional MOEX bond history.",
    )
    parser.add_argument(
        "--with-cbonds",
        action="store_true",
        help="Enable optional Cbonds loader when credentials are configured via environment variables.",
    )
    parser.add_argument(
        "--with-tinvest",
        action="store_true",
        help="Enable optional T-Invest loader when TINVEST_TOKEN is configured via environment variables.",
    )
    parser.add_argument(
        "--tinvest-coupons",
        action="store_true",
        help="Load optional T-Invest bond coupon schedules for matched bonds.",
    )
    parser.add_argument(
        "--tinvest-coupon-from",
        default="2010-01-01",
        help="Start date for optional T-Invest bond coupon schedule.",
    )
    parser.add_argument(
        "--tinvest-coupon-to",
        default="2035-12-31",
        help="End date for optional T-Invest bond coupon schedule.",
    )
    args = parser.parse_args()

    results = run_pipeline(
        shortlist_path=args.shortlist,
        output_path=args.output,
        processed_dir=args.processed_dir,
        raw_dir=args.raw_dir,
        max_companies=args.max_companies,
        load_moex_history=args.moex_history,
        moex_history_from=args.moex_history_from,
        moex_history_to=args.moex_history_to,
        with_cbonds=args.with_cbonds,
        with_tinvest=args.with_tinvest,
        load_tinvest_coupons=args.tinvest_coupons,
        tinvest_coupon_from=args.tinvest_coupon_from,
        tinvest_coupon_to=args.tinvest_coupon_to,
    )

    print(f"Workbook created: {args.output}")
    for name, frame in results.items():
        print(f"{name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
