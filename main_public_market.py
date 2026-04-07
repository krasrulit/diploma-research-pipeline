from __future__ import annotations

import argparse
import os
from pathlib import Path

import pandas as pd

from src.apply_manual_validation import apply_manual_validation
from src.build_public_market_workbook import build_public_market_workbook
from src.load_tinvest import load_tinvest_optional
from src.merge_public_data import run_pipeline
from src.utils import create_session, load_env_file, save_dataframe_csv


PROJECT_ROOT = Path(__file__).resolve().parent


def main() -> None:
    load_env_file(PROJECT_ROOT / ".env")

    parser = argparse.ArgumentParser(
        description="Build the full public-market workbook without SPARK financial statements.",
    )
    parser.add_argument("--shortlist", type=Path, required=True, help="Path to the shortlist Excel file.")
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "public_market_data_full.xlsx",
        help="Path to the final public-market Excel workbook.",
    )
    parser.add_argument(
        "--public-output",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "additional_public_data.xlsx",
        help="Intermediate workbook with raw public-source outputs.",
    )
    parser.add_argument(
        "--validated-output",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "additional_public_data_validated.xlsx",
        help="Intermediate workbook after applying manual validation.",
    )
    parser.add_argument(
        "--validation-workbook",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "additional_public_data_manual_review_validated.xlsx",
        help="Manual validation workbook. If missing, the final workbook is built without validated mapping.",
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
    parser.add_argument("--max-companies", type=int, default=None, help="Optional cap for smoke tests.")
    parser.add_argument(
        "--skip-download",
        action="store_true",
        help="Reuse existing additional_public_data.xlsx instead of calling public APIs again.",
    )
    parser.add_argument(
        "--refresh-tinvest",
        action="store_true",
        help="With --skip-download, refresh only T-Invest sheets in the existing public workbook.",
    )
    parser.add_argument(
        "--disable-tinvest-ssl-verify",
        action="store_true",
        help="Disable SSL verification only for T-Invest in this run. Useful with local SSL interception.",
    )
    parser.add_argument(
        "--no-tinvest",
        action="store_true",
        help="Disable T-Invest even if TINVEST_TOKEN is configured.",
    )
    parser.add_argument(
        "--no-tinvest-coupons",
        action="store_true",
        help="Do not load the T-Invest coupon schedule for matched bonds.",
    )
    parser.add_argument(
        "--no-moex-history",
        action="store_true",
        help="Do not load MOEX bond trading history.",
    )
    parser.add_argument(
        "--no-moex-share-history",
        action="store_true",
        help="Do not load MOEX share trading history.",
    )
    parser.add_argument("--moex-history-from", default="2019-01-01", help="Start date for MOEX bond history.")
    parser.add_argument("--moex-history-to", default=None, help="End date for MOEX bond history. Defaults to today.")
    parser.add_argument("--moex-share-history-from", default=None, help="Start date for MOEX share history. Defaults to --moex-history-from.")
    parser.add_argument("--moex-share-history-to", default=None, help="End date for MOEX share history. Defaults to today.")
    parser.add_argument(
        "--with-cbonds",
        action="store_true",
        help="Enable optional Cbonds loader when credentials are configured.",
    )
    args = parser.parse_args()

    if args.disable_tinvest_ssl_verify:
        os.environ["TINVEST_SSL_VERIFY"] = "false"

    if not args.skip_download:
        run_pipeline(
            shortlist_path=args.shortlist,
            output_path=args.public_output,
            processed_dir=args.processed_dir,
            raw_dir=args.raw_dir,
            max_companies=args.max_companies,
            load_moex_history=not args.no_moex_history,
            moex_history_from=args.moex_history_from,
            moex_history_to=args.moex_history_to,
            load_moex_share_history=not args.no_moex_share_history,
            moex_share_history_from=args.moex_share_history_from,
            moex_share_history_to=args.moex_share_history_to,
            with_cbonds=args.with_cbonds,
            with_tinvest=not args.no_tinvest,
            load_tinvest_coupons=not args.no_tinvest_coupons,
        )
    elif args.refresh_tinvest and not args.no_tinvest:
        public_sheets = pd.read_excel(args.public_output, sheet_name=None)
        companies_master = public_sheets.get("companies_master", pd.DataFrame())
        tinvest_result = load_tinvest_optional(
            companies_master=companies_master,
            output_dir=args.processed_dir,
            session=create_session(),
            load_coupons=not args.no_tinvest_coupons,
        )
        public_sheets["tinvest_instruments"] = tinvest_result.get("tinvest_instruments", pd.DataFrame())
        public_sheets["tinvest_bonds"] = tinvest_result.get("tinvest_bonds", pd.DataFrame())
        public_sheets["tinvest_bond_coupons"] = tinvest_result.get("tinvest_bond_coupons", pd.DataFrame())

        existing_mapping_log = public_sheets.get("mapping_log", pd.DataFrame())
        if not existing_mapping_log.empty and "source" in existing_mapping_log.columns:
            existing_mapping_log = existing_mapping_log.loc[~existing_mapping_log["source"].eq("tinvest")].copy()
        public_sheets["mapping_log"] = pd.concat(
            [existing_mapping_log, tinvest_result.get("mapping_log", pd.DataFrame())],
            ignore_index=True,
            sort=False,
        )
        public_sheets["download_log"] = pd.concat(
            [public_sheets.get("download_log", pd.DataFrame()), tinvest_result.get("download_log", pd.DataFrame())],
            ignore_index=True,
            sort=False,
        )

        with pd.ExcelWriter(args.public_output, engine="openpyxl") as writer:
            for sheet_name, frame in public_sheets.items():
                frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        for sheet_name in ["tinvest_instruments", "tinvest_bonds", "tinvest_bond_coupons", "mapping_log", "download_log"]:
            save_dataframe_csv(public_sheets[sheet_name], args.processed_dir / f"{sheet_name}.csv")

    if args.validation_workbook.exists():
        apply_manual_validation(
            main_workbook=args.public_output,
            validation_workbook=args.validation_workbook,
            output_workbook=args.validated_output,
        )
        validated_workbook = args.validated_output
    else:
        validated_workbook = None

    outputs = build_public_market_workbook(
        public_workbook=args.public_output,
        validated_workbook=validated_workbook,
        output_workbook=args.output,
        processed_dir=args.processed_dir,
    )

    print(f"Public-market workbook created: {args.output}")
    for name, frame in outputs.items():
        print(f"{name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
