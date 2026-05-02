from __future__ import annotations

import argparse
import os
from pathlib import Path

import pandas as pd

from src.build_analysis_panel import (
    DEFAULT_METALLURGY_SHORTLIST,
    DEFAULT_OIL_GAS_SHORTLIST,
    combine_shortlists,
)
from src.build_public_securities_workbook import (
    build_public_securities_outputs,
    normalize_share_manual_review,
    read_csv_if_exists,
    write_public_securities_workbook,
)
from src.load_moex import load_moex_data
from src.load_tinvest import load_tinvest_optional
from src.utils import clean_text, create_session, ensure_directory, load_env_file, normalize_identifier, normalize_isin, normalize_ticker, save_dataframe_csv


PROJECT_ROOT = Path(__file__).resolve().parent


def read_share_manual_review(path: Path | None) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="review_queue")
    except ValueError:
        return pd.DataFrame()


def read_moex_review_workbook(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

    quote_sheet = "quotes_clean" if "quotes_clean" in xl.sheet_names else ""
    if not quote_sheet and "quotes_clean_all" in xl.sheet_names:
        quote_sheet = "quotes_clean_all"
    summary_sheet = "security_summary" if "security_summary" in xl.sheet_names else ""
    if not summary_sheet and "security_summary_all" in xl.sheet_names:
        summary_sheet = "security_summary_all"

    quotes = pd.read_excel(path, sheet_name=quote_sheet) if quote_sheet else pd.DataFrame()
    summary = pd.read_excel(path, sheet_name=summary_sheet) if summary_sheet else pd.DataFrame()
    return quotes, summary


def normalize_moex_review_quotes(quotes: pd.DataFrame) -> pd.DataFrame:
    if quotes.empty:
        return pd.DataFrame()
    out = quotes.copy()
    column_map = {
        "ticker": "SECID",
        "board": "BOARDID",
        "shortname": "SHORTNAME",
    }
    for source_col, target_col in column_map.items():
        if target_col not in out.columns and source_col in out.columns:
            out[target_col] = out[source_col]
    for column in ["OPEN", "LOW", "HIGH", "CLOSE", "LEGALCLOSEPRICE", "WAPRICE", "VOLUME", "VALUE", "NUMTRADES"]:
        if column not in out.columns:
            out[column] = pd.NA
    out["SECID"] = out.get("SECID", "").map(normalize_ticker)
    out["BOARDID"] = out.get("BOARDID", "").map(clean_text)
    out["SHORTNAME"] = out.get("SHORTNAME", "").map(clean_text)
    out["TRADEDATE"] = pd.to_datetime(out.get("TRADEDATE"), errors="coerce")
    out = out.loc[out["SECID"].ne("") & out["TRADEDATE"].notna()].copy()
    return out


def build_moex_share_review_history(
    quotes_dir: Path | None,
    share_manual_review: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    manual = normalize_share_manual_review(share_manual_review)
    if quotes_dir is None or not quotes_dir.exists() or manual.empty:
        return pd.DataFrame(), pd.DataFrame()

    manual = manual.loc[manual["manual_share_usable_flag"]].copy()
    if manual.empty:
        return pd.DataFrame(), pd.DataFrame()

    history_rows: list[pd.DataFrame] = []
    summary_rows: list[pd.DataFrame] = []

    for workbook_path in sorted(quotes_dir.glob("*_moex_quotes_review.xlsx")):
        if workbook_path.name.startswith("all_") or workbook_path.name.startswith("~$"):
            continue
        quotes_raw, summary_raw = read_moex_review_workbook(workbook_path)
        quotes = normalize_moex_review_quotes(quotes_raw)
        if quotes.empty:
            continue

        summary = summary_raw.copy()
        summary_tickers = set()
        summary_isins = set()
        if not summary.empty:
            if "ticker" in summary.columns:
                summary["ticker"] = summary["ticker"].map(normalize_ticker)
                summary_tickers.update(summary["ticker"].dropna().astype(str).tolist())
            if "isin" in summary.columns:
                summary["isin"] = summary["isin"].map(normalize_isin)
                summary_isins.update(summary["isin"].dropna().astype(str).tolist())

        quote_tickers = set(quotes["SECID"].dropna().astype(str).map(normalize_ticker).tolist())
        if not summary_tickers:
            summary_tickers.update(quote_tickers)

        matched = manual.loc[
            manual["ticker"].isin(summary_tickers | quote_tickers)
            | manual["isin"].isin(summary_isins)
            | manual["ticker"].isin(summary_isins)
        ].copy()
        if matched.empty:
            continue

        for _, review_row in matched.iterrows():
            expanded = quotes.copy()
            expanded["company_id"] = clean_text(review_row.get("company_id", ""))
            expanded["company_name"] = clean_text(review_row.get("company_name", ""))
            expanded["company_inn"] = normalize_identifier(review_row.get("company_inn", ""), length=10)
            expanded["sector"] = clean_text(review_row.get("sector", ""))
            expanded["sample_flag"] = clean_text(review_row.get("sample_flag", ""))
            expanded["match_score"] = 1_000
            expanded["match_confidence"] = "manual_usable"
            expanded["selected_flag"] = True
            expanded["exact_identifier_match"] = True
            expanded["match_reasons"] = "manual_share_review_usable"
            expanded["manual_share_review_source"] = clean_text(review_row.get("manual_share_review_source", ""))
            expanded["manual_review_ticker"] = clean_text(review_row.get("ticker", ""))
            expanded["manual_review_isin"] = clean_text(review_row.get("isin", ""))
            expanded["ticker"] = clean_text(review_row.get("ticker", ""))
            expanded["isin"] = clean_text(review_row.get("isin", ""))
            expanded["instrument_name"] = clean_text(review_row.get("instrument_name", ""))
            expanded["canonical_security_key"] = clean_text(review_row.get("canonical_security_key", ""))
            history_rows.append(expanded)

            if not summary.empty:
                summary_expanded = summary.copy()
                summary_expanded["source_file"] = workbook_path.name
                summary_expanded["company_id"] = clean_text(review_row.get("company_id", ""))
                summary_expanded["company_inn"] = normalize_identifier(review_row.get("company_inn", ""), length=10)
                summary_expanded["manual_review_ticker"] = clean_text(review_row.get("ticker", ""))
                summary_expanded["manual_review_isin"] = clean_text(review_row.get("isin", ""))
                summary_rows.append(summary_expanded)

    history = pd.concat(history_rows, ignore_index=True, sort=False) if history_rows else pd.DataFrame()
    summary_out = pd.concat(summary_rows, ignore_index=True, sort=False) if summary_rows else pd.DataFrame()
    if not history.empty:
        history = history.drop_duplicates(
            subset=["company_id", "company_inn", "SECID", "BOARDID", "TRADEDATE"],
            keep="last",
        )
    if not summary_out.empty:
        summary_out = summary_out.drop_duplicates(
            subset=[column for column in ["company_id", "company_inn", "ticker", "isin"] if column in summary_out.columns],
            keep="last",
        )
    return history, summary_out


def main() -> None:
    load_env_file(PROJECT_ROOT / ".env")

    parser = argparse.ArgumentParser(
        description="Collect MOEX and T-Invest securities data for all oil&gas and metallurgy companies.",
    )
    parser.add_argument(
        "--oil-gas-shortlist",
        type=Path,
        default=DEFAULT_OIL_GAS_SHORTLIST,
        help="Path to the oil&gas shortlist workbook.",
    )
    parser.add_argument(
        "--metallurgy-shortlist",
        type=Path,
        default=DEFAULT_METALLURGY_SHORTLIST,
        help="Path to the metallurgy shortlist workbook.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "public_securities_all_companies.xlsx",
        help="Final workbook with unified securities metadata and coverage.",
    )
    parser.add_argument(
        "--processed-dir",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "public_securities_all",
        help="Directory for processed sidecar CSV files.",
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
        help="Reuse already downloaded sidecar CSVs from --processed-dir and only rebuild the workbook.",
    )
    parser.add_argument(
        "--disable-tinvest-ssl-verify",
        action="store_true",
        help="Disable SSL verification for T-Invest in this run.",
    )
    parser.add_argument(
        "--no-tinvest",
        action="store_true",
        help="Disable T-Invest loading even when token is configured.",
    )
    parser.add_argument(
        "--no-tinvest-history",
        action="store_true",
        help="Skip T-Invest historical candles.",
    )
    parser.add_argument(
        "--no-tinvest-coupons",
        action="store_true",
        help="Skip T-Invest coupon schedule.",
    )
    parser.add_argument(
        "--no-moex-bond-history",
        action="store_true",
        help="Skip MOEX bond history.",
    )
    parser.add_argument(
        "--no-moex-share-history",
        action="store_true",
        help="Skip MOEX share history.",
    )
    parser.add_argument(
        "--history-from",
        default="2014-07-01",
        help="History start date for both MOEX and T-Invest.",
    )
    parser.add_argument(
        "--history-to",
        default=None,
        help="History end date. Defaults to today.",
    )
    parser.add_argument(
        "--share-manual-review-workbook",
        type=Path,
        default=PROJECT_ROOT / "data_processed" / "share_manual_review_quotes.xlsx",
        help="Workbook with manually validated share mappings; sheet review_queue, column usable.",
    )
    parser.add_argument(
        "--moex-share-review-dir",
        type=Path,
        default=PROJECT_ROOT / "moex_quotes_review",
        help="Directory with MOEX share quote review workbooks.",
    )
    parser.add_argument(
        "--no-share-manual-review",
        action="store_true",
        help="Do not apply manually validated share mappings.",
    )
    args = parser.parse_args()

    if args.disable_tinvest_ssl_verify:
        os.environ["TINVEST_SSL_VERIFY"] = "false"

    ensure_directory(args.processed_dir)
    ensure_directory(args.raw_dir)
    ensure_directory(args.output.parent)

    companies_master, shortlist_log = combine_shortlists(
        oil_gas_shortlist=args.oil_gas_shortlist,
        metallurgy_shortlist=args.metallurgy_shortlist,
    )
    if args.max_companies is not None:
        companies_master = companies_master.head(args.max_companies).copy()

    save_dataframe_csv(companies_master, args.processed_dir / "companies_master.csv")
    save_dataframe_csv(shortlist_log, args.processed_dir / "shortlist_log.csv")

    history_to = args.history_to or pd.Timestamp.now("UTC").date().isoformat()

    if args.skip_download:
        existing_download_log = read_csv_if_exists(args.processed_dir / "download_log.csv")
        moex_result = {
            "moex_instruments": read_csv_if_exists(args.processed_dir / "moex_instruments.csv"),
            "moex_bonds": read_csv_if_exists(args.processed_dir / "moex_bonds.csv"),
            "moex_bond_history": read_csv_if_exists(args.processed_dir / "moex_bond_history.csv"),
            "moex_share_history": read_csv_if_exists(args.processed_dir / "moex_share_history.csv"),
            "mapping_log": read_csv_if_exists(args.processed_dir / "mapping_log_moex.csv"),
            "download_log": pd.DataFrame(),
        }
        tinvest_result = {
            "tinvest_instruments": read_csv_if_exists(args.processed_dir / "tinvest_instruments.csv"),
            "tinvest_bonds": read_csv_if_exists(args.processed_dir / "tinvest_bonds.csv"),
            "tinvest_bond_coupons": read_csv_if_exists(args.processed_dir / "tinvest_bond_coupons.csv"),
            "tinvest_history": read_csv_if_exists(args.processed_dir / "tinvest_history.csv"),
            "mapping_log": read_csv_if_exists(args.processed_dir / "mapping_log_tinvest.csv"),
            "download_log": read_csv_if_exists(args.processed_dir / "download_log_tinvest.csv"),
        }
        mapping_log = pd.concat(
            [
                moex_result.get("mapping_log", pd.DataFrame()),
                tinvest_result.get("mapping_log", pd.DataFrame()),
            ],
            ignore_index=True,
            sort=False,
        )
        if existing_download_log.empty:
            download_log = shortlist_log.copy()
        else:
            download_log = existing_download_log.drop_duplicates().reset_index(drop=True)
    else:
        session = create_session()

        moex_result = load_moex_data(
            companies_master=companies_master,
            output_dir=args.processed_dir,
            raw_dir=args.raw_dir,
            session=session,
            load_history=not args.no_moex_bond_history,
            history_from=args.history_from,
            history_to=history_to,
            load_share_history=not args.no_moex_share_history,
            share_history_from=args.history_from,
            share_history_to=history_to,
            bond_detail_scope="history_candidates",
            bond_history_scope="history_candidates",
            share_history_scope="history_candidates",
        )

        tinvest_result = load_tinvest_optional(
            companies_master=companies_master,
            output_dir=args.processed_dir,
            session=session,
            load_coupons=not args.no_tinvest_coupons,
            coupon_from=args.history_from,
            coupon_to=history_to,
            load_history=(not args.no_tinvest) and (not args.no_tinvest_history),
            history_from=args.history_from,
            history_to=history_to,
        ) if not args.no_tinvest else {
            "tinvest_instruments": pd.DataFrame(),
            "tinvest_bonds": pd.DataFrame(),
            "tinvest_bond_coupons": pd.DataFrame(),
            "tinvest_history": pd.DataFrame(),
            "mapping_log": pd.DataFrame(),
            "download_log": pd.DataFrame(),
        }

        mapping_log = pd.concat(
            [
                moex_result.get("mapping_log", pd.DataFrame()),
                tinvest_result.get("mapping_log", pd.DataFrame()),
            ],
            ignore_index=True,
            sort=False,
        )
        download_log = pd.concat(
            [
                shortlist_log,
                moex_result.get("download_log", pd.DataFrame()),
                tinvest_result.get("download_log", pd.DataFrame()),
            ],
            ignore_index=True,
            sort=False,
        )

        save_dataframe_csv(mapping_log, args.processed_dir / "mapping_log.csv")
        save_dataframe_csv(download_log, args.processed_dir / "download_log.csv")
        if not tinvest_result.get("mapping_log", pd.DataFrame()).empty:
            save_dataframe_csv(tinvest_result["mapping_log"], args.processed_dir / "mapping_log_tinvest.csv")
        if not tinvest_result.get("download_log", pd.DataFrame()).empty:
            save_dataframe_csv(tinvest_result["download_log"], args.processed_dir / "download_log_tinvest.csv")

    share_manual_review = pd.DataFrame()
    moex_share_review_history = pd.DataFrame()
    moex_share_review_security_summary = pd.DataFrame()
    if not args.no_share_manual_review:
        share_manual_review = read_share_manual_review(args.share_manual_review_workbook)
        moex_share_review_history, moex_share_review_security_summary = build_moex_share_review_history(
            args.moex_share_review_dir,
            share_manual_review,
        )
        if not share_manual_review.empty:
            save_dataframe_csv(normalize_share_manual_review(share_manual_review), args.processed_dir / "share_manual_review.csv")
        if not moex_share_review_history.empty:
            save_dataframe_csv(moex_share_review_history, args.processed_dir / "moex_share_review_history.csv")
        if not moex_share_review_security_summary.empty:
            save_dataframe_csv(
                moex_share_review_security_summary,
                args.processed_dir / "moex_share_review_security_summary.csv",
            )

    outputs = build_public_securities_outputs(
        companies_master=companies_master,
        moex_instruments=moex_result.get("moex_instruments", pd.DataFrame()),
        moex_bonds=moex_result.get("moex_bonds", pd.DataFrame()),
        moex_bond_history=moex_result.get("moex_bond_history", pd.DataFrame()),
        moex_share_history=moex_result.get("moex_share_history", pd.DataFrame()),
        tinvest_instruments=tinvest_result.get("tinvest_instruments", pd.DataFrame()),
        tinvest_bonds=tinvest_result.get("tinvest_bonds", pd.DataFrame()),
        tinvest_bond_coupons=tinvest_result.get("tinvest_bond_coupons", pd.DataFrame()),
        tinvest_history=tinvest_result.get("tinvest_history", pd.DataFrame()),
        mapping_log=mapping_log,
        download_log=download_log,
        share_manual_review=share_manual_review,
        moex_share_review_history=moex_share_review_history,
        moex_share_review_security_summary=moex_share_review_security_summary,
    )
    workbook_outputs, inventory = write_public_securities_workbook(
        outputs=outputs,
        output_workbook=args.output,
        processed_dir=args.processed_dir,
    )

    print(f"Public securities workbook created: {args.output}")
    for sheet_name, frame in workbook_outputs.items():
        print(f"{sheet_name}: {len(frame)} rows")
    print(f"Sheets inventory rows: {len(inventory)}")


if __name__ == "__main__":
    main()
