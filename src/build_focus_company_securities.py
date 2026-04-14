from __future__ import annotations

import argparse
import json
import re
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_PROCESSED = PROJECT_ROOT / "data_processed"
DEFAULT_SNAPSHOT = DATA_PROCESSED / "focus_company_moex_snapshot.xlsx"
DEFAULT_WATCHLIST = DATA_PROCESSED / "focus_company_securities_watchlist.xlsx"
DEFAULT_QUEUE = DATA_PROCESSED / "cbonds_pending_download_queue.xlsx"
CBONDS_CARDS_PATH = DATA_PROCESSED / "cbonds_bond_cards_validated.xlsx"
CBONDS_QUOTES_VALIDATED_PATH = DATA_PROCESSED / "cbonds_quotes_validated.xlsx"
CBONDS_QUOTES_CAPTURE_PATH = DATA_PROCESSED / "cbonds_quotes_capture_output.xlsx"
ANALYSIS_PANEL_PATH = DATA_PROCESSED / "analysis_panel.xlsx"

MOEX_RELEVANT_TYPES = {
    "common_share",
    "preferred_share",
    "exchange_bond",
    "corporate_bond",
}

QUEUE_PRIORITY_ORDER = {
    "high": 1,
    "legacy_missing_strong_candidate": 2,
    "group_related": 3,
    "secondary": 4,
}


@dataclass(frozen=True)
class FocusCompany:
    focus_company_name: str
    focus_company_inn: str
    sector: str
    sample_flag: str
    relation_mode: str
    moex_query_inn: str
    issuer_name_expected: str
    issuer_keyword: str
    note: str


FOCUS_COMPANIES: tuple[FocusCompany, ...] = (
    FocusCompany(
        focus_company_name="ГАЗПРОМ НЕФТЬ, ПАО",
        focus_company_inn="5504036333",
        sector="oil_gas",
        sample_flag="main",
        relation_mode="direct_current_company",
        moex_query_inn="5504036333",
        issuer_name_expected='Публичное акционерное общество "Газпром нефть"',
        issuer_keyword="Газпром нефть",
        note="Manual sample addition: large public company omitted in original shortlist export.",
    ),
    FocusCompany(
        focus_company_name="МЕЧЕЛ-ТРАНС, ООО",
        focus_company_inn="7728246919",
        sector="metallurgy",
        sample_flag="main",
        relation_mode="parent_group_public_securities",
        moex_query_inn="7703370008",
        issuer_name_expected='Публичное акционерное общество "Мечел"',
        issuer_keyword="Мечел",
        note="Company itself has no confirmed direct public securities; watch parent-group public securities for market context.",
    ),
)


def fetch_moex_snapshot(query_value: str) -> pd.DataFrame:
    url = (
        "https://iss.moex.com/iss/securities.json?iss.meta=off&q="
        + urllib.parse.quote(query_value)
    )
    with urllib.request.urlopen(url, timeout=60) as response:
        payload = json.load(response)
    return pd.DataFrame(payload["securities"]["data"], columns=payload["securities"]["columns"])


def build_snapshot(refresh: bool, snapshot_path: Path) -> dict[str, pd.DataFrame]:
    if snapshot_path.exists() and not refresh:
        xl = pd.ExcelFile(snapshot_path)
        return {
            sheet: pd.read_excel(snapshot_path, sheet_name=sheet)
            for sheet in xl.sheet_names
        }

    sheets: dict[str, pd.DataFrame] = {}
    for focus in FOCUS_COMPANIES:
        sheet_name = f"{focus.focus_company_inn}_{focus.moex_query_inn}"
        df = fetch_moex_snapshot(focus.moex_query_inn)
        df.insert(0, "focus_company_name", focus.focus_company_name)
        df.insert(1, "focus_company_inn", focus.focus_company_inn)
        df.insert(2, "focus_relation_mode", focus.relation_mode)
        sheets[sheet_name] = df

    with pd.ExcelWriter(snapshot_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return sheets


def extract_downloaded_quote_queries(workbook_path: Path) -> set[str]:
    if not workbook_path.exists():
        return set()
    xl = pd.ExcelFile(workbook_path)
    queries: set[str] = set()
    for sheet_name in xl.sheet_names:
        match = re.match(r"r\d+_(.+?)(?:_\d+)?$", sheet_name)
        if match:
            queries.add(match.group(1))
    return queries


def load_existing_card_maps() -> tuple[pd.DataFrame, set[str], set[str]]:
    cards = pd.read_excel(CBONDS_CARDS_PATH, sheet_name="issue_master_validated_all")
    card_isins = set(cards["parsed_isin"].dropna().astype(str).str.strip())
    card_regs = set(cards["reg_number"].dropna().astype(str).str.strip())
    return cards, card_isins, card_regs


def build_focus_watchlist(
    snapshot_sheets: dict[str, pd.DataFrame],
    cards: pd.DataFrame,
    downloaded_quote_queries: set[str],
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for focus in FOCUS_COMPANIES:
        matching_sheet = next(
            sheet for sheet in snapshot_sheets if sheet.startswith(focus.focus_company_inn)
        )
        df = snapshot_sheets[matching_sheet].copy()
        df = df[df["type"].isin(MOEX_RELEVANT_TYPES)].copy()
        df = df[
            [
                "focus_company_name",
                "focus_company_inn",
                "focus_relation_mode",
                "secid",
                "shortname",
                "name",
                "isin",
                "emitent_title",
                "emitent_inn",
                "type",
                "group",
            ]
        ].drop_duplicates(subset=["secid", "isin", "type"])
        df["sector"] = focus.sector
        df["sample_flag"] = focus.sample_flag
        df["focus_note"] = focus.note
        df["issuer_name_expected"] = focus.issuer_name_expected
        df["security_relation_label"] = focus.relation_mode
        df["instrument_type"] = df["type"].map(
            {
                "common_share": "share_common",
                "preferred_share": "share_preferred",
                "exchange_bond": "bond",
                "corporate_bond": "bond",
            }
        )
        df["cbonds_query"] = df["isin"].fillna(df["secid"])
        df["priority_bucket"] = "reference"
        df.loc[df["instrument_type"].eq("bond"), "priority_bucket"] = "bond_focus"
        df.loc[df["type"].isin({"common_share", "preferred_share"}), "priority_bucket"] = "equity_focus"
        df.loc[
            (df["focus_company_name"].eq("ГАЗПРОМ НЕФТЬ, ПАО"))
            & (df["secid"].eq("SIBN")),
            "priority_bucket",
        ] = "equity_primary"
        df.loc[
            (df["focus_company_name"].eq("МЕЧЕЛ-ТРАНС, ООО"))
            & (df["secid"].isin({"MTLR", "MTLRP"})),
            "priority_bucket",
        ] = "equity_group_focus"

        cards_match = cards[["parsed_isin", "reg_number", "borrower_name", "validated_relation_label", "company_name"]].copy()
        cards_match["parsed_isin"] = cards_match["parsed_isin"].astype(str).str.strip()
        df = df.merge(
            cards_match.rename(
                columns={
                    "parsed_isin": "isin",
                    "borrower_name": "existing_card_borrower_name",
                    "validated_relation_label": "existing_card_relation_label",
                    "company_name": "existing_card_company_name",
                }
            ),
            on="isin",
            how="left",
        )
        df["issue_card_downloaded_flag"] = df["existing_card_borrower_name"].notna()
        df["quotes_downloaded_flag"] = df["cbonds_query"].astype(str).isin(downloaded_quote_queries)
        df["issue_card_needed_flag"] = df["instrument_type"].eq("bond") & ~df["issue_card_downloaded_flag"]
        df["quotes_needed_flag"] = df["instrument_type"].eq("bond") & ~df["quotes_downloaded_flag"]
        df["download_action"] = "no_download_needed"
        df.loc[
            df["issue_card_needed_flag"] & df["quotes_needed_flag"], "download_action"
        ] = "download_card_and_quotes"
        df.loc[
            df["issue_card_needed_flag"] & ~df["quotes_needed_flag"], "download_action"
        ] = "download_card_only"
        df.loc[
            ~df["issue_card_needed_flag"] & df["quotes_needed_flag"], "download_action"
        ] = "download_quotes_only"
        frames.append(df)

    watchlist = pd.concat(frames, ignore_index=True)

    supplemental_frames: list[pd.DataFrame] = []
    cards_copy = cards.copy()
    cards_copy["parsed_isin"] = cards_copy["parsed_isin"].astype(str).str.strip()
    for focus in FOCUS_COMPANIES:
        existing_isins = set(
            watchlist.loc[
                watchlist["focus_company_name"].eq(focus.focus_company_name), "isin"
            ]
            .dropna()
            .astype(str)
        )
        extra = cards_copy[
            cards_copy["borrower_name"].astype(str).str.contains(
                focus.issuer_keyword, case=False, na=False
            )
        ].copy()
        extra = extra[~extra["parsed_isin"].isin(existing_isins)].copy()
        if extra.empty:
            continue
        extra_df = pd.DataFrame(
            {
                "focus_company_name": focus.focus_company_name,
                "focus_company_inn": focus.focus_company_inn,
                "focus_relation_mode": focus.relation_mode,
                "secid": extra["parsed_isin"],
                "shortname": extra["issue_name"],
                "name": extra["issue_name"],
                "isin": extra["parsed_isin"],
                "emitent_title": extra["borrower_name"],
                "emitent_inn": pd.NA,
                "type": "cbonds_existing_bond",
                "group": "cbonds_existing",
                "sector": focus.sector,
                "sample_flag": focus.sample_flag,
                "focus_note": focus.note,
                "issuer_name_expected": focus.issuer_name_expected,
                "security_relation_label": focus.relation_mode,
                "instrument_type": "bond",
                "cbonds_query": extra["parsed_isin"],
                "priority_bucket": "bond_focus_existing_card",
                "existing_card_borrower_name": extra["borrower_name"],
                "existing_card_relation_label": extra["validated_relation_label"],
                "existing_card_company_name": extra["company_name"],
                "issue_card_downloaded_flag": True,
                "quotes_downloaded_flag": extra["parsed_isin"].astype(str).isin(downloaded_quote_queries),
                "issue_card_needed_flag": False,
                "quotes_needed_flag": ~extra["parsed_isin"].astype(str).isin(downloaded_quote_queries),
            }
        )
        extra_df["download_action"] = "no_download_needed"
        extra_df.loc[extra_df["quotes_needed_flag"], "download_action"] = "download_quotes_only"
        supplemental_frames.append(extra_df)

    if supplemental_frames:
        watchlist = pd.concat([watchlist, *supplemental_frames], ignore_index=True)
        watchlist = watchlist.drop_duplicates(
            subset=["focus_company_name", "cbonds_query", "instrument_type"], keep="first"
        )

    watchlist["download_rank"] = (
        watchlist["issue_card_needed_flag"].astype(int) * 1000
        + watchlist["quotes_needed_flag"].astype(int) * 100
        + watchlist["instrument_type"].eq("bond").astype(int) * 10
    )
    watchlist = watchlist.sort_values(
        by=["download_rank", "focus_company_name", "instrument_type", "isin", "secid"],
        ascending=[False, True, True, True, True],
    ).reset_index(drop=True)
    return watchlist


def build_pending_queue(watchlist: pd.DataFrame) -> pd.DataFrame:
    queue = watchlist[
        watchlist["instrument_type"].eq("bond")
        & (watchlist["issue_card_needed_flag"] | watchlist["quotes_needed_flag"])
    ].copy()
    queue["queue_priority"] = "secondary"
    queue.loc[
        queue["focus_company_name"].eq("ГАЗПРОМ НЕФТЬ, ПАО"),
        "queue_priority",
    ] = "high"
    queue.loc[
        queue["focus_company_name"].eq("МЕЧЕЛ-ТРАНС, ООО"),
        "queue_priority",
    ] = "group_related"
    queue["queue_note"] = queue.apply(
        lambda row: (
            "Direct issuer bond for newly added sample company."
            if row["focus_company_name"] == "ГАЗПРОМ НЕФТЬ, ПАО"
            else "Parent-group issuer bond for Mechel-Trans market context."
        ),
        axis=1,
    )
    queue["queue_origin"] = "focus_company_watchlist"
    queue = queue.sort_values(
        by=["queue_priority", "issue_card_needed_flag", "quotes_needed_flag", "isin", "secid"],
        key=lambda col: col.map(QUEUE_PRIORITY_ORDER).fillna(999) if col.name == "queue_priority" else col,
        ascending=[True, False, False, True, True],
    ).reset_index(drop=True)
    queue.insert(0, "queue_rank", range(1, len(queue) + 1))
    return queue


def build_legacy_pending_queue(downloaded_quote_queries: set[str]) -> pd.DataFrame:
    legacy = pd.read_excel(CBONDS_QUOTES_VALIDATED_PATH, sheet_name="missing_issue_cards_strong").copy()
    legacy["cbonds_quote_run_query"] = legacy["cbonds_quote_run_query"].astype(str).str.strip()
    legacy["quotes_downloaded_flag"] = legacy["cbonds_quote_run_query"].isin(downloaded_quote_queries)
    legacy = legacy[~legacy["quotes_downloaded_flag"]].copy()
    if legacy.empty:
        return pd.DataFrame()

    queue = pd.DataFrame(
        {
            "queue_rank": range(1, len(legacy) + 1),
            "focus_company_name": legacy["company_name"],
            "focus_company_inn": legacy["company_inn"],
            "focus_relation_mode": legacy["validated_relation_label"].fillna(legacy["quote_collection_scope"]),
            "secid": pd.NA,
            "shortname": legacy["issue_name"],
            "name": legacy["issue_name"],
            "isin": legacy["cbonds_quote_run_query"],
            "emitent_title": pd.NA,
            "emitent_inn": pd.NA,
            "type": "bond",
            "group": "legacy_pending_missing_issue_cards",
            "sector": legacy["sector"],
            "sample_flag": legacy["sample_flag"],
            "focus_note": legacy["quote_collection_note_validated"],
            "issuer_name_expected": pd.NA,
            "security_relation_label": legacy["validated_relation_label"],
            "instrument_type": "bond",
            "cbonds_query": legacy["cbonds_quote_run_query"],
            "priority_bucket": legacy["quote_collection_bucket_validated"],
            "reg_number": pd.NA,
            "existing_card_borrower_name": pd.NA,
            "existing_card_relation_label": pd.NA,
            "existing_card_company_name": pd.NA,
            "issue_card_downloaded_flag": False,
            "quotes_downloaded_flag": False,
            "issue_card_needed_flag": True,
            "quotes_needed_flag": True,
            "download_action": "download_card_and_quotes",
            "download_rank": 1110,
            "queue_priority": "legacy_missing_strong_candidate",
            "queue_note": legacy["quote_collection_note_validated"],
            "queue_origin": "existing_validated_universe_missing_card",
        }
    )
    return queue


def build_summary(watchlist: pd.DataFrame, pending: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for focus_company, group in watchlist.groupby("focus_company_name", sort=False):
        rows.append(
            {
                "focus_company_name": focus_company,
                "n_total_relevant_securities": len(group),
                "n_relevant_bonds": int(group["instrument_type"].eq("bond").sum()),
                "n_relevant_shares": int(group["instrument_type"].str.startswith("share").sum()),
                "n_cards_downloaded": int(group["issue_card_downloaded_flag"].sum()),
                "n_quotes_downloaded": int(group["quotes_downloaded_flag"].sum()),
                "n_cards_needed": int(group["issue_card_needed_flag"].sum()),
                "n_quotes_needed": int(group["quotes_needed_flag"].sum()),
            }
        )
    rows.append(
        {
            "focus_company_name": "TOTAL",
            "n_total_relevant_securities": len(watchlist),
            "n_relevant_bonds": int(watchlist["instrument_type"].eq("bond").sum()),
            "n_relevant_shares": int(watchlist["instrument_type"].str.startswith("share").sum()),
            "n_cards_downloaded": int(watchlist["issue_card_downloaded_flag"].sum()),
            "n_quotes_downloaded": int(watchlist["quotes_downloaded_flag"].sum()),
            "n_cards_needed": int(watchlist["issue_card_needed_flag"].sum()),
            "n_quotes_needed": int(watchlist["quotes_needed_flag"].sum()),
        }
    )
    rows.append(
        {
            "focus_company_name": "PENDING_BOND_QUEUE",
            "n_total_relevant_securities": len(pending),
            "n_relevant_bonds": len(pending),
            "n_relevant_shares": 0,
            "n_cards_downloaded": int((~pending["issue_card_needed_flag"]).sum()),
            "n_quotes_downloaded": int((~pending["quotes_needed_flag"]).sum()),
            "n_cards_needed": int(pending["issue_card_needed_flag"].sum()),
            "n_quotes_needed": int(pending["quotes_needed_flag"].sum()),
        }
    )
    return pd.DataFrame(rows)


def build_queue_summary(all_pending: pd.DataFrame) -> pd.DataFrame:
    if all_pending.empty:
        return pd.DataFrame(columns=["bucket", "metric", "value"])
    rows = []
    rows.append({"bucket": "all_pending", "metric": "n_rows", "value": len(all_pending)})
    rows.append(
        {
            "bucket": "all_pending",
            "metric": "n_companies",
            "value": all_pending["focus_company_name"].nunique(),
        }
    )
    for origin, group in all_pending.groupby("queue_origin", dropna=False):
        rows.append({"bucket": str(origin), "metric": "n_rows", "value": len(group)})
        rows.append(
            {
                "bucket": str(origin),
                "metric": "n_companies",
                "value": group["focus_company_name"].nunique(),
            }
        )
    for company, group in all_pending.groupby("focus_company_name", dropna=False):
        rows.append({"bucket": str(company), "metric": "n_rows", "value": len(group)})
    return pd.DataFrame(rows)


def build_focus_companies_sheet() -> pd.DataFrame:
    rows = []
    companies_master = pd.read_excel(ANALYSIS_PANEL_PATH, sheet_name="companies_master")
    master_subset = companies_master[
        companies_master["company_name"].isin([company.focus_company_name for company in FOCUS_COMPANIES])
    ].copy()
    for focus in FOCUS_COMPANIES:
        row = master_subset.loc[
            master_subset["company_name"] == focus.focus_company_name
        ].iloc[0].to_dict()
        row["focus_relation_mode"] = focus.relation_mode
        row["issuer_name_expected"] = focus.issuer_name_expected
        row["focus_note"] = focus.note
        rows.append(row)
    return pd.DataFrame(rows)


def write_workbooks(
    watchlist: pd.DataFrame,
    pending: pd.DataFrame,
    all_pending: pd.DataFrame,
    legacy_pending: pd.DataFrame,
    snapshot_sheets: dict[str, pd.DataFrame],
    watchlist_path: Path,
    queue_path: Path,
) -> None:
    focus_companies = build_focus_companies_sheet()
    summary = build_summary(watchlist, pending)
    queue_summary = build_queue_summary(all_pending)
    shares = watchlist[watchlist["instrument_type"].str.startswith("share")].copy()
    bonds = watchlist[watchlist["instrument_type"].eq("bond")].copy()

    with pd.ExcelWriter(watchlist_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        focus_companies.to_excel(writer, sheet_name="focus_companies", index=False)
        watchlist.to_excel(writer, sheet_name="securities_watchlist", index=False)
        bonds.to_excel(writer, sheet_name="bonds_watchlist", index=False)
        shares.to_excel(writer, sheet_name="shares_watchlist", index=False)
        for sheet_name, df in snapshot_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    with pd.ExcelWriter(queue_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        queue_summary.to_excel(writer, sheet_name="queue_summary", index=False)
        all_pending.to_excel(writer, sheet_name="all_pending_download_queue", index=False)
        pending.to_excel(writer, sheet_name="focus_company_pending", index=False)
        legacy_pending.to_excel(writer, sheet_name="legacy_pending", index=False)
        pending[pending["issue_card_needed_flag"]].to_excel(
            writer, sheet_name="cards_needed", index=False
        )
        all_pending[all_pending["quotes_needed_flag"]].to_excel(
            writer, sheet_name="quotes_needed", index=False
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build focused securities watchlists and pending Cbonds download queue."
    )
    parser.add_argument(
        "--refresh-moex",
        action="store_true",
        help="Refresh cached MOEX focus-company snapshot via official ISS.",
    )
    parser.add_argument(
        "--snapshot-path",
        type=Path,
        default=DEFAULT_SNAPSHOT,
    )
    parser.add_argument(
        "--watchlist-output",
        type=Path,
        default=DEFAULT_WATCHLIST,
    )
    parser.add_argument(
        "--queue-output",
        type=Path,
        default=DEFAULT_QUEUE,
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    snapshot_sheets = build_snapshot(refresh=args.refresh_moex, snapshot_path=args.snapshot_path)
    cards, _, _ = load_existing_card_maps()
    downloaded_quote_queries = extract_downloaded_quote_queries(CBONDS_QUOTES_CAPTURE_PATH)
    watchlist = build_focus_watchlist(snapshot_sheets, cards, downloaded_quote_queries)
    pending = build_pending_queue(watchlist)
    legacy_pending = build_legacy_pending_queue(downloaded_quote_queries)
    all_pending = pd.concat([legacy_pending, pending], ignore_index=True)
    if not all_pending.empty:
        all_pending = all_pending.sort_values(
            by=["queue_priority", "focus_company_name", "isin", "secid"],
            key=lambda col: col.map(QUEUE_PRIORITY_ORDER).fillna(999) if col.name == "queue_priority" else col,
            ascending=[True, True, True, True],
        ).reset_index(drop=True)
        all_pending["queue_rank"] = range(1, len(all_pending) + 1)
    write_workbooks(
        watchlist=watchlist,
        pending=pending,
        all_pending=all_pending,
        legacy_pending=legacy_pending,
        snapshot_sheets=snapshot_sheets,
        watchlist_path=args.watchlist_output,
        queue_path=args.queue_output,
    )
    print(f"Watchlist rows: {len(watchlist)}")
    print(f"Pending queue rows: {len(pending)}")
    print(f"Legacy pending rows: {len(legacy_pending)}")
    print(f"All pending rows: {len(all_pending)}")
    print(f"Cards needed: {int(pending['issue_card_needed_flag'].sum())}")
    print(f"Quotes needed: {int(pending['quotes_needed_flag'].sum())}")


if __name__ == "__main__":
    main()
