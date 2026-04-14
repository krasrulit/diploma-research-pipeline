from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from .apply_cbonds_bond_review import (
    DEFAULT_HELPER_WORKBOOK,
    build_missing_strong_candidates,
    build_relationship_summary,
)
from .build_cbonds_bond_cards import (
    DEFAULT_PROCESSED_OUTPUT,
    add_model_use_flag,
    add_quotes_priority,
    build_company_bond_quarterly,
    build_company_bond_summary,
    dedupe_issue_master,
    match_bucket,
    parse_capture_buckets,
    write_workbook,
)
from .utils import normalize_identifier


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PENDING_QUEUE = PROJECT_ROOT / "data_processed" / "cbonds_pending_download_queue.xlsx"
DEFAULT_PENDING_ISSUE_CARDS = PROJECT_ROOT / "data_processed" / "cbonds_pending_issue_cards_output.xlsx"
DEFAULT_VALIDATED_WORKBOOK = PROJECT_ROOT / "data_processed" / "cbonds_bond_cards_validated.xlsx"
DEFAULT_QUOTES_WORKBOOK = PROJECT_ROOT / "data_processed" / "cbonds_quotes_validated.xlsx"
DEFAULT_AUDIT_OUTPUT = PROJECT_ROOT / "data_processed" / "cbonds_pending_integration_audit.xlsx"


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).replace("\xa0", " ").strip()


def borrower_name_match_score(expected_name: object, borrower_name: object) -> float:
    expected = clean_text(expected_name).lower()
    borrower = clean_text(borrower_name).lower()
    if not expected or not borrower:
        return 0.0
    if expected == borrower:
        return 1.0
    if expected in borrower or borrower in expected:
        return 0.95
    expected_tokens = set(expected.replace('"', "").split())
    borrower_tokens = set(borrower.replace('"', "").split())
    if not expected_tokens or not borrower_tokens:
        return 0.0
    overlap = len(expected_tokens & borrower_tokens)
    union = len(expected_tokens | borrower_tokens)
    return overlap / union if union else 0.0


def normalize_pending_queue(queue_workbook: Path) -> pd.DataFrame:
    queue = pd.read_excel(queue_workbook, sheet_name="all_pending_download_queue").copy()
    out = queue.copy()
    out["company_name"] = out["focus_company_name"]
    out["company_inn"] = out["focus_company_inn"].map(lambda value: normalize_identifier(value, length=10))
    out["company_ticker"] = ""
    out["company_isin"] = ""
    out["ticker"] = out["secid"].map(clean_text)
    out["instrument_name"] = out["name"].map(clean_text)
    out["query_text"] = out["cbonds_query"].map(clean_text)
    out["best_source"] = "moex"
    out["best_source_instrument_id"] = out["secid"].map(clean_text)
    out["best_n_trade_dates"] = 0
    out["best_history_start"] = pd.NaT
    out["best_history_end"] = pd.NaT
    out["best_match_score"] = np.nan
    out["best_match_confidence"] = ""
    out["manual_review_needed_flag"] = False
    out["history_available_flag"] = False
    out["usable_history_flag"] = False
    out["reliable_mapping_flag"] = out["focus_relation_mode"].eq("direct_current_company")
    out["market_relevance_bucket"] = out["priority_bucket"].fillna("").map(clean_text)
    out["company_id"] = out["sector"].map(clean_text) + ":" + out["company_inn"].map(clean_text)
    return out


def classify_pending(issue_master_pending: pd.DataFrame) -> pd.DataFrame:
    out = issue_master_pending.copy()
    out["queue_isin"] = out["security_isin"].fillna("").map(clean_text)
    out["parsed_isin"] = out["parsed_isin"].fillna("").map(clean_text)
    out["queue_isin_match_flag"] = out["queue_isin"].eq(out["parsed_isin"])
    out["expected_issuer_match_score"] = out.apply(
        lambda row: borrower_name_match_score(row.get("issuer_name_expected"), row.get("borrower_name")),
        axis=1,
    )
    out["expected_issuer_match_bucket"] = out["expected_issuer_match_score"].map(match_bucket)

    high_match = out["borrower_match_bucket"].eq("high") | out["expected_issuer_match_bucket"].eq("high")
    direct_mode = out["focus_relation_mode"].isin(["direct_current_company", "missing_issue_card_not_captured"])
    related_mode = out["focus_relation_mode"].eq("parent_group_public_securities")

    out["pending_include_direct_flag"] = out["queue_isin_match_flag"] & high_match & direct_mode
    out["pending_include_related_flag"] = out["queue_isin_match_flag"] & high_match & related_mode
    out["pending_exclude_flag"] = ~(out["pending_include_direct_flag"] | out["pending_include_related_flag"])

    out["validated_manual_code"] = pd.Series(pd.NA, index=out.index, dtype="Int64")
    out.loc[out["pending_include_direct_flag"], "validated_manual_code"] = 1
    out.loc[out["pending_include_related_flag"], "validated_manual_code"] = 6
    out.loc[out["pending_exclude_flag"], "validated_manual_code"] = 4
    out["validated_manual_code_effective"] = out["validated_manual_code"]

    out["validated_code_source"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            "auto_pending_direct_match",
            "auto_pending_same_group_match",
            "auto_pending_excluded",
        ],
        default="auto_pending_unreviewed",
    )
    out["validated_relation_label"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            "direct_current_company",
            "same_group_other_legal_entity",
            "exclude_unrelated",
        ],
        default="",
    )

    out["validated_keep_for_current_company_flag"] = out["pending_include_direct_flag"]
    out["validated_group_financing_agent_flag"] = False
    out["validated_unrelated_exclude_flag"] = out["pending_exclude_flag"]
    out["validated_structured_linked_flag"] = False
    out["validated_same_group_other_legal_flag"] = out["pending_include_related_flag"]
    out["validated_reassign_not_in_sample_flag"] = False
    out["validated_keep_related_optional_flag"] = out["pending_include_related_flag"]
    out["validated_drop_from_current_company_flag"] = out["pending_exclude_flag"]
    out["validated_any_keep_flag"] = out["pending_include_direct_flag"] | out["pending_include_related_flag"]

    out["quote_collection_bucket_validated"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            "current_company_direct",
            "same_group_optional",
            "exclude_unrelated",
        ],
        default="unreviewed",
    )
    out["quote_collection_rank_validated"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            1,
            3,
            9,
        ],
        default=7,
    ).astype(int)

    out["quote_collection_note_validated"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            "Новая pending-карточка уверенно подтверждает прямой выпуск компании из выборки.",
            "Новая pending-карточка подтверждает выпуск другой компании той же группы; учитывать как связанный, но не прямой долг текущего юрлица.",
            "Pending-карточка не подтверждает выпуск компании из очереди; в модель не включать.",
        ],
        default="Требует дополнительной проверки.",
    )
    out["pending_integration_bucket"] = np.select(
        [
            out["pending_include_direct_flag"],
            out["pending_include_related_flag"],
            out["pending_exclude_flag"],
        ],
        [
            "include_current_company",
            "include_related_group_only",
            "exclude_from_model",
        ],
        default="needs_review",
    )
    return out


def build_quotes_workbook(
    issue_master_validated_all: pd.DataFrame,
    missing_strong: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    quotes = issue_master_validated_all.copy()
    quotes["quote_collect_now_flag"] = quotes["validated_keep_for_current_company_flag"] | quotes["validated_keep_related_optional_flag"]
    quotes["quote_collect_primary_flag"] = quotes["validated_keep_for_current_company_flag"]

    current_company = quotes.loc[quotes["validated_keep_for_current_company_flag"]].copy()
    related_optional = quotes.loc[quotes["validated_keep_related_optional_flag"]].copy()
    excluded = quotes.loc[quotes["validated_drop_from_current_company_flag"]].copy()

    missing_for_quotes = missing_strong.copy()
    if not missing_for_quotes.empty:
        missing_for_quotes = missing_for_quotes.rename(
            columns={
                "cbonds_paper_query_primary": "cbonds_quote_run_query",
                "instrument_name": "issue_name",
            }
        )
        missing_for_quotes["quote_query_type"] = "ISIN_or_primary_id"
        missing_for_quotes["quote_date_from"] = "01.07.2014"
        missing_for_quotes["quote_date_to"] = "31.12.2025"
        missing_for_quotes["quote_exchange"] = ""
        missing_for_quotes["quote_collect_now_flag"] = True
        missing_for_quotes["quote_collect_primary_flag"] = True
        missing_for_quotes["validated_code_source"] = "missing_issue_card_audit"
        missing_for_quotes["validated_manual_code_effective"] = 1
        missing_for_quotes["quote_collection_scope"] = "missing_strong_candidate"

    full_collect = pd.concat(
        [current_company, related_optional, missing_for_quotes],
        ignore_index=True,
        sort=False,
    )
    sort_columns = [col for col in ["quote_collection_rank_validated", "company_name", "issue_name"] if col in full_collect.columns]
    if sort_columns:
        full_collect = full_collect.sort_values(sort_columns, kind="stable").reset_index(drop=True)

    return {
        "summary": pd.DataFrame(
            [
                {"metric": "quotes_direct_current_company", "value": int(len(current_company))},
                {"metric": "quotes_related_optional", "value": int(len(related_optional))},
                {"metric": "quotes_excluded", "value": int(len(excluded))},
                {"metric": "missing_strong_issue_cards", "value": int(len(missing_for_quotes))},
                {"metric": "quotes_full_collect_rows", "value": int(len(full_collect))},
            ]
        ),
        "quotes_full_collect": full_collect,
        "quotes_current_company": current_company,
        "quotes_related_optional": related_optional,
        "quotes_excluded": excluded,
        "missing_issue_cards_strong": missing_for_quotes,
    }


def build_validated_workbook(
    helper_workbook: Path,
    validated_workbook: Path,
    pending_queue_workbook: Path,
    pending_issue_cards_workbook: Path,
    quotes_output: Path,
    audit_output: Path,
) -> dict[str, pd.DataFrame]:
    normalized_queue = normalize_pending_queue(pending_queue_workbook)
    queue_lookup = normalized_queue.reset_index(drop=True).copy()
    queue_lookup["source_row"] = queue_lookup.index + 2
    pending_raw, pending_coupon, pending_offer, pending_quote, pending_parse_log = parse_capture_buckets(
        [pending_issue_cards_workbook],
        normalized_queue,
    )
    pending_raw = pending_raw.merge(
        queue_lookup[
            [
                "source_row",
                "focus_relation_mode",
                "issuer_name_expected",
                "queue_origin",
                "priority_bucket",
                "queue_priority",
            ]
        ],
        on="source_row",
        how="left",
    )
    pending_clean = add_model_use_flag(add_quotes_priority(dedupe_issue_master(pending_raw)))
    pending_classified = classify_pending(pending_clean)

    existing_all = pd.read_excel(validated_workbook, sheet_name="issue_master_validated_all")
    existing_quote_snapshot = pd.read_excel(validated_workbook, sheet_name="quote_snapshot")
    existing_parse_log = pd.read_excel(validated_workbook, sheet_name="parse_log")
    existing_coupon = pd.read_excel(validated_workbook, sheet_name="coupon_schedule_validated")
    existing_offer = pd.read_excel(validated_workbook, sheet_name="offer_terms_validated")

    pending_current = pending_classified.loc[pending_classified["validated_keep_for_current_company_flag"]].copy()
    pending_related = pending_classified.loc[pending_classified["validated_keep_related_optional_flag"]].copy()
    pending_excluded = pending_classified.loc[pending_classified["validated_drop_from_current_company_flag"]].copy()

    issue_master_validated_all = pd.concat(
        [existing_all, pending_classified],
        ignore_index=True,
        sort=False,
    )
    issue_master_validated_all = issue_master_validated_all.sort_values(
        ["issue_key", "validated_keep_for_current_company_flag", "validated_keep_related_optional_flag", "source_workbook", "sheet_name"],
        ascending=[True, False, False, True, True],
        kind="stable",
    ).drop_duplicates(subset=["issue_key"], keep="first").reset_index(drop=True)

    issue_master_validated_current = issue_master_validated_all.loc[
        issue_master_validated_all["validated_keep_for_current_company_flag"]
    ].copy()
    issue_master_validated_related = issue_master_validated_all.loc[
        issue_master_validated_all["validated_keep_related_optional_flag"]
    ].copy()
    issue_master_validated_excluded = issue_master_validated_all.loc[
        issue_master_validated_all["validated_drop_from_current_company_flag"]
    ].copy()

    pending_current_issue_keys = set(pending_current["issue_key"].dropna().astype(str))
    pending_coupon_current = pending_coupon.loc[pending_coupon["issue_key"].isin(pending_current_issue_keys)].copy() if not pending_coupon.empty else pd.DataFrame()
    pending_offer_current = pending_offer.loc[pending_offer["issue_key"].isin(pending_current_issue_keys)].copy() if not pending_offer.empty else pd.DataFrame()

    coupon_schedule_validated = pd.concat([existing_coupon, pending_coupon_current], ignore_index=True, sort=False)
    offer_terms_validated = pd.concat([existing_offer, pending_offer_current], ignore_index=True, sort=False)
    quote_snapshot = pd.concat([existing_quote_snapshot, pending_quote], ignore_index=True, sort=False)
    parse_log = pd.concat([existing_parse_log, pending_parse_log], ignore_index=True, sort=False)

    relationship_summary = build_relationship_summary(issue_master_validated_all)
    company_bond_summary_validated = build_company_bond_summary(issue_master_validated_current)
    company_bond_quarterly_validated = build_company_bond_quarterly(issue_master_validated_current, offer_terms_validated)
    missing_strong = build_missing_strong_candidates(helper_workbook, issue_master_validated_all)

    if company_bond_summary_validated.empty:
        company_bond_summary_validated = relationship_summary.copy()
    else:
        company_bond_summary_validated = company_bond_summary_validated.merge(
            relationship_summary,
            on="company_id",
            how="outer",
        )

    if not missing_strong.empty:
        missing_counts = (
            missing_strong.groupby(["company_name", "company_inn"], dropna=False)
            .size()
            .reset_index(name="cbonds_missing_strong_issue_count")
        )
        current_company_index = issue_master_validated_all[["company_id", "company_name", "company_inn"]].drop_duplicates()
        current_company_index = current_company_index.sort_values(
            ["company_id", "company_name", "company_inn"],
            kind="stable",
        ).drop_duplicates(subset=["company_id"], keep="first")
        missing_counts = current_company_index.merge(
            missing_counts,
            on=["company_name", "company_inn"],
            how="left",
        )
        missing_counts["cbonds_missing_strong_issue_count"] = missing_counts["cbonds_missing_strong_issue_count"].fillna(0).astype(int)
        company_bond_summary_validated = company_bond_summary_validated.merge(
            missing_counts[["company_id", "cbonds_missing_strong_issue_count"]],
            on="company_id",
            how="left",
        )
    else:
        company_bond_summary_validated["cbonds_missing_strong_issue_count"] = 0

    company_bond_summary_validated = company_bond_summary_validated.sort_values(
        ["company_id"],
        kind="stable",
    ).drop_duplicates(subset=["company_id"], keep="first").reset_index(drop=True)

    fill_zero_cols = [
        "cbonds_issue_count",
        "cbonds_issue_count_high_conf",
        "cbonds_issue_count_outstanding",
        "cbonds_issue_count_redeemed",
        "cbonds_issue_count_default",
        "cbonds_issue_count_need_quotes_high",
        "cbonds_issue_count_need_quotes_lower",
        "cbonds_issue_count_manual_review",
        "cbonds_issue_currency_count",
        "cbonds_validated_direct_issue_count",
        "cbonds_group_agent_issue_count",
        "cbonds_same_group_other_legal_issue_count",
        "cbonds_structured_linked_issue_count",
        "cbonds_reassign_not_in_sample_issue_count",
        "cbonds_unrelated_excluded_issue_count",
        "cbonds_missing_strong_issue_count",
    ]
    fill_false_cols = [
        "cbonds_issue_card_found_flag",
        "cbonds_has_rating_issue_flag",
        "cbonds_has_listing_issue_flag",
        "cbonds_has_offer_issue_flag",
        "cbonds_borrower_match_high_conf_flag",
        "cbonds_validated_direct_issue_flag",
        "cbonds_group_agent_issue_flag",
        "cbonds_same_group_other_legal_issue_flag",
        "cbonds_structured_linked_issue_flag",
        "cbonds_reassign_not_in_sample_issue_flag",
        "cbonds_unrelated_excluded_issue_flag",
        "cbonds_any_related_issue_flag",
        "cbonds_any_nonexcluded_issue_flag",
    ]
    for column in fill_zero_cols:
        if column in company_bond_summary_validated.columns:
            company_bond_summary_validated[column] = pd.to_numeric(
                company_bond_summary_validated[column],
                errors="coerce",
            ).fillna(0)
    for column in fill_false_cols:
        if column in company_bond_summary_validated.columns:
            company_bond_summary_validated[column] = company_bond_summary_validated[column].fillna(False).astype(bool)

    review_summary = (
        issue_master_validated_all.groupby(
            ["validated_relation_label", "validated_code_source"],
            dropna=False,
        )
        .size()
        .reset_index(name="issue_count")
        .sort_values(["validated_relation_label", "validated_code_source"], kind="stable")
        .reset_index(drop=True)
    )

    validated_sheets = {
        "summary": pd.DataFrame(
            [
                {"metric": "issues_all_captured_cards", "value": int(len(issue_master_validated_all))},
                {"metric": "issues_current_company_validated", "value": int(len(issue_master_validated_current))},
                {"metric": "issues_related_optional", "value": int(len(issue_master_validated_related))},
                {"metric": "issues_excluded", "value": int(len(issue_master_validated_excluded))},
                {"metric": "companies_current_company_validated", "value": int(issue_master_validated_current["company_id"].nunique()) if not issue_master_validated_current.empty else 0},
                {"metric": "companies_any_related_optional", "value": int(issue_master_validated_related["company_id"].nunique()) if not issue_master_validated_related.empty else 0},
                {"metric": "missing_strong_issue_cards", "value": int(len(missing_strong))},
                {"metric": "pending_cards_auto_direct_added", "value": int(len(pending_current))},
                {"metric": "pending_cards_auto_related_added", "value": int(len(pending_related))},
                {"metric": "pending_cards_auto_excluded", "value": int(len(pending_excluded))},
            ]
        ),
        "review_summary": review_summary,
        "issue_master_validated_all": issue_master_validated_all,
        "issue_master_validated_current": issue_master_validated_current,
        "issue_master_validated_related": issue_master_validated_related,
        "issue_master_validated_excluded": issue_master_validated_excluded,
        "company_bond_summary_validated": company_bond_summary_validated,
        "company_bond_quarterly_validated": company_bond_quarterly_validated,
        "company_bond_relationships": relationship_summary,
        "coupon_schedule_validated": coupon_schedule_validated,
        "offer_terms_validated": offer_terms_validated,
        "quote_snapshot": quote_snapshot,
        "parse_log": parse_log,
        "missing_strong_candidates": missing_strong,
    }
    write_workbook(validated_workbook, validated_sheets)
    write_workbook(quotes_output, build_quotes_workbook(issue_master_validated_all, missing_strong))

    audit_sheets = {
        "summary": pd.DataFrame(
            [
                {"metric": "pending_rows_parsed", "value": int(len(pending_classified))},
                {"metric": "pending_include_current", "value": int(len(pending_current))},
                {"metric": "pending_include_related", "value": int(len(pending_related))},
                {"metric": "pending_excluded", "value": int(len(pending_excluded))},
            ]
        ),
        "pending_issue_master_classified": pending_classified,
        "pending_include_current": pending_current,
        "pending_include_related": pending_related,
        "pending_excluded": pending_excluded,
        "pending_coupon_schedule": pending_coupon,
        "pending_offer_terms": pending_offer,
        "pending_quote_snapshot": pending_quote,
        "pending_parse_log": pending_parse_log,
    }
    write_workbook(audit_output, audit_sheets)

    return {
        "issue_master_validated_all": issue_master_validated_all,
        "issue_master_validated_current": issue_master_validated_current,
        "issue_master_validated_related": issue_master_validated_related,
        "issue_master_validated_excluded": issue_master_validated_excluded,
        "company_bond_summary_validated": company_bond_summary_validated,
        "company_bond_quarterly_validated": company_bond_quarterly_validated,
        "pending_issue_master_classified": pending_classified,
        "pending_include_current": pending_current,
        "pending_include_related": pending_related,
        "pending_excluded": pending_excluded,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Integrate newly captured pending Cbonds issue cards into validated Cbonds outputs and rebuild model-facing bond layers."
    )
    parser.add_argument("--helper", type=Path, default=DEFAULT_HELPER_WORKBOOK)
    parser.add_argument("--validated", type=Path, default=DEFAULT_VALIDATED_WORKBOOK)
    parser.add_argument("--pending-queue", type=Path, default=DEFAULT_PENDING_QUEUE)
    parser.add_argument("--pending-issue-cards", type=Path, default=DEFAULT_PENDING_ISSUE_CARDS)
    parser.add_argument("--quotes-output", type=Path, default=DEFAULT_QUOTES_WORKBOOK)
    parser.add_argument("--audit-output", type=Path, default=DEFAULT_AUDIT_OUTPUT)
    args = parser.parse_args()

    outputs = build_validated_workbook(
        helper_workbook=args.helper,
        validated_workbook=args.validated,
        pending_queue_workbook=args.pending_queue,
        pending_issue_cards_workbook=args.pending_issue_cards,
        quotes_output=args.quotes_output,
        audit_output=args.audit_output,
    )

    print(f"Updated validated workbook: {args.validated}")
    print(f"Updated quotes workbook: {args.quotes_output}")
    print(f"Pending integration audit workbook: {args.audit_output}")
    print(f"Pending direct issues added: {len(outputs['pending_include_current'])}")
    print(f"Pending related issues added: {len(outputs['pending_include_related'])}")
    print(f"Pending issues excluded: {len(outputs['pending_excluded'])}")


if __name__ == "__main__":
    main()
