from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from .build_cbonds_bond_cards import (
    build_company_bond_quarterly,
    build_company_bond_summary,
    write_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PROCESSED_WORKBOOK = PROJECT_ROOT / "data_processed" / "cbonds_bond_cards_processed.xlsx"
DEFAULT_REVIEW_WORKBOOK = PROJECT_ROOT / "data_processed" / "cbonds_quotes_full_run.xlsx"
DEFAULT_HELPER_WORKBOOK = PROJECT_ROOT / "data_processed" / "cbonds_manual_helper.xlsx"
DEFAULT_VALIDATED_OUTPUT = PROJECT_ROOT / "data_processed" / "cbonds_bond_cards_validated.xlsx"
DEFAULT_QUOTES_OUTPUT = PROJECT_ROOT / "data_processed" / "cbonds_quotes_validated.xlsx"
DEFAULT_MISSING_OUTPUT = PROJECT_ROOT / "data_processed" / "cbonds_missing_strong_candidates.xlsx"

REVIEW_CODE_LABELS = {
    1: "direct_current_company",
    2: "group_financing_agent",
    4: "exclude_unrelated",
    5: "structured_sber_linked_note",
    6: "same_group_other_legal_entity",
    7: "different_company_not_in_sample",
}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).replace("\xa0", " ").strip()


def read_sheet_or_empty(workbook: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(workbook, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()


def load_review_table(review_workbook: Path) -> pd.DataFrame:
    review = pd.read_excel(review_workbook, sheet_name="quotes_full_run").copy()
    review["source_workbook"] = review["source_workbook"].map(clean_text)
    review["sheet_name"] = review["sheet_name"].map(clean_text)
    review["validated_manual_code"] = pd.to_numeric(review["Результат_проверки"], errors="coerce").astype("Int64")

    bad_codes = sorted(
        {
            int(code)
            for code in review["validated_manual_code"].dropna().astype(int).tolist()
            if int(code) not in REVIEW_CODE_LABELS
        }
    )
    if bad_codes:
        raise ValueError(f"Unknown manual review codes in quotes_full_run: {bad_codes}")

    auto_direct_mask = (
        review["validated_manual_code"].isna()
        & review["quotes_priority_bucket"].eq("definitely_needed")
        & review["borrower_match_bucket"].eq("high")
        & (~review["manual_review_needed_flag"].fillna(False))
    )

    review["validated_manual_code_effective"] = review["validated_manual_code"]
    review.loc[auto_direct_mask, "validated_manual_code_effective"] = 1
    review["validated_code_source"] = np.select(
        [
            auto_direct_mask,
            review["validated_manual_code"].notna(),
        ],
        [
            "auto_from_definitely_needed",
            "manual_user_review",
        ],
        default="unreviewed",
    )
    review["validated_relation_label"] = (
        review["validated_manual_code_effective"]
        .map(lambda value: REVIEW_CODE_LABELS.get(int(value), "") if pd.notna(value) else "")
        .astype(str)
    )
    review["validated_keep_for_current_company_flag"] = review["validated_manual_code_effective"].eq(1)
    review["validated_group_financing_agent_flag"] = review["validated_manual_code_effective"].eq(2)
    review["validated_unrelated_exclude_flag"] = review["validated_manual_code_effective"].eq(4)
    review["validated_structured_linked_flag"] = review["validated_manual_code_effective"].eq(5)
    review["validated_same_group_other_legal_flag"] = review["validated_manual_code_effective"].eq(6)
    review["validated_reassign_not_in_sample_flag"] = review["validated_manual_code_effective"].eq(7)
    review["validated_keep_related_optional_flag"] = review["validated_manual_code_effective"].isin([2, 6, 7])
    review["validated_drop_from_current_company_flag"] = review["validated_manual_code_effective"].isin([4, 5])
    review["validated_any_keep_flag"] = review["validated_keep_for_current_company_flag"] | review["validated_keep_related_optional_flag"]
    review["quote_collection_bucket_validated"] = np.select(
        [
            review["validated_keep_for_current_company_flag"],
            review["validated_group_financing_agent_flag"],
            review["validated_same_group_other_legal_flag"],
            review["validated_reassign_not_in_sample_flag"],
            review["validated_structured_linked_flag"],
            review["validated_unrelated_exclude_flag"],
        ],
        [
            "current_company_direct",
            "group_agent_optional",
            "same_group_optional",
            "reassign_optional",
            "exclude_structured",
            "exclude_unrelated",
        ],
        default="unreviewed",
    )
    review["quote_collection_rank_validated"] = (
        review["quote_collection_bucket_validated"]
        .map(
            {
                "current_company_direct": 1,
                "group_agent_optional": 2,
                "same_group_optional": 3,
                "reassign_optional": 4,
                "exclude_structured": 8,
                "exclude_unrelated": 9,
                "unreviewed": 7,
            }
        )
        .fillna(9)
        .astype(int)
    )
    return review


def merge_review(issue_master: pd.DataFrame, review: pd.DataFrame) -> pd.DataFrame:
    review_cols = [
        "source_workbook",
        "sheet_name",
        "validated_manual_code",
        "validated_manual_code_effective",
        "validated_code_source",
        "validated_relation_label",
        "validated_keep_for_current_company_flag",
        "validated_group_financing_agent_flag",
        "validated_unrelated_exclude_flag",
        "validated_structured_linked_flag",
        "validated_same_group_other_legal_flag",
        "validated_reassign_not_in_sample_flag",
        "validated_keep_related_optional_flag",
        "validated_drop_from_current_company_flag",
        "validated_any_keep_flag",
        "quote_collection_bucket_validated",
        "quote_collection_rank_validated",
    ]
    out = issue_master.copy()
    out["source_workbook"] = out["source_workbook"].map(clean_text)
    out["sheet_name"] = out["sheet_name"].map(clean_text)
    out = out.merge(review[review_cols], on=["source_workbook", "sheet_name"], how="left")
    for column in [
        "validated_keep_for_current_company_flag",
        "validated_group_financing_agent_flag",
        "validated_unrelated_exclude_flag",
        "validated_structured_linked_flag",
        "validated_same_group_other_legal_flag",
        "validated_reassign_not_in_sample_flag",
        "validated_keep_related_optional_flag",
        "validated_drop_from_current_company_flag",
        "validated_any_keep_flag",
    ]:
        out[column] = out[column].fillna(False)
    return out


def build_relationship_summary(issue_master_validated: pd.DataFrame) -> pd.DataFrame:
    if issue_master_validated.empty:
        return pd.DataFrame(columns=["company_id"])
    grouped = issue_master_validated.groupby("company_id", dropna=False)
    summary = grouped.agg(
        cbonds_validated_direct_issue_count=("validated_keep_for_current_company_flag", "sum"),
        cbonds_group_agent_issue_count=("validated_group_financing_agent_flag", "sum"),
        cbonds_same_group_other_legal_issue_count=("validated_same_group_other_legal_flag", "sum"),
        cbonds_structured_linked_issue_count=("validated_structured_linked_flag", "sum"),
        cbonds_reassign_not_in_sample_issue_count=("validated_reassign_not_in_sample_flag", "sum"),
        cbonds_unrelated_excluded_issue_count=("validated_unrelated_exclude_flag", "sum"),
    ).reset_index()
    for count_col, flag_col in [
        ("cbonds_validated_direct_issue_count", "cbonds_validated_direct_issue_flag"),
        ("cbonds_group_agent_issue_count", "cbonds_group_agent_issue_flag"),
        ("cbonds_same_group_other_legal_issue_count", "cbonds_same_group_other_legal_issue_flag"),
        ("cbonds_structured_linked_issue_count", "cbonds_structured_linked_issue_flag"),
        ("cbonds_reassign_not_in_sample_issue_count", "cbonds_reassign_not_in_sample_issue_flag"),
        ("cbonds_unrelated_excluded_issue_count", "cbonds_unrelated_excluded_issue_flag"),
    ]:
        summary[flag_col] = summary[count_col].fillna(0).astype(int).gt(0)
    summary["cbonds_any_related_issue_flag"] = summary[
        [
            "cbonds_group_agent_issue_flag",
            "cbonds_same_group_other_legal_issue_flag",
            "cbonds_reassign_not_in_sample_issue_flag",
        ]
    ].any(axis=1)
    summary["cbonds_any_nonexcluded_issue_flag"] = (
        summary["cbonds_validated_direct_issue_count"].fillna(0).astype(int)
        + summary["cbonds_group_agent_issue_count"].fillna(0).astype(int)
        + summary["cbonds_same_group_other_legal_issue_count"].fillna(0).astype(int)
        + summary["cbonds_reassign_not_in_sample_issue_count"].fillna(0).astype(int)
    ) > 0
    return summary


def build_missing_strong_candidates(
    helper_workbook: Path,
    issue_master_validated: pd.DataFrame,
) -> pd.DataFrame:
    helper = pd.read_excel(helper_workbook, sheet_name="bond_queries").copy()
    helper["best_n_trade_dates"] = pd.to_numeric(helper["best_n_trade_dates"], errors="coerce")
    strong = helper.loc[
        helper["instrument_type"].eq("bond")
        & (
            helper["reliable_mapping_flag"].fillna(False)
            | helper["best_n_trade_dates"].fillna(0).ge(100)
        )
    ].copy()

    existing_queries = set(issue_master_validated["query_text"].dropna().astype(str))
    existing_isins = set(issue_master_validated["parsed_isin"].dropna().astype(str))
    strong["capture_lookup_key"] = strong["cbonds_paper_query_primary"].astype(str)
    strong["captured_issue_card_flag"] = strong["capture_lookup_key"].isin(existing_queries | existing_isins)
    missing = strong.loc[~strong["captured_issue_card_flag"]].copy()
    missing["public_company_flag"] = missing["company_name"].astype(str).str.contains(r"\bПАО\b", case=False, na=False)
    missing["capture_issue_card_needed_flag"] = True
    missing["quote_collection_bucket_validated"] = "missing_issue_card_capture_first"
    missing["quote_collection_rank_validated"] = 5
    missing["validated_relation_label"] = "missing_issue_card_not_captured"
    missing["validated_keep_for_current_company_flag"] = True
    missing["validated_keep_related_optional_flag"] = False
    missing["validated_drop_from_current_company_flag"] = False
    missing["quote_collection_note_validated"] = (
        "Сильный кандидат из helper/MOEX/T-Invest слоя отсутствует среди захваченных карточек Cbonds; "
        "нужно отдельно добрать карточку выпуска, после чего можно собирать котировки."
    )
    selected = [
        "company_name",
        "company_inn",
        "sector",
        "sample_flag",
        "cbonds_paper_query_primary",
        "instrument_name",
        "best_n_trade_dates",
        "reliable_mapping_flag",
        "best_match_score",
        "public_company_flag",
        "capture_issue_card_needed_flag",
        "quote_collection_bucket_validated",
        "quote_collection_rank_validated",
        "validated_relation_label",
        "quote_collection_note_validated",
    ]
    selected = [column for column in selected if column in missing.columns]
    return missing[selected].sort_values(
        ["public_company_flag", "company_name", "best_n_trade_dates"],
        ascending=[False, True, False],
        kind="stable",
    ).reset_index(drop=True)


def build_quotes_outputs(
    review: pd.DataFrame,
    missing_strong: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    quotes = review.copy()
    quotes["quote_collection_note_validated"] = np.select(
        [
            quotes["validated_keep_for_current_company_flag"],
            quotes["validated_group_financing_agent_flag"],
            quotes["validated_same_group_other_legal_flag"],
            quotes["validated_reassign_not_in_sample_flag"],
            quotes["validated_structured_linked_flag"],
            quotes["validated_unrelated_exclude_flag"],
        ],
        [
            "Собирать котировки как прямой выпуск компании из выборки.",
            "Собирать как связанный выпуск финансового агента группы; в модель как выпуск текущего юрлица не тащить.",
            "Собирать опционально как выпуск другой компании той же группы; в модель текущего юрлица не тащить.",
            "Собирать опционально: выпуск относится к компании вне текущего shortlist, возможно пригодится для расширения выборки.",
            "Не собирать в модельный слой: структурный продукт банка, привязанный к акции компании.",
            "Не собирать в модельный слой: ручная проверка показала отсутствие связи с компанией.",
        ],
        default="Требует дополнительной ручной проверки.",
    )
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
        [
            current_company,
            related_optional,
            missing_for_quotes,
        ],
        ignore_index=True,
        sort=False,
    )
    sort_columns = [column for column in ["quote_collection_rank_validated", "company_name", "issue_name"] if column in full_collect.columns]
    if sort_columns:
        full_collect = full_collect.sort_values(sort_columns, kind="stable").reset_index(drop=True)

    quotes_sheets = {
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
    return quotes, quotes_sheets


def build_validated_outputs(
    processed_workbook: Path,
    review_workbook: Path,
    helper_workbook: Path,
    validated_output: Path,
    quotes_output: Path,
    missing_output: Path,
) -> dict[str, pd.DataFrame]:
    issue_master_clean = pd.read_excel(processed_workbook, sheet_name="issue_master_clean")
    coupon_schedule = read_sheet_or_empty(processed_workbook, "coupon_schedule")
    offer_terms = read_sheet_or_empty(processed_workbook, "offer_terms")
    quote_snapshot = read_sheet_or_empty(processed_workbook, "quote_snapshot")
    parse_log = read_sheet_or_empty(processed_workbook, "parse_log")

    review = load_review_table(review_workbook)
    issue_master_validated_all = merge_review(issue_master_clean, review)
    issue_master_validated_current = issue_master_validated_all.loc[
        issue_master_validated_all["validated_keep_for_current_company_flag"]
    ].copy()
    issue_master_validated_related = issue_master_validated_all.loc[
        issue_master_validated_all["validated_keep_related_optional_flag"]
    ].copy()
    issue_master_validated_excluded = issue_master_validated_all.loc[
        issue_master_validated_all["validated_drop_from_current_company_flag"]
    ].copy()

    coupon_schedule_validated = (
        coupon_schedule.loc[coupon_schedule["issue_key"].isin(issue_master_validated_current["issue_key"])].copy()
        if not coupon_schedule.empty
        else pd.DataFrame()
    )
    offer_terms_validated = (
        offer_terms.loc[offer_terms["issue_key"].isin(issue_master_validated_current["issue_key"])].copy()
        if not offer_terms.empty
        else pd.DataFrame()
    )

    company_bond_summary_validated = build_company_bond_summary(issue_master_validated_current)
    company_bond_quarterly_validated = build_company_bond_quarterly(issue_master_validated_current, offer_terms_validated)
    relationship_summary = build_relationship_summary(issue_master_validated_all)
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
        current_company_index = (
            issue_master_validated_all[["company_id", "company_name", "company_inn"]]
            .drop_duplicates()
        )
        missing_counts = current_company_index.merge(
            missing_counts,
            on=["company_name", "company_inn"],
            how="left",
        )
        missing_counts["cbonds_missing_strong_issue_count"] = (
            missing_counts["cbonds_missing_strong_issue_count"].fillna(0).astype(int)
        )
        company_bond_summary_validated = company_bond_summary_validated.merge(
            missing_counts[["company_id", "cbonds_missing_strong_issue_count"]],
            on="company_id",
            how="left",
        )
    else:
        company_bond_summary_validated["cbonds_missing_strong_issue_count"] = 0

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
                company_bond_summary_validated[column], errors="coerce"
            ).fillna(0)
    for column in fill_false_cols:
        if column in company_bond_summary_validated.columns:
            company_bond_summary_validated[column] = company_bond_summary_validated[column].fillna(False).astype(bool)

    review_summary = (
        issue_master_validated_all.groupby(
            ["validated_relation_label", "validated_code_source"], dropna=False
        )
        .size()
        .reset_index(name="issue_count")
        .sort_values(["validated_relation_label", "validated_code_source"], kind="stable")
        .reset_index(drop=True)
    )

    quotes_validated_all, quotes_sheets = build_quotes_outputs(review, missing_strong)

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

    write_workbook(validated_output, validated_sheets)
    write_workbook(quotes_output, quotes_sheets)
    write_workbook(
        missing_output,
        {
            "missing_strong_candidates": missing_strong,
            "public_company_focus": missing_strong.loc[missing_strong["public_company_flag"].fillna(False)].copy()
            if not missing_strong.empty
            else pd.DataFrame(),
        },
    )

    return {
        "issue_master_validated_all": issue_master_validated_all,
        "issue_master_validated_current": issue_master_validated_current,
        "issue_master_validated_related": issue_master_validated_related,
        "issue_master_validated_excluded": issue_master_validated_excluded,
        "company_bond_summary_validated": company_bond_summary_validated,
        "company_bond_quarterly_validated": company_bond_quarterly_validated,
        "missing_strong_candidates": missing_strong,
        "quotes_validated_all": quotes_validated_all,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Apply manual Cbonds bond review codes to captured issue cards and build validated outputs."
    )
    parser.add_argument("--processed", type=Path, default=DEFAULT_PROCESSED_WORKBOOK)
    parser.add_argument("--review", type=Path, default=DEFAULT_REVIEW_WORKBOOK)
    parser.add_argument("--helper", type=Path, default=DEFAULT_HELPER_WORKBOOK)
    parser.add_argument("--validated-output", type=Path, default=DEFAULT_VALIDATED_OUTPUT)
    parser.add_argument("--quotes-output", type=Path, default=DEFAULT_QUOTES_OUTPUT)
    parser.add_argument("--missing-output", type=Path, default=DEFAULT_MISSING_OUTPUT)
    args = parser.parse_args()

    outputs = build_validated_outputs(
        processed_workbook=args.processed,
        review_workbook=args.review,
        helper_workbook=args.helper,
        validated_output=args.validated_output,
        quotes_output=args.quotes_output,
        missing_output=args.missing_output,
    )

    print(f"Validated workbook: {args.validated_output}")
    print(f"Validated quotes workbook: {args.quotes_output}")
    print(f"Missing-strong audit workbook: {args.missing_output}")
    print(f"Current-company validated issues: {len(outputs['issue_master_validated_current'])}")
    print(f"Companies with current-company validated issues: {outputs['issue_master_validated_current']['company_id'].nunique() if not outputs['issue_master_validated_current'].empty else 0}")
    print(f"Strong missing issue cards: {len(outputs['missing_strong_candidates'])}")


if __name__ == "__main__":
    main()
