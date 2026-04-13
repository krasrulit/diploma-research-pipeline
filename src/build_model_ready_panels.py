from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ANALYSIS_PANEL = PROJECT_ROOT / "data_processed" / "analysis_panel.xlsx"
DEFAULT_QUARTERLY_OUTPUT = PROJECT_ROOT / "data_processed" / "model_ready_quarterly.xlsx"
DEFAULT_ANNUAL_OUTPUT = PROJECT_ROOT / "data_processed" / "model_ready_annual.xlsx"

BASE_IDENTIFIER_COLUMNS = [
    "company_id",
    "sector",
    "sample_flag",
    "sample_main_flag",
    "sample_extended_flag",
    "company_name_shortlist",
    "company",
    "inn",
    "ogrn",
    "ticker",
    "isin",
    "industry",
    "report_year",
    "period_type",
    "quarter_label",
    "quarter_end_date",
]

TRACEABILITY_COLUMNS = [
    "source_file",
    "source_report_period",
    "source_period_type",
    "data_source",
    "unit",
    "coverage_source_files",
]

AVAILABILITY_COLUMNS = [
    "reporting_observed_flag",
    "available_core_metrics_count",
    "has_balance",
    "has_pnl",
    "has_cash_flow",
    "n_tables",
    "n_metrics_current",
    "n_values_total",
    "shortlist_match_flag",
    "public_market_enriched_flag",
]

FINANCIAL_LEVEL_COLUMNS = [
    "assets_total",
    "ppe",
    "cash",
    "equity",
    "debt_lt",
    "debt_st",
    "total_debt",
    "revenue",
    "interest_expense",
    "net_income",
    "ebit",
    "cfo",
    "capex",
    "investing_payments",
    "financing_inflows",
    "loans_received",
    "bonds_issued",
    "debt_repayment",
    "interest_paid",
]

RATIO_COLUMNS = [
    "debt_to_assets",
    "debt_lt_share",
    "cash_to_assets",
    "ppe_to_assets",
    "equity_to_assets",
    "net_margin",
    "ebit_margin",
    "net_debt",
    "net_debt_to_assets",
    "cfo_to_assets",
    "capex_to_assets",
    "capex_to_cfo",
    "interest_coverage",
    "net_margin_winsor_1_99",
    "ebit_margin_winsor_1_99",
    "log_assets",
    "log_revenue",
]

MARKET_FLAG_COLUMNS = [
    "market_n_validated_instruments",
    "market_n_validated_bonds",
    "market_n_validated_shares",
    "market_has_validated_public_instrument",
    "market_has_validated_public_bond",
    "market_has_validated_public_share",
    "market_n_moex_selected_rows",
    "market_has_moex_selected",
    "market_n_tinvest_selected_rows",
    "market_has_tinvest_selected",
    "market_n_moex_bonds_rows",
    "market_has_moex_bonds",
    "market_n_tinvest_bonds_rows",
    "market_has_tinvest_bonds",
]

MARKET_ACCESS_COLUMNS = [
    "market_access_status",
    "market_primary_source",
    "market_has_any_candidate_flag",
    "market_has_history_flag",
    "market_has_usable_history_flag",
    "market_has_reliable_mapping_flag",
    "market_has_usable_share_flag",
    "market_has_usable_bond_flag",
    "market_has_reliable_share_flag",
    "market_has_reliable_bond_flag",
    "market_needs_manual_review_flag",
    "n_security_groups_total",
    "n_groups_with_history",
    "n_groups_usable_history",
    "n_groups_reliable",
    "n_share_groups_usable",
    "n_bond_groups_usable",
    "n_share_groups_reliable",
    "n_bond_groups_reliable",
    "n_groups_manual_review",
    "n_groups_best_source_moex",
    "n_groups_best_source_tinvest",
    "max_trade_dates_any",
    "max_trade_dates_usable",
    "market_history_start_min",
    "market_history_end_max",
]

OWNERSHIP_COLUMNS = [
    "ownership_report_found_flag",
    "ownership_head_company_name",
    "ownership_n_subsidiaries",
    "ownership_has_head_company_field_flag",
    "ownership_has_subsidiaries_flag",
    "ownership_is_subsidiary_flag",
    "ownership_is_parent_flag",
    "ownership_ownership_role",
    "ownership_group_name_inferred",
    "ownership_state_owned_flag",
    "ownership_state_bucket",
    "ownership_group_inference_source",
    "ownership_group_inference_confidence",
    "ownership_ownership_review_needed_flag",
]

CBONDS_STATIC_COLUMNS = [
    "cbonds_issue_card_found_flag",
    "cbonds_issue_count",
    "cbonds_issue_count_high_conf",
    "cbonds_issue_count_outstanding",
    "cbonds_issue_count_redeemed",
    "cbonds_issue_count_default",
    "cbonds_issue_count_need_quotes_high",
    "cbonds_issue_count_need_quotes_lower",
    "cbonds_issue_count_manual_review",
    "cbonds_issue_currency_count",
    "cbonds_issue_volume_sum",
    "cbonds_issue_volume_rub_sum",
    "cbonds_circulation_volume_sum",
    "cbonds_circulation_volume_rub_sum",
    "cbonds_outstanding_issue_volume_rub_sum",
    "cbonds_has_rating_issue_flag",
    "cbonds_has_listing_issue_flag",
    "cbonds_has_offer_issue_flag",
    "cbonds_first_registration_date",
    "cbonds_first_placement_end_date",
    "cbonds_last_maturity_date",
    "cbonds_coupon_rate_avg_pct",
    "cbonds_coupon_rate_max_pct",
    "cbonds_borrower_match_high_conf_flag",
]

CBONDS_QUARTERLY_COLUMNS = [
    "cbonds_issue_registration_q_count",
    "cbonds_issue_registration_q_volume_rub",
    "cbonds_issue_placement_q_count",
    "cbonds_issue_placement_q_volume_rub",
    "cbonds_issue_maturity_q_count",
    "cbonds_issue_maturity_q_volume_rub",
    "cbonds_issue_offer_q_count",
    "cbonds_issue_offer_q_price_avg",
    "cbonds_issue_outstanding_q_count",
    "cbonds_issue_outstanding_q_volume_rub",
]

MACRO_COLUMNS = [
    "macro_key_rate_avg_q",
    "macro_key_rate_end_q",
    "macro_key_rate_max_q",
    "macro_key_rate_min_q",
    "macro_inflation_yoy_avg_q",
    "macro_inflation_target_avg_q",
    "macro_key_rate_month_end_avg_q",
    "macro_usdrub_avg_q",
    "macro_usdrub_end_q",
    "macro_usdrub_max_q",
    "macro_usdrub_min_q",
    "covid_flag",
    "post_2022_flag",
]


def ensure_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = frame.copy()
    for column in columns:
        if column not in out.columns:
            out[column] = np.nan
    return out


def commodity_columns(frame: pd.DataFrame) -> list[str]:
    return sorted([column for column in frame.columns if column.startswith("cmd_") and column.endswith("_avg_q")])


def winsorize_series(series: pd.Series, lower_q: float = 0.01, upper_q: float = 0.99) -> pd.Series:
    out = pd.to_numeric(series, errors="coerce")
    valid = out.dropna()
    if valid.empty:
        return out
    lower = valid.quantile(lower_q)
    upper = valid.quantile(upper_q)
    return out.clip(lower=lower, upper=upper)


def as_bool_series(series: pd.Series) -> pd.Series:
    if series.dtype == bool:
        return series.fillna(False)
    text = series.fillna("").astype(str).str.strip().str.lower()
    return text.isin({"1", "1.0", "true", "yes"})


def add_engineered_columns(panel: pd.DataFrame) -> pd.DataFrame:
    out = panel.copy()
    out["net_debt"] = out["total_debt"] - out["cash"]
    out["net_debt_to_assets"] = out["net_debt"] / out["assets_total"]
    out["cfo_to_assets"] = out["cfo"] / out["assets_total"]
    out["capex_to_assets"] = out["capex"] / out["assets_total"]
    out["capex_to_cfo"] = out["capex"] / out["cfo"]
    out["interest_coverage"] = out["ebit"] / out["interest_expense"]
    out["net_margin_winsor_1_99"] = winsorize_series(out["net_margin"])
    out["ebit_margin_winsor_1_99"] = winsorize_series(out["ebit_margin"])
    return out


def select_model_columns(frame: pd.DataFrame) -> pd.DataFrame:
    columns = (
        BASE_IDENTIFIER_COLUMNS
        + TRACEABILITY_COLUMNS
        + AVAILABILITY_COLUMNS
        + FINANCIAL_LEVEL_COLUMNS
        + RATIO_COLUMNS
        + MARKET_FLAG_COLUMNS
        + MARKET_ACCESS_COLUMNS
        + OWNERSHIP_COLUMNS
        + CBONDS_STATIC_COLUMNS
        + CBONDS_QUARTERLY_COLUMNS
        + MACRO_COLUMNS
        + commodity_columns(frame)
    )
    out = ensure_columns(frame, columns)
    return out[columns]


def build_summary(panel: pd.DataFrame, frequency: str) -> pd.DataFrame:
    companies_with_data = (
        panel.groupby(["sector", "company_id"], dropna=False)["reporting_observed_flag"]
        .max()
        .reset_index()
    )
    market_history_mask = as_bool_series(panel["market_has_history_flag"])
    market_usable_mask = as_bool_series(panel["market_has_usable_history_flag"])
    cbonds_issue_mask = as_bool_series(panel["cbonds_issue_card_found_flag"])
    cbonds_outstanding_mask = pd.to_numeric(panel["cbonds_issue_outstanding_q_count"], errors="coerce").fillna(0).gt(0)
    rows = [
        {"frequency": frequency, "metric": "n_rows", "value": len(panel)},
        {"frequency": frequency, "metric": "n_unique_companies", "value": int(panel["company_id"].nunique())},
        {"frequency": frequency, "metric": "n_oil_gas_companies", "value": int(panel.loc[panel["sector"].eq("oil_gas"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_metallurgy_companies", "value": int(panel.loc[panel["sector"].eq("metallurgy"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_main_companies", "value": int(panel.loc[panel["sample_flag"].eq("main"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_extended_companies", "value": int(panel.loc[panel["sample_flag"].eq("extended"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "rows_with_reporting_data", "value": int(panel["reporting_observed_flag"].fillna(False).sum())},
        {"frequency": frequency, "metric": "rows_with_reporting_data_share", "value": float(panel["reporting_observed_flag"].fillna(False).mean())},
        {"frequency": frequency, "metric": "companies_with_reporting_data", "value": int(companies_with_data["reporting_observed_flag"].fillna(False).sum())},
        {"frequency": frequency, "metric": "rows_with_usdrub", "value": int(panel["macro_usdrub_avg_q"].notna().sum())},
        {"frequency": frequency, "metric": "companies_with_state_bucket", "value": int(panel.loc[panel["ownership_state_bucket"].notna(), "company_id"].nunique())},
        {"frequency": frequency, "metric": "companies_state_owned", "value": int(panel.loc[pd.to_numeric(panel["ownership_state_owned_flag"], errors="coerce").eq(1), "company_id"].nunique())},
        {"frequency": frequency, "metric": "companies_with_market_history", "value": int(panel.loc[market_history_mask, "company_id"].nunique())},
        {"frequency": frequency, "metric": "companies_with_usable_market_history", "value": int(panel.loc[market_usable_mask, "company_id"].nunique())},
        {"frequency": frequency, "metric": "companies_with_cbonds_issue_cards", "value": int(panel.loc[cbonds_issue_mask, "company_id"].nunique())},
        {"frequency": frequency, "metric": "rows_with_cbonds_outstanding_issue", "value": int(cbonds_outstanding_mask.sum())},
        {
            "frequency": frequency,
            "metric": "oil_gas_companies_with_reporting_data",
            "value": int(
                companies_with_data.loc[companies_with_data["sector"].eq("oil_gas"), "reporting_observed_flag"].fillna(False).sum()
            ),
        },
        {
            "frequency": frequency,
            "metric": "metallurgy_companies_with_reporting_data",
            "value": int(
                companies_with_data.loc[companies_with_data["sector"].eq("metallurgy"), "reporting_observed_flag"].fillna(False).sum()
            ),
        },
        {"frequency": frequency, "metric": "rows_with_assets", "value": int(panel["assets_total"].notna().sum())},
        {"frequency": frequency, "metric": "rows_with_revenue", "value": int(panel["revenue"].notna().sum())},
    ]
    return pd.DataFrame(rows)


def variable_group(column: str) -> str:
    if column in BASE_IDENTIFIER_COLUMNS:
        return "identifier"
    if column in TRACEABILITY_COLUMNS:
        return "traceability"
    if column in AVAILABILITY_COLUMNS:
        return "availability"
    if column in FINANCIAL_LEVEL_COLUMNS:
        return "financial_level"
    if column in RATIO_COLUMNS:
        return "ratio"
    if column in MARKET_FLAG_COLUMNS:
        return "public_market_flag"
    if column in MARKET_ACCESS_COLUMNS:
        return "market_access"
    if column in OWNERSHIP_COLUMNS:
        return "ownership"
    if column in CBONDS_STATIC_COLUMNS:
        return "cbonds_static"
    if column in CBONDS_QUARTERLY_COLUMNS:
        return "cbonds_quarterly"
    if column in MACRO_COLUMNS:
        return "macro"
    if column.startswith("cmd_"):
        return "commodity"
    return "other"


def variable_description(column: str) -> str:
    descriptions = {
        "company_id": "Stable company key built from sector and INN.",
        "sector": "Sector flag: oil_gas or metallurgy.",
        "sample_flag": "Main vs extended shortlist flag.",
        "sample_main_flag": "1 if company belongs to the main sample.",
        "sample_extended_flag": "1 if company belongs to the extended sample.",
        "company_name_shortlist": "Company name from shortlist.",
        "company": "Company name from SPARK report filename.",
        "inn": "Taxpayer identification number.",
        "ogrn": "Primary state registration number.",
        "ticker": "Exchange ticker from shortlist when available.",
        "isin": "ISIN from shortlist when available.",
        "industry": "Industry description from shortlist.",
        "report_year": "Calendar year of the observation.",
        "period_type": "Quarter label inside the year: Q1/Q2/Q3/Q4.",
        "quarter_label": "Canonical quarter label, e.g. 2022Q3.",
        "quarter_end_date": "Quarter end date.",
        "reporting_observed_flag": "1 if SPARK reports at least one financial form for the period.",
        "available_core_metrics_count": "Number of non-missing core financial metrics in the row.",
        "has_balance": "1 if balance sheet is present.",
        "has_pnl": "1 if income statement is present.",
        "has_cash_flow": "1 if cash flow statement is present.",
        "total_debt": "Debt long-term plus debt short-term.",
        "debt_to_assets": "Total debt divided by total assets.",
        "debt_lt_share": "Long-term debt share in total debt.",
        "cash_to_assets": "Cash divided by total assets.",
        "ppe_to_assets": "Property, plant and equipment divided by total assets.",
        "equity_to_assets": "Equity divided by total assets.",
        "net_margin": "Net income divided by revenue.",
        "ebit_margin": "EBIT divided by revenue.",
        "net_debt": "Total debt minus cash.",
        "net_debt_to_assets": "Net debt divided by total assets.",
        "cfo_to_assets": "Cash flow from operations divided by total assets.",
        "capex_to_assets": "Capital expenditures divided by total assets.",
        "capex_to_cfo": "Capital expenditures divided by cash flow from operations.",
        "interest_coverage": "EBIT divided by interest expense.",
        "net_margin_winsor_1_99": "Net margin winsorized at the 1st and 99th percentiles within the panel frequency.",
        "ebit_margin_winsor_1_99": "EBIT margin winsorized at the 1st and 99th percentiles within the panel frequency.",
        "log_assets": "Natural log of total assets.",
        "log_revenue": "Natural log of revenue.",
        "public_market_enriched_flag": "1 if public-market enrichment from the oil-gas pipeline is available.",
        "market_access_status": "Company-level market access bucket from the cleaned securities layer.",
        "market_primary_source": "Preferred public market source for the company-level access layer.",
        "market_has_any_candidate_flag": "1 if at least one public security candidate was found.",
        "market_has_history_flag": "1 if at least one resolved security has any history.",
        "market_has_usable_history_flag": "1 if at least one resolved security has usable history coverage.",
        "market_has_reliable_mapping_flag": "1 if at least one resolved security has reliable issuer mapping.",
        "market_has_usable_share_flag": "1 if the company has a usable share history.",
        "market_has_usable_bond_flag": "1 if the company has a usable bond history.",
        "market_has_reliable_share_flag": "1 if the company has a reliably matched share.",
        "market_has_reliable_bond_flag": "1 if the company has a reliably matched bond.",
        "market_needs_manual_review_flag": "1 if the cleaned market-access layer still needs manual review.",
        "n_security_groups_total": "Number of resolved public security groups linked to the company.",
        "n_groups_with_history": "Number of resolved security groups with any history.",
        "n_groups_usable_history": "Number of resolved security groups with usable history.",
        "n_groups_reliable": "Number of resolved security groups with reliable issuer mapping.",
        "n_share_groups_usable": "Number of usable share groups.",
        "n_bond_groups_usable": "Number of usable bond groups.",
        "n_share_groups_reliable": "Number of reliably matched share groups.",
        "n_bond_groups_reliable": "Number of reliably matched bond groups.",
        "n_groups_manual_review": "Number of resolved groups still marked for manual review.",
        "n_groups_best_source_moex": "Number of resolved groups where MOEX is the preferred source.",
        "n_groups_best_source_tinvest": "Number of resolved groups where T-Invest is the preferred source.",
        "max_trade_dates_any": "Maximum number of trade dates across resolved securities for the company.",
        "max_trade_dates_usable": "Maximum number of trade dates across usable resolved securities.",
        "market_history_start_min": "Earliest market-history date across resolved securities.",
        "market_history_end_max": "Latest market-history date across resolved securities.",
        "ownership_report_found_flag": "1 if an ownership/source SPARK DOCX report was found for the company.",
        "ownership_head_company_name": "Head company name parsed from SPARK ownership metadata.",
        "ownership_n_subsidiaries": "Count of subsidiaries parsed from SPARK ownership metadata.",
        "ownership_has_head_company_field_flag": "1 if the SPARK card explicitly lists a head company.",
        "ownership_has_subsidiaries_flag": "1 if the SPARK card explicitly lists subsidiaries.",
        "ownership_is_subsidiary_flag": "1 if the company appears to be a subsidiary.",
        "ownership_is_parent_flag": "1 if the company appears to be a parent company.",
        "ownership_ownership_role": "Ownership role inferred from SPARK card: subsidiary, parent, both, or unknown.",
        "ownership_group_name_inferred": "Inferred business group name from ownership metadata or domain.",
        "ownership_state_owned_flag": "1 if the company was heuristically inferred as state-owned, 0 if private.",
        "ownership_state_bucket": "Simplified ownership class: state, private, unknown.",
        "ownership_group_inference_source": "Field used for ownership/group inference.",
        "ownership_group_inference_confidence": "Confidence level of ownership/group inference.",
        "ownership_ownership_review_needed_flag": "1 if ownership inference should be manually reviewed.",
        "cbonds_issue_card_found_flag": "1 if at least one Cbonds issue card was parsed for the company.",
        "cbonds_issue_count": "Number of deduplicated Cbonds bond issues linked to the company.",
        "cbonds_issue_count_high_conf": "Number of Cbonds bond issues with high-confidence mapping after borrower/card validation.",
        "cbonds_issue_count_outstanding": "Number of Cbonds bond issues currently marked as outstanding.",
        "cbonds_issue_count_redeemed": "Number of Cbonds bond issues currently marked as redeemed.",
        "cbonds_issue_count_default": "Number of Cbonds bond issues currently marked as defaulted.",
        "cbonds_issue_count_need_quotes_high": "Number of bond issues ranked as definitely needing quote history collection.",
        "cbonds_issue_count_need_quotes_lower": "Number of bond issues ranked as lower-priority quote history collection.",
        "cbonds_issue_count_manual_review": "Number of bond issues that should be manually reviewed before collecting quotes.",
        "cbonds_issue_currency_count": "Number of distinct currencies across Cbonds bond issues for the company.",
        "cbonds_issue_volume_sum": "Total issue volume across parsed Cbonds bond issues, regardless of currency.",
        "cbonds_issue_volume_rub_sum": "Total issue volume across RUB-denominated Cbonds bond issues.",
        "cbonds_circulation_volume_sum": "Total circulation volume across parsed Cbonds bond issues, regardless of currency.",
        "cbonds_circulation_volume_rub_sum": "Total circulation volume across RUB-denominated Cbonds bond issues.",
        "cbonds_outstanding_issue_volume_rub_sum": "Approximate total RUB circulation volume of issues still outstanding.",
        "cbonds_has_rating_issue_flag": "1 if any parsed Cbonds issue card contains issue ratings.",
        "cbonds_has_listing_issue_flag": "1 if any parsed Cbonds issue card contains listing information.",
        "cbonds_has_offer_issue_flag": "1 if any parsed Cbonds issue card contains offer or early redemption conditions.",
        "cbonds_first_registration_date": "Earliest bond registration date across parsed Cbonds issue cards for the company.",
        "cbonds_first_placement_end_date": "Earliest placement-end date across parsed Cbonds issue cards for the company.",
        "cbonds_last_maturity_date": "Latest maturity date across parsed Cbonds issue cards for the company.",
        "cbonds_coupon_rate_avg_pct": "Average coupon rate across parsed coupon schedules from Cbonds issue cards.",
        "cbonds_coupon_rate_max_pct": "Maximum coupon rate across parsed coupon schedules from Cbonds issue cards.",
        "cbonds_borrower_match_high_conf_flag": "1 if at least one parsed Cbonds borrower name strongly matches the target company.",
        "cbonds_issue_registration_q_count": "Number of bond issues registered in the quarter according to Cbonds issue cards.",
        "cbonds_issue_registration_q_volume_rub": "RUB issue volume of bond issues registered in the quarter.",
        "cbonds_issue_placement_q_count": "Number of bond issues with placement ending in the quarter according to Cbonds issue cards.",
        "cbonds_issue_placement_q_volume_rub": "RUB issue volume of bond issues placed in the quarter.",
        "cbonds_issue_maturity_q_count": "Number of bond issues maturing in the quarter according to Cbonds issue cards.",
        "cbonds_issue_maturity_q_volume_rub": "RUB circulation volume of bond issues maturing in the quarter.",
        "cbonds_issue_offer_q_count": "Number of bond offer / early redemption events in the quarter according to Cbonds issue cards.",
        "cbonds_issue_offer_q_price_avg": "Average offer price across bond offer events in the quarter.",
        "cbonds_issue_outstanding_q_count": "Approximate number of issues outstanding in the quarter based on placement and maturity dates.",
        "cbonds_issue_outstanding_q_volume_rub": "Approximate RUB circulation volume of issues outstanding in the quarter based on placement and maturity dates.",
        "macro_key_rate_avg_q": "Average Bank of Russia key rate within quarter.",
        "macro_key_rate_end_q": "Key rate at the end of quarter.",
        "macro_key_rate_max_q": "Maximum Bank of Russia key rate within quarter.",
        "macro_key_rate_min_q": "Minimum Bank of Russia key rate within quarter.",
        "macro_inflation_yoy_avg_q": "Average year-on-year inflation within quarter.",
        "macro_inflation_target_avg_q": "Average inflation target within quarter.",
        "macro_key_rate_month_end_avg_q": "Average month-end key rate within quarter.",
        "macro_usdrub_avg_q": "Average USD/RUB exchange rate within quarter.",
        "macro_usdrub_end_q": "USD/RUB exchange rate at quarter end.",
        "macro_usdrub_max_q": "Maximum USD/RUB exchange rate within quarter.",
        "macro_usdrub_min_q": "Minimum USD/RUB exchange rate within quarter.",
        "covid_flag": "1 for 2020-2021 quarters.",
        "post_2022_flag": "1 for quarters from 2022Q1 onward.",
    }
    if column in descriptions:
        return descriptions[column]
    if column.startswith("cmd_") and column.endswith("_avg_q"):
        commodity = column.removeprefix("cmd_").removesuffix("_avg_q")
        return f"Quarter-average commodity price for {commodity}."
    return column.replace("_", " ")


def build_variable_dictionary(panel: pd.DataFrame, frequency: str) -> pd.DataFrame:
    rows = []
    for column in panel.columns:
        rows.append(
            {
                "frequency": frequency,
                "column": column,
                "group": variable_group(column),
                "description": variable_description(column),
            }
        )
    return pd.DataFrame(rows)


def write_workbook(path: Path, panel: pd.DataFrame, summary: pd.DataFrame, variable_dictionary: pd.DataFrame) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        panel.to_excel(writer, sheet_name="panel", index=False)
        summary.to_excel(writer, sheet_name="summary", index=False)
        variable_dictionary.to_excel(writer, sheet_name="variable_dictionary", index=False)


def build_model_ready_panels(
    analysis_panel_workbook: Path = DEFAULT_ANALYSIS_PANEL,
    quarterly_output: Path = DEFAULT_QUARTERLY_OUTPUT,
    annual_output: Path = DEFAULT_ANNUAL_OUTPUT,
) -> dict[str, pd.DataFrame]:
    panel = pd.read_excel(analysis_panel_workbook, sheet_name="analysis_panel_quarterly")
    quarterly_panel = add_engineered_columns(select_model_columns(panel).copy())
    annual_panel = quarterly_panel.loc[quarterly_panel["period_type"].eq("Q4")].copy()

    quarterly_summary = build_summary(quarterly_panel, "quarterly")
    annual_summary = build_summary(annual_panel, "annual")
    quarterly_dictionary = build_variable_dictionary(quarterly_panel, "quarterly")
    annual_dictionary = build_variable_dictionary(annual_panel, "annual")

    write_workbook(quarterly_output, quarterly_panel, quarterly_summary, quarterly_dictionary)
    write_workbook(annual_output, annual_panel, annual_summary, annual_dictionary)

    return {
        "quarterly_panel": quarterly_panel,
        "annual_panel": annual_panel,
        "quarterly_summary": quarterly_summary,
        "annual_summary": annual_summary,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build model-ready quarterly and annual panels from analysis_panel.xlsx.")
    parser.add_argument("--analysis-panel", type=Path, default=DEFAULT_ANALYSIS_PANEL)
    parser.add_argument("--quarterly-output", type=Path, default=DEFAULT_QUARTERLY_OUTPUT)
    parser.add_argument("--annual-output", type=Path, default=DEFAULT_ANNUAL_OUTPUT)
    args = parser.parse_args()

    outputs = build_model_ready_panels(
        analysis_panel_workbook=args.analysis_panel,
        quarterly_output=args.quarterly_output,
        annual_output=args.annual_output,
    )
    print(f"Quarterly workbook: {args.quarterly_output}")
    print(f"Annual workbook: {args.annual_output}")
    print(f"Quarterly rows: {len(outputs['quarterly_panel'])}")
    print(f"Annual rows: {len(outputs['annual_panel'])}")


if __name__ == "__main__":
    main()
