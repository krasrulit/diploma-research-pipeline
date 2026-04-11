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

MACRO_COLUMNS = [
    "macro_key_rate_avg_q",
    "macro_key_rate_end_q",
    "macro_key_rate_max_q",
    "macro_key_rate_min_q",
    "macro_inflation_yoy_avg_q",
    "macro_inflation_target_avg_q",
    "macro_key_rate_month_end_avg_q",
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


def select_model_columns(frame: pd.DataFrame) -> pd.DataFrame:
    columns = (
        BASE_IDENTIFIER_COLUMNS
        + TRACEABILITY_COLUMNS
        + AVAILABILITY_COLUMNS
        + FINANCIAL_LEVEL_COLUMNS
        + RATIO_COLUMNS
        + MARKET_FLAG_COLUMNS
        + MACRO_COLUMNS
        + commodity_columns(frame)
    )
    out = ensure_columns(frame, columns)
    return out[columns]


def build_summary(panel: pd.DataFrame, frequency: str) -> pd.DataFrame:
    rows = [
        {"frequency": frequency, "metric": "n_rows", "value": len(panel)},
        {"frequency": frequency, "metric": "n_unique_companies", "value": int(panel["company_id"].nunique())},
        {"frequency": frequency, "metric": "n_oil_gas_companies", "value": int(panel.loc[panel["sector"].eq("oil_gas"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_metallurgy_companies", "value": int(panel.loc[panel["sector"].eq("metallurgy"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_main_companies", "value": int(panel.loc[panel["sample_flag"].eq("main"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "n_extended_companies", "value": int(panel.loc[panel["sample_flag"].eq("extended"), "company_id"].nunique())},
        {"frequency": frequency, "metric": "rows_with_reporting_data", "value": int(panel["reporting_observed_flag"].fillna(False).sum())},
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
        "log_assets": "Natural log of total assets.",
        "log_revenue": "Natural log of revenue.",
        "public_market_enriched_flag": "1 if public-market enrichment from the oil-gas pipeline is available.",
        "macro_key_rate_avg_q": "Average Bank of Russia key rate within quarter.",
        "macro_key_rate_end_q": "Key rate at the end of quarter.",
        "macro_inflation_yoy_avg_q": "Average year-on-year inflation within quarter.",
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
    quarterly_panel = select_model_columns(panel).copy()
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
