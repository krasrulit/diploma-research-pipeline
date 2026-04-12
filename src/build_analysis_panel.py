from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

from .load_shortlist import load_shortlist
from .utils import normalize_identifier


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OIL_GAS_SHORTLIST = Path("/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx")
DEFAULT_METALLURGY_SHORTLIST = Path("/Users/grigorijkrasovickij/4 крус/Диплом/Металлургические_компании_shortlist.xlsx")
DEFAULT_OIL_GAS_SPARK = PROJECT_ROOT / "data_processed" / "spark_neftegaz_report_2014q3_2025q4.xlsx"
DEFAULT_METALLURGY_SPARK = PROJECT_ROOT / "data_processed" / "spark_metallurgy_report_2014q3_2025q4.xlsx"
DEFAULT_PUBLIC_MARKET = PROJECT_ROOT / "data_processed" / "public_market_data_full.xlsx"
DEFAULT_OWNERSHIP = PROJECT_ROOT / "data_processed" / "ownership_state_table.xlsx"
DEFAULT_SECURITIES = PROJECT_ROOT / "data_processed" / "public_securities_all_companies_optimized_lite.xlsx"
DEFAULT_SPARK_COMBINED = PROJECT_ROOT / "data_processed" / "spark_sector_combined_2014q3_2025q4.xlsx"
DEFAULT_ANALYSIS_OUTPUT = PROJECT_ROOT / "data_processed" / "analysis_panel.xlsx"

SPARK_SHEETS = ["input_inventory", "panel_core", "panel_quarterly", "coverage", "parse_log"]
SHORTLIST_META_COLUMNS = [
    "company_name",
    "company_name_short",
    "company_name_full",
    "company_name_en",
    "company_name_core",
    "inn",
    "ogrn",
    "ticker",
    "isin",
    "spark_id",
    "industry",
    "sample_flag",
    "sample_membership",
    "selection_result",
    "source_sheet",
    "records_merged",
]


def quarter_num(period_type: object) -> object:
    text = str(period_type).strip().upper()
    return {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}.get(text, np.nan)


def quarter_end_date(quarter_label: object) -> pd.Timestamp | pd.NaT:
    text = str(quarter_label).strip().upper()
    match = pd.Series([text]).str.extract(r"(?P<year>\d{4})Q(?P<quarter>[1-4])").iloc[0]
    if match.isna().any():
        return pd.NaT
    year = int(match["year"])
    quarter = int(match["quarter"])
    month = quarter * 3
    return pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)


def first_notna(values: Iterable[object]) -> object:
    for value in values:
        if pd.notna(value):
            return value
    return np.nan


def join_unique(values: pd.Series) -> str:
    seen: set[str] = set()
    result: list[str] = []
    for value in values.dropna().astype(str):
        text = value.strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return " | ".join(result)


def add_quarter_label_from_date(frame: pd.DataFrame, date_column: str) -> pd.DataFrame:
    out = frame.copy()
    out[date_column] = pd.to_datetime(out[date_column], errors="coerce")
    out = out.loc[out[date_column].notna()].copy()
    out["report_year"] = out[date_column].dt.year.astype(int)
    out["quarter_num"] = ((out[date_column].dt.month - 1) // 3 + 1).astype(int)
    out["period_type"] = out["quarter_num"].map(lambda value: f"Q{value}")
    out["quarter_label"] = out["report_year"].astype(str) + out["period_type"]
    out["quarter_end_date"] = out["quarter_label"].map(quarter_end_date)
    return out


def add_sector_to_shortlist(shortlist_path: Path, sector: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    companies_master, log = load_shortlist(shortlist_path)
    companies_master = companies_master.copy()
    log = log.copy()
    companies_master["sector"] = sector
    companies_master["sample_main_flag"] = companies_master["sample_flag"].eq("main")
    companies_master["sample_extended_flag"] = companies_master["sample_flag"].eq("extended")
    companies_master["sector_company_id"] = companies_master.apply(
        lambda row: f"{sector}:{row['inn']}" if str(row["inn"]).strip() else f"{sector}:{row['company_name']}",
        axis=1,
    )
    log["sector"] = sector
    log["shortlist_path"] = str(shortlist_path)
    return companies_master, log


def combine_shortlists(oil_gas_shortlist: Path, metallurgy_shortlist: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    oil_gas_master, oil_gas_log = add_sector_to_shortlist(oil_gas_shortlist, "oil_gas")
    metallurgy_master, metallurgy_log = add_sector_to_shortlist(metallurgy_shortlist, "metallurgy")
    companies_master = pd.concat([oil_gas_master, metallurgy_master], ignore_index=True, sort=False)
    shortlist_log = pd.concat([oil_gas_log, metallurgy_log], ignore_index=True, sort=False)
    return companies_master, shortlist_log


def read_spark_workbook(path: Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(path, sheet_name=None)
    return {sheet_name: sheets.get(sheet_name, pd.DataFrame()) for sheet_name in SPARK_SHEETS}


def enrich_with_shortlist(frame: pd.DataFrame, companies_master: pd.DataFrame, sector: str) -> pd.DataFrame:
    if frame.empty:
        return frame

    sector_master = companies_master.loc[companies_master["sector"].eq(sector), SHORTLIST_META_COLUMNS + ["sector", "sample_main_flag", "sample_extended_flag", "sector_company_id"]].copy()
    out = frame.copy()
    out["sector"] = sector
    out["inn"] = out["inn"].map(lambda value: normalize_identifier(value, length=10))
    sector_master["inn"] = sector_master["inn"].map(lambda value: normalize_identifier(value, length=10))

    merged = out.merge(
        sector_master.rename(columns={"company_name": "company_name_shortlist"}),
        on=["inn", "sector"],
        how="left",
    )
    merged["shortlist_match_flag"] = merged["company_name_shortlist"].notna()
    return merged


def combine_spark_workbooks(
    oil_gas_workbook: Path,
    metallurgy_workbook: Path,
    companies_master: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    oil_gas = read_spark_workbook(oil_gas_workbook)
    metallurgy = read_spark_workbook(metallurgy_workbook)

    combined: dict[str, pd.DataFrame] = {}
    for sheet_name in SPARK_SHEETS:
        oil_frame = enrich_with_shortlist(oil_gas.get(sheet_name, pd.DataFrame()), companies_master, "oil_gas")
        metallurgy_frame = enrich_with_shortlist(metallurgy.get(sheet_name, pd.DataFrame()), companies_master, "metallurgy")
        combined[sheet_name] = pd.concat([oil_frame, metallurgy_frame], ignore_index=True, sort=False)

    return combined


def build_shortlist_coverage(companies_master: pd.DataFrame, spark_input_inventory: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    inventory = spark_input_inventory.copy()
    inventory = inventory.loc[inventory["source_ext"].eq(".docx")].copy()
    inventory["inn"] = inventory["inn"].map(lambda value: normalize_identifier(value, length=10))

    matched_reports = (
        inventory.groupby(["sector", "inn"], dropna=False)
        .agg(
            reports_found=("source_file", "size"),
            matched_report_files=("source_file", join_unique),
            matched_report_paths=("source_path", join_unique),
            matched_report_companies=("company", join_unique),
        )
        .reset_index()
    )

    detailed = companies_master.merge(matched_reports, on=["sector", "inn"], how="left")
    detailed["reports_found"] = detailed["reports_found"].fillna(0).astype(int)
    detailed["reports_found_flag"] = detailed["reports_found"].gt(0)
    for column in ["matched_report_files", "matched_report_paths", "matched_report_companies"]:
        detailed[column] = detailed[column].fillna("")

    summary_rows = []
    for sector_value in ["oil_gas", "metallurgy"]:
        sector_slice = detailed.loc[detailed["sector"].eq(sector_value)]
        summary_rows.extend(
            [
                {"sector": sector_value, "metric": "shortlist_total", "value": len(sector_slice)},
                {"sector": sector_value, "metric": "main_total", "value": int(sector_slice["sample_flag"].eq("main").sum())},
                {"sector": sector_value, "metric": "extended_total", "value": int(sector_slice["sample_flag"].eq("extended").sum())},
                {"sector": sector_value, "metric": "covered_total", "value": int(sector_slice["reports_found_flag"].sum())},
                {"sector": sector_value, "metric": "missing_total", "value": int((~sector_slice["reports_found_flag"]).sum())},
            ]
        )
    summary_rows.extend(
        [
            {"sector": "all", "metric": "shortlist_total", "value": len(detailed)},
            {"sector": "all", "metric": "covered_total", "value": int(detailed["reports_found_flag"].sum())},
            {"sector": "all", "metric": "missing_total", "value": int((~detailed["reports_found_flag"]).sum())},
        ]
    )
    summary = pd.DataFrame(summary_rows)
    return detailed, summary


def build_macro_quarterly(public_market_workbook: Path) -> pd.DataFrame:
    macro = pd.read_excel(public_market_workbook, sheet_name="macro_cbr")
    commodities = pd.read_excel(public_market_workbook, sheet_name="commodity_prices_monthly")

    macro["obs_date"] = pd.to_datetime(macro["date"], errors="coerce")
    if "period" in macro.columns:
        period_dates = pd.to_datetime(macro["period"].astype(str) + "-01", errors="coerce")
        macro["obs_date"] = macro["obs_date"].fillna(period_dates)
    macro = add_quarter_label_from_date(macro, "obs_date")

    key_rate = macro.loc[macro["series"].eq("key_rate")].copy()
    key_rate_quarterly = (
        key_rate.sort_values("obs_date")
        .groupby("quarter_label", dropna=False)
        .agg(
            macro_key_rate_avg_q=("value", "mean"),
            macro_key_rate_end_q=("value", "last"),
            macro_key_rate_max_q=("value", "max"),
            macro_key_rate_min_q=("value", "min"),
            report_year=("report_year", "first"),
            period_type=("period_type", "first"),
            quarter_end_date=("quarter_end_date", "first"),
        )
        .reset_index()
    )

    usdrub = macro.loc[macro["series"].eq("usdrub")].copy()
    usdrub_quarterly = (
        usdrub.sort_values("obs_date")
        .groupby("quarter_label", dropna=False)
        .agg(
            macro_usdrub_avg_q=("value", "mean"),
            macro_usdrub_end_q=("value", "last"),
            macro_usdrub_max_q=("value", "max"),
            macro_usdrub_min_q=("value", "min"),
        )
        .reset_index()
    )

    monthly = macro.loc[macro["series"].isin(["inflation_yoy", "inflation_target", "key_rate_month_end"])].copy()
    monthly_pivot = (
        monthly.pivot_table(
            index="quarter_label",
            columns="series",
            values="value",
            aggfunc="mean",
        )
        .reset_index()
        .rename(
            columns={
                "inflation_yoy": "macro_inflation_yoy_avg_q",
                "inflation_target": "macro_inflation_target_avg_q",
                "key_rate_month_end": "macro_key_rate_month_end_avg_q",
            }
        )
    )

    commodities = add_quarter_label_from_date(commodities, "date")
    commodity_quarterly = (
        commodities.pivot_table(
            index="quarter_label",
            columns="commodity_code",
            values="value",
            aggfunc="mean",
        )
        .reset_index()
    )
    commodity_quarterly = commodity_quarterly.rename(
        columns={column: f"cmd_{column}_avg_q" for column in commodity_quarterly.columns if column != "quarter_label"}
    )

    macro_quarterly = (
        key_rate_quarterly
        .merge(monthly_pivot, on="quarter_label", how="left")
        .merge(usdrub_quarterly, on="quarter_label", how="left")
        .merge(commodity_quarterly, on="quarter_label", how="left")
    )
    macro_quarterly["quarter_num"] = macro_quarterly["period_type"].map(quarter_num).astype("Int64")
    macro_quarterly["covid_flag"] = (
        (macro_quarterly["report_year"].between(2020, 2021))
    ).astype(int)
    macro_quarterly["post_2022_flag"] = (
        (macro_quarterly["report_year"] > 2022)
        | ((macro_quarterly["report_year"] == 2022) & (macro_quarterly["quarter_num"] >= 1))
    ).astype(int)

    return macro_quarterly.sort_values(["report_year", "quarter_num"], kind="stable").reset_index(drop=True)


def load_public_market_flags(public_market_workbook: Path) -> pd.DataFrame:
    flags = pd.read_excel(public_market_workbook, sheet_name="firm_market_flags")
    flags = flags.copy()
    flags["inn"] = flags["inn"].map(lambda value: normalize_identifier(value, length=10))
    flag_columns = [
        "inn",
        "n_validated_instruments",
        "n_validated_bonds",
        "n_validated_shares",
        "has_validated_public_instrument",
        "has_validated_public_bond",
        "has_validated_public_share",
        "n_moex_selected_rows",
        "has_moex_selected",
        "n_tinvest_selected_rows",
        "has_tinvest_selected",
        "n_moex_bonds_rows",
        "has_moex_bonds",
        "n_tinvest_bonds_rows",
        "has_tinvest_bonds",
    ]
    out = flags[flag_columns].copy()
    rename_map = {column: f"market_{column}" for column in out.columns if column != "inn"}
    out = out.rename(columns=rename_map)
    out["public_market_enriched_flag"] = True
    return out


def load_clean_market_access(securities_workbook: Path) -> pd.DataFrame:
    access = pd.read_excel(securities_workbook, sheet_name="company_market_access_clean")
    access = access.copy()
    access["company_id"] = access["company_id"].astype(str).str.strip()
    selected_columns = [
        "company_id",
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
    selected_columns = [column for column in selected_columns if column in access.columns]
    out = access[selected_columns].copy()
    for date_column in ["market_history_start_min", "market_history_end_max"]:
        if date_column in out.columns:
            out[date_column] = pd.to_datetime(out[date_column], errors="coerce")
    return out


def load_ownership_state(ownership_workbook: Path) -> pd.DataFrame:
    ownership = pd.read_excel(ownership_workbook, sheet_name="ownership_state_table")
    ownership = ownership.copy()
    ownership["sector"] = ownership["sector"].astype(str).str.strip()
    ownership["inn"] = ownership["inn"].map(lambda value: normalize_identifier(value, length=10))
    ownership["company_id"] = ownership.apply(
        lambda row: f"{row['sector']}:{row['inn']}" if str(row["inn"]).strip() else f"{row['sector']}:{row.get('company_name', '')}",
        axis=1,
    )
    selected_columns = [
        "company_id",
        "report_found_flag",
        "source_file",
        "source_path",
        "company_from_file",
        "full_company_name_report",
        "web_domain",
        "head_company_name",
        "n_subsidiaries",
        "has_head_company_field_flag",
        "has_subsidiaries_flag",
        "is_subsidiary_flag",
        "is_parent_flag",
        "ownership_role",
        "group_name_inferred",
        "state_owned_flag",
        "state_bucket",
        "group_inference_source",
        "group_inference_confidence",
        "ownership_review_needed_flag",
        "head_company_snippet",
        "text_length",
        "parse_status",
    ]
    selected_columns = [column for column in selected_columns if column in ownership.columns]
    out = ownership[selected_columns].copy()
    rename_map = {
        column: f"ownership_{column}" for column in out.columns if column not in {"company_id"}
    }
    return out.rename(columns=rename_map)


def with_min_count_sum(frame: pd.DataFrame, columns: list[str]) -> pd.Series:
    return frame[columns].sum(axis=1, min_count=1)


def build_analysis_panel_quarterly(
    spark_panel_quarterly: pd.DataFrame,
    companies_master: pd.DataFrame,
    public_market_workbook: Path,
    ownership_workbook: Path,
    securities_workbook: Path,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    macro_quarterly = build_macro_quarterly(public_market_workbook)
    public_market_flags = load_public_market_flags(public_market_workbook)
    clean_market_access = load_clean_market_access(securities_workbook)
    ownership_state = load_ownership_state(ownership_workbook)
    macro_for_merge = macro_quarterly.drop(
        columns=["report_year", "period_type", "quarter_num", "quarter_end_date"],
        errors="ignore",
    )

    panel = spark_panel_quarterly.copy()
    panel["quarter_num"] = panel["period_type"].map(quarter_num).astype("Int64")
    panel["quarter_end_date"] = panel["quarter_label"].map(quarter_end_date)
    panel["company_id"] = panel.apply(
        lambda row: f"{row['sector']}:{row['inn']}" if str(row["inn"]).strip() else f"{row['sector']}:{row['company_name_shortlist']}",
        axis=1,
    )
    panel["sample_main_flag"] = panel["sample_flag"].eq("main")
    panel["sample_extended_flag"] = panel["sample_flag"].eq("extended")
    panel["reporting_observed_flag"] = (
        panel[["has_balance", "has_pnl", "has_cash_flow"]].fillna(0).astype(int).sum(axis=1).gt(0)
    )
    panel["available_core_metrics_count"] = panel[
        ["assets_total", "ppe", "cash", "equity", "debt_lt", "debt_st", "revenue", "interest_expense", "net_income", "ebit"]
    ].notna().sum(axis=1)

    panel = (
        panel.merge(macro_for_merge, on="quarter_label", how="left")
        .merge(public_market_flags, on="inn", how="left")
        .merge(clean_market_access, on="company_id", how="left")
        .merge(ownership_state, on="company_id", how="left")
    )
    panel["public_market_enriched_flag"] = panel["public_market_enriched_flag"].fillna(False)

    panel["total_debt"] = with_min_count_sum(panel, ["debt_lt", "debt_st"])
    panel["debt_to_assets"] = panel["total_debt"] / panel["assets_total"]
    panel["debt_lt_share"] = panel["debt_lt"] / panel["total_debt"]
    panel["cash_to_assets"] = panel["cash"] / panel["assets_total"]
    panel["ppe_to_assets"] = panel["ppe"] / panel["assets_total"]
    panel["equity_to_assets"] = panel["equity"] / panel["assets_total"]
    panel["net_margin"] = panel["net_income"] / panel["revenue"]
    panel["ebit_margin"] = panel["ebit"] / panel["revenue"]
    panel["log_assets"] = np.where(panel["assets_total"] > 0, np.log(panel["assets_total"]), np.nan)
    panel["log_revenue"] = np.where(panel["revenue"] > 0, np.log(panel["revenue"]), np.nan)

    sort_columns = ["sector", "sample_flag", "company_name_shortlist", "report_year", "quarter_num"]
    panel = panel.sort_values(sort_columns, kind="stable").reset_index(drop=True)
    return panel, macro_quarterly


def build_analysis_summary(
    companies_master: pd.DataFrame,
    spark_panel_quarterly: pd.DataFrame,
    analysis_panel_quarterly: pd.DataFrame,
    shortlist_coverage: pd.DataFrame,
) -> pd.DataFrame:
    rows = [
        {"metric": "companies_total", "value": len(companies_master)},
        {"metric": "companies_oil_gas", "value": int(companies_master["sector"].eq("oil_gas").sum())},
        {"metric": "companies_metallurgy", "value": int(companies_master["sector"].eq("metallurgy").sum())},
        {"metric": "main_total", "value": int(companies_master["sample_flag"].eq("main").sum())},
        {"metric": "extended_total", "value": int(companies_master["sample_flag"].eq("extended").sum())},
        {"metric": "shortlist_missing_total", "value": int((~shortlist_coverage["reports_found_flag"]).sum())},
        {"metric": "spark_panel_quarterly_rows", "value": len(spark_panel_quarterly)},
        {"metric": "analysis_panel_quarterly_rows", "value": len(analysis_panel_quarterly)},
        {"metric": "analysis_unique_company_ids", "value": int(analysis_panel_quarterly["company_id"].nunique())},
        {"metric": "analysis_rows_with_reporting_data", "value": int(analysis_panel_quarterly["reporting_observed_flag"].sum())},
        {"metric": "analysis_rows_with_usdrub", "value": int(analysis_panel_quarterly["macro_usdrub_avg_q"].notna().sum())},
        {"metric": "analysis_companies_with_ownership_state", "value": int(analysis_panel_quarterly.loc[analysis_panel_quarterly["ownership_state_bucket"].notna(), "company_id"].nunique())},
        {"metric": "analysis_companies_with_clean_market_access", "value": int(analysis_panel_quarterly.loc[analysis_panel_quarterly["market_access_status"].notna(), "company_id"].nunique())},
    ]
    return pd.DataFrame(rows)


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def build_analysis_outputs(
    oil_gas_shortlist: Path = DEFAULT_OIL_GAS_SHORTLIST,
    metallurgy_shortlist: Path = DEFAULT_METALLURGY_SHORTLIST,
    oil_gas_spark: Path = DEFAULT_OIL_GAS_SPARK,
    metallurgy_spark: Path = DEFAULT_METALLURGY_SPARK,
    public_market_workbook: Path = DEFAULT_PUBLIC_MARKET,
    ownership_workbook: Path = DEFAULT_OWNERSHIP,
    securities_workbook: Path = DEFAULT_SECURITIES,
    spark_combined_output: Path = DEFAULT_SPARK_COMBINED,
    analysis_output: Path = DEFAULT_ANALYSIS_OUTPUT,
) -> dict[str, pd.DataFrame]:
    companies_master, shortlist_log = combine_shortlists(oil_gas_shortlist, metallurgy_shortlist)
    spark_outputs = combine_spark_workbooks(oil_gas_spark, metallurgy_spark, companies_master)
    shortlist_coverage, shortlist_coverage_summary = build_shortlist_coverage(companies_master, spark_outputs["input_inventory"])
    analysis_panel_quarterly, macro_quarterly = build_analysis_panel_quarterly(
        spark_panel_quarterly=spark_outputs["panel_quarterly"],
        companies_master=companies_master,
        public_market_workbook=public_market_workbook,
        ownership_workbook=ownership_workbook,
        securities_workbook=securities_workbook,
    )
    analysis_summary = build_analysis_summary(
        companies_master=companies_master,
        spark_panel_quarterly=spark_outputs["panel_quarterly"],
        analysis_panel_quarterly=analysis_panel_quarterly,
        shortlist_coverage=shortlist_coverage,
    )

    spark_combined_sheets = {
        "companies_master": companies_master,
        "input_inventory": spark_outputs["input_inventory"],
        "panel_core": spark_outputs["panel_core"],
        "panel_quarterly": spark_outputs["panel_quarterly"],
        "coverage": spark_outputs["coverage"],
        "parse_log": spark_outputs["parse_log"],
        "shortlist_coverage": shortlist_coverage,
        "shortlist_coverage_summary": shortlist_coverage_summary,
        "shortlist_log": shortlist_log,
    }
    write_excel(spark_combined_output, spark_combined_sheets)

    analysis_sheets = {
        "analysis_summary": analysis_summary,
        "companies_master": companies_master,
        "analysis_panel_quarterly": analysis_panel_quarterly,
        "ownership_state_table": load_ownership_state(ownership_workbook),
        "company_market_access_clean": load_clean_market_access(securities_workbook),
        "spark_panel_quarterly": spark_outputs["panel_quarterly"],
        "spark_panel_core": spark_outputs["panel_core"],
        "spark_input_inventory": spark_outputs["input_inventory"],
        "spark_coverage": spark_outputs["coverage"],
        "macro_quarterly": macro_quarterly,
        "shortlist_coverage": shortlist_coverage,
        "shortlist_coverage_summary": shortlist_coverage_summary,
        "parse_log": spark_outputs["parse_log"],
    }
    write_excel(analysis_output, analysis_sheets)

    return {
        "companies_master": companies_master,
        "spark_input_inventory": spark_outputs["input_inventory"],
        "spark_panel_core": spark_outputs["panel_core"],
        "spark_panel_quarterly": spark_outputs["panel_quarterly"],
        "spark_coverage": spark_outputs["coverage"],
        "parse_log": spark_outputs["parse_log"],
        "shortlist_coverage": shortlist_coverage,
        "shortlist_coverage_summary": shortlist_coverage_summary,
        "macro_quarterly": macro_quarterly,
        "analysis_panel_quarterly": analysis_panel_quarterly,
        "analysis_summary": analysis_summary,
        "ownership_state_table": load_ownership_state(ownership_workbook),
        "company_market_access_clean": load_clean_market_access(securities_workbook),
        "shortlist_log": shortlist_log,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build combined SPARK sector panel and final analysis panel.")
    parser.add_argument("--oil-gas-shortlist", type=Path, default=DEFAULT_OIL_GAS_SHORTLIST)
    parser.add_argument("--metallurgy-shortlist", type=Path, default=DEFAULT_METALLURGY_SHORTLIST)
    parser.add_argument("--oil-gas-spark", type=Path, default=DEFAULT_OIL_GAS_SPARK)
    parser.add_argument("--metallurgy-spark", type=Path, default=DEFAULT_METALLURGY_SPARK)
    parser.add_argument("--public-market", type=Path, default=DEFAULT_PUBLIC_MARKET)
    parser.add_argument("--ownership", type=Path, default=DEFAULT_OWNERSHIP)
    parser.add_argument("--securities", type=Path, default=DEFAULT_SECURITIES)
    parser.add_argument("--spark-output", type=Path, default=DEFAULT_SPARK_COMBINED)
    parser.add_argument("--analysis-output", type=Path, default=DEFAULT_ANALYSIS_OUTPUT)
    args = parser.parse_args()

    outputs = build_analysis_outputs(
        oil_gas_shortlist=args.oil_gas_shortlist,
        metallurgy_shortlist=args.metallurgy_shortlist,
        oil_gas_spark=args.oil_gas_spark,
        metallurgy_spark=args.metallurgy_spark,
        public_market_workbook=args.public_market,
        ownership_workbook=args.ownership,
        securities_workbook=args.securities,
        spark_combined_output=args.spark_output,
        analysis_output=args.analysis_output,
    )

    print(f"Combined SPARK workbook: {args.spark_output}")
    print(f"Analysis workbook: {args.analysis_output}")
    print(f"Companies master: {len(outputs['companies_master'])}")
    print(f"SPARK quarterly rows: {len(outputs['spark_panel_quarterly'])}")
    print(f"Analysis quarterly rows: {len(outputs['analysis_panel_quarterly'])}")


if __name__ == "__main__":
    main()
