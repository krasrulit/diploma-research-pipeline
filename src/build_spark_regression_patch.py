from __future__ import annotations

import argparse
import re
from pathlib import Path

import numpy as np
import pandas as pd

from .build_model_ready_panels import (
    FINANCIAL_LEVEL_COLUMNS,
    RATIO_COLUMNS,
    build_summary,
    winsorize_series,
)
from .build_spark_recovery_audit import DEFAULT_OUTPUT as DEFAULT_RECOVERY_AUDIT


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_REGRESSION_READY = PROJECT_ROOT / "data_processed" / "regression_ready_final.xlsx"
DEFAULT_PATCH_OUTPUT = PROJECT_ROOT / "data_processed" / "spark_regression_ready_patch.xlsx"

DIRECT_SPARK_COLUMNS = [
    "assets_total",
    "ppe",
    "cash",
    "equity",
    "debt_lt",
    "debt_st",
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

CORE_COUNT_COLUMNS = [
    "assets_total",
    "ppe",
    "cash",
    "equity",
    "debt_lt",
    "debt_st",
    "revenue",
    "interest_expense",
    "net_income",
    "ebit",
]

PATCH_COMPARE_COLUMNS = (
    ["reporting_observed_flag", "available_core_metrics_count"]
    + FINANCIAL_LEVEL_COLUMNS
    + RATIO_COLUMNS
)


def period_to_quarter(period_type: object) -> str:
    text = str(period_type).strip().upper()
    if text == "FY":
        return "Q4"
    return text


def normalize_identifier(value: object, length: int | None = None) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if re.fullmatch(r"\d+(\.0+)?", text):
        text = text.split(".", 1)[0]
    digits = re.sub(r"\D+", "", text)
    if length and digits:
        digits = digits.zfill(length)
    return digits


def normalize_patch_keys(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    out["inn_key"] = out["inn"].map(lambda value: normalize_identifier(value, length=10))
    out["report_year_key"] = pd.to_numeric(out["report_year"], errors="coerce").astype("Int64")
    out["period_type_key"] = out["period_type"].map(period_to_quarter)
    out["quarter_label_key"] = out["report_year_key"].astype(str) + out["period_type_key"]
    out.loc[out["report_year_key"].isna(), "quarter_label_key"] = ""
    return out


def read_recovery_wide(recovery_audit: Path) -> pd.DataFrame:
    recovery = pd.read_excel(recovery_audit, sheet_name="recovery_needed_wide")
    if recovery.empty:
        return pd.DataFrame()

    recovery = normalize_patch_keys(recovery)
    value_columns = [column for column in DIRECT_SPARK_COLUMNS if column in recovery.columns]
    recovery[value_columns] = recovery[value_columns].apply(pd.to_numeric, errors="coerce")
    recovery = recovery.dropna(subset=value_columns, how="all")
    if recovery.empty:
        return recovery

    key_columns = ["inn_key", "report_year_key", "period_type_key", "quarter_label_key"]
    meta_columns = ["company", "source_file", "report_period", "data_source", "unit"]
    aggregations: dict[str, object] = {column: "first" for column in value_columns}
    for column in meta_columns:
        if column in recovery.columns:
            aggregations[column] = lambda values: " | ".join(
                dict.fromkeys(str(value).strip() for value in values.dropna() if str(value).strip())
            )
    grouped = recovery[key_columns + meta_columns + value_columns].groupby(
        key_columns,
        dropna=False,
        as_index=False,
    ).agg(aggregations)
    return grouped


def with_min_count_sum(frame: pd.DataFrame, columns: list[str]) -> pd.Series:
    return frame[columns].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    num = pd.to_numeric(numerator, errors="coerce")
    den = pd.to_numeric(denominator, errors="coerce")
    return num / den.replace({0: np.nan})


def recompute_financial_derivatives(panel: pd.DataFrame) -> pd.DataFrame:
    out = panel.copy()
    for column in DIRECT_SPARK_COLUMNS + ["total_debt"]:
        if column in out.columns:
            out[column] = pd.to_numeric(out[column], errors="coerce")

    out["available_core_metrics_count"] = out[
        [column for column in CORE_COUNT_COLUMNS if column in out.columns]
    ].notna().sum(axis=1)
    if {"has_balance", "has_pnl", "has_cash_flow"}.issubset(out.columns):
        observed_from_forms = out[["has_balance", "has_pnl", "has_cash_flow"]].fillna(0).astype(int).sum(axis=1).gt(0)
    else:
        observed_from_forms = pd.Series(False, index=out.index)
    observed_from_values = out["available_core_metrics_count"].gt(0)
    out["reporting_observed_flag"] = observed_from_forms | observed_from_values

    out["total_debt"] = with_min_count_sum(out, ["debt_lt", "debt_st"])
    out["debt_to_assets"] = safe_divide(out["total_debt"], out["assets_total"])
    out["debt_lt_share"] = safe_divide(out["debt_lt"], out["total_debt"])
    out["cash_to_assets"] = safe_divide(out["cash"], out["assets_total"])
    out["ppe_to_assets"] = safe_divide(out["ppe"], out["assets_total"])
    out["equity_to_assets"] = safe_divide(out["equity"], out["assets_total"])
    out["net_margin"] = safe_divide(out["net_income"], out["revenue"])
    out["ebit_margin"] = safe_divide(out["ebit"], out["revenue"])
    out["net_debt"] = out["total_debt"] - out["cash"]
    out["net_debt_to_assets"] = safe_divide(out["net_debt"], out["assets_total"])
    out["cfo_to_assets"] = safe_divide(out["cfo"], out["assets_total"])
    out["capex_to_assets"] = safe_divide(out["capex"], out["assets_total"])
    out["capex_to_cfo"] = safe_divide(out["capex"], out["cfo"])
    out["interest_coverage"] = safe_divide(out["ebit"], out["interest_expense"])
    out["net_margin_winsor_1_99"] = winsorize_series(out["net_margin"])
    out["ebit_margin_winsor_1_99"] = winsorize_series(out["ebit_margin"])
    out["log_assets"] = np.where(out["assets_total"] > 0, np.log(out["assets_total"]), np.nan)
    out["log_revenue"] = np.where(out["revenue"] > 0, np.log(out["revenue"]), np.nan)
    return out


def value_changed(before: pd.Series, after: pd.Series) -> pd.Series:
    before_num = pd.to_numeric(before, errors="coerce")
    after_num = pd.to_numeric(after, errors="coerce")
    numeric_equal = np.isclose(before_num, after_num, equal_nan=True)
    both_numeric_or_na = before_num.notna() | after_num.notna() | (before.isna() & after.isna())

    before_text = before.fillna("").astype(str)
    after_text = after.fillna("").astype(str)
    text_changed = before_text.ne(after_text)
    return pd.Series(np.where(both_numeric_or_na, ~numeric_equal, text_changed), index=before.index)


def apply_recovery_patch(panel: pd.DataFrame, recovery_wide: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if recovery_wide.empty:
        return panel.copy(), pd.DataFrame()

    original = panel.copy()
    out = normalize_patch_keys(panel)
    recovery_columns = [column for column in DIRECT_SPARK_COLUMNS if column in recovery_wide.columns]
    merge_columns = ["inn_key", "report_year_key", "period_type_key", "quarter_label_key"]
    merged = out.merge(
        recovery_wide[merge_columns + recovery_columns],
        on=merge_columns,
        how="left",
        suffixes=("", "_recovered"),
    )

    filled_columns: list[str] = []
    for column in recovery_columns:
        recovered_column = f"{column}_recovered"
        if recovered_column not in merged.columns:
            continue
        old_value = pd.to_numeric(merged[column], errors="coerce") if column in merged.columns else pd.Series(np.nan, index=merged.index)
        recovered_value = pd.to_numeric(merged[recovered_column], errors="coerce")
        fill_mask = old_value.isna() & recovered_value.notna()
        if fill_mask.any():
            merged.loc[fill_mask, column] = recovered_value.loc[fill_mask]
            filled_columns.append(column)

    patched = recompute_financial_derivatives(
        merged.drop(columns=[column for column in merged.columns if column.endswith("_recovered")], errors="ignore")
    )
    patched = patched[original.columns]

    compare_columns = [column for column in PATCH_COMPARE_COLUMNS if column in original.columns and column in patched.columns]
    any_change = pd.Series(False, index=original.index)
    for column in compare_columns:
        any_change = any_change | value_changed(original[column], patched[column])

    changed_rows = patched.loc[any_change].copy()
    direct_patch_rows = out.loc[any_change, ["company_id", "quarter_label"]].copy()
    direct_patch_rows["direct_recovered_columns"] = ""
    if filled_columns:
        direct_changed = []
        for idx in original.index[any_change]:
            columns = [
                column
                for column in filled_columns
                if pd.isna(original.at[idx, column]) and pd.notna(patched.at[idx, column])
            ]
            direct_changed.append(", ".join(columns))
        direct_patch_rows["direct_recovered_columns"] = direct_changed
    return changed_rows, direct_patch_rows


def build_changes_long(
    original: pd.DataFrame,
    patched_rows: pd.DataFrame,
    evidence: pd.DataFrame,
    frequency: str,
) -> pd.DataFrame:
    if patched_rows.empty:
        return pd.DataFrame()

    original_keyed = original.set_index(["company_id", "quarter_label"], drop=False)
    changed_keyed = patched_rows.set_index(["company_id", "quarter_label"], drop=False)
    compare_columns = [column for column in PATCH_COMPARE_COLUMNS if column in original.columns and column in patched_rows.columns]
    rows: list[dict[str, object]] = []
    for key, patched_row in changed_keyed.iterrows():
        original_row = original_keyed.loc[key]
        for column in compare_columns:
            before = original_row[column]
            after = patched_row[column]
            if not bool(value_changed(pd.Series([before]), pd.Series([after])).iloc[0]):
                continue
            rows.append(
                {
                    "frequency": frequency,
                    "company_id": patched_row.get("company_id"),
                    "sector": patched_row.get("sector"),
                    "company": patched_row.get("company"),
                    "company_name_shortlist": patched_row.get("company_name_shortlist"),
                    "inn": normalize_identifier(patched_row.get("inn"), length=10),
                    "report_year": patched_row.get("report_year"),
                    "period_type": patched_row.get("period_type"),
                    "quarter_label": patched_row.get("quarter_label"),
                    "column": column,
                    "old_value": before,
                    "new_value": after,
                    "change_type": "direct_recovered" if column in DIRECT_SPARK_COLUMNS else "derived_recomputed",
                }
            )
    changes = pd.DataFrame(rows)
    if changes.empty or evidence.empty:
        return changes

    evidence = evidence.copy()
    evidence["inn"] = evidence["inn"].map(lambda value: normalize_identifier(value, length=10))
    evidence["period_type"] = evidence["period_type"].map(period_to_quarter)
    evidence = evidence.rename(columns={"metric_std": "column", "recovered_value": "evidence_recovered_value"})
    evidence_columns = [
        "inn",
        "report_year",
        "period_type",
        "column",
        "source_file",
        "table_title",
        "statement_section",
        "metric_name",
        "metric_code",
        "value_raw",
        "value_num",
        "metric_match_method",
        "metric_match_confidence",
        "source_workbook",
    ]
    evidence = evidence[[column for column in evidence_columns if column in evidence.columns]].drop_duplicates()
    return changes.merge(
        evidence,
        on=["inn", "report_year", "period_type", "column"],
        how="left",
    )


def build_patch_summary(
    quarterly_patch: pd.DataFrame,
    annual_patch: pd.DataFrame,
    quarterly_changes: pd.DataFrame,
    annual_changes: pd.DataFrame,
    regression_ready: Path,
    recovery_audit: Path,
) -> pd.DataFrame:
    rows = [
        {"metric": "source_regression_ready", "value": str(regression_ready)},
        {"metric": "source_recovery_audit", "value": str(recovery_audit)},
        {"metric": "quarterly_patch_rows", "value": len(quarterly_patch)},
        {"metric": "annual_patch_rows", "value": len(annual_patch)},
        {
            "metric": "quarterly_patch_companies",
            "value": int(quarterly_patch["company_id"].nunique()) if not quarterly_patch.empty else 0,
        },
        {
            "metric": "annual_patch_companies",
            "value": int(annual_patch["company_id"].nunique()) if not annual_patch.empty else 0,
        },
        {"metric": "quarterly_changed_cells", "value": len(quarterly_changes)},
        {"metric": "annual_changed_cells", "value": len(annual_changes)},
        {
            "metric": "direct_recovered_columns",
            "value": ", ".join(
                sorted(
                    set(
                        quarterly_changes.loc[quarterly_changes["change_type"].eq("direct_recovered"), "column"].dropna().astype(str)
                    )
                    | set(
                        annual_changes.loc[annual_changes["change_type"].eq("direct_recovered"), "column"].dropna().astype(str)
                    )
                )
            ),
        },
        {
            "metric": "derived_recomputed_columns",
            "value": ", ".join(
                sorted(
                    set(
                        quarterly_changes.loc[quarterly_changes["change_type"].eq("derived_recomputed"), "column"].dropna().astype(str)
                    )
                    | set(
                        annual_changes.loc[annual_changes["change_type"].eq("derived_recomputed"), "column"].dropna().astype(str)
                    )
                )
            ),
        },
    ]
    return pd.DataFrame(rows)


def build_readme_sheet() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "section": "what_this_file_is",
                "text": (
                    "Patch workbook for regression_ready_final.xlsx. It fills only missing SPARK financial-level "
                    "values recovered from rows without metric_code and recomputes affected ratios/logs."
                ),
            },
            {
                "section": "how_to_paste_quarterly",
                "text": (
                    "Use quarterly_patch_same_columns. Match rows by company_id + quarter_label, then replace "
                    "the matching rows/columns in quarterly_panel. Column names and order are identical to regression_ready_final.xlsx."
                ),
            },
            {
                "section": "how_to_paste_annual",
                "text": (
                    "Use annual_patch_same_columns. Match rows by company_id + quarter_label, then replace "
                    "the matching rows/columns in annual_panel."
                ),
            },
            {
                "section": "audit",
                "text": (
                    "Use quarterly_changes_long and annual_changes_long to see every changed cell, direct raw evidence "
                    "from SPARK, and whether the change is direct_recovered or derived_recomputed."
                ),
            },
            {
                "section": "important",
                "text": (
                    "The script does not overwrite existing non-missing financial values. It only fills missing values "
                    "and recalculates dependent variables."
                ),
            },
        ]
    )


def write_patch_workbook(output: Path, sheets: dict[str, pd.DataFrame]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def build_spark_regression_patch(
    regression_ready: Path = DEFAULT_REGRESSION_READY,
    recovery_audit: Path = DEFAULT_RECOVERY_AUDIT,
    output: Path = DEFAULT_PATCH_OUTPUT,
) -> dict[str, pd.DataFrame]:
    quarterly = pd.read_excel(regression_ready, sheet_name="quarterly_panel")
    annual = pd.read_excel(regression_ready, sheet_name="annual_panel")
    evidence = pd.read_excel(recovery_audit, sheet_name="recovery_needed_long")
    recovery_wide = read_recovery_wide(recovery_audit)

    quarterly_patch, _ = apply_recovery_patch(quarterly, recovery_wide)
    annual_patch, _ = apply_recovery_patch(annual, recovery_wide)
    quarterly_changes = build_changes_long(quarterly, quarterly_patch, evidence, "quarterly")
    annual_changes = build_changes_long(annual, annual_patch, evidence, "annual")

    recovered_financial_columns = [
        "company_id",
        "sector",
        "company",
        "company_name_shortlist",
        "inn",
        "report_year",
        "period_type",
        "quarter_label",
    ] + [column for column in FINANCIAL_LEVEL_COLUMNS if column in quarterly_patch.columns]
    recovered_financial_wide = quarterly_patch[
        [column for column in recovered_financial_columns if column in quarterly_patch.columns]
    ].copy()

    summary = build_patch_summary(
        quarterly_patch=quarterly_patch,
        annual_patch=annual_patch,
        quarterly_changes=quarterly_changes,
        annual_changes=annual_changes,
        regression_ready=regression_ready,
        recovery_audit=recovery_audit,
    )
    if not quarterly_patch.empty:
        summary = pd.concat([summary, build_summary(quarterly_patch, "quarterly_patch")], ignore_index=True, sort=False)
    if not annual_patch.empty:
        summary = pd.concat([summary, build_summary(annual_patch, "annual_patch")], ignore_index=True, sort=False)

    sheets = {
        "readme": build_readme_sheet(),
        "summary": summary,
        "quarterly_patch_same_columns": quarterly_patch,
        "annual_patch_same_columns": annual_patch,
        "quarterly_changes_long": quarterly_changes,
        "annual_changes_long": annual_changes,
        "recovered_financial_wide": recovered_financial_wide,
        "source_recovery_wide": recovery_wide,
    }
    write_patch_workbook(output, sheets)
    sheets["output"] = pd.DataFrame([{"path": str(output)}])
    return sheets


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Build a same-column regression patch workbook from SPARK metric-code recovery audit output."
        ),
    )
    parser.add_argument("--regression-ready", type=Path, default=DEFAULT_REGRESSION_READY)
    parser.add_argument("--recovery-audit", type=Path, default=DEFAULT_RECOVERY_AUDIT)
    parser.add_argument("--output", type=Path, default=DEFAULT_PATCH_OUTPUT)
    args = parser.parse_args(argv)

    outputs = build_spark_regression_patch(
        regression_ready=args.regression_ready,
        recovery_audit=args.recovery_audit,
        output=args.output,
    )
    print(f"SPARK regression patch workbook: {args.output}")
    for sheet_name, frame in outputs.items():
        print(f"{sheet_name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
