from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from .spark_docx_parser import (
    CORE_METRIC_COLUMNS,
    build_panel_core,
    classify_metric,
    clean_text,
    normalize_code,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_WORKBOOKS = [
    PROJECT_ROOT / "data_processed" / "spark_neftegaz_report_2014q3_2025q4.xlsx",
    PROJECT_ROOT / "data_processed" / "spark_metallurgy_report_2014q3_2025q4.xlsx",
]
DEFAULT_OUTPUT = PROJECT_ROOT / "data_processed" / "spark_metric_recovery_audit.xlsx"


def existing_default_workbooks() -> list[Path]:
    return [path for path in DEFAULT_INPUT_WORKBOOKS if path.exists()]


def read_sheet_if_exists(workbook: Path, sheet_name: str) -> pd.DataFrame:
    try:
        frame = pd.read_excel(workbook, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()
    frame["source_workbook"] = str(workbook)
    return frame


def classify_raw_current_values(raw_long: pd.DataFrame) -> pd.DataFrame:
    if raw_long.empty:
        return pd.DataFrame()

    current = raw_long.loc[pd.to_numeric(raw_long["value_col_index"], errors="coerce").eq(0)].copy()
    matches = current.apply(
        lambda row: classify_metric(
            table_title=clean_text(row.get("table_title", "")),
            metric_code=normalize_code(row.get("metric_code", "")),
            metric_name=clean_text(row.get("metric_name", "")),
            statement_section=clean_text(row.get("statement_section", "")),
        ),
        axis=1,
    )
    current["metric_std"] = matches.map(lambda value: value[0])
    current["metric_match_method"] = matches.map(lambda value: value[1])
    current["metric_match_confidence"] = matches.map(lambda value: value[2])
    current["metric_code_missing_flag"] = current["metric_code"].map(normalize_code).eq("")
    return current


def panel_to_long(panel: pd.DataFrame, value_name: str) -> pd.DataFrame:
    if panel.empty:
        return pd.DataFrame()
    id_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "data_source",
        "unit",
    ]
    for column in id_columns:
        if column not in panel.columns:
            panel[column] = ""
    value_columns = [column for column in CORE_METRIC_COLUMNS if column in panel.columns]
    return panel.melt(
        id_vars=id_columns,
        value_vars=value_columns,
        var_name="metric_std",
        value_name=value_name,
    )


def build_panel_comparison(existing_panel: pd.DataFrame, recovered_panel: pd.DataFrame) -> pd.DataFrame:
    old_long = panel_to_long(existing_panel.copy(), "old_value")
    new_long = panel_to_long(recovered_panel.copy(), "recovered_value")
    if old_long.empty and new_long.empty:
        return pd.DataFrame()

    key_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "data_source",
        "unit",
        "metric_std",
    ]
    comparison = old_long.merge(new_long, on=key_columns, how="outer")
    comparison["old_value"] = pd.to_numeric(comparison["old_value"], errors="coerce")
    comparison["recovered_value"] = pd.to_numeric(comparison["recovered_value"], errors="coerce")
    comparison["recovery_needed_flag"] = comparison["old_value"].isna() & comparison["recovered_value"].notna()
    comparison["value_changed_flag"] = (
        comparison["old_value"].notna()
        & comparison["recovered_value"].notna()
        & ~np.isclose(comparison["old_value"], comparison["recovered_value"], equal_nan=True)
    )
    return comparison.sort_values(
        ["company", "report_year", "period_type", "metric_std", "source_file"],
        kind="stable",
    )


def add_raw_evidence(recovery_needed: pd.DataFrame, classified_current: pd.DataFrame) -> pd.DataFrame:
    if recovery_needed.empty or classified_current.empty:
        return recovery_needed

    evidence_columns = [
        "source_workbook",
        "source_file",
        "company",
        "inn",
        "report_year",
        "period_type",
        "metric_std",
        "table_title",
        "statement_section",
        "metric_name",
        "metric_code",
        "value_raw",
        "value_num",
        "metric_match_method",
        "metric_match_confidence",
    ]
    evidence = classified_current.loc[
        classified_current["metric_std"].notna()
        & classified_current["metric_code_missing_flag"]
        & classified_current["value_num"].notna(),
        [column for column in evidence_columns if column in classified_current.columns],
    ].copy()
    evidence = evidence.drop_duplicates(
        subset=["source_file", "company", "inn", "report_year", "period_type", "metric_std", "value_num"],
        keep="first",
    )
    merge_columns = ["source_file", "company", "inn", "report_year", "period_type", "metric_std"]
    return recovery_needed.merge(
        evidence,
        on=merge_columns,
        how="left",
        suffixes=("", "_raw"),
    )


def build_recovery_wide(recovery_needed: pd.DataFrame) -> pd.DataFrame:
    if recovery_needed.empty:
        return pd.DataFrame()
    index_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "data_source",
        "unit",
    ]
    wide = (
        recovery_needed.pivot_table(
            index=index_columns,
            columns="metric_std",
            values="recovered_value",
            aggfunc="first",
        )
        .reset_index()
    )
    wide.columns.name = None
    return wide


def build_audit(input_workbooks: list[Path]) -> dict[str, pd.DataFrame]:
    raw_frames: list[pd.DataFrame] = []
    panel_frames: list[pd.DataFrame] = []
    input_rows: list[dict[str, object]] = []

    for workbook in input_workbooks:
        raw = read_sheet_if_exists(workbook, "raw_long")
        panel = read_sheet_if_exists(workbook, "panel_core")
        raw_frames.append(raw)
        panel_frames.append(panel)
        input_rows.append(
            {
                "input_workbook": str(workbook),
                "exists": workbook.exists(),
                "raw_long_rows": len(raw),
                "panel_core_rows": len(panel),
            }
        )

    raw_long = pd.concat(raw_frames, ignore_index=True, sort=False) if raw_frames else pd.DataFrame()
    existing_panel = pd.concat(panel_frames, ignore_index=True, sort=False) if panel_frames else pd.DataFrame()
    classified_current = classify_raw_current_values(raw_long)
    recovered_panel = build_panel_core(raw_long.drop(columns=["source_workbook"], errors="ignore"))
    comparison = build_panel_comparison(existing_panel.drop(columns=["source_workbook"], errors="ignore"), recovered_panel)

    recovered_raw_values = classified_current.loc[
        classified_current["metric_code_missing_flag"]
        & classified_current["metric_std"].notna()
        & classified_current["value_num"].notna()
    ].copy()
    recovery_needed = comparison.loc[comparison["recovery_needed_flag"]].copy()
    recovery_needed = add_raw_evidence(recovery_needed, classified_current)
    direct_recovery_needed = recovery_needed.loc[recovery_needed["metric_match_method"].notna()].copy()
    recovery_wide = build_recovery_wide(direct_recovery_needed)
    debt_recovery_needed = direct_recovery_needed.loc[
        direct_recovery_needed["metric_std"].isin(["debt_lt", "debt_st"])
    ].copy()
    debt_recovery_wide = build_recovery_wide(debt_recovery_needed)

    if recovered_raw_values.empty:
        rule_summary = pd.DataFrame()
    else:
        rule_summary = (
            recovered_raw_values.groupby(
                ["metric_std", "metric_match_method", "metric_match_confidence"],
                dropna=False,
            )
            .agg(
                n_rows=("metric_std", "size"),
                n_companies=("company", "nunique"),
                n_source_files=("source_file", "nunique"),
                n_nonmissing_values=("value_num", lambda values: int(pd.to_numeric(values, errors="coerce").notna().sum())),
            )
            .reset_index()
            .sort_values(["metric_std", "metric_match_method"], kind="stable")
        )

    return {
        "recovery_needed_long": direct_recovery_needed,
        "recovery_needed_wide": recovery_wide,
        "debt_recovery_long": debt_recovery_needed,
        "debt_recovery_wide": debt_recovery_wide,
        "recovered_raw_values": recovered_raw_values,
        "panel_before_after": comparison,
        "rule_summary": rule_summary,
        "input_workbooks": pd.DataFrame(input_rows),
    }


def write_audit_workbook(outputs: dict[str, pd.DataFrame], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in outputs.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Build an audit workbook for SPARK rows recovered from metric_name/statement_section fallback rules.",
    )
    parser.add_argument(
        "--input-workbook",
        type=Path,
        action="append",
        dest="input_workbooks",
        help="SPARK workbook with raw_long and panel_core sheets. Can be passed multiple times.",
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)

    input_workbooks = args.input_workbooks or existing_default_workbooks()
    if not input_workbooks:
        raise FileNotFoundError("No SPARK workbooks found. Pass --input-workbook explicitly.")

    outputs = build_audit(input_workbooks)
    write_audit_workbook(outputs, args.output)
    print(f"SPARK metric recovery audit workbook: {args.output}")
    for sheet_name, frame in outputs.items():
        print(f"{sheet_name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
