from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = PROJECT_ROOT / "data_processed" / "regression_ready_final.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "data_processed" / "h1_regression_results.xlsx"

Y_VAR = "debt_to_assets"
MAIN_X = "cfo_to_assets"
BASE_CONTROLS = ["log_assets", "ppe_to_assets", "capex_to_assets"]
EBIT_CONTROL = "ebit_margin"
KEY_VARS = [Y_VAR, MAIN_X, *BASE_CONTROLS, EBIT_CONTROL]


@dataclass
class ModelSpec:
    name: str
    firm_fe: bool = False
    year_fe: bool = False


MODEL_SPECS = [
    ModelSpec("M1 pooled OLS"),
    ModelSpec("M2 pooled OLS + year FE", year_fe=True),
    ModelSpec("M3 firm FE", firm_fe=True),
    ModelSpec("M4 firm FE + year FE", firm_fe=True, year_fe=True),
]


def normal_pvalue(z_value: float) -> float:
    if not np.isfinite(z_value):
        return np.nan
    cdf = 0.5 * (1.0 + math.erf(abs(float(z_value)) / math.sqrt(2.0)))
    return 2.0 * (1.0 - cdf)


def star(p_value: float) -> str:
    if not np.isfinite(p_value):
        return ""
    if p_value < 0.01:
        return "***"
    if p_value < 0.05:
        return "**"
    if p_value < 0.10:
        return "*"
    return ""


def find_sheet(sheet_names: list[str], candidates: list[str]) -> str:
    lower_map = {name.lower(): name for name in sheet_names}
    for candidate in candidates:
        if candidate.lower() in lower_map:
            return lower_map[candidate.lower()]
    for name in sheet_names:
        low = name.lower()
        if any(candidate.lower() in low for candidate in candidates):
            return name
    raise ValueError(f"Could not find a matching sheet among {candidates}")


def year_column(frame: pd.DataFrame) -> str:
    if "report_year" in frame.columns:
        return "report_year"
    if "year" in frame.columns:
        return "year"
    raise ValueError("No year column found: expected report_year or year")


def company_column(frame: pd.DataFrame) -> str:
    for column in ["company_id", "company_name", "company_name_shortlist", "company"]:
        if column in frame.columns:
            return column
    raise ValueError("No company identifier column found")


def coerce_numeric(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = frame.copy()
    for column in columns:
        if column in out.columns:
            out[column] = pd.to_numeric(out[column], errors="coerce")
    return out


def missingness(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    rows = []
    total = len(frame)
    for column in columns:
        missing = int(frame[column].isna().sum()) if column in frame.columns else total
        rows.append(
            {
                "variable": column,
                "n_total": total,
                "n_missing": missing,
                "missing_share": missing / total if total else np.nan,
                "n_nonmissing": total - missing,
            }
        )
    return pd.DataFrame(rows)


def descriptive_stats(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    rows = []
    for column in columns:
        series = pd.to_numeric(frame[column], errors="coerce").dropna()
        rows.append(
            {
                "variable": column,
                "n": int(series.size),
                "mean": float(series.mean()) if series.size else np.nan,
                "median": float(series.median()) if series.size else np.nan,
                "std": float(series.std(ddof=1)) if series.size > 1 else np.nan,
                "min": float(series.min()) if series.size else np.nan,
                "max": float(series.max()) if series.size else np.nan,
                "p1": float(series.quantile(0.01)) if series.size else np.nan,
                "p99": float(series.quantile(0.99)) if series.size else np.nan,
            }
        )
    return pd.DataFrame(rows)


def build_sample(frame: pd.DataFrame, x_vars: list[str], years: tuple[int, int]) -> pd.DataFrame:
    ycol = year_column(frame)
    ccol = company_column(frame)
    needed = [ccol, ycol, Y_VAR, *x_vars]
    sample = frame.copy()
    sample[ycol] = pd.to_numeric(sample[ycol], errors="coerce")
    sample = coerce_numeric(sample, [Y_VAR, *x_vars])
    sample = sample.loc[sample[ycol].between(years[0], years[1])].copy()
    sample = sample.dropna(subset=needed)
    sample[ccol] = sample[ccol].astype(str)
    sample[ycol] = sample[ycol].astype(int)
    return sample


def make_design(sample: pd.DataFrame, x_vars: list[str], spec: ModelSpec) -> tuple[np.ndarray, list[str]]:
    ycol = year_column(sample)
    ccol = company_column(sample)
    parts = [pd.Series(1.0, index=sample.index, name="const"), sample[x_vars].astype(float)]
    if spec.firm_fe:
        firm_dummies = pd.get_dummies(sample[ccol].astype(str), prefix="firm", drop_first=True, dtype=float)
        parts.append(firm_dummies)
    if spec.year_fe:
        year_dummies = pd.get_dummies(sample[ycol].astype(str), prefix="year", drop_first=True, dtype=float)
        parts.append(year_dummies)
    design = pd.concat(parts, axis=1)
    return design.to_numpy(dtype=float), list(design.columns)


def fit_ols_cluster(sample: pd.DataFrame, x_vars: list[str], spec: ModelSpec) -> dict[str, object]:
    y = sample[Y_VAR].to_numpy(dtype=float)
    x, names = make_design(sample, x_vars, spec)
    ccol = company_column(sample)
    clusters = sample[ccol].astype(str).to_numpy()
    nobs, k = x.shape

    xtx_inv = np.linalg.pinv(x.T @ x)
    beta = xtx_inv @ (x.T @ y)
    resid = y - x @ beta

    meat = np.zeros((k, k), dtype=float)
    unique_clusters = np.unique(clusters)
    for cluster in unique_clusters:
        mask = clusters == cluster
        xu = x[mask].T @ resid[mask]
        meat += np.outer(xu, xu)

    vcov = xtx_inv @ meat @ xtx_inv
    g = len(unique_clusters)
    if g > 1 and nobs > k:
        vcov *= (g / (g - 1)) * ((nobs - 1) / (nobs - k))
    se = np.sqrt(np.clip(np.diag(vcov), 0, np.inf))
    tstat = beta / se
    pvals = np.array([normal_pvalue(value) for value in tstat])

    y_centered = y - y.mean()
    r2 = 1.0 - float(np.sum(resid**2) / np.sum(y_centered**2)) if np.sum(y_centered**2) > 0 else np.nan
    return {
        "model": spec.name,
        "nobs": int(nobs),
        "n_companies": int(g),
        "n_years": int(sample[year_column(sample)].nunique()),
        "k": int(k),
        "r2": r2,
        "firm_fe": spec.firm_fe,
        "year_fe": spec.year_fe,
        "names": names,
        "beta": beta,
        "se": se,
        "tstat": tstat,
        "pval": pvals,
    }


def model_table(results: list[dict[str, object]], variables: list[str]) -> pd.DataFrame:
    rows = []
    for variable in variables:
        row = {"term": variable}
        for result in results:
            names = result["names"]
            if variable in names:
                idx = names.index(variable)
                coef = float(result["beta"][idx])
                se = float(result["se"][idx])
                pval = float(result["pval"][idx])
                row[result["model"]] = f"{coef:.4f}{star(pval)} ({se:.4f})"
            else:
                row[result["model"]] = ""
        rows.append(row)

    for stat_label, key in [
        ("N", "nobs"),
        ("Companies", "n_companies"),
        ("Years", "n_years"),
        ("R2", "r2"),
        ("Firm FE", "firm_fe"),
        ("Year FE", "year_fe"),
    ]:
        row = {"term": stat_label}
        for result in results:
            value = result[key]
            if isinstance(value, bool):
                row[result["model"]] = "Yes" if value else "No"
            elif key == "r2":
                row[result["model"]] = f"{float(value):.3f}"
            else:
                row[result["model"]] = int(value)
        rows.append(row)
    return pd.DataFrame(rows)


def coefficient_summary(results: list[dict[str, object]], sample_label: str) -> pd.DataFrame:
    rows = []
    for result in results:
        idx = result["names"].index(MAIN_X)
        rows.append(
            {
                "sample": sample_label,
                "model": result["model"],
                "coef_cfo_to_assets": float(result["beta"][idx]),
                "se_cluster_company": float(result["se"][idx]),
                "z_stat": float(result["tstat"][idx]),
                "p_value_normal_approx": float(result["pval"][idx]),
                "significant_10pct": bool(result["pval"][idx] < 0.10),
                "significant_5pct": bool(result["pval"][idx] < 0.05),
                "nobs": int(result["nobs"]),
                "n_companies": int(result["n_companies"]),
                "n_years": int(result["n_years"]),
                "r2": float(result["r2"]),
            }
        )
    return pd.DataFrame(rows)


def outlier_sensitivity(frame: pd.DataFrame, x_vars: list[str]) -> pd.DataFrame:
    base = build_sample(frame, x_vars, (2014, 2025))
    cfo_low = base[MAIN_X].quantile(0.01)
    cfo_high = base[MAIN_X].quantile(0.99)
    trimmed = base.loc[
        base[Y_VAR].between(0, 1)
        & base[MAIN_X].between(cfo_low, cfo_high)
    ].copy()
    results = [fit_ols_cluster(trimmed, x_vars, MODEL_SPECS[-1])]
    out = coefficient_summary(results, "annual_2014_2025_trim_debt_0_1_cfo_p1_p99")
    out["dropped_rows_vs_base"] = len(base) - len(trimmed)
    out["base_rows"] = len(base)
    out["cfo_p1_cutoff"] = float(cfo_low)
    out["cfo_p99_cutoff"] = float(cfo_high)
    out["debt_filter"] = "0 <= debt_to_assets <= 1"
    return out


def quarterly_missingness_report(frame: pd.DataFrame) -> pd.DataFrame:
    qcols = [Y_VAR, MAIN_X, *BASE_CONTROLS]
    ycol = year_column(frame)
    frame = frame.copy()
    frame[ycol] = pd.to_numeric(frame[ycol], errors="coerce")
    frame = coerce_numeric(frame, qcols)
    frame = frame.loc[frame[ycol].between(2014, 2025)].copy()
    group_column = "quarter_label" if "quarter_label" in frame.columns else "period_type"
    rows = []
    for period, group in frame.groupby(group_column, dropna=False):
        complete = group[qcols].notna().all(axis=1)
        row = {"period": period, "n_rows": len(group), "complete_base_rows": int(complete.sum())}
        for col in qcols:
            row[f"{col}_nonmissing"] = int(group[col].notna().sum())
        rows.append(row)
    return pd.DataFrame(rows).sort_values("period").reset_index(drop=True)


def build_report(input_path: Path = DEFAULT_INPUT, output_path: Path = DEFAULT_OUTPUT) -> dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(input_path)
    annual_sheet = find_sheet(excel.sheet_names, ["annual_panel", "annual", "panel_annual", "model_ready_annual"])
    quarterly_sheet = find_sheet(excel.sheet_names, ["quarterly_panel", "quarterly", "panel_quarterly", "model_ready_quarterly"])

    annual = pd.read_excel(input_path, sheet_name=annual_sheet)
    quarterly = pd.read_excel(input_path, sheet_name=quarterly_sheet)
    ycol = year_column(annual)
    ccol = company_column(annual)
    annual[ycol] = pd.to_numeric(annual[ycol], errors="coerce")
    annual = coerce_numeric(annual, KEY_VARS)
    annual_2014_2025 = annual.loc[annual[ycol].between(2014, 2025)].copy()

    base_x = [MAIN_X, *BASE_CONTROLS]
    ebit_x = [MAIN_X, *BASE_CONTROLS, EBIT_CONTROL]
    base_sample = build_sample(annual, base_x, (2014, 2025))
    ebit_sample = build_sample(annual, ebit_x, (2014, 2025))
    late_sample = build_sample(annual, base_x, (2019, 2025))

    base_results = [fit_ols_cluster(base_sample, base_x, spec) for spec in MODEL_SPECS]
    ebit_results = [fit_ols_cluster(ebit_sample, ebit_x, spec) for spec in MODEL_SPECS]
    late_results = [fit_ols_cluster(late_sample, base_x, spec) for spec in MODEL_SPECS]

    diagnostics = pd.DataFrame(
        [
            {"metric": "input_file", "value": str(input_path)},
            {"metric": "sheets", "value": ", ".join(excel.sheet_names)},
            {"metric": "annual_sheet", "value": annual_sheet},
            {"metric": "quarterly_sheet", "value": quarterly_sheet},
            {"metric": "annual_rows_2014_2025", "value": len(annual_2014_2025)},
            {"metric": "annual_companies_2014_2025", "value": int(annual_2014_2025[ccol].nunique())},
            {"metric": "annual_year_min", "value": int(annual_2014_2025[ycol].min())},
            {"metric": "annual_year_max", "value": int(annual_2014_2025[ycol].max())},
            {"metric": "base_complete_rows", "value": len(base_sample)},
            {"metric": "base_complete_companies", "value": int(base_sample[ccol].nunique())},
            {"metric": "ebit_complete_rows", "value": len(ebit_sample)},
            {"metric": "ebit_complete_companies", "value": int(ebit_sample[ccol].nunique())},
            {"metric": "late_2019_2025_complete_rows", "value": len(late_sample)},
            {"metric": "late_2019_2025_complete_companies", "value": int(late_sample[ccol].nunique())},
            {
                "metric": "corr_debt_to_assets_cfo_to_assets_base_sample",
                "value": float(base_sample[Y_VAR].corr(base_sample[MAIN_X])),
            },
        ]
    )

    outputs = {
        "diagnostics": diagnostics,
        "annual_missingness": missingness(annual_2014_2025, KEY_VARS),
        "annual_descriptive_base": descriptive_stats(base_sample, [Y_VAR, MAIN_X, *BASE_CONTROLS, EBIT_CONTROL]),
        "regression_table_base": model_table(base_results, [MAIN_X, *BASE_CONTROLS]),
        "cfo_coefficients_base": coefficient_summary(base_results, "annual_2014_2025_base"),
        "regression_table_with_ebit": model_table(ebit_results, [MAIN_X, *BASE_CONTROLS, EBIT_CONTROL]),
        "cfo_coefficients_with_ebit": coefficient_summary(ebit_results, "annual_2014_2025_with_ebit"),
        "cfo_coefficients_2019_2025": coefficient_summary(late_results, "annual_2019_2025_base"),
        "outlier_sensitivity": outlier_sensitivity(annual, base_x),
        "quarterly_missingness": quarterly_missingness_report(quarterly),
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, frame in outputs.items():
            frame.to_excel(writer, sheet_name=sheet[:31], index=False)
    return outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Run H1 pecking-order regression smoke test.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    outputs = build_report(args.input, args.output)
    print(f"H1 regression report: {args.output}")
    print(outputs["diagnostics"].to_string(index=False))
    print(outputs["regression_table_base"].to_string(index=False))


if __name__ == "__main__":
    main()
