from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .utils import (
    clean_text,
    company_name_variants,
    make_log_entry,
    normalize_column_name,
    normalize_identifier,
    normalize_isin,
    normalize_text,
    normalize_ticker,
    pick_preferred_value,
    strip_legal_form,
)


SHORTLIST_COLUMN_CANDIDATES = {
    "company_name": [
        "наименование",
        "company_name",
        "компания",
    ],
    "company_name_short": [
        "краткое наименование",
        "short_name",
    ],
    "company_name_full": [
        "наименование полное",
        "полное наименование",
        "full_name",
    ],
    "company_name_en": [
        "наименование на английском",
        "english_name",
    ],
    "inn": [
        "код налогоплательщика",
        "инн",
        "inn",
    ],
    "ogrn": [
        "регистрационный номер",
        "огрн",
        "ogrn",
    ],
    "ticker": [
        "тикер биржевой",
        "ticker",
    ],
    "isin": [
        "isin",
    ],
    "spark_id": [
        "спарк код",
        "spark_id",
    ],
    "industry": [
        "вид деятельности отрасль",
        "вид деятельности отрасль",
        "вид деятельности отрасль",
        "вид деятельности отрасль",
    ],
    "selection_result": [
        "selection_result",
        "результат отбора",
        "выборка",
    ],
}


def choose_column(columns: list[str], aliases: list[str]) -> str | None:
    normalized = {normalize_column_name(col): col for col in columns}
    for alias in aliases:
        match = normalized.get(normalize_column_name(alias))
        if match:
            return match
    return None


def detect_relevant_sheets(excel_file: pd.ExcelFile) -> tuple[list[str], list[dict[str, object]]]:
    logs: list[dict[str, object]] = []
    candidate_sheets: list[str] = []
    explicit_sample_sheets: list[str] = []

    for sheet_name in excel_file.sheet_names:
        preview = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=3)
        columns = list(preview.columns)
        normalized_columns = {normalize_column_name(col) for col in columns}
        has_company = any(
            alias in normalized_columns
            for alias in ["наименование", "краткое наименование", "наименование полное"]
        )
        has_inn = any(alias in normalized_columns for alias in ["код налогоплательщика", "инн"])
        has_selection = "selection_result" in normalized_columns

        logs.append(
            make_log_entry(
                source="shortlist",
                status="INFO",
                message="Sheet inspected",
                sheet_name=sheet_name,
                columns_preview=", ".join(columns[:10]),
                has_company=has_company,
                has_inn=has_inn,
                has_selection=has_selection,
            )
        )

        if has_company or has_inn or has_selection:
            candidate_sheets.append(sheet_name)

        lower_name = sheet_name.lower()
        if "main" in lower_name or "extended" in lower_name:
            explicit_sample_sheets.append(sheet_name)

    if explicit_sample_sheets:
        return explicit_sample_sheets, logs

    return candidate_sheets, logs


def infer_sample_flag(sheet_name: str, frame: pd.DataFrame) -> str:
    lower_name = sheet_name.lower()
    if "main" in lower_name:
        return "main"
    if "extended" in lower_name:
        return "extended"

    if "selection_result" in frame.columns:
        values = {normalize_text(val) for val in frame["selection_result"].dropna().unique()}
        if any("основ" in value for value in values):
            return "main"
        if any("расшир" in value for value in values):
            return "extended"

    return "unknown"


def standardize_shortlist_frame(frame: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    frame = frame.copy()
    rename_map: dict[str, str] = {}
    columns = list(frame.columns)

    for standard_name, aliases in SHORTLIST_COLUMN_CANDIDATES.items():
        column = choose_column(columns, aliases)
        if column:
            rename_map[column] = standard_name

    frame = frame.rename(columns=rename_map)

    for required in [
        "company_name",
        "company_name_short",
        "company_name_full",
        "company_name_en",
        "inn",
        "ogrn",
        "ticker",
        "isin",
        "spark_id",
        "industry",
        "selection_result",
    ]:
        if required not in frame.columns:
            frame[required] = np.nan

    frame["source_sheet"] = sheet_name
    frame["sample_flag"] = infer_sample_flag(sheet_name, frame)

    frame["company_name"] = frame["company_name"].map(clean_text)
    frame["company_name_short"] = frame["company_name_short"].map(clean_text)
    frame["company_name_full"] = frame["company_name_full"].map(clean_text)
    frame["company_name_en"] = frame["company_name_en"].map(clean_text)
    frame["inn"] = frame["inn"].map(lambda x: normalize_identifier(x, length=10))
    frame["ogrn"] = frame["ogrn"].map(lambda x: normalize_identifier(x, length=13))
    frame["ticker"] = frame["ticker"].map(normalize_ticker)
    frame["isin"] = frame["isin"].map(normalize_isin)
    frame["spark_id"] = frame["spark_id"].map(lambda x: normalize_identifier(x))
    frame["selection_result"] = frame["selection_result"].map(clean_text)
    frame["company_name_core"] = frame["company_name"].map(strip_legal_form)
    frame["name_variants"] = frame.apply(company_name_variants, axis=1)

    return frame


def deduplicate_companies(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame

    frame = frame.copy()
    frame["dedupe_key"] = frame.apply(
        lambda row: pick_preferred_value(
            [
                f"inn:{row['inn']}" if clean_text(row["inn"]) else "",
                f"ogrn:{row['ogrn']}" if clean_text(row["ogrn"]) else "",
                f"ticker:{row['ticker']}" if clean_text(row["ticker"]) else "",
                f"name:{normalize_text(row['company_name'])}" if clean_text(row["company_name"]) else "",
            ]
        ),
        axis=1,
    )

    grouped_rows = []
    for _, group in frame.groupby("dedupe_key", dropna=False, sort=False):
        sample_membership = ",".join(sorted(flag for flag in group["sample_flag"].dropna().unique() if flag))
        preferred_flag = "main" if "main" in set(group["sample_flag"]) else pick_preferred_value(group["sample_flag"])
        grouped_rows.append(
            {
                "company_name": pick_preferred_value(group["company_name"]),
                "company_name_short": pick_preferred_value(group["company_name_short"]),
                "company_name_full": pick_preferred_value(group["company_name_full"]),
                "company_name_en": pick_preferred_value(group["company_name_en"]),
                "company_name_core": pick_preferred_value(group["company_name_core"]),
                "inn": pick_preferred_value(group["inn"]),
                "ogrn": pick_preferred_value(group["ogrn"]),
                "ticker": pick_preferred_value(group["ticker"]),
                "isin": pick_preferred_value(group["isin"]),
                "spark_id": pick_preferred_value(group["spark_id"]),
                "industry": pick_preferred_value(group["industry"]),
                "sample_flag": preferred_flag,
                "sample_membership": sample_membership,
                "selection_result": pick_preferred_value(group["selection_result"]),
                "source_sheet": ",".join(sorted(group["source_sheet"].astype(str).unique())),
                "records_merged": len(group),
                "name_variants": company_name_variants(group.iloc[0]),
            }
        )

    result = pd.DataFrame(grouped_rows)
    return result.sort_values(["sample_flag", "company_name"], kind="stable").reset_index(drop=True)


def load_shortlist(shortlist_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    excel_file = pd.ExcelFile(shortlist_path)
    selected_sheets, logs = detect_relevant_sheets(excel_file)

    sheet_frames = []
    for sheet_name in selected_sheets:
        frame = pd.read_excel(excel_file, sheet_name=sheet_name)
        standardized = standardize_shortlist_frame(frame, sheet_name)
        if standardized["sample_flag"].eq("unknown").all() and "selection_result" in standardized.columns:
            standardized = standardized.loc[standardized["selection_result"].map(clean_text).ne("")]
        sheet_frames.append(standardized)
        logs.append(
            make_log_entry(
                source="shortlist",
                status="INFO",
                message="Sheet loaded",
                sheet_name=sheet_name,
                rows=len(standardized),
                sample_flag=pick_preferred_value(standardized["sample_flag"]),
            )
        )

    if not sheet_frames:
        raise ValueError(
            f"Не удалось найти релевантные листы в shortlist-файле. Доступные листы: {excel_file.sheet_names}"
        )

    combined = pd.concat(sheet_frames, ignore_index=True)

    # Если explicit main/extended листов нет, а есть full_with_filter, режем по selection_result.
    if "sample_flag" in combined.columns and combined["sample_flag"].eq("unknown").all():
        normalized_result = combined["selection_result"].map(normalize_text)
        combined.loc[normalized_result.str.contains("основ", na=False), "sample_flag"] = "main"
        combined.loc[normalized_result.str.contains("расшир", na=False), "sample_flag"] = "extended"
        combined = combined.loc[combined["sample_flag"].isin(["main", "extended"])].copy()

    companies_master = deduplicate_companies(combined)
    logs.append(
        make_log_entry(
            source="shortlist",
            status="INFO",
            message="Shortlist consolidated",
            shortlist_path=str(shortlist_path),
            available_sheets=", ".join(excel_file.sheet_names),
            selected_sheets=", ".join(selected_sheets),
            combined_rows=len(combined),
            unique_companies=len(companies_master),
        )
    )

    return companies_master, pd.DataFrame(logs)
