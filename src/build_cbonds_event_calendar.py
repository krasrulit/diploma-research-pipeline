from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

from .utils import clean_text, normalize_identifier, normalize_isin, normalize_text, strip_legal_form, to_float


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_PROCESSED = PROJECT_ROOT / "data_processed"
DEFAULT_CBONDS_WORKBOOK = DATA_PROCESSED / "cbonds_bond_cards_validated.xlsx"
DEFAULT_QUOTES_WORKBOOK = DATA_PROCESSED / "cbonds_quotes_validated.xlsx"
DEFAULT_EVENT_OUTPUT = DATA_PROCESSED / "cbonds_event_calendar_processed.xlsx"
DEFAULT_EVENT_MISSING_OUTPUT = DATA_PROCESSED / "cbonds_event_calendar_missing_check.xlsx"

CALENDAR_COLUMN_RENAME = {
    "Дата": "event_date",
    "Тип события": "event_type",
    "Страна": "event_country",
    "Эмитент": "event_issuer",
    "Бумага": "event_issue_name",
    "ISIN": "event_isin",
    "Регистрационный номер": "event_registration_number",
    "Тип бумаги": "event_security_type",
    "Анонсированный объём": "event_announced_volume",
    "Объем эмиссии": "event_issue_volume",
    "Объем в обращении": "event_circulation_volume",
    "Объем доразмещения": "event_additional_placement_volume",
    "Валюта": "event_currency",
    "Номинал / Минимальный торговый лот": "event_nominal_or_lot",
    "Начало периода предъявления": "event_offer_window_start",
    "Окончание периода предъявления": "event_offer_window_end",
    "Период уведомления": "event_notification_period",
    "Цена оферты": "event_offer_price_pct",
    "Ставка купона, %": "event_coupon_rate_pct",
    "Размер купона": "event_coupon_amount",
    "Объем выплат": "event_payment_volume",
    "Дата фиксации списка держателей": "event_record_date",
    "Дата фактической выплаты": "event_actual_payment_date",
    "Погашение номинала": "event_principal_repayment",
    "Выкупленный объем": "event_repurchased_volume",
    "Дата погашения": "event_maturity_date",
    "Организаторы": "event_organizers",
    "Цена размещения": "event_placement_price_pct",
    "Доходность при размещении": "event_placement_yield_pct",
    "Время открытия книги": "event_book_open_time",
    "Время закрытия книги": "event_book_close_time",
    "Ориентир по купону (нижняя граница)": "event_coupon_guidance_low_pct",
    "Ориентир по купону (верхняя граница)": "event_coupon_guidance_high_pct",
    "Ориентир по доходности (нижняя граница)": "event_yield_guidance_low_pct",
    "Ориентир по доходности (верхняя граница)": "event_yield_guidance_high_pct",
    "Статус дефолта": "event_default_status",
    "Причина дефолта": "event_default_reason",
    "Дата окончания grace period": "event_grace_period_end_date",
    "Фактическая дата исполнения обязательств": "event_obligation_actual_date",
    "Объем неисполненных обязательств": "event_defaulted_obligation_volume",
}

DATE_COLUMNS = [
    "event_date",
    "event_offer_window_start",
    "event_offer_window_end",
    "event_record_date",
    "event_actual_payment_date",
    "event_maturity_date",
    "event_grace_period_end_date",
    "event_obligation_actual_date",
]

NUMERIC_COLUMNS = [
    "event_announced_volume",
    "event_issue_volume",
    "event_circulation_volume",
    "event_additional_placement_volume",
    "event_nominal_or_lot",
    "event_offer_price_pct",
    "event_coupon_rate_pct",
    "event_coupon_amount",
    "event_payment_volume",
    "event_principal_repayment",
    "event_repurchased_volume",
    "event_placement_price_pct",
    "event_placement_yield_pct",
    "event_coupon_guidance_low_pct",
    "event_coupon_guidance_high_pct",
    "event_yield_guidance_low_pct",
    "event_yield_guidance_high_pct",
    "event_defaulted_obligation_volume",
]

EVENT_TYPE_GROUPS = [
    "bookbuilding",
    "placement",
    "additional_placement",
    "coupon_payment",
    "offer_put",
    "offer_call",
    "offer_other",
    "maturity",
    "early_redemption",
    "amortization",
    "default",
    "other",
]

STATIC_EVENT_COUNT_COLUMNS = [
    "cbonds_calendar_event_count",
    "cbonds_calendar_issue_count",
    "cbonds_calendar_bookbuilding_count",
    "cbonds_calendar_placement_count",
    "cbonds_calendar_additional_placement_count",
    "cbonds_calendar_coupon_payment_count",
    "cbonds_calendar_offer_put_count",
    "cbonds_calendar_offer_call_count",
    "cbonds_calendar_offer_other_count",
    "cbonds_calendar_maturity_count",
    "cbonds_calendar_early_redemption_count",
    "cbonds_calendar_amortization_count",
    "cbonds_calendar_default_count",
    "cbonds_calendar_other_count",
    "cbonds_calendar_placement_volume_rub_sum",
    "cbonds_calendar_coupon_payment_sum",
    "cbonds_calendar_principal_repayment_sum",
]

QUARTERLY_EVENT_COLUMNS = [
    "cbonds_calendar_event_q_count",
    "cbonds_calendar_issue_q_count",
    "cbonds_calendar_bookbuilding_q_count",
    "cbonds_calendar_placement_q_count",
    "cbonds_calendar_additional_placement_q_count",
    "cbonds_calendar_coupon_payment_q_count",
    "cbonds_calendar_offer_put_q_count",
    "cbonds_calendar_offer_call_q_count",
    "cbonds_calendar_offer_other_q_count",
    "cbonds_calendar_maturity_q_count",
    "cbonds_calendar_early_redemption_q_count",
    "cbonds_calendar_amortization_q_count",
    "cbonds_calendar_default_q_count",
    "cbonds_calendar_other_q_count",
    "cbonds_calendar_placement_volume_rub_q",
    "cbonds_calendar_coupon_payment_sum_q",
    "cbonds_calendar_principal_repayment_sum_q",
]


def find_default_event_calendar(project_root: Path = PROJECT_ROOT) -> Path | None:
    candidates = [
        path
        for path in project_root.glob("Календарь*событ*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not candidates:
        return None
    return sorted(candidates, key=lambda path: path.stat().st_mtime, reverse=True)[0]


DEFAULT_EVENT_CALENDAR = find_default_event_calendar()


def safe_read_excel(path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()


def compact_sheet_name(value: object) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = re.sub(r"[^\w]+", "_", text, flags=re.UNICODE)
    return re.sub(r"_+", "_", text).strip("_").lower()


def name_key(value: object) -> str:
    text = strip_legal_form(value)
    text = normalize_text(text)
    text = re.sub(r"\bанк\b", " ", text)
    text = re.sub(r"\bим\b", " ", text)
    text = re.sub(r"\bв\b|\bд\b|\bшашина\b", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def first_notna(values: Iterable[object]) -> object:
    for value in values:
        if pd.notna(value) and clean_text(value):
            return value
    return ""


def join_unique(values: pd.Series, limit: int | None = None) -> str:
    seen: set[str] = set()
    result: list[str] = []
    for value in values.dropna().astype(str):
        text = clean_text(value)
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
        if limit and len(result) >= limit:
            break
    return " | ".join(result)


def quarter_label_from_date(value: object) -> str:
    date = pd.to_datetime(value, errors="coerce")
    if pd.isna(date):
        return ""
    quarter = (int(date.month) - 1) // 3 + 1
    return f"{int(date.year)}Q{quarter}"


def extract_calendar_metadata(raw: pd.DataFrame) -> dict[str, str]:
    metadata = {
        "sheet_issuer": "",
        "issue_types": "",
        "selected_events": "",
        "calendar_period": "",
    }
    for _, row in raw.iloc[:8].iterrows():
        for value in row.dropna().astype(str):
            text = clean_text(value)
            if text.startswith("Эмитент:"):
                metadata["sheet_issuer"] = text.replace("Эмитент:", "", 1).strip()
            elif text.startswith("Типы эмиссии:"):
                metadata["issue_types"] = text.replace("Типы эмиссии:", "", 1).strip()
            elif text.startswith("События:"):
                metadata["selected_events"] = text.replace("События:", "", 1).strip()
            elif text.startswith("Период:"):
                metadata["calendar_period"] = text.replace("Период:", "", 1).strip()
    return metadata


def find_header_row(raw: pd.DataFrame) -> int | None:
    for idx, row in raw.iterrows():
        values = {clean_text(value) for value in row.tolist() if clean_text(value)}
        if {"Дата", "Тип события", "ISIN"}.issubset(values):
            return int(idx)
    return None


def event_type_group(value: object) -> str:
    text = normalize_text(value)
    if "книга заявок" in text:
        return "bookbuilding"
    if "доразмещ" in text:
        return "additional_placement"
    if "размещ" in text:
        return "placement"
    if "выплата купона" in text or text == "купон":
        return "coupon_payment"
    if "оферта" in text and "put" in text:
        return "offer_put"
    if "call" in text or ("оферта" in text and "call" in text):
        return "offer_call"
    if "оферта" in text:
        return "offer_other"
    if "досрочное погаш" in text:
        return "early_redemption"
    if "погаш" in text and "дефолт" not in text:
        return "maturity"
    if "амортиз" in text:
        return "amortization"
    if "дефолт" in text:
        return "default"
    return "other"


def load_raw_event_calendar(calendar_workbook: Path | None) -> tuple[pd.DataFrame, pd.DataFrame]:
    if calendar_workbook is None or not calendar_workbook.exists():
        log = pd.DataFrame(
            [
                {
                    "source": "cbonds_event_calendar",
                    "status": "MISSING",
                    "message": "Cbonds event calendar workbook was not found.",
                    "calendar_workbook": "" if calendar_workbook is None else str(calendar_workbook),
                }
            ]
        )
        return pd.DataFrame(), log

    workbook = pd.ExcelFile(calendar_workbook)
    frames: list[pd.DataFrame] = []
    log_rows: list[dict[str, object]] = []

    for sheet_name in workbook.sheet_names:
        raw_head = pd.read_excel(calendar_workbook, sheet_name=sheet_name, header=None, nrows=40)
        header_row = find_header_row(raw_head)
        metadata = extract_calendar_metadata(raw_head)
        if header_row is None:
            log_rows.append(
                {
                    "source": "cbonds_event_calendar",
                    "status": "SKIP",
                    "message": "Calendar table header not found.",
                    "calendar_workbook": str(calendar_workbook),
                    "sheet_name": sheet_name,
                    **metadata,
                }
            )
            continue

        frame = pd.read_excel(calendar_workbook, sheet_name=sheet_name, header=header_row)
        frame = frame.loc[:, frame.columns.notna()].copy()
        frame.columns = [clean_text(column) for column in frame.columns]
        frame = frame.dropna(how="all")
        if "Дата" not in frame.columns:
            log_rows.append(
                {
                    "source": "cbonds_event_calendar",
                    "status": "SKIP",
                    "message": "Date column not found after header parsing.",
                    "calendar_workbook": str(calendar_workbook),
                    "sheet_name": sheet_name,
                    "header_row": header_row + 1,
                    **metadata,
                }
            )
            continue

        frame = frame.loc[pd.to_datetime(frame["Дата"], errors="coerce").notna()].copy()
        frame = frame.rename(columns=CALENDAR_COLUMN_RENAME)
        for column in CALENDAR_COLUMN_RENAME.values():
            if column not in frame.columns:
                frame[column] = np.nan

        frame["calendar_source_file"] = calendar_workbook.name
        frame["calendar_source_sheet"] = sheet_name
        frame["calendar_source_sheet_key"] = compact_sheet_name(sheet_name)
        for key, value in metadata.items():
            frame[key] = value
        frames.append(frame)

        log_rows.append(
            {
                "source": "cbonds_event_calendar",
                "status": "OK",
                "message": "Calendar sheet parsed.",
                "calendar_workbook": str(calendar_workbook),
                "sheet_name": sheet_name,
                "header_row": header_row + 1,
                "event_rows": len(frame),
                **metadata,
            }
        )

    if not frames:
        return pd.DataFrame(), pd.DataFrame(log_rows)

    events = pd.concat(frames, ignore_index=True, sort=False)
    for column in DATE_COLUMNS:
        events[column] = pd.to_datetime(events[column], errors="coerce")
    for column in NUMERIC_COLUMNS:
        events[column] = events[column].map(to_float)

    text_columns = [
        "event_type",
        "event_country",
        "event_issuer",
        "event_issue_name",
        "event_registration_number",
        "event_security_type",
        "event_currency",
        "event_notification_period",
        "event_organizers",
        "event_default_status",
        "event_default_reason",
        "sheet_issuer",
        "issue_types",
        "selected_events",
        "calendar_period",
    ]
    for column in text_columns:
        events[column] = events[column].map(clean_text)

    events["event_isin"] = events["event_isin"].map(normalize_isin)
    events["event_type_group"] = events["event_type"].map(event_type_group)
    events["quarter_label"] = events["event_date"].map(quarter_label_from_date)
    events["event_issue_key"] = np.where(
        events["event_isin"].astype(str).str.len().gt(0),
        events["event_isin"],
        events["event_issue_name"].map(name_key),
    )
    events["event_issue_volume_rub"] = np.where(
        events["event_currency"].str.upper().eq("RUB"),
        events["event_issue_volume"],
        np.nan,
    )
    events["event_circulation_volume_rub"] = np.where(
        events["event_currency"].str.upper().eq("RUB"),
        events["event_circulation_volume"],
        np.nan,
    )

    return events, pd.DataFrame(log_rows)


def read_issue_scope(cbonds_workbook: Path) -> pd.DataFrame:
    scopes = [
        ("issue_master_validated_current", "current"),
        ("issue_master_validated_related", "related"),
        ("issue_master_validated_excluded", "excluded"),
    ]
    frames: list[pd.DataFrame] = []
    for sheet_name, scope in scopes:
        frame = safe_read_excel(cbonds_workbook, sheet_name)
        if frame.empty:
            continue
        frame = frame.copy()
        frame["cbonds_issue_scope"] = scope
        frame["parsed_isin"] = frame.get("parsed_isin", "").map(normalize_isin)
        frame["scope_priority"] = {"current": 0, "related": 1, "excluded": 2}[scope]
        frames.append(frame)
    if not frames:
        return pd.DataFrame()

    issue_scope = pd.concat(frames, ignore_index=True, sort=False)
    issue_scope = issue_scope.sort_values("scope_priority", kind="stable")
    issue_scope = issue_scope.drop_duplicates(subset=["parsed_isin"], keep="first")
    selected = [
        "parsed_isin",
        "cbonds_issue_scope",
        "company_id",
        "company_name",
        "company_inn",
        "sector",
        "sample_flag",
        "issue_name",
        "borrower_name",
        "issuer_name",
        "validated_relation_label",
        "validated_code_source",
        "quote_collection_bucket_validated",
        "quote_collection_rank_validated",
        "best_n_trade_dates",
    ]
    selected = [column for column in selected if column in issue_scope.columns]
    return issue_scope[selected].rename(
        columns={
            "company_id": "card_company_id",
            "company_name": "card_company_name",
            "company_inn": "card_company_inn",
            "sector": "card_sector",
            "sample_flag": "card_sample_flag",
            "issue_name": "card_issue_name",
        }
    )


def read_quote_scope(quotes_workbook: Path | None) -> pd.DataFrame:
    if quotes_workbook is None or not quotes_workbook.exists():
        return pd.DataFrame(columns=["parsed_isin", "has_validated_quote_row_flag", "quote_scope"])

    frames: list[pd.DataFrame] = []
    for sheet_name, scope in [
        ("quotes_current_company", "current"),
        ("quotes_related_optional", "related"),
        ("quotes_excluded", "excluded"),
    ]:
        frame = safe_read_excel(quotes_workbook, sheet_name)
        if frame.empty or "parsed_isin" not in frame.columns:
            continue
        out = frame[["parsed_isin"]].copy()
        out["parsed_isin"] = out["parsed_isin"].map(normalize_isin)
        out["quote_scope"] = scope
        out["quote_priority"] = {"current": 0, "related": 1, "excluded": 2}[scope]
        frames.append(out)
    if not frames:
        return pd.DataFrame(columns=["parsed_isin", "has_validated_quote_row_flag", "quote_scope"])

    quotes = pd.concat(frames, ignore_index=True, sort=False)
    quotes = quotes.loc[quotes["parsed_isin"].astype(str).str.len().gt(0)].copy()
    quotes = quotes.sort_values("quote_priority", kind="stable").drop_duplicates(subset=["parsed_isin"], keep="first")
    quotes["has_validated_quote_row_flag"] = True
    return quotes[["parsed_isin", "has_validated_quote_row_flag", "quote_scope"]]


def build_issuer_company_map(companies_master: pd.DataFrame, issue_scope: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    current_issues = issue_scope.loc[issue_scope.get("cbonds_issue_scope", "").eq("current")].copy()
    for _, row in current_issues.iterrows():
        for source_column in ["borrower_name", "issuer_name", "card_company_name"]:
            if source_column not in current_issues.columns:
                continue
            key = name_key(row.get(source_column, ""))
            if key:
                rows.append(
                    {
                        "issuer_key": key,
                        "company_id": row.get("card_company_id", ""),
                        "map_company_name": row.get("card_company_name", ""),
                        "issuer_map_source": f"cbonds_current_issue_{source_column}",
                    }
                )

    master = companies_master.copy()
    master["company_id"] = master["sector_company_id"]
    for _, row in master.iterrows():
        for source_column in ["company_name", "company_name_short", "company_name_full", "company_name_core"]:
            key = name_key(row.get(source_column, ""))
            if key:
                rows.append(
                    {
                        "issuer_key": key,
                        "company_id": row.get("company_id", ""),
                        "map_company_name": row.get("company_name", ""),
                        "issuer_map_source": f"companies_master_{source_column}",
                    }
                )

    if not rows:
        return pd.DataFrame(columns=["issuer_key", "company_id", "map_company_name", "issuer_map_source"])

    mapping = pd.DataFrame(rows)
    mapping = mapping.loc[mapping["company_id"].astype(str).str.len().gt(0)].copy()
    unique_counts = mapping.groupby("issuer_key")["company_id"].nunique().reset_index(name="n_company_ids")
    mapping = mapping.merge(unique_counts, on="issuer_key", how="left")
    mapping = mapping.loc[mapping["n_company_ids"].eq(1)].copy()
    mapping = mapping.sort_values(["issuer_key", "issuer_map_source"], kind="stable")
    mapping = mapping.drop_duplicates(subset=["issuer_key"], keep="first")
    return mapping[["issuer_key", "company_id", "map_company_name", "issuer_map_source"]]


def map_events_to_companies(
    events: pd.DataFrame,
    companies_master: pd.DataFrame,
    issue_scope: pd.DataFrame,
    quote_scope: pd.DataFrame,
) -> pd.DataFrame:
    if events.empty:
        return events.copy()

    issuer_map = build_issuer_company_map(companies_master, issue_scope)
    mapped = events.copy()
    mapped["event_issuer_key"] = mapped["event_issuer"].map(name_key)

    if not issue_scope.empty:
        mapped = mapped.merge(issue_scope, left_on="event_isin", right_on="parsed_isin", how="left")
    else:
        mapped["cbonds_issue_scope"] = ""
        mapped["card_company_id"] = ""

    if not quote_scope.empty:
        mapped = mapped.merge(quote_scope, left_on="event_isin", right_on="parsed_isin", how="left", suffixes=("", "_quote"))
    else:
        mapped["has_validated_quote_row_flag"] = False
        mapped["quote_scope"] = ""

    mapped = mapped.merge(issuer_map, left_on="event_issuer_key", right_on="issuer_key", how="left")
    mapped["company_id"] = np.where(
        mapped["cbonds_issue_scope"].eq("current") & mapped["card_company_id"].notna(),
        mapped["card_company_id"],
        mapped["company_id"],
    )
    mapped["event_mapping_source"] = np.select(
        [
            mapped["cbonds_issue_scope"].eq("current") & mapped["card_company_id"].notna(),
            mapped["company_id"].notna() & mapped["issuer_map_source"].notna(),
        ],
        ["validated_current_issue_isin", "calendar_issuer_name"],
        default="unmapped",
    )
    mapped["event_mapping_status"] = np.where(mapped["company_id"].notna() & mapped["company_id"].astype(str).str.len().gt(0), "mapped", "unmapped")
    mapped["has_validated_current_issue_card_flag"] = mapped["cbonds_issue_scope"].eq("current")
    mapped["has_any_issue_card_flag"] = mapped["cbonds_issue_scope"].isin(["current", "related", "excluded"])
    mapped["has_validated_quote_row_flag"] = mapped["has_validated_quote_row_flag"].fillna(False).astype(bool)
    return mapped


def type_counts(frame: pd.DataFrame, index_columns: list[str], prefix: str, suffix: str = "_count") -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=index_columns + [f"{prefix}_{group}{suffix}" for group in EVENT_TYPE_GROUPS])
    pivot = (
        frame.pivot_table(
            index=index_columns,
            columns="event_type_group",
            values="event_date",
            aggfunc="size",
            fill_value=0,
        )
        .reset_index()
    )
    for group in EVENT_TYPE_GROUPS:
        if group not in pivot.columns:
            pivot[group] = 0
    rename = {group: f"{prefix}_{group}{suffix}" for group in EVENT_TYPE_GROUPS}
    return pivot[index_columns + EVENT_TYPE_GROUPS].rename(columns=rename)


def build_company_summary(companies_master: pd.DataFrame, mapped_events: pd.DataFrame) -> pd.DataFrame:
    base = companies_master[["sector_company_id"]].rename(columns={"sector_company_id": "company_id"}).copy()
    if mapped_events.empty:
        for column in STATIC_EVENT_COUNT_COLUMNS:
            base[column] = 0
        base["cbonds_calendar_found_flag"] = False
        base["cbonds_calendar_first_event_date"] = pd.NaT
        base["cbonds_calendar_last_event_date"] = pd.NaT
        return base

    events = mapped_events.loc[mapped_events["event_mapping_status"].eq("mapped")].copy()
    grouped = (
        events.groupby("company_id", dropna=False)
        .agg(
            cbonds_calendar_event_count=("event_date", "size"),
            cbonds_calendar_issue_count=("event_issue_key", pd.Series.nunique),
            cbonds_calendar_first_event_date=("event_date", "min"),
            cbonds_calendar_last_event_date=("event_date", "max"),
            cbonds_calendar_placement_volume_rub_sum=("event_issue_volume_rub", "sum"),
            cbonds_calendar_coupon_payment_sum=("event_payment_volume", "sum"),
            cbonds_calendar_principal_repayment_sum=("event_principal_repayment", "sum"),
        )
        .reset_index()
    )
    counts = type_counts(events, ["company_id"], "cbonds_calendar", "_count")
    out = base.merge(grouped, on="company_id", how="left").merge(counts, on="company_id", how="left")
    out["cbonds_calendar_found_flag"] = out["cbonds_calendar_event_count"].fillna(0).gt(0)
    for column in STATIC_EVENT_COUNT_COLUMNS:
        if column in out.columns:
            out[column] = pd.to_numeric(out[column], errors="coerce").fillna(0)
    return out


def build_company_quarterly(mapped_events: pd.DataFrame) -> pd.DataFrame:
    if mapped_events.empty:
        return pd.DataFrame(columns=["company_id", "quarter_label"] + QUARTERLY_EVENT_COLUMNS)

    events = mapped_events.loc[
        mapped_events["event_mapping_status"].eq("mapped")
        & mapped_events["quarter_label"].astype(str).str.match(r"^\d{4}Q[1-4]$")
    ].copy()
    if events.empty:
        return pd.DataFrame(columns=["company_id", "quarter_label"] + QUARTERLY_EVENT_COLUMNS)

    grouped = (
        events.groupby(["company_id", "quarter_label"], dropna=False)
        .agg(
            cbonds_calendar_event_q_count=("event_date", "size"),
            cbonds_calendar_issue_q_count=("event_issue_key", pd.Series.nunique),
            cbonds_calendar_placement_volume_rub_q=("event_issue_volume_rub", "sum"),
            cbonds_calendar_coupon_payment_sum_q=("event_payment_volume", "sum"),
            cbonds_calendar_principal_repayment_sum_q=("event_principal_repayment", "sum"),
        )
        .reset_index()
    )
    counts = type_counts(events, ["company_id", "quarter_label"], "cbonds_calendar", "_q_count")
    out = grouped.merge(counts, on=["company_id", "quarter_label"], how="left")
    for column in QUARTERLY_EVENT_COLUMNS:
        if column not in out.columns:
            out[column] = 0
        out[column] = pd.to_numeric(out[column], errors="coerce").fillna(0)
    return out[["company_id", "quarter_label"] + QUARTERLY_EVENT_COLUMNS]


def build_issue_coverage(mapped_events: pd.DataFrame) -> pd.DataFrame:
    if mapped_events.empty:
        return pd.DataFrame()

    issue_coverage = (
        mapped_events.groupby(["event_isin", "event_issue_name", "event_issuer", "company_id"], dropna=False)
        .agg(
            event_count=("event_date", "size"),
            first_event_date=("event_date", "min"),
            last_event_date=("event_date", "max"),
            event_types=("event_type", join_unique),
            event_type_groups=("event_type_group", join_unique),
            calendar_source_sheets=("calendar_source_sheet", join_unique),
            has_validated_current_issue_card_flag=("has_validated_current_issue_card_flag", "max"),
            has_any_issue_card_flag=("has_any_issue_card_flag", "max"),
            cbonds_issue_scope=("cbonds_issue_scope", first_notna),
            card_company_id=("card_company_id", first_notna),
            card_company_name=("card_company_name", first_notna),
            card_issue_name=("card_issue_name", first_notna),
            validated_relation_label=("validated_relation_label", first_notna),
            has_validated_quote_row_flag=("has_validated_quote_row_flag", "max"),
            quote_scope=("quote_scope", first_notna),
            event_mapping_status=("event_mapping_status", first_notna),
            event_mapping_source=("event_mapping_source", first_notna),
            event_issue_volume=("event_issue_volume", "max"),
            event_circulation_volume=("event_circulation_volume", "max"),
            event_currency=("event_currency", first_notna),
        )
        .reset_index()
    )
    issue_coverage["needs_issue_card_check_flag"] = ~issue_coverage["has_validated_current_issue_card_flag"].fillna(False).astype(bool)
    issue_coverage["needs_quote_check_flag"] = ~issue_coverage["has_validated_quote_row_flag"].fillna(False).astype(bool)
    issue_coverage["event_isin_missing_flag"] = issue_coverage["event_isin"].astype(str).str.len().eq(0)
    return issue_coverage.sort_values(["event_mapping_status", "company_id", "event_issuer", "event_issue_name"], kind="stable")


def build_company_coverage(
    companies_master: pd.DataFrame,
    company_summary: pd.DataFrame,
    cbonds_workbook: Path,
) -> pd.DataFrame:
    master = companies_master.copy()
    master["company_id"] = master["sector_company_id"]
    selected = [
        "company_id",
        "sector",
        "sample_flag",
        "company_name",
        "company_name_short",
        "company_name_full",
        "inn",
        "ticker",
        "manual_inclusion_flag",
        "manual_inclusion_reason",
    ]
    selected = [column for column in selected if column in master.columns]
    out = master[selected].copy()
    out = out.merge(company_summary, on="company_id", how="left")

    cbonds_summary = safe_read_excel(cbonds_workbook, "company_bond_summary_validated")
    if not cbonds_summary.empty:
        columns = [
            "company_id",
            "cbonds_issue_card_found_flag",
            "cbonds_issue_count",
            "cbonds_validated_direct_issue_count",
            "cbonds_validated_direct_issue_flag",
            "cbonds_any_nonexcluded_issue_flag",
            "cbonds_missing_strong_issue_count",
        ]
        columns = [column for column in columns if column in cbonds_summary.columns]
        out = out.merge(cbonds_summary[columns], on="company_id", how="left")

    for column in [
        "cbonds_calendar_event_count",
        "cbonds_calendar_issue_count",
        "cbonds_issue_count",
        "cbonds_validated_direct_issue_count",
        "cbonds_missing_strong_issue_count",
    ]:
        if column in out.columns:
            out[column] = pd.to_numeric(out[column], errors="coerce").fillna(0)

    for column in [
        "cbonds_calendar_found_flag",
        "cbonds_issue_card_found_flag",
        "cbonds_validated_direct_issue_flag",
        "cbonds_any_nonexcluded_issue_flag",
    ]:
        if column in out.columns:
            out[column] = out[column].fillna(False).astype(bool)

    out["calendar_missing_flag"] = ~out["cbonds_calendar_found_flag"].fillna(False).astype(bool)
    out["calendar_missing_suspicious_flag"] = (
        out["calendar_missing_flag"]
        & (
            out.get("cbonds_validated_direct_issue_count", 0).fillna(0).astype(float).gt(0)
            | out.get("cbonds_missing_strong_issue_count", 0).fillna(0).astype(float).gt(0)
        )
    )
    out["calendar_missing_reason"] = np.select(
        [
            out["cbonds_calendar_found_flag"].fillna(False).astype(bool),
            out.get("cbonds_validated_direct_issue_count", 0).fillna(0).astype(float).gt(0),
            out.get("cbonds_missing_strong_issue_count", 0).fillna(0).astype(float).gt(0),
        ],
        [
            "calendar_events_found",
            "validated_cbonds_issue_cards_exist_but_calendar_absent",
            "strong_bond_candidates_exist_but_calendar_absent",
        ],
        default="no_calendar_and_no_validated_bond_evidence",
    )
    return out.sort_values(["calendar_missing_suspicious_flag", "sector", "company_name"], ascending=[False, True, True], kind="stable")


def build_cbonds_event_layers(
    calendar_workbook: Path | None,
    companies_master: pd.DataFrame,
    cbonds_workbook: Path = DEFAULT_CBONDS_WORKBOOK,
    quotes_workbook: Path | None = DEFAULT_QUOTES_WORKBOOK,
) -> dict[str, pd.DataFrame]:
    events, parse_log = load_raw_event_calendar(calendar_workbook)
    issue_scope = read_issue_scope(cbonds_workbook) if cbonds_workbook.exists() else pd.DataFrame()
    quote_scope = read_quote_scope(quotes_workbook)
    mapped_events = map_events_to_companies(events, companies_master, issue_scope, quote_scope)
    company_summary = build_company_summary(companies_master, mapped_events)
    company_quarterly = build_company_quarterly(mapped_events)
    issue_coverage = build_issue_coverage(mapped_events)
    company_coverage = build_company_coverage(companies_master, company_summary, cbonds_workbook)
    missing_suspicious = company_coverage.loc[company_coverage["calendar_missing_suspicious_flag"].fillna(False)].copy()

    return {
        "cbonds_event_raw": mapped_events,
        "cbonds_event_company_summary": company_summary,
        "cbonds_event_company_quarterly": company_quarterly,
        "cbonds_event_issue_coverage": issue_coverage,
        "cbonds_event_company_coverage": company_coverage,
        "cbonds_event_missing": missing_suspicious,
        "cbonds_event_parse_log": parse_log,
    }


def write_event_workbooks(
    layers: dict[str, pd.DataFrame],
    processed_output: Path = DEFAULT_EVENT_OUTPUT,
    missing_output: Path = DEFAULT_EVENT_MISSING_OUTPUT,
) -> None:
    processed_output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(processed_output, engine="openpyxl") as writer:
        for sheet_name, frame in layers.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    with pd.ExcelWriter(missing_output, engine="openpyxl") as writer:
        for sheet_name in [
            "cbonds_event_missing",
            "cbonds_event_company_coverage",
            "cbonds_event_issue_coverage",
            "cbonds_event_parse_log",
        ]:
            layers.get(sheet_name, pd.DataFrame()).to_excel(writer, sheet_name=sheet_name[:31], index=False)


def main() -> None:
    import argparse

    from .build_analysis_panel import DEFAULT_METALLURGY_SHORTLIST, DEFAULT_OIL_GAS_SHORTLIST, combine_shortlists

    parser = argparse.ArgumentParser(description="Parse Cbonds event calendar and build company/issue coverage layers.")
    parser.add_argument("--calendar", type=Path, default=DEFAULT_EVENT_CALENDAR)
    parser.add_argument("--cbonds", type=Path, default=DEFAULT_CBONDS_WORKBOOK)
    parser.add_argument("--quotes", type=Path, default=DEFAULT_QUOTES_WORKBOOK)
    parser.add_argument("--processed-output", type=Path, default=DEFAULT_EVENT_OUTPUT)
    parser.add_argument("--missing-output", type=Path, default=DEFAULT_EVENT_MISSING_OUTPUT)
    args = parser.parse_args()

    companies_master, _ = combine_shortlists(DEFAULT_OIL_GAS_SHORTLIST, DEFAULT_METALLURGY_SHORTLIST)
    layers = build_cbonds_event_layers(
        calendar_workbook=args.calendar,
        companies_master=companies_master,
        cbonds_workbook=args.cbonds,
        quotes_workbook=args.quotes,
    )
    write_event_workbooks(layers, args.processed_output, args.missing_output)
    print(f"Cbonds event calendar workbook: {args.processed_output}")
    print(f"Cbonds calendar missing-check workbook: {args.missing_output}")
    print(f"Event rows: {len(layers['cbonds_event_raw'])}")
    print(f"Companies with calendar: {int(layers['cbonds_event_company_summary']['cbonds_calendar_found_flag'].sum())}")
    print(f"Suspicious missing-calendar companies: {len(layers['cbonds_event_missing'])}")


if __name__ == "__main__":
    main()
