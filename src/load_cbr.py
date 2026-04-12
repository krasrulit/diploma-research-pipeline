from __future__ import annotations

import xml.etree.ElementTree as ET
from datetime import datetime
from io import StringIO
from pathlib import Path

import pandas as pd
import requests

from .utils import create_session, make_log_entry, request_text, save_dataframe_csv, to_scaled_percent


CBR_KEY_RATE_URL = "https://www.cbr.ru/hd_base/KeyRate/"
CBR_INFLATION_URL = "https://www.cbr.ru/hd_base/infl/"
CBR_USDRUB_XML_URL = "https://www.cbr.ru/scripts/XML_dynamic.asp"
CBR_USDRUB_DYNAMICS_URL = "https://www.cbr.ru/currency_base/dynamics/"
CBR_USD_VAL_NM_RQ = "R01235"


def to_cbr_xml_date(value: str | None) -> str | None:
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return None


def fetch_key_rate_history(
    session: requests.Session,
    date_from: str = "17.09.2013",
    date_to: str | None = None,
) -> pd.DataFrame:
    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": date_from,
    }
    if date_to:
        params["UniDbQuery.To"] = date_to

    html = request_text(session, CBR_KEY_RATE_URL, params=params)
    table = pd.read_html(StringIO(html))[0]
    table = table.rename(columns={"Дата": "date", "Ставка": "value_raw"})
    table["date"] = pd.to_datetime(table["date"], dayfirst=True, errors="coerce")
    table["value"] = table["value_raw"].map(to_scaled_percent)
    table["series"] = "key_rate"
    table["frequency"] = "daily"
    table["unit"] = "percent"
    table["source_url"] = CBR_KEY_RATE_URL
    return table[["date", "series", "frequency", "value", "unit", "source_url"]].sort_values(
        "date", kind="stable"
    )


def fetch_inflation_and_monthly_key_rate(
    session: requests.Session,
    date_from: str = "17.09.2013",
    date_to: str | None = None,
) -> pd.DataFrame:
    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.From": date_from,
    }
    if date_to:
        params["UniDbQuery.To"] = date_to

    html = request_text(session, CBR_INFLATION_URL, params=params)
    table = pd.read_html(StringIO(html))[0].rename(
        columns={
            "Дата": "period_raw",
            "Ключевая ставка, % годовых": "key_rate_raw",
            "Инфляция, % г/г": "inflation_yoy_raw",
            "Цель по инфляции, %": "inflation_target_raw",
        }
    )
    period_parts = table["period_raw"].astype(str).str.extract(r"(?P<month>\d{1,2})\.(?P<year>\d{4})")
    table = table.loc[period_parts["year"].notna() & period_parts["month"].notna()].copy()
    period_parts = period_parts.loc[table.index].copy()
    month_stamp = (
        period_parts["year"].astype(int).astype(str)
        + "-"
        + period_parts["month"].astype(int).astype(str).str.zfill(2)
        + "-01"
    )
    table["period"] = pd.to_datetime(month_stamp, errors="coerce").dt.to_period("M")
    table["date"] = table["period"].dt.to_timestamp("M")
    table["key_rate_month_end"] = table["key_rate_raw"].map(to_scaled_percent)
    table["inflation_yoy"] = table["inflation_yoy_raw"].map(to_scaled_percent)
    table["inflation_target"] = table["inflation_target_raw"].map(to_scaled_percent)

    melted = table.melt(
        id_vars=["date", "period"],
        value_vars=["key_rate_month_end", "inflation_yoy", "inflation_target"],
        var_name="series",
        value_name="value",
    )
    melted["frequency"] = "monthly"
    melted["unit"] = "percent"
    melted["source_url"] = CBR_INFLATION_URL
    return melted[["date", "period", "series", "frequency", "value", "unit", "source_url"]].sort_values(
        ["date", "series"], kind="stable"
    )


def fetch_usdrub_history(
    session: requests.Session,
    date_from: str = "17.09.2013",
    date_to: str | None = None,
) -> pd.DataFrame:
    normalized_from = to_cbr_xml_date(date_from) or "17/09/2013"
    normalized_to = to_cbr_xml_date(date_to) or pd.Timestamp.utcnow().strftime("%d/%m/%Y")
    params = {
        "date_req1": normalized_from,
        "date_req2": normalized_to,
        "VAL_NM_RQ": CBR_USD_VAL_NM_RQ,
    }

    xml_text = request_text(session, CBR_USDRUB_XML_URL, params=params)
    rows: list[dict[str, object]] = []
    try:
        root = ET.fromstring(xml_text)
        for record in root.findall(".//Record"):
            date_text = record.attrib.get("Date", "")
            nominal_text = record.findtext("Nominal", default="")
            value_text = record.findtext("Value", default="")
            date_value = pd.to_datetime(date_text, dayfirst=True, errors="coerce")
            nominal_value = pd.to_numeric(str(nominal_text).replace(",", "."), errors="coerce")
            rate_value = pd.to_numeric(str(value_text).replace(",", "."), errors="coerce")
            if pd.isna(date_value) or pd.isna(nominal_value) or pd.isna(rate_value) or nominal_value == 0:
                continue
            rows.append(
                {
                    "date": date_value,
                    "series": "usdrub",
                    "frequency": "daily",
                    "value": float(rate_value) / float(nominal_value),
                    "unit": "rub_per_usd",
                    "source_url": CBR_USDRUB_DYNAMICS_URL,
                }
            )
    except ET.ParseError:
        html_text = request_text(
            session,
            CBR_USDRUB_DYNAMICS_URL,
            params={
                "UniDbQuery.Posted": "True",
                "UniDbQuery.so": "1",
                "UniDbQuery.mode": "1",
                "UniDbQuery.VAL_NM_RQ": CBR_USD_VAL_NM_RQ,
                "UniDbQuery.From": normalized_from.replace("/", "."),
                "UniDbQuery.To": normalized_to.replace("/", "."),
            },
        )
        table = pd.read_html(StringIO(html_text))[0]
        table = table.rename(
            columns={
                "Дата": "date",
                "Date": "date",
                "Единиц": "nominal",
                "Unit": "nominal",
                "Курс": "value_raw",
                "Rate": "value_raw",
            }
        )
        table["date"] = pd.to_datetime(table["date"], dayfirst=True, errors="coerce")
        table["nominal"] = pd.to_numeric(table["nominal"], errors="coerce")
        table["value_raw"] = pd.to_numeric(table["value_raw"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
        table["value"] = table["value_raw"] / table["nominal"]
        rows = (
            table.loc[table["date"].notna() & table["value"].notna(), ["date", "value"]]
            .assign(series="usdrub", frequency="daily", unit="rub_per_usd", source_url=CBR_USDRUB_DYNAMICS_URL)
            .to_dict("records")
        )

    return pd.DataFrame(rows).sort_values("date", kind="stable").reset_index(drop=True)


def load_cbr_macro(
    output_dir: Path,
    session: requests.Session | None = None,
    date_from: str = "17.09.2013",
    date_to: str | None = None,
) -> dict[str, pd.DataFrame]:
    output_dir.mkdir(parents=True, exist_ok=True)
    session = session or create_session()
    logs: list[dict[str, object]] = []

    key_rate_daily = fetch_key_rate_history(session, date_from=date_from, date_to=date_to)
    logs.append(
        make_log_entry(
            source="cbr",
            status="INFO",
            message="Daily key rate loaded",
            url=CBR_KEY_RATE_URL,
            rows=len(key_rate_daily),
        )
    )

    monthly_macro = fetch_inflation_and_monthly_key_rate(session, date_from=date_from, date_to=date_to)
    logs.append(
        make_log_entry(
            source="cbr",
            status="INFO",
            message="Monthly inflation/key-rate table loaded",
            url=CBR_INFLATION_URL,
            rows=len(monthly_macro),
        )
    )

    usdrub_daily = fetch_usdrub_history(session, date_from=date_from, date_to=date_to)
    logs.append(
        make_log_entry(
            source="cbr",
            status="INFO",
            message="Daily USDRUB loaded",
            url=CBR_USDRUB_DYNAMICS_URL,
            rows=len(usdrub_daily),
        )
    )

    macro_cbr = pd.concat([key_rate_daily, monthly_macro, usdrub_daily], ignore_index=True, sort=False)
    save_dataframe_csv(macro_cbr, output_dir / "macro_cbr.csv")

    return {
        "macro_cbr": macro_cbr,
        "download_log": pd.DataFrame(logs),
    }
