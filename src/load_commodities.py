from __future__ import annotations

import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import requests

from .utils import create_session, make_log_entry, normalize_text, request_content, request_text, save_dataframe_csv


WORLD_BANK_COMMODITY_PAGE = "https://www.worldbank.org/en/research/commodity-markets"

TARGET_COMMODITIES_RAW = {
    "crude oil, average": "crude_oil_avg",
    "crude oil, brent": "brent",
    "crude oil, dubai": "dubai_crude",
    "crude oil, wti": "wti",
    "coal, australian": "coal_australia",
    "coal, south african": "coal_south_africa",
    "natural gas, us": "natgas_us",
    "natural gas, europe": "natgas_europe",
    "liquefied natural gas, japan": "lng_japan",
    "aluminum": "aluminum",
    "iron ore, cfr spot": "iron_ore",
    "copper": "copper",
    "lead": "lead",
    "tin": "tin",
    "nickel": "nickel",
    "zinc": "zinc",
    "gold": "gold",
    "platinum": "platinum",
    "silver": "silver",
}

TARGET_COMMODITIES = {
    normalize_text(commodity_name): commodity_code
    for commodity_name, commodity_code in TARGET_COMMODITIES_RAW.items()
}


def resolve_world_bank_file_links(session: requests.Session) -> dict[str, str]:
    html = request_text(session, WORLD_BANK_COMMODITY_PAGE)
    patterns = {
        "monthly": r"https://[^\"']+CMO-Historical-Data-Monthly\.xlsx",
        "annual": r"https://[^\"']+CMO-Historical-Data-Annual\.xlsx",
    }
    links = {}
    for key, pattern in patterns.items():
        matches = re.findall(pattern, html)
        if not matches:
            raise ValueError(f"Не удалось найти ссылку World Bank для {key}")
        links[key] = matches[0]
    return links


def find_header_row(preview: pd.DataFrame) -> int:
    required_labels = {"crude oil brent", "copper"}
    for idx in range(len(preview)):
        values = {normalize_text(val) for val in preview.iloc[idx].tolist()}
        if required_labels.issubset(values):
            return idx
    raise ValueError("Не удалось определить строку заголовка в Pink Sheet")


def melt_commodity_sheet(
    content: bytes,
    sheet_name: str,
    frequency: str,
    source_url: str,
) -> pd.DataFrame:
    preview = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None, nrows=12)
    header_row = find_header_row(preview)
    unit_row = header_row + 1

    raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
    header = raw.iloc[header_row].tolist()
    units = raw.iloc[unit_row].tolist()
    data = raw.iloc[unit_row + 1 :].copy()
    data.columns = header

    first_column = data.columns[0]
    data = data.rename(columns={first_column: "period_raw"})
    data = data.loc[data["period_raw"].notna()].copy()

    melted = data.melt(id_vars=["period_raw"], var_name="commodity_name", value_name="value")
    melted["commodity_name"] = melted["commodity_name"].map(lambda x: str(x).strip() if pd.notna(x) else "")
    melted = melted.loc[melted["commodity_name"].ne("")].copy()
    melted["commodity_key"] = melted["commodity_name"].map(normalize_text)
    melted["commodity_code"] = melted["commodity_key"].map(TARGET_COMMODITIES)
    melted = melted.loc[melted["commodity_code"].notna()].copy()

    unit_map = {
        str(header[idx]).strip(): str(units[idx]).strip()
        for idx in range(len(header))
        if idx < len(units) and pd.notna(header[idx])
    }
    melted["unit"] = melted["commodity_name"].map(unit_map)
    melted["value"] = (
        melted["value"]
        .replace({"…": np.nan, "..": np.nan, "—": np.nan, "-": np.nan})
        .map(lambda x: pd.to_numeric(str(x).replace(",", "."), errors="coerce") if pd.notna(x) else np.nan)
    )

    if frequency == "monthly":
        period_parts = melted["period_raw"].astype(str).str.extract(r"(?P<year>\d{4})M(?P<month>\d{2})")
        melted = melted.loc[period_parts["year"].notna() & period_parts["month"].notna()].copy()
        period_parts = period_parts.loc[melted.index].copy()
        month_stamp = (
            period_parts["year"].astype(int).astype(str)
            + "-"
            + period_parts["month"].astype(int).astype(str).str.zfill(2)
            + "-01"
        )
        melted["period"] = pd.to_datetime(month_stamp, errors="coerce").dt.to_period("M")
        melted["date"] = melted["period"].dt.to_timestamp("M")
        melted["year"] = melted["period"].dt.year
        melted["month"] = melted["period"].dt.month
    else:
        year = pd.to_numeric(melted["period_raw"], errors="coerce")
        melted["year"] = year.astype("Int64")
        melted["date"] = pd.to_datetime(melted["year"].astype(str) + "-12-31", errors="coerce")
        melted["month"] = pd.NA
        melted["period"] = melted["year"].astype(str)

    melted["frequency"] = frequency
    melted["source_url"] = source_url
    return melted[
        [
            "date",
            "period",
            "year",
            "month",
            "commodity_code",
            "commodity_name",
            "unit",
            "value",
            "frequency",
            "source_url",
        ]
    ].sort_values(["commodity_code", "date"], kind="stable")


def load_world_bank_commodities(
    output_dir: Path,
    raw_dir: Path,
    session: requests.Session | None = None,
) -> dict[str, pd.DataFrame]:
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)
    session = session or create_session()

    logs: list[dict[str, object]] = []
    links = resolve_world_bank_file_links(session)

    datasets = {}
    for frequency, url in links.items():
        content = request_content(session, url)
        raw_path = raw_dir / f"world_bank_commodity_{frequency}.xlsx"
        raw_path.write_bytes(content)
        logs.append(
            make_log_entry(
                source="world_bank",
                status="INFO",
                message="Commodity workbook downloaded",
                frequency=frequency,
                url=url,
                saved_path=str(raw_path),
                bytes=len(content),
            )
        )

        if frequency == "monthly":
            frame = melt_commodity_sheet(
                content=content,
                sheet_name="Monthly Prices",
                frequency="monthly",
                source_url=url,
            )
            save_dataframe_csv(frame, output_dir / "commodity_prices_monthly.csv")
            datasets["commodity_prices_monthly"] = frame
        else:
            frame = melt_commodity_sheet(
                content=content,
                sheet_name="Annual Prices (Nominal)",
                frequency="annual",
                source_url=url,
            )
            save_dataframe_csv(frame, output_dir / "commodity_prices_annual.csv")
            datasets["commodity_prices_annual"] = frame

    datasets["download_log"] = pd.DataFrame(logs)
    return datasets
