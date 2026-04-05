from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import requests


REQUEST_TIMEOUT = 60
DEFAULT_USER_AGENT = "diploma-research-pipeline/1.0"

RUS_LEGAL_FORM_PATTERNS = [
    r"\bпубличное акционерное общество\b",
    r"\bакционерное общество\b",
    r"\bнепубличное акционерное общество\b",
    r"\bобщество с ограниченной ответственностью\b",
    r"\bпao\b",
    r"\bпао\b",
    r"\bпао нк\b",
    r"\bоао\b",
    r"\bзао\b",
    r"\bао\b",
    r"\bооо\b",
    r"\bнк\b",
]


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = text.replace("\xa0", " ").replace("\u202f", " ").replace("\u2009", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(value: object) -> str:
    text = clean_text(value).lower().replace("ё", "е")
    text = text.replace('"', " ").replace("«", " ").replace("»", " ")
    text = text.replace("'", " ")
    text = re.sub(r"[^\w\s]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_column_name(value: object) -> str:
    return normalize_text(value)


def normalize_ticker(value: object) -> str:
    text = clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text


def normalize_isin(value: object) -> str:
    text = clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text


def normalize_identifier(value: object, length: int | None = None) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d+(\.0+)?", text):
        text = text.split(".", 1)[0]
    digits = re.sub(r"\D+", "", text)
    if length and digits:
        digits = digits.zfill(length)
    return digits


def strip_legal_form(value: object) -> str:
    text = normalize_text(value)
    for pattern in RUS_LEGAL_FORM_PATTERNS:
        text = re.sub(pattern, " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bим\b", " ", text)
    text = re.sub(r"\bимени\b", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def tokenize_name(value: object) -> list[str]:
    text = strip_legal_form(value)
    tokens = [token for token in text.split() if len(token) > 1]
    return tokens


def company_name_variants(record: pd.Series | dict[str, object]) -> list[str]:
    values = []
    for key in [
        "company_name",
        "company_name_short",
        "company_name_full",
        "company_name_en",
    ]:
        if isinstance(record, dict):
            values.append(record.get(key, ""))
        else:
            values.append(record.get(key, ""))

    variants: list[str] = []
    for value in values:
        raw = clean_text(value)
        core = strip_legal_form(value)
        if raw:
            variants.append(raw)
        if core and core != normalize_text(raw):
            variants.append(core)
    deduped = []
    seen = set()
    for item in variants:
        key = normalize_text(item)
        if key and key not in seen:
            seen.add(key)
            deduped.append(item)
    return deduped


def pick_preferred_value(values: Iterable[object]) -> object:
    for value in values:
        if isinstance(value, float) and np.isnan(value):
            continue
        if clean_text(value):
            return value
    return ""


def to_float(value: object) -> float:
    text = clean_text(value)
    if text in {"", "-", "—", "–", "…"}:
        return np.nan
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = text.replace(" ", "").replace(",", ".")
    result = pd.to_numeric(text, errors="coerce")
    if pd.isna(result):
        return np.nan
    result = float(result)
    return -result if negative else result


def to_scaled_percent(value: object, scale: float = 100.0) -> float:
    number = to_float(value)
    if pd.isna(number):
        return np.nan
    return float(number) / scale


def json_dumps(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


def make_log_entry(
    source: str,
    status: str,
    message: str,
    **extra: object,
) -> dict[str, object]:
    payload = {
        "timestamp_utc": now_utc_iso(),
        "source": source,
        "status": status,
        "message": message,
    }
    payload.update(extra)
    return payload


def create_session(user_agent: str = DEFAULT_USER_AGENT) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": user_agent,
            "Accept": "*/*",
        }
    )
    return session


def request_text(
    session: requests.Session,
    url: str,
    params: dict[str, object] | None = None,
    timeout: int = REQUEST_TIMEOUT,
) -> str:
    response = session.get(url, params=params, timeout=timeout)
    response.raise_for_status()
    return response.text


def request_content(
    session: requests.Session,
    url: str,
    params: dict[str, object] | None = None,
    timeout: int = REQUEST_TIMEOUT,
) -> bytes:
    response = session.get(url, params=params, timeout=timeout)
    response.raise_for_status()
    return response.content


def request_json(
    session: requests.Session,
    url: str,
    params: dict[str, object] | None = None,
    timeout: int = REQUEST_TIMEOUT,
) -> dict[str, object]:
    response = session.get(url, params=params, timeout=timeout)
    response.raise_for_status()
    return response.json()


def dataframe_from_iss_block(payload: dict[str, object], block_name: str) -> pd.DataFrame:
    block = payload.get(block_name, {})
    if not isinstance(block, dict):
        return pd.DataFrame()
    columns = block.get("columns", [])
    data = block.get("data", [])
    if not columns or not data:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(data, columns=columns)


def save_dataframe_csv(df: pd.DataFrame, path: Path) -> Path:
    ensure_directory(path.parent)
    df.to_csv(path, index=False)
    return path


def maybe_limit_rows(df: pd.DataFrame, max_rows: int | None = None) -> pd.DataFrame:
    if max_rows is None or df.empty:
        return df
    return df.head(max_rows).copy()
