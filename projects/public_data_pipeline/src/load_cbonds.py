from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
import requests

from .utils import create_session, make_log_entry, save_dataframe_csv


DEFAULT_CBONDS_BASE = "https://cbonds.com/api/services/json/"


def load_cbonds_optional(
    output_dir: Path,
    session: requests.Session | None = None,
) -> dict[str, pd.DataFrame]:
    output_dir.mkdir(parents=True, exist_ok=True)
    session = session or create_session()

    api_url = os.getenv("CBONDS_API_URL", "").strip()
    api_login = os.getenv("CBONDS_API_LOGIN", "").strip()
    api_password = os.getenv("CBONDS_API_PASSWORD", "").strip()
    endpoint_path = os.getenv("CBONDS_ENDPOINT_PATH", "").strip()

    if not api_url:
        api_url = DEFAULT_CBONDS_BASE

    if not (api_login and api_password and endpoint_path):
        log_df = pd.DataFrame(
            [
                make_log_entry(
                    source="cbonds",
                    status="SKIPPED",
                    message="Cbonds skipped: credentials or endpoint path are not configured",
                    api_url=api_url,
                    endpoint_path=endpoint_path,
                )
            ]
        )
        return {"cbonds_raw": pd.DataFrame(), "download_log": log_df}

    url = api_url.rstrip("/") + "/" + endpoint_path.lstrip("/")
    try:
        response = session.get(url, auth=(api_login, api_password), timeout=60)
        response.raise_for_status()
        payload = response.json()

        if isinstance(payload, dict):
            for key in ["items", "data", "results", "response"]:
                if key in payload and isinstance(payload[key], list):
                    frame = pd.DataFrame(payload[key])
                    break
            else:
                frame = pd.json_normalize(payload)
        elif isinstance(payload, list):
            frame = pd.DataFrame(payload)
        else:
            frame = pd.DataFrame([{"raw_payload": payload}])

        if not frame.empty:
            save_dataframe_csv(frame, output_dir / "cbonds_raw.csv")

        log_df = pd.DataFrame(
            [
                make_log_entry(
                    source="cbonds",
                    status="INFO",
                    message="Cbonds response loaded",
                    url=url,
                    rows=len(frame),
                )
            ]
        )
        return {"cbonds_raw": frame, "download_log": log_df}

    except Exception as exc:
        log_df = pd.DataFrame(
            [
                make_log_entry(
                    source="cbonds",
                    status="ERROR",
                    message="Cbonds request failed",
                    url=url,
                    error=str(exc),
                )
            ]
        )
        return {"cbonds_raw": pd.DataFrame(), "download_log": log_df}
