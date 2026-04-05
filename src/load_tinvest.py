from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests

from .utils import (
    camel_to_snake,
    clean_text,
    company_name_variants,
    create_session,
    json_dumps,
    make_log_entry,
    normalize_isin,
    normalize_text,
    normalize_ticker,
    request_json_post,
    save_dataframe_csv,
    strip_legal_form,
)


TINVEST_TOKEN_ENV = "TINVEST_TOKEN"
DEFAULT_TINVEST_API_URL = "https://invest-public-api.tbank.ru/rest/tinkoff.public.invest.api.contract.v1.InstrumentsService"

TINVEST_SHARES_URL = DEFAULT_TINVEST_API_URL + "/Shares"
TINVEST_BONDS_URL = DEFAULT_TINVEST_API_URL + "/Bonds"
TINVEST_BOND_COUPONS_URL = DEFAULT_TINVEST_API_URL + "/GetBondCoupons"


def get_tinvest_token() -> str:
    return os.getenv(TINVEST_TOKEN_ENV, "").strip()


def build_auth_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def quotation_to_float(value: dict[str, Any]) -> float | None:
    if not isinstance(value, dict):
        return None
    units_raw = value.get("units", 0)
    nano_raw = value.get("nano", 0)
    try:
        units = int(str(units_raw))
        nano = int(str(nano_raw))
    except Exception:
        return None
    return units + nano / 1_000_000_000


def flatten_tinvest_value(prefix: str, value: Any, out: dict[str, Any]) -> None:
    key = camel_to_snake(prefix)
    if isinstance(value, dict):
        normalized_keys = {camel_to_snake(k) for k in value.keys()}
        if normalized_keys.issubset({"units", "nano"}) or normalized_keys.issubset({"currency", "units", "nano"}):
            out[key] = quotation_to_float(value)
            currency = clean_text(value.get("currency", ""))
            if currency:
                out[f"{key}_currency"] = currency
            return

        for nested_key, nested_value in value.items():
            nested_prefix = f"{key}_{camel_to_snake(str(nested_key))}" if key else camel_to_snake(str(nested_key))
            flatten_tinvest_value(nested_prefix, nested_value, out)
        return

    if isinstance(value, list):
        out[key] = json_dumps(value)
        return

    out[key] = value


def flatten_tinvest_record(record: dict[str, Any]) -> dict[str, Any]:
    flattened: dict[str, Any] = {}
    for key, value in record.items():
        flatten_tinvest_value(str(key), value, flattened)
    return flattened


def normalize_instrument_frame(
    records: list[dict[str, Any]],
    instrument_kind: str,
    source_endpoint: str,
) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()

    rows = []
    for record in records:
        row = flatten_tinvest_record(record)
        row["instrument_kind_source"] = instrument_kind
        row["source_endpoint"] = source_endpoint
        row["ticker"] = normalize_ticker(row.get("ticker", ""))
        row["isin"] = normalize_isin(row.get("isin", ""))
        row["figi"] = clean_text(row.get("figi", ""))
        row["uid"] = clean_text(row.get("uid", ""))
        row["position_uid"] = clean_text(row.get("position_uid", ""))
        row["name"] = clean_text(row.get("name", ""))
        row["class_code"] = clean_text(row.get("class_code", ""))
        row["exchange"] = clean_text(row.get("exchange", ""))
        row["currency"] = clean_text(row.get("currency", ""))
        row["sector"] = clean_text(row.get("sector", ""))
        row["country_of_risk"] = clean_text(row.get("country_of_risk", ""))
        row["country_of_risk_name"] = clean_text(row.get("country_of_risk_name", ""))
        row["name_norm"] = normalize_text(row.get("name", ""))
        rows.append(row)

    frame = pd.DataFrame(rows)
    if "instrument_kind" not in frame.columns:
        frame["instrument_kind"] = instrument_kind.upper()
    return frame


def fetch_tinvest_instruments(
    session: requests.Session,
    token: str,
    instrument_kind: str,
) -> pd.DataFrame:
    url = TINVEST_SHARES_URL if instrument_kind == "share" else TINVEST_BONDS_URL
    payload = request_json_post(
        session,
        url,
        json_body={"instrumentStatus": "INSTRUMENT_STATUS_ALL"},
        headers=build_auth_headers(token),
    )
    instruments = payload.get("instruments", [])
    if not isinstance(instruments, list):
        return pd.DataFrame()
    return normalize_instrument_frame(
        records=instruments,
        instrument_kind=instrument_kind,
        source_endpoint=url,
    )


def token_overlap_score(company_name: str, candidate_text: str) -> float:
    company_tokens = set(normalize_text(strip_legal_form(company_name)).split())
    candidate_tokens = set(normalize_text(candidate_text).split())
    if not company_tokens or not candidate_tokens:
        return 0.0
    overlap = company_tokens & candidate_tokens
    return len(overlap) / max(1, len(company_tokens))


def company_variants(company_row: pd.Series) -> list[str]:
    raw_variants = company_row.get("name_variants")
    if isinstance(raw_variants, list) and raw_variants:
        return [clean_text(item) for item in raw_variants if clean_text(item)]
    return company_name_variants(company_row)


def score_tinvest_candidate(company_row: pd.Series, candidate_row: pd.Series) -> tuple[float, float, list[str]]:
    score = 0.0
    reasons: list[str] = []
    variants = company_variants(company_row)

    company_ticker = normalize_ticker(company_row.get("ticker", ""))
    company_isin = normalize_isin(company_row.get("isin", ""))
    candidate_ticker = normalize_ticker(candidate_row.get("ticker", ""))
    candidate_isin = normalize_isin(candidate_row.get("isin", ""))

    if company_isin and company_isin == candidate_isin:
        score += 260
        reasons.append("isin_exact")

    if company_ticker and company_ticker == candidate_ticker:
        score += 210
        reasons.append("ticker_exact")

    name_overlap = max(
        token_overlap_score(variant, candidate_row.get("name", ""))
        for variant in variants
    ) if variants else 0.0
    if name_overlap > 0:
        score += round(name_overlap * 180, 2)
        reasons.append(f"name_overlap:{name_overlap:.2f}")

    return score, name_overlap, reasons


def confidence_from_match(score: float, overlap: float, reasons: list[str]) -> str:
    if "isin_exact" in reasons or "ticker_exact" in reasons or overlap >= 0.9:
        return "high"
    if score >= 140 and overlap >= 0.6:
        return "medium"
    return "low"


def is_selected_candidate(score: float, overlap: float, reasons: list[str]) -> bool:
    if "isin_exact" in reasons or "ticker_exact" in reasons:
        return True
    return overlap >= 0.8 and score >= 140


def match_tinvest_instruments(
    companies_master: pd.DataFrame,
    universe_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if companies_master.empty or universe_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    instrument_rows: list[dict[str, Any]] = []
    mapping_rows: list[dict[str, Any]] = []
    universe_records = universe_df.to_dict(orient="records")

    for _, company_row in companies_master.iterrows():
        company_name = clean_text(company_row.get("company_name", ""))
        company_inn = clean_text(company_row.get("inn", ""))

        company_candidates: list[dict[str, Any]] = []
        for candidate in universe_records:
            score, overlap, reasons = score_tinvest_candidate(company_row, pd.Series(candidate))
            exact_id = "isin_exact" in reasons or "ticker_exact" in reasons
            if score <= 0:
                continue
            if not exact_id and overlap < 0.35:
                continue

            candidate_row = dict(candidate)
            candidate_row["company_name"] = company_name
            candidate_row["company_inn"] = company_inn
            candidate_row["company_ticker"] = normalize_ticker(company_row.get("ticker", ""))
            candidate_row["company_isin"] = normalize_isin(company_row.get("isin", ""))
            candidate_row["sample_flag"] = clean_text(company_row.get("sample_flag", ""))
            candidate_row["match_score"] = score
            candidate_row["name_overlap"] = overlap
            candidate_row["match_reasons"] = json_dumps(reasons)
            candidate_row["match_confidence"] = confidence_from_match(score, overlap, reasons)
            candidate_row["selected_flag"] = is_selected_candidate(score, overlap, reasons)
            candidate_row["exact_identifier_match"] = exact_id
            company_candidates.append(candidate_row)

        company_candidates = sorted(
            company_candidates,
            key=lambda row: (row["selected_flag"], row["match_score"], row.get("name_overlap", 0.0)),
            reverse=True,
        )
        instrument_rows.extend(company_candidates)

        selected = [row for row in company_candidates if row["selected_flag"]]
        best_score = company_candidates[0]["match_score"] if company_candidates else None

        if selected:
            status = "INFO" if any(row["exact_identifier_match"] for row in selected) else "REVIEW"
            message = "T-Invest matches selected"
        elif company_candidates:
            status = "REVIEW"
            message = "T-Invest candidates found but no high-confidence selection"
        else:
            status = "WARN"
            message = "No T-Invest candidates matched"

        mapping_rows.append(
            make_log_entry(
                source="tinvest",
                status=status,
                message=message,
                company_name=company_name,
                inn=company_inn,
                company_ticker=normalize_ticker(company_row.get("ticker", "")),
                company_isin=normalize_isin(company_row.get("isin", "")),
                n_candidates=len(company_candidates),
                n_selected=len(selected),
                best_score=best_score,
                selected_tickers=", ".join(sorted({clean_text(row.get("ticker", "")) for row in selected if clean_text(row.get("ticker", ""))})),
                selected_isins=", ".join(sorted({clean_text(row.get("isin", "")) for row in selected if clean_text(row.get("isin", ""))})),
            )
        )

    instrument_df = pd.DataFrame(instrument_rows)
    mapping_df = pd.DataFrame(mapping_rows)
    if not instrument_df.empty:
        instrument_df = instrument_df.sort_values(
            ["company_name", "selected_flag", "match_score", "instrument_kind_source", "ticker", "isin"],
            ascending=[True, False, False, True, True, True],
            kind="stable",
        ).reset_index(drop=True)
    return instrument_df, mapping_df


def to_utc_timestamp(value: str, hour: int) -> str:
    ts = pd.Timestamp(value)
    if ts.tzinfo is None:
        ts = ts.tz_localize("UTC")
    else:
        ts = ts.tz_convert("UTC")
    ts = ts.normalize() + pd.Timedelta(hours=hour)
    return ts.isoformat().replace("+00:00", "Z")


def fetch_bond_coupons(
    session: requests.Session,
    token: str,
    figi: str,
    coupon_from: str,
    coupon_to: str,
) -> pd.DataFrame:
    payload = request_json_post(
        session,
        TINVEST_BOND_COUPONS_URL,
        json_body={
            "figi": figi,
            "from": to_utc_timestamp(coupon_from, hour=0),
            "to": to_utc_timestamp(coupon_to, hour=23),
        },
        headers=build_auth_headers(token),
    )
    events = payload.get("events", [])
    if not isinstance(events, list) or not events:
        return pd.DataFrame()

    rows = []
    for event in events:
        row = flatten_tinvest_record(event)
        row["figi"] = clean_text(row.get("figi", figi))
        rows.append(row)
    return pd.DataFrame(rows)


def load_tinvest_optional(
    companies_master: pd.DataFrame,
    output_dir: Path,
    session: requests.Session | None = None,
    load_coupons: bool = False,
    coupon_from: str = "2010-01-01",
    coupon_to: str = "2035-12-31",
) -> dict[str, pd.DataFrame]:
    output_dir.mkdir(parents=True, exist_ok=True)
    session = session or create_session()

    token = get_tinvest_token()
    if not token:
        log_df = pd.DataFrame(
            [
                make_log_entry(
                    source="tinvest",
                    status="SKIPPED",
                    message="T-Invest loader skipped: TINVEST_TOKEN is not configured",
                    token_env=TINVEST_TOKEN_ENV,
                )
            ]
        )
        return {
            "tinvest_instruments": pd.DataFrame(),
            "tinvest_bonds": pd.DataFrame(),
            "tinvest_bond_coupons": pd.DataFrame(),
            "mapping_log": pd.DataFrame(),
            "download_log": log_df,
        }

    download_rows: list[dict[str, Any]] = []
    try:
        shares_df = fetch_tinvest_instruments(session, token=token, instrument_kind="share")
        download_rows.append(
            make_log_entry(
                source="tinvest",
                status="INFO",
                message="T-Invest shares loaded",
                rows=len(shares_df),
                endpoint=TINVEST_SHARES_URL,
            )
        )
    except Exception as exc:
        shares_df = pd.DataFrame()
        download_rows.append(
            make_log_entry(
                source="tinvest",
                status="ERROR",
                message="Failed to load T-Invest shares",
                endpoint=TINVEST_SHARES_URL,
                error=str(exc),
            )
        )

    try:
        bonds_df = fetch_tinvest_instruments(session, token=token, instrument_kind="bond")
        download_rows.append(
            make_log_entry(
                source="tinvest",
                status="INFO",
                message="T-Invest bonds loaded",
                rows=len(bonds_df),
                endpoint=TINVEST_BONDS_URL,
            )
        )
    except Exception as exc:
        bonds_df = pd.DataFrame()
        download_rows.append(
            make_log_entry(
                source="tinvest",
                status="ERROR",
                message="Failed to load T-Invest bonds",
                endpoint=TINVEST_BONDS_URL,
                error=str(exc),
            )
        )

    universe_df = pd.concat([shares_df, bonds_df], ignore_index=True, sort=False)
    instruments_df, mapping_df = match_tinvest_instruments(companies_master, universe_df)

    selected_bonds = instruments_df.loc[
        instruments_df["selected_flag"].fillna(False) & instruments_df["instrument_kind_source"].eq("bond")
    ].copy() if not instruments_df.empty else pd.DataFrame()

    if not instruments_df.empty:
        save_dataframe_csv(instruments_df, output_dir / "tinvest_instruments.csv")
    if not selected_bonds.empty:
        save_dataframe_csv(selected_bonds, output_dir / "tinvest_bonds.csv")

    coupon_frames = []
    if load_coupons and not selected_bonds.empty:
        unique_bonds = (
            selected_bonds.sort_values(["company_name", "match_score"], ascending=[True, False], kind="stable")
            .drop_duplicates(subset=["figi"], keep="first")
        )
        for _, bond_row in unique_bonds.iterrows():
            figi = clean_text(bond_row.get("figi", ""))
            if not figi:
                continue
            try:
                coupon_df = fetch_bond_coupons(
                    session=session,
                    token=token,
                    figi=figi,
                    coupon_from=coupon_from,
                    coupon_to=coupon_to,
                )
                if not coupon_df.empty:
                    coupon_df["company_name"] = clean_text(bond_row.get("company_name", ""))
                    coupon_df["company_inn"] = clean_text(bond_row.get("company_inn", ""))
                    coupon_df["bond_name"] = clean_text(bond_row.get("name", ""))
                    coupon_df["bond_ticker"] = clean_text(bond_row.get("ticker", ""))
                    coupon_df["bond_isin"] = clean_text(bond_row.get("isin", ""))
                    coupon_df["bond_uid"] = clean_text(bond_row.get("uid", ""))
                    coupon_frames.append(coupon_df)
                download_rows.append(
                    make_log_entry(
                        source="tinvest",
                        status="INFO",
                        message="T-Invest bond coupons loaded",
                        figi=figi,
                        rows=len(coupon_df),
                        endpoint=TINVEST_BOND_COUPONS_URL,
                    )
                )
            except Exception as exc:
                download_rows.append(
                    make_log_entry(
                        source="tinvest",
                        status="ERROR",
                        message="Failed to load T-Invest bond coupons",
                        figi=figi,
                        endpoint=TINVEST_BOND_COUPONS_URL,
                        error=str(exc),
                    )
                )
            time.sleep(0.05)

    coupons_df = pd.concat(coupon_frames, ignore_index=True, sort=False) if coupon_frames else pd.DataFrame()
    if load_coupons and not coupons_df.empty:
        save_dataframe_csv(coupons_df, output_dir / "tinvest_bond_coupons.csv")

    download_log_df = pd.DataFrame(download_rows)
    return {
        "tinvest_instruments": instruments_df,
        "tinvest_bonds": selected_bonds,
        "tinvest_bond_coupons": coupons_df,
        "mapping_log": mapping_df,
        "download_log": download_log_df,
    }
