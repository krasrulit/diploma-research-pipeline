from __future__ import annotations

import time
from pathlib import Path

import numpy as np
import pandas as pd
import requests

from .utils import (
    clean_text,
    company_name_variants,
    create_session,
    dataframe_from_iss_block,
    make_log_entry,
    normalize_identifier,
    normalize_isin,
    normalize_text,
    normalize_ticker,
    request_json,
    save_dataframe_csv,
    strip_legal_form,
)


MOEX_SECURITY_SEARCH_URL = "https://iss.moex.com/iss/securities.json"
MOEX_BOND_DETAIL_URL = "https://iss.moex.com/iss/engines/stock/markets/bonds/securities/{secid}.json"
MOEX_SECURITY_DESCRIPTION_URL = "https://iss.moex.com/iss/securities/{secid}.json"
MOEX_BOND_HISTORY_URL = "https://iss.moex.com/iss/history/engines/stock/markets/bonds/securities/{secid}.json"
MOEX_SHARE_HISTORY_URL = "https://iss.moex.com/iss/history/engines/stock/markets/shares/securities/{secid}.json"


def build_search_terms(company_row: pd.Series) -> list[dict[str, str]]:
    terms: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    candidates = [
        ("isin", normalize_isin(company_row.get("isin", ""))),
        ("ticker", normalize_ticker(company_row.get("ticker", ""))),
        ("inn", normalize_identifier(company_row.get("inn", ""), length=10)),
    ]
    for variant in company_name_variants(company_row):
        candidates.append(("name", clean_text(variant)))

    for term_type, term in candidates:
        key = (term_type, normalize_text(term) if term_type == "name" else term)
        if not term or key in seen:
            continue
        seen.add(key)
        terms.append({"query_type": term_type, "query": term})
    return terms


def fetch_search_candidates(
    session: requests.Session,
    query: str,
    page_size: int = 100,
    max_pages: int = 5,
) -> pd.DataFrame:
    frames = []
    seen_rows = set()

    for page_idx in range(max_pages):
        start = page_idx * page_size
        payload = request_json(
            session,
            MOEX_SECURITY_SEARCH_URL,
            params={
                "q": query,
                "iss.meta": "off",
                "limit": page_size,
                "start": start,
            },
        )
        frame = dataframe_from_iss_block(payload, "securities")
        if frame.empty:
            break

        frame = frame.copy()
        frame["secid"] = frame["secid"].map(normalize_ticker)
        frame["isin"] = frame["isin"].map(normalize_isin)
        frame["emitent_inn"] = frame["emitent_inn"].map(lambda x: normalize_identifier(x, length=10))

        deduped_rows = []
        for row in frame.to_dict(orient="records"):
            key = (row.get("secid"), row.get("isin"), row.get("emitent_inn"), row.get("type"))
            if key in seen_rows:
                continue
            seen_rows.add(key)
            deduped_rows.append(row)

        if not deduped_rows:
            break
        frames.append(pd.DataFrame(deduped_rows))

        if len(frame) < page_size:
            break

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def token_overlap_score(company_name: str, candidate_text: str) -> float:
    company_tokens = set(normalize_text(strip_legal_form(company_name)).split())
    candidate_tokens = set(normalize_text(candidate_text).split())
    if not company_tokens or not candidate_tokens:
        return 0.0
    overlap = company_tokens & candidate_tokens
    return len(overlap) / max(1, len(company_tokens))


def score_candidate(company_row: pd.Series, candidate: pd.Series) -> tuple[float, list[str]]:
    score = 0.0
    reasons: list[str] = []

    company_inn = normalize_identifier(company_row.get("inn", ""), length=10)
    company_ticker = normalize_ticker(company_row.get("ticker", ""))
    company_isin = normalize_isin(company_row.get("isin", ""))

    secid = normalize_ticker(candidate.get("secid", ""))
    isin = normalize_isin(candidate.get("isin", ""))
    emitent_inn = normalize_identifier(candidate.get("emitent_inn", ""), length=10)

    if company_inn and company_inn == emitent_inn:
        score += 250
        reasons.append("emitent_inn_exact")

    if company_ticker and company_ticker == secid:
        score += 180
        reasons.append("ticker_exact")

    if company_isin and company_isin == isin:
        score += 180
        reasons.append("isin_exact")

    name_candidates = [
        candidate.get("emitent_title", ""),
        candidate.get("name", ""),
        candidate.get("shortname", ""),
    ]
    best_overlap = max(
        token_overlap_score(company_row.get("company_name", ""), item) for item in name_candidates
    )
    if best_overlap > 0:
        score += round(best_overlap * 120, 2)
        reasons.append(f"name_overlap:{best_overlap:.2f}")

    group = normalize_text(candidate.get("group", ""))
    type_name = normalize_text(candidate.get("type", ""))
    if "bond" in group or "bond" in type_name or "облигац" in type_name:
        score += 15
        reasons.append("bond_type")

    if "stock_shares" == group or "share" in type_name:
        score += 10
        reasons.append("share_type")

    return score, reasons


def confidence_from_score(score: float) -> str:
    if score >= 250:
        return "high"
    if score >= 150:
        return "medium"
    return "low"


def fetch_bond_detail(
    session: requests.Session,
    secid: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    payload = request_json(
        session,
        MOEX_BOND_DETAIL_URL.format(secid=secid),
        params={"iss.meta": "off"},
    )
    security_df = dataframe_from_iss_block(payload, "securities")
    marketdata_df = dataframe_from_iss_block(payload, "marketdata")
    yields_df = dataframe_from_iss_block(payload, "marketdata_yields")

    description_payload = request_json(
        session,
        MOEX_SECURITY_DESCRIPTION_URL.format(secid=secid),
        params={"iss.meta": "off"},
    )
    description_df = dataframe_from_iss_block(description_payload, "description")

    description_row: dict[str, object] = {}
    if not description_df.empty:
        for _, row in description_df.iterrows():
            key = clean_text(row.get("name", "")).lower()
            value = row.get("value")
            if key:
                description_row[f"descr_{key}"] = value

    main_row = {}
    if not security_df.empty:
        main_row.update(security_df.iloc[0].to_dict())
    if not marketdata_df.empty:
        market_row = {f"market_{k}": v for k, v in marketdata_df.iloc[0].to_dict().items()}
        main_row.update(market_row)
    if not yields_df.empty:
        yield_row = {f"yield_{k}": v for k, v in yields_df.iloc[0].to_dict().items()}
        main_row.update(yield_row)
    main_row.update(description_row)
    if "SECID" not in main_row and "descr_secid" in main_row:
        main_row["SECID"] = main_row["descr_secid"]
    if "SHORTNAME" not in main_row and "descr_shortname" in main_row:
        main_row["SHORTNAME"] = main_row["descr_shortname"]
    if "ISIN" not in main_row and "descr_isin" in main_row:
        main_row["ISIN"] = main_row["descr_isin"]
    if "MATDATE" not in main_row and "descr_matdate" in main_row:
        main_row["MATDATE"] = main_row["descr_matdate"]
    if "FACEVALUE" not in main_row and "descr_facevalue" in main_row:
        main_row["FACEVALUE"] = main_row["descr_facevalue"]
    if "COUPONFREQUENCY" not in main_row and "descr_couponfrequency" in main_row:
        main_row["COUPONFREQUENCY"] = main_row["descr_couponfrequency"]
    if "SECTYPE" not in main_row and "descr_type" in main_row:
        main_row["SECTYPE"] = main_row["descr_type"]
    if "GROUP" not in main_row and "descr_group" in main_row:
        main_row["GROUP"] = main_row["descr_group"]

    boards_df = dataframe_from_iss_block(description_payload, "boards")
    if not boards_df.empty:
        boards_df["secid"] = secid

    return pd.DataFrame([main_row]), boards_df


def fetch_bond_history(
    session: requests.Session,
    secid: str,
    history_from: str,
    history_to: str,
    page_size: int = 100,
) -> pd.DataFrame:
    frames = []
    for start in range(0, 100000, page_size):
        payload = request_json(
            session,
            MOEX_BOND_HISTORY_URL.format(secid=secid),
            params={
                "iss.meta": "off",
                "from": history_from,
                "till": history_to,
                "limit": page_size,
                "start": start,
            },
        )
        frame = dataframe_from_iss_block(payload, "history")
        if frame.empty:
            break
        frame["SECID"] = secid
        frames.append(frame)
        if len(frame) < page_size:
            break
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def fetch_share_history(
    session: requests.Session,
    secid: str,
    history_from: str,
    history_to: str,
    page_size: int = 100,
) -> pd.DataFrame:
    frames = []
    for start in range(0, 100000, page_size):
        payload = request_json(
            session,
            MOEX_SHARE_HISTORY_URL.format(secid=secid),
            params={
                "iss.meta": "off",
                "from": history_from,
                "till": history_to,
                "limit": page_size,
                "start": start,
            },
        )
        frame = dataframe_from_iss_block(payload, "history")
        if frame.empty:
            break
        frame["SECID"] = secid
        frames.append(frame)
        if len(frame) < page_size:
            break
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def load_moex_data(
    companies_master: pd.DataFrame,
    output_dir: Path,
    raw_dir: Path | None = None,
    session: requests.Session | None = None,
    load_history: bool = False,
    history_from: str = "2013-01-01",
    history_to: str | None = None,
    load_share_history: bool = False,
    share_history_from: str | None = None,
    share_history_to: str | None = None,
    bond_detail_scope: str = "selected",
    bond_history_scope: str = "selected",
    share_history_scope: str = "selected",
) -> dict[str, pd.DataFrame]:
    output_dir.mkdir(parents=True, exist_ok=True)
    if raw_dir:
        raw_dir.mkdir(parents=True, exist_ok=True)

    session = session or create_session()
    mapping_logs: list[dict[str, object]] = []
    download_logs: list[dict[str, object]] = []
    instrument_rows: list[pd.DataFrame] = []
    bond_detail_rows: list[pd.DataFrame] = []
    history_rows: list[pd.DataFrame] = []
    share_history_rows: list[pd.DataFrame] = []

    for _, company_row in companies_master.iterrows():
        search_terms = build_search_terms(company_row)
        company_candidates = []
        company_name = company_row.get("company_name", "")
        company_inn = company_row.get("inn", "")

        for term in search_terms:
            query = term["query"]
            try:
                frame = fetch_search_candidates(session, query=query)
                if not frame.empty:
                    frame["query"] = query
                    frame["query_type"] = term["query_type"]
                    frame["company_name"] = company_name
                    frame["company_inn"] = company_inn
                    company_candidates.append(frame)
                download_logs.append(
                    make_log_entry(
                        source="moex",
                        status="INFO",
                        message="MOEX search completed",
                        company_name=company_name,
                        inn=company_inn,
                        query=query,
                        query_type=term["query_type"],
                        rows=0 if frame.empty else len(frame),
                        url=MOEX_SECURITY_SEARCH_URL,
                    )
                )
            except Exception as exc:
                download_logs.append(
                    make_log_entry(
                        source="moex",
                        status="ERROR",
                        message="MOEX search failed",
                        company_name=company_name,
                        inn=company_inn,
                        query=query,
                        query_type=term["query_type"],
                        url=MOEX_SECURITY_SEARCH_URL,
                        error=str(exc),
                    )
                )
            time.sleep(0.05)

        if company_candidates:
            candidates = pd.concat(company_candidates, ignore_index=True)
            candidates = candidates.drop_duplicates(subset=["secid", "isin", "emitent_inn", "type"], keep="first")
            scores_and_reasons = candidates.apply(lambda row: score_candidate(company_row, row), axis=1)
            candidates["match_score"] = [item[0] for item in scores_and_reasons]
            candidates["match_reasons"] = [",".join(item[1]) for item in scores_and_reasons]
            candidates["confidence"] = candidates["match_score"].map(confidence_from_score)
            candidates["emitent_inn_exact"] = (
                candidates["emitent_inn"].fillna("") == normalize_identifier(company_inn, length=10)
            ) & candidates["emitent_inn"].fillna("").ne("")
            candidates["ticker_exact"] = candidates["secid"].fillna("").eq(
                normalize_ticker(company_row.get("ticker", ""))
            ) & candidates["secid"].fillna("").ne("")
            candidates["isin_exact"] = candidates["isin"].fillna("").eq(
                normalize_isin(company_row.get("isin", ""))
            ) & candidates["isin"].fillna("").ne("")
            candidates["bond_flag"] = candidates["group"].map(normalize_text).str.contains("bond", na=False) | candidates[
                "type"
            ].map(normalize_text).str.contains("bond", na=False)
            candidates["share_flag"] = candidates["group"].map(normalize_text).str.contains("share", na=False) | candidates[
                "type"
            ].map(normalize_text).str.contains("share", na=False)
            candidates["eligible_universe"] = candidates["share_flag"] | candidates["bond_flag"]
            candidates["selected_flag"] = candidates["eligible_universe"] & (
                (candidates["match_score"] >= 150)
                | (
                    candidates["emitent_inn_exact"]
                )
            )
            candidates["exact_identifier_match"] = (
                candidates["emitent_inn_exact"] | candidates["ticker_exact"] | candidates["isin_exact"]
            )
            candidates["history_candidate_flag"] = candidates["eligible_universe"] & (
                candidates["selected_flag"]
                | candidates["exact_identifier_match"]
                | candidates["match_score"].ge(120)
            )
            candidates["source"] = "moex"
            candidates["source_instrument_id"] = candidates["secid"]
            candidates["instrument_kind_source"] = np.where(
                candidates["bond_flag"],
                "bond",
                np.where(candidates["share_flag"], "share", ""),
            )
            candidates["sample_flag"] = company_row.get("sample_flag", "")
            candidates["sample_membership"] = company_row.get("sample_membership", "")
            instrument_rows.append(candidates)

            selected = candidates.loc[candidates["selected_flag"]].copy()
            selected_bonds = selected.loc[selected["bond_flag"]].copy()
            selected_shares = selected.loc[selected["share_flag"] & ~selected["bond_flag"]].copy()
            history_candidate_bonds = candidates.loc[
                candidates["history_candidate_flag"] & candidates["bond_flag"]
            ].copy()
            history_candidate_shares = candidates.loc[
                candidates["history_candidate_flag"] & candidates["share_flag"] & ~candidates["bond_flag"]
            ].copy()

            bond_detail_targets = (
                history_candidate_bonds if bond_detail_scope == "history_candidates" else selected_bonds
            )
            bond_history_targets = (
                history_candidate_bonds if bond_history_scope == "history_candidates" else selected_bonds
            )
            share_history_targets = (
                history_candidate_shares if share_history_scope == "history_candidates" else selected_shares
            )

            if selected.empty:
                top_row = candidates.sort_values("match_score", ascending=False, kind="stable").head(1)
                if not top_row.empty:
                    mapping_logs.append(
                        make_log_entry(
                            source="moex_mapping",
                            status="REVIEW",
                            message="No high-confidence match; manual review recommended",
                            company_name=company_name,
                            inn=company_inn,
                            best_secid=top_row.iloc[0].get("secid", ""),
                            best_score=top_row.iloc[0].get("match_score", np.nan),
                            best_emitent_inn=top_row.iloc[0].get("emitent_inn", ""),
                            best_type=top_row.iloc[0].get("type", ""),
                        )
                    )
            else:
                mapping_logs.append(
                    make_log_entry(
                        source="moex_mapping",
                        status="INFO",
                        message="Company mapped to MOEX instruments",
                        company_name=company_name,
                        inn=company_inn,
                        selected_instruments=len(selected),
                        selected_bonds=len(selected_bonds),
                        best_score=float(selected["match_score"].max()),
                        matched_secids=", ".join(selected["secid"].astype(str).head(10)),
                    )
                )

            for _, bond_row in bond_detail_targets.drop_duplicates(subset=["secid"], keep="first").iterrows():
                secid = bond_row["secid"]
                try:
                    bond_detail_df, boards_df = fetch_bond_detail(session, secid=secid)
                    if not bond_detail_df.empty:
                        for meta_column in [
                            "company_name",
                            "company_inn",
                            "sample_flag",
                            "sample_membership",
                            "matched_via_query",
                            "match_score",
                            "match_reasons",
                            "confidence",
                        ]:
                            if meta_column == "matched_via_query":
                                bond_detail_df[meta_column] = bond_row.get("query", "")
                            else:
                                source_col = meta_column.replace("company_", "")
                                bond_detail_df[meta_column] = bond_row.get(meta_column, company_row.get(source_col, ""))
                        bond_detail_rows.append(bond_detail_df)
                    if (
                        load_history
                        and history_to
                        and secid in set(bond_history_targets["secid"].astype(str))
                    ):
                        history_df = fetch_bond_history(
                            session,
                            secid=secid,
                            history_from=history_from,
                            history_to=history_to,
                        )
                        if not history_df.empty:
                            history_df["company_name"] = company_name
                            history_df["company_inn"] = company_inn
                            history_rows.append(history_df)
                    download_logs.append(
                        make_log_entry(
                            source="moex_bonds",
                            status="INFO",
                            message="Bond detail loaded",
                            company_name=company_name,
                            inn=company_inn,
                            secid=secid,
                            url=MOEX_BOND_DETAIL_URL.format(secid=secid),
                        )
                    )
                except Exception as exc:
                    download_logs.append(
                        make_log_entry(
                            source="moex_bonds",
                            status="ERROR",
                            message="Bond detail failed",
                            company_name=company_name,
                            inn=company_inn,
                            secid=secid,
                            url=MOEX_BOND_DETAIL_URL.format(secid=secid),
                            error=str(exc),
                        )
                    )
                time.sleep(0.05)

            if load_share_history:
                share_from = share_history_from or history_from
                share_to = share_history_to or history_to
                if share_to:
                    unique_shares = share_history_targets.drop_duplicates(subset=["secid"], keep="first")
                    for _, share_row in unique_shares.iterrows():
                        secid = share_row["secid"]
                        try:
                            share_history_df = fetch_share_history(
                                session=session,
                                secid=secid,
                                history_from=share_from,
                                history_to=share_to,
                            )
                            if not share_history_df.empty:
                                share_history_df["company_name"] = company_name
                                share_history_df["company_inn"] = company_inn
                                share_history_df["sample_flag"] = company_row.get("sample_flag", "")
                                share_history_df["match_score"] = share_row.get("match_score", np.nan)
                                share_history_df["match_reasons"] = share_row.get("match_reasons", "")
                                share_history_rows.append(share_history_df)
                            download_logs.append(
                                make_log_entry(
                                    source="moex_share_history",
                                    status="INFO",
                                    message="Share history loaded",
                                    company_name=company_name,
                                    inn=company_inn,
                                    secid=secid,
                                    rows=len(share_history_df),
                                    url=MOEX_SHARE_HISTORY_URL.format(secid=secid),
                                )
                            )
                        except Exception as exc:
                            download_logs.append(
                                make_log_entry(
                                    source="moex_share_history",
                                    status="ERROR",
                                    message="Share history failed",
                                    company_name=company_name,
                                    inn=company_inn,
                                    secid=secid,
                                    url=MOEX_SHARE_HISTORY_URL.format(secid=secid),
                                    error=str(exc),
                                )
                            )
                        time.sleep(0.05)

        else:
            mapping_logs.append(
                make_log_entry(
                    source="moex_mapping",
                    status="REVIEW",
                    message="No MOEX candidates found",
                    company_name=company_name,
                    inn=company_inn,
                )
            )

    moex_instruments = pd.concat(instrument_rows, ignore_index=True) if instrument_rows else pd.DataFrame()
    moex_bonds = pd.concat(bond_detail_rows, ignore_index=True) if bond_detail_rows else pd.DataFrame()
    moex_history = pd.concat(history_rows, ignore_index=True) if history_rows else pd.DataFrame()
    moex_share_history = pd.concat(share_history_rows, ignore_index=True) if share_history_rows else pd.DataFrame()
    mapping_log = pd.DataFrame(mapping_logs)
    download_log = pd.DataFrame(download_logs)

    if not moex_instruments.empty:
        save_dataframe_csv(moex_instruments, output_dir / "moex_instruments.csv")
    if not moex_bonds.empty:
        save_dataframe_csv(moex_bonds, output_dir / "moex_bonds.csv")
    if load_history and not moex_history.empty:
        save_dataframe_csv(moex_history, output_dir / "moex_bond_history.csv")
    if load_share_history and not moex_share_history.empty:
        save_dataframe_csv(moex_share_history, output_dir / "moex_share_history.csv")
    if not mapping_log.empty:
        save_dataframe_csv(mapping_log, output_dir / "mapping_log_moex.csv")

    return {
        "moex_instruments": moex_instruments,
        "moex_bonds": moex_bonds,
        "moex_bond_history": moex_history,
        "moex_share_history": moex_share_history,
        "mapping_log": mapping_log,
        "download_log": download_log,
    }
