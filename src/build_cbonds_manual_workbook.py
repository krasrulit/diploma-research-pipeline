from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from .utils import clean_text, ensure_directory, normalize_identifier


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ANALYSIS_PANEL = PROJECT_ROOT / "data_processed" / "analysis_panel.xlsx"
DEFAULT_SECURITIES = PROJECT_ROOT / "data_processed" / "public_securities_all_companies_optimized_lite.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "data_processed" / "cbonds_manual_helper.xlsx"

DATE_FROM = "01.07.2014"
DATE_TO = "31.12.2025"


def _as_text(value: object) -> str:
    text = clean_text(value)
    return text


def _first_nonempty(*values: object) -> str:
    for value in values:
        text = _as_text(value)
        if text:
            return text
    return ""


def _format_inn(value: object) -> str:
    return normalize_identifier(value, length=10)


def _format_ogrn(value: object) -> str:
    return normalize_identifier(value, length=13)


def _priority_bucket(frame: pd.DataFrame) -> pd.Series:
    score = (
        frame["reliable_mapping_flag"].fillna(False).astype(int) * 100
        + frame["usable_history_flag"].fillna(False).astype(int) * 60
        + frame["history_available_flag"].fillna(False).astype(int) * 20
        + pd.to_numeric(frame["best_n_trade_dates"], errors="coerce").fillna(0).clip(upper=5000) / 100
        - frame["manual_review_needed_flag"].fillna(False).astype(int) * 40
    )
    return pd.cut(
        score,
        bins=[-np.inf, 19.999, 79.999, np.inf],
        labels=["low", "medium", "high"],
    ).astype(str)


def build_settings_guide() -> pd.DataFrame:
    rows = [
        {
            "section_order": 1,
            "cbonds_block": "Фин. отчетность организации по РСБУ и МСФО",
            "mode": "ПО ЭМИТЕНТУ",
            "use_now_flag": 1,
            "what_for": "Получать финансовую отчетность эмитента и быстро сверять наличие РСБУ/МСФО.",
            "identifier_field": "Название, ИНН, ОГРН, LEI, ISIN, CUSIP, гос. рег. номер",
            "what_to_paste_primary": "ИНН компании",
            "what_to_paste_fallback_1": "ОГРН компании",
            "what_to_paste_fallback_2": "Полное наименование компании",
            "recommended_check_1": "Бухгалтерский баланс РСБУ (тыс руб.)",
            "recommended_check_2": "Отчет о финансовых результатах РСБУ (тыс руб.)",
            "recommended_check_3": "Отчет о движении денежных средств РСБУ (тыс руб.)",
            "recommended_check_4": "Отчетность МСФО (реальный сектор)",
            "recommended_date_from": DATE_FROM,
            "recommended_date_to": DATE_TO,
            "notes": "МСФО финансового сектора для наших эмитентов обычно не нужно.",
        },
        {
            "section_order": 2,
            "cbonds_block": "Котировки по выпуску",
            "mode": "ПО БУМАГЕ",
            "use_now_flag": 1,
            "what_for": "Получать историю по облигационному выпуску.",
            "identifier_field": "Название / ISIN / CUSIP / Гос. рег. номер",
            "what_to_paste_primary": "ISIN выпуска",
            "what_to_paste_fallback_1": "Название выпуска",
            "what_to_paste_fallback_2": "Гос. рег. номер выпуска, если есть",
            "recommended_check_1": "Биржа: оставить пусто на первом проходе",
            "recommended_check_2": f"Дата с: {DATE_FROM}",
            "recommended_check_3": f"Дата по: {DATE_TO}",
            "recommended_check_4": "",
            "recommended_date_from": DATE_FROM,
            "recommended_date_to": DATE_TO,
            "notes": "Если по ISIN находится несколько записей, сначала сравнивать по ISIN, затем по названию выпуска и эмитенту.",
        },
        {
            "section_order": 3,
            "cbonds_block": "Поиск акций",
            "mode": "Название / Тикер / ISIN",
            "use_now_flag": 1,
            "what_for": "Поиск акций эмитента и базовая сверка идентификаторов.",
            "identifier_field": "Название / Тикер / ISIN",
            "what_to_paste_primary": "ISIN акции, если есть",
            "what_to_paste_fallback_1": "Ticker акции",
            "what_to_paste_fallback_2": "Название бумаги",
            "recommended_check_1": "",
            "recommended_check_2": "",
            "recommended_check_3": "",
            "recommended_check_4": "",
            "recommended_date_from": "",
            "recommended_date_to": "",
            "notes": "Для акций сначала лучше пробовать ticker, если ISIN отсутствует.",
        },
        {
            "section_order": 4,
            "cbonds_block": "Календарь событий",
            "mode": "ПО ОРГАНИЗАЦИИ",
            "use_now_flag": 1,
            "what_for": "Искать размещения, погашения, оферты и другие облигационные события по компании.",
            "identifier_field": "Название организации, ИНН, ОГРН, LEI",
            "what_to_paste_primary": "ИНН компании",
            "what_to_paste_fallback_1": "ОГРН компании",
            "what_to_paste_fallback_2": "Полное наименование компании",
            "recommended_check_1": "Облигации = да",
            "recommended_check_2": "Еврооблигации = да",
            "recommended_check_3": "События: сначала без узкого фильтра",
            "recommended_check_4": "",
            "recommended_date_from": DATE_FROM,
            "recommended_date_to": DATE_TO,
            "notes": "Сначала лучше брать широкий календарь, а потом фильтровать выгрузку в Excel.",
        },
        {
            "section_order": 5,
            "cbonds_block": "Калькулятор / Watchlist / Карты рынка / Индексы",
            "mode": "не использовать на первом проходе",
            "use_now_flag": 0,
            "what_for": "Не обязательны для первичного сбора bond issuance / history / event data.",
            "identifier_field": "",
            "what_to_paste_primary": "",
            "what_to_paste_fallback_1": "",
            "what_to_paste_fallback_2": "",
            "recommended_check_1": "",
            "recommended_check_2": "",
            "recommended_check_3": "",
            "recommended_check_4": "",
            "recommended_date_from": "",
            "recommended_date_to": "",
            "notes": "Можно использовать позже для точечных проверок, но не для массового первого прохода.",
        },
    ]
    return pd.DataFrame(rows)


def build_issuer_queries(companies_master: pd.DataFrame) -> pd.DataFrame:
    out = companies_master.copy()
    out["inn_str"] = out["inn"].map(_format_inn)
    out["ogrn_str"] = out["ogrn"].map(_format_ogrn)
    out["company_name_full"] = out["company_name_full"].map(_as_text)
    out["company_name"] = out["company_name"].map(_as_text)
    out["cbonds_fin_query_primary"] = out["inn_str"]
    out["cbonds_fin_query_secondary"] = out["ogrn_str"]
    out["cbonds_fin_query_fallback"] = out["company_name_full"].where(out["company_name_full"].ne(""), out["company_name"])
    out["cbonds_event_query_primary"] = out["inn_str"]
    out["cbonds_event_query_secondary"] = out["ogrn_str"]
    out["cbonds_event_query_fallback"] = out["cbonds_fin_query_fallback"]
    out["cbonds_date_from"] = DATE_FROM
    out["cbonds_date_to"] = DATE_TO
    out["cbonds_fin_report_types"] = (
        "Баланс РСБУ; ОФР РСБУ; ОДДС РСБУ; МСФО (реальный сектор)"
    )
    out["cbonds_event_filters"] = "Облигации=да; Еврооблигации=да; События=без узкого фильтра"
    keep = [
        "company_id",
        "sector",
        "sample_flag",
        "company_name",
        "company_name_short",
        "company_name_full",
        "inn_str",
        "ogrn_str",
        "ticker",
        "isin",
        "industry",
        "cbonds_fin_query_primary",
        "cbonds_fin_query_secondary",
        "cbonds_fin_query_fallback",
        "cbonds_event_query_primary",
        "cbonds_event_query_secondary",
        "cbonds_event_query_fallback",
        "cbonds_fin_report_types",
        "cbonds_event_filters",
        "cbonds_date_from",
        "cbonds_date_to",
    ]
    return out[keep].sort_values(["sector", "sample_flag", "company_name"], kind="stable").reset_index(drop=True)


def build_security_queries(securities: pd.DataFrame, companies_master: pd.DataFrame, instrument_type: str) -> pd.DataFrame:
    subset = securities.loc[securities["instrument_type"].eq(instrument_type)].copy()
    if subset.empty:
        return subset

    company_meta = companies_master.copy()
    company_meta["inn_str"] = company_meta["inn"].map(_format_inn)
    company_meta["ogrn_str"] = company_meta["ogrn"].map(_format_ogrn)
    company_meta = company_meta.rename(columns={"ticker": "company_ticker", "isin": "company_isin"})
    subset = subset.merge(
        company_meta[["company_id", "company_name_full", "company_name_short", "inn_str", "ogrn_str", "company_ticker", "company_isin"]],
        on="company_id",
        how="left",
    )

    subset["isin"] = subset["isin"].map(_as_text)
    subset["ticker"] = subset["ticker"].map(_as_text)
    subset["instrument_name"] = subset["instrument_name"].map(_as_text)
    subset["company_name"] = subset["company_name"].map(_as_text)

    if instrument_type == "bond":
        subset["cbonds_paper_query_primary"] = subset["isin"]
        subset["cbonds_paper_query_secondary"] = subset["instrument_name"]
        subset["cbonds_paper_query_tertiary"] = subset["ticker"]
        subset["cbonds_block"] = "Котировки по выпуску -> ПО БУМАГЕ"
    else:
        subset["cbonds_paper_query_primary"] = subset["isin"].where(subset["isin"].ne(""), subset["ticker"])
        subset["cbonds_paper_query_secondary"] = subset["ticker"]
        subset["cbonds_paper_query_tertiary"] = subset["instrument_name"]
        subset["cbonds_block"] = "Поиск акций"

    subset["cbonds_issuer_query_primary"] = subset["company_inn"].map(_format_inn)
    subset["cbonds_issuer_query_secondary"] = subset["ogrn_str"]
    subset["cbonds_issuer_query_fallback"] = subset["company_name_full"].where(
        subset["company_name_full"].map(_as_text).ne(""),
        subset["company_name"],
    )
    subset["cbonds_date_from"] = DATE_FROM
    subset["cbonds_date_to"] = DATE_TO
    subset["cbonds_priority_bucket"] = _priority_bucket(subset)
    subset["cbonds_priority_rank"] = (
        subset["cbonds_priority_bucket"].map({"high": 1, "medium": 2, "low": 3}).fillna(9).astype(int)
    )
    subset["cbonds_notes"] = np.select(
        [
            subset["manual_review_needed_flag"].fillna(False),
            subset["reliable_mapping_flag"].fillna(False),
            subset["usable_history_flag"].fillna(False),
        ],
        [
            "Нужна ручная сверка эмитента/бумаги",
            "Хороший кандидат: reliable mapping",
            "Есть usable history, но mapping не fully reliable",
        ],
        default="Использовать как низкоприоритетный резервный запрос",
    )

    keep = [
        "company_id",
        "sector",
        "sample_flag",
        "company_name",
        "company_inn",
        "company_ticker",
        "company_isin",
        "instrument_type",
        "ticker",
        "isin",
        "instrument_name",
        "best_source",
        "best_source_instrument_id",
        "best_n_trade_dates",
        "best_history_start",
        "best_history_end",
        "best_match_score",
        "best_match_confidence",
        "manual_review_needed_flag",
        "history_available_flag",
        "usable_history_flag",
        "reliable_mapping_flag",
        "market_relevance_bucket",
        "cbonds_block",
        "cbonds_paper_query_primary",
        "cbonds_paper_query_secondary",
        "cbonds_paper_query_tertiary",
        "cbonds_issuer_query_primary",
        "cbonds_issuer_query_secondary",
        "cbonds_issuer_query_fallback",
        "cbonds_date_from",
        "cbonds_date_to",
        "cbonds_priority_bucket",
        "cbonds_priority_rank",
        "cbonds_notes",
    ]
    out = subset[keep].sort_values(
        ["cbonds_priority_rank", "company_name", "best_n_trade_dates", "best_match_score"],
        ascending=[True, True, False, False],
        kind="stable",
    )
    return out.reset_index(drop=True)


def build_summary(
    companies_master: pd.DataFrame,
    bond_queries: pd.DataFrame,
    share_queries: pd.DataFrame,
) -> pd.DataFrame:
    rows = [
        {"metric": "companies_total", "value": len(companies_master)},
        {"metric": "companies_oil_gas", "value": int(companies_master["sector"].eq("oil_gas").sum())},
        {"metric": "companies_metallurgy", "value": int(companies_master["sector"].eq("metallurgy").sum())},
        {"metric": "issuer_queries_total", "value": len(companies_master)},
        {"metric": "bond_queries_total", "value": len(bond_queries)},
        {"metric": "share_queries_total", "value": len(share_queries)},
        {"metric": "bond_queries_high_priority", "value": int(bond_queries["cbonds_priority_bucket"].eq("high").sum())},
        {"metric": "share_queries_high_priority", "value": int(share_queries["cbonds_priority_bucket"].eq("high").sum())},
        {"metric": "bond_queries_manual_review", "value": int(bond_queries["manual_review_needed_flag"].fillna(False).sum())},
        {"metric": "share_queries_manual_review", "value": int(share_queries["manual_review_needed_flag"].fillna(False).sum())},
    ]
    return pd.DataFrame(rows)


def build_cbonds_manual_outputs(
    analysis_panel_workbook: Path = DEFAULT_ANALYSIS_PANEL,
    securities_workbook: Path = DEFAULT_SECURITIES,
) -> dict[str, pd.DataFrame]:
    companies_master = pd.read_excel(analysis_panel_workbook, sheet_name="companies_master")
    companies_master["company_id"] = companies_master["sector_company_id"].astype(str).str.strip()
    securities = pd.read_excel(securities_workbook, sheet_name="security_resolution_clean")
    securities["company_id"] = securities["company_id"].astype(str).str.strip()

    issuer_queries = build_issuer_queries(companies_master)
    bond_queries = build_security_queries(securities, companies_master, "bond")
    share_queries = build_security_queries(securities, companies_master, "share")
    security_manual_review = pd.read_excel(securities_workbook, sheet_name="security_manual_review")
    summary = build_summary(companies_master, bond_queries, share_queries)

    return {
        "summary": summary,
        "settings_guide": build_settings_guide(),
        "issuer_queries": issuer_queries,
        "bond_queries": bond_queries,
        "share_queries": share_queries,
        "manual_review": security_manual_review,
    }


def write_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    ensure_directory(path.parent)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build manual Cbonds helper workbook with copy-ready query values.")
    parser.add_argument("--analysis-panel", type=Path, default=DEFAULT_ANALYSIS_PANEL)
    parser.add_argument("--securities", type=Path, default=DEFAULT_SECURITIES)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    outputs = build_cbonds_manual_outputs(
        analysis_panel_workbook=args.analysis_panel,
        securities_workbook=args.securities,
    )
    write_workbook(args.output, outputs)

    print(f"Cbonds helper workbook: {args.output}")
    for sheet_name, frame in outputs.items():
        print(f"{sheet_name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
