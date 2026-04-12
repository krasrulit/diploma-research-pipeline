from __future__ import annotations

import argparse
import html
import re
from pathlib import Path
from zipfile import ZipFile

import numpy as np
import pandas as pd

from .build_analysis_panel import (
    DEFAULT_METALLURGY_SHORTLIST,
    DEFAULT_OIL_GAS_SHORTLIST,
    combine_shortlists,
)
from .spark_docx_parser import SPARK_REPORT_STEM_RE, parse_filename
from .utils import clean_text, ensure_directory, normalize_identifier, normalize_text, save_dataframe_csv, strip_legal_form


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_OIL_GAS_REPORTS_DIR = PROJECT_ROOT / "Сборка отчетностей"
DEFAULT_METALLURGY_REPORTS_DIR = PROJECT_ROOT / "Отчетности_металлургия"
DEFAULT_OUTPUT = PROJECT_ROOT / "data_processed" / "ownership_state_table.xlsx"

TAG_RE = re.compile(r"<[^>]+>")
SPACE_RE = re.compile(r"\s+")

FULL_NAME_STOP_TOKENS = [
    "Прежние наименования",
    "Другие наименования",
    "На английском",
    "Адрес",
    "Телефон",
    "E-mail",
    "Web",
    "ОКОПФ",
    "Отрасль",
]

HEAD_COMPANY_STOP_TOKENS = [
    "Размер предприятия",
    "Среднесписочная численность",
    "Уставный капитал",
    "Дочерние компании",
    "Филиалы и представительства",
    "Доходы",
    "Расходы",
    "Выручка от продажи",
    "Чистая прибыль",
    "Чистые активы",
    "Основные средства",
    "Регистрационные данные",
    "Сведения о регистрации",
    "Код СПАРК",
    "ОГРН",
    "ИНН",
    "КПП",
    "ОКПО",
]

GROUP_RULES = [
    {"group_name": "rosneft", "state_owned_flag": 1, "patterns": ["роснефть", "rosneft", "башнефть", "bashneft", "реестр-рн"]},
    {"group_name": "gazprom", "state_owned_flag": 1, "patterns": ["газпром", "gazprom", "газпром нефть", "gazprom-neft", "vostokgazprom", "славнефть"]},
    {"group_name": "transneft", "state_owned_flag": 1, "patterns": ["транснефть", "transneft"]},
    {"group_name": "alrosa", "state_owned_flag": 1, "patterns": ["алроса", "alrosa"]},
    {"group_name": "zarubezhneft", "state_owned_flag": 1, "patterns": ["зарубежнефть", "zarubezhneft"]},
    {"group_name": "tatneft", "state_owned_flag": 1, "patterns": ["татнефть", "tatneft"]},
    {"group_name": "lukoil", "state_owned_flag": 0, "patterns": ["лукойл", "lukoil"]},
    {"group_name": "novatek", "state_owned_flag": 0, "patterns": ["новатэк", "novatek"]},
    {"group_name": "nornickel", "state_owned_flag": 0, "patterns": ["норильский никель", "норникель", "nornickel"]},
    {"group_name": "rusal", "state_owned_flag": 0, "patterns": ["русал", "rusal"]},
    {"group_name": "severstal", "state_owned_flag": 0, "patterns": ["северсталь", "severstal"]},
    {"group_name": "mmk", "state_owned_flag": 0, "patterns": ["ммк", "magnitogorsk iron and steel", "magnitogorsk", "mmk"]},
    {"group_name": "mechel", "state_owned_flag": 0, "patterns": ["мечел", "mechel"]},
    {"group_name": "evraz", "state_owned_flag": 0, "patterns": ["евраз", "evraz"]},
    {"group_name": "nlmk", "state_owned_flag": 0, "patterns": ["нлмк", "nlmk"]},
    {"group_name": "tmk", "state_owned_flag": 0, "patterns": ["тмк", "tmk"]},
    {"group_name": "metalloinvest", "state_owned_flag": 0, "patterns": ["металлоинвест", "metalloinvest"]},
]


def extract_docx_text(path: Path) -> str:
    with ZipFile(path) as archive:
        xml_text = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    return SPACE_RE.sub(" ", html.unescape(TAG_RE.sub(" ", xml_text))).strip()


def extract_labeled_text(text: str, label: str, stop_tokens: list[str]) -> str:
    if not text:
        return ""
    stop_pattern = "|".join(re.escape(token) for token in stop_tokens)
    pattern = re.compile(
        rf"{re.escape(label)}\s+(?P<value>.+?)(?=(?:{stop_pattern})\b|$)",
        flags=re.IGNORECASE,
    )
    match = pattern.search(text)
    if not match:
        return ""
    value = clean_text(match.group("value"))
    value = re.sub(r"^[·\-\s]+", "", value).strip()
    return value


def extract_snippet(text: str, phrase: str, radius: int = 180) -> str:
    lower = text.lower()
    idx = lower.find(phrase.lower())
    if idx == -1:
        return ""
    return clean_text(text[max(0, idx - radius): min(len(text), idx + radius)])


def extract_web_domain(text: str) -> str:
    match = re.search(r"Web\s+([^\s]+)", text, flags=re.IGNORECASE)
    if not match:
        return ""
    raw = clean_text(match.group(1)).rstrip("·,;")
    raw = raw.replace("http://", "").replace("https://", "")
    return raw.strip("/")


def clean_head_company_name(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\s+Реорганизации\s+.+$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+Санкции\s+.+$", "", text, flags=re.IGNORECASE)
    return clean_text(text)


def extract_subsidiaries_count(text: str) -> int:
    match = re.search(r"Дочерние компании\s+(\d+)", text, flags=re.IGNORECASE)
    if not match:
        return 0
    return int(match.group(1))


def parse_report_inventory(report_dir: Path, sector: str) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    if not report_dir.exists():
        return pd.DataFrame()
    for path in sorted(report_dir.iterdir()):
        if not path.is_file():
            continue
        if path.suffix.lower() != ".docx":
            continue
        if not SPARK_REPORT_STEM_RE.match(path.stem):
            continue
        company, inn = parse_filename(path)
        rows.append(
            {
                "sector": sector,
                "source_file": path.name,
                "source_path": str(path.resolve()),
                "company_from_file": company,
                "inn": normalize_identifier(inn, length=10),
                "selected_latest_flag": False,
            }
        )
    inventory = pd.DataFrame(rows)
    if inventory.empty:
        return inventory
    inventory = inventory.sort_values(["sector", "inn", "source_file"], kind="stable").reset_index(drop=True)
    latest_index = inventory.groupby(["sector", "inn"], dropna=False)["source_file"].idxmax()
    inventory.loc[latest_index, "selected_latest_flag"] = True
    return inventory


def infer_group_and_state(
    company_name: str,
    company_name_full: str,
    head_company_name: str,
    web_domain: str,
) -> tuple[str, object, str, str, str]:
    candidates = [
        ("head_company_name", clean_text(head_company_name), "high"),
        ("web_domain", clean_text(web_domain), "high"),
        ("company_name_full", clean_text(company_name_full), "medium"),
        ("company_name", clean_text(company_name), "medium"),
    ]

    for source_name, raw_value, confidence in candidates:
        normalized_value = normalize_text(raw_value)
        if not normalized_value:
            continue
        for rule in GROUP_RULES:
            for pattern in rule["patterns"]:
                if normalize_text(pattern) in normalized_value:
                    state_flag = pd.NA if rule["state_owned_flag"] is None else int(rule["state_owned_flag"])
                    state_bucket = (
                        "unknown"
                        if pd.isna(state_flag)
                        else ("state" if int(state_flag) == 1 else "private")
                    )
                    return rule["group_name"], state_flag, state_bucket, source_name, confidence

    return "", pd.NA, "unknown", "", "low"


def parse_docx_ownership(path: Path) -> dict[str, object]:
    text = extract_docx_text(path)
    full_company_name = extract_labeled_text(text, "Полное наименование", FULL_NAME_STOP_TOKENS)
    head_company_name = clean_head_company_name(extract_labeled_text(text, "Головная компания", HEAD_COMPANY_STOP_TOKENS))
    web_domain = extract_web_domain(text)
    n_subsidiaries = extract_subsidiaries_count(text)
    return {
        "source_file": path.name,
        "source_path": str(path.resolve()),
        "full_company_name_report": full_company_name,
        "web_domain": web_domain,
        "head_company_name": head_company_name,
        "head_company_snippet": extract_snippet(text, "Головная компания"),
        "n_subsidiaries": n_subsidiaries,
        "has_head_company_field_flag": bool(head_company_name),
        "has_subsidiaries_flag": n_subsidiaries > 0,
        "ownership_text_available_flag": bool(head_company_name or n_subsidiaries > 0),
        "text_length": len(text),
        "parse_status": "parsed",
    }


def build_ownership_outputs(
    oil_gas_shortlist: Path = DEFAULT_OIL_GAS_SHORTLIST,
    metallurgy_shortlist: Path = DEFAULT_METALLURGY_SHORTLIST,
    oil_gas_reports_dir: Path = DEFAULT_OIL_GAS_REPORTS_DIR,
    metallurgy_reports_dir: Path = DEFAULT_METALLURGY_REPORTS_DIR,
    max_companies: int | None = None,
) -> dict[str, pd.DataFrame]:
    companies_master, shortlist_log = combine_shortlists(
        oil_gas_shortlist=oil_gas_shortlist,
        metallurgy_shortlist=metallurgy_shortlist,
    )
    if max_companies is not None:
        companies_master = companies_master.head(max_companies).copy()

    inventory = pd.concat(
        [
            parse_report_inventory(oil_gas_reports_dir, "oil_gas"),
            parse_report_inventory(metallurgy_reports_dir, "metallurgy"),
        ],
        ignore_index=True,
        sort=False,
    )
    selected_inventory = inventory.loc[inventory["selected_latest_flag"].fillna(False)].copy()

    parsed_rows: list[dict[str, object]] = []
    parse_logs: list[dict[str, object]] = []
    for _, row in selected_inventory.iterrows():
        path = Path(row["source_path"])
        try:
            parsed = parse_docx_ownership(path)
            parsed["sector"] = row["sector"]
            parsed["inn"] = normalize_identifier(row["inn"], length=10)
            parsed["company_from_file"] = clean_text(row["company_from_file"])
            parsed_rows.append(parsed)
            parse_logs.append(
                {
                    "source_file": row["source_file"],
                    "inn": normalize_identifier(row["inn"], length=10),
                    "sector": row["sector"],
                    "status": "OK",
                    "message": "Ownership metadata parsed",
                }
            )
        except Exception as exc:
            parse_logs.append(
                {
                    "source_file": row["source_file"],
                    "inn": normalize_identifier(row["inn"], length=10),
                    "sector": row["sector"],
                    "status": "ERROR",
                    "message": str(exc),
                }
            )

    parsed_df = pd.DataFrame(parsed_rows)
    parse_log = pd.DataFrame(parse_logs)

    ownership = companies_master.copy()
    ownership["inn"] = ownership["inn"].map(lambda value: normalize_identifier(value, length=10))
    ownership = ownership.merge(
        parsed_df,
        on=["sector", "inn"],
        how="left",
    )
    ownership["report_found_flag"] = ownership["source_file"].notna()
    ownership["n_subsidiaries"] = pd.to_numeric(ownership.get("n_subsidiaries"), errors="coerce").fillna(0).astype(int)
    ownership["has_head_company_field_flag"] = ownership.get("has_head_company_field_flag", False).fillna(False)
    ownership["has_subsidiaries_flag"] = ownership.get("has_subsidiaries_flag", False).fillna(False)
    ownership["is_subsidiary_flag"] = ownership["has_head_company_field_flag"]
    ownership["is_parent_flag"] = ownership["has_subsidiaries_flag"]
    is_subsidiary = ownership["is_subsidiary_flag"].fillna(False).astype(bool)
    is_parent = ownership["is_parent_flag"].fillna(False).astype(bool)
    ownership["ownership_role"] = np.select(
        [
            is_subsidiary & is_parent,
            is_subsidiary,
            is_parent,
        ],
        [
            "parent_and_subsidiary",
            "subsidiary",
            "parent",
        ],
        default="standalone_or_unknown",
    )

    inference_rows = ownership.apply(
        lambda row: infer_group_and_state(
            company_name=clean_text(row.get("company_name", "")),
            company_name_full=clean_text(row.get("company_name_full", "")),
            head_company_name=clean_text(row.get("head_company_name", "")),
            web_domain=clean_text(row.get("web_domain", "")),
        ),
        axis=1,
        result_type="expand",
    )
    inference_rows.columns = [
        "group_name_inferred",
        "state_owned_flag",
        "state_bucket",
        "group_inference_source",
        "group_inference_confidence",
    ]
    ownership = pd.concat([ownership, inference_rows], axis=1)
    ownership["ownership_review_needed_flag"] = (
        ownership["report_found_flag"]
        & (
            ownership["head_company_name"].fillna("").astype(str).ne("")
            | ownership["n_subsidiaries"].gt(0)
            | ownership["group_name_inferred"].fillna("").astype(str).ne("")
        )
        & ownership["group_inference_confidence"].eq("low")
    )

    ownership_links = ownership.loc[
        ownership["head_company_name"].fillna("").astype(str).ne(""),
        [
            "company_name",
            "inn",
            "sector",
            "sample_flag",
            "head_company_name",
            "group_name_inferred",
            "state_owned_flag",
            "state_bucket",
            "group_inference_source",
            "group_inference_confidence",
            "source_file",
        ],
    ].copy().rename(columns={"company_name": "child_company_name", "inn": "child_inn", "head_company_name": "parent_company_name"})

    summary_rows = [
        {"metric": "companies_total", "value": len(ownership)},
        {"metric": "reports_found", "value": int(ownership["report_found_flag"].sum())},
        {"metric": "with_head_company", "value": int(ownership["has_head_company_field_flag"].sum())},
        {"metric": "with_subsidiaries", "value": int(ownership["has_subsidiaries_flag"].sum())},
        {"metric": "subsidiaries_detected", "value": int(ownership["is_subsidiary_flag"].sum())},
        {"metric": "parents_detected", "value": int(ownership["is_parent_flag"].sum())},
        {"metric": "state_owned_inferred", "value": int(pd.to_numeric(ownership["state_owned_flag"], errors="coerce").fillna(0).eq(1).sum())},
        {"metric": "private_owned_inferred", "value": int(pd.to_numeric(ownership["state_owned_flag"], errors="coerce").fillna(-1).eq(0).sum())},
        {"metric": "ownership_review_needed", "value": int(ownership["ownership_review_needed_flag"].sum())},
    ]
    for sector_value in ["oil_gas", "metallurgy"]:
        sector_slice = ownership.loc[ownership["sector"].eq(sector_value)]
        summary_rows.extend(
            [
                {"metric": f"{sector_value}_companies", "value": len(sector_slice)},
                {"metric": f"{sector_value}_reports_found", "value": int(sector_slice["report_found_flag"].sum())},
                {"metric": f"{sector_value}_state_owned_inferred", "value": int(pd.to_numeric(sector_slice["state_owned_flag"], errors="coerce").fillna(0).eq(1).sum())},
            ]
        )
    summary = pd.DataFrame(summary_rows)

    keep_cols = [
        "company_name",
        "company_name_short",
        "company_name_full",
        "inn",
        "ogrn",
        "ticker",
        "isin",
        "spark_id",
        "industry",
        "sector",
        "sample_flag",
        "sample_main_flag",
        "sample_extended_flag",
        "sector_company_id",
        "report_found_flag",
        "source_file",
        "source_path",
        "company_from_file",
        "full_company_name_report",
        "web_domain",
        "head_company_name",
        "n_subsidiaries",
        "has_head_company_field_flag",
        "has_subsidiaries_flag",
        "is_subsidiary_flag",
        "is_parent_flag",
        "ownership_role",
        "group_name_inferred",
        "state_owned_flag",
        "state_bucket",
        "group_inference_source",
        "group_inference_confidence",
        "ownership_review_needed_flag",
        "head_company_snippet",
        "text_length",
        "parse_status",
    ]
    keep_cols = [column for column in keep_cols if column in ownership.columns]
    ownership = ownership[keep_cols].sort_values(["sector", "sample_flag", "company_name"], kind="stable").reset_index(drop=True)

    return {
        "ownership_state_table": ownership,
        "ownership_links": ownership_links.sort_values(["sector", "child_company_name"], kind="stable").reset_index(drop=True),
        "ownership_doc_inventory": inventory.sort_values(["sector", "inn", "source_file"], kind="stable").reset_index(drop=True),
        "ownership_parse_log": parse_log.sort_values(["status", "sector", "inn"], kind="stable").reset_index(drop=True),
        "ownership_summary": summary,
        "shortlist_log": shortlist_log,
    }


def write_ownership_outputs(outputs: dict[str, pd.DataFrame], output_workbook: Path, processed_dir: Path | None = None) -> None:
    ensure_directory(output_workbook.parent)
    with pd.ExcelWriter(output_workbook, engine="openpyxl") as writer:
        for sheet_name, frame in outputs.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    if processed_dir is not None:
        ensure_directory(processed_dir)
        for sheet_name, frame in outputs.items():
            save_dataframe_csv(frame, processed_dir / f"{sheet_name}.csv")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build ownership/state table from raw SPARK DOCX reports.")
    parser.add_argument("--oil-gas-shortlist", type=Path, default=DEFAULT_OIL_GAS_SHORTLIST)
    parser.add_argument("--metallurgy-shortlist", type=Path, default=DEFAULT_METALLURGY_SHORTLIST)
    parser.add_argument("--oil-gas-reports-dir", type=Path, default=DEFAULT_OIL_GAS_REPORTS_DIR)
    parser.add_argument("--metallurgy-reports-dir", type=Path, default=DEFAULT_METALLURGY_REPORTS_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--processed-dir", type=Path, default=PROJECT_ROOT / "data_processed")
    parser.add_argument("--max-companies", type=int, default=None)
    args = parser.parse_args()

    outputs = build_ownership_outputs(
        oil_gas_shortlist=args.oil_gas_shortlist,
        metallurgy_shortlist=args.metallurgy_shortlist,
        oil_gas_reports_dir=args.oil_gas_reports_dir,
        metallurgy_reports_dir=args.metallurgy_reports_dir,
        max_companies=args.max_companies,
    )
    write_ownership_outputs(outputs, output_workbook=args.output, processed_dir=args.processed_dir)

    print(f"Ownership workbook created: {args.output}")
    for sheet_name, frame in outputs.items():
        print(f"{sheet_name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
