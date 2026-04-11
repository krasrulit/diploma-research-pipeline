from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile

import numpy as np
import pandas as pd


W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SPARK_REPORT_STEM_RE = re.compile(r"СПАРК[-_]Отчет_(.+)_(\d{10})_\d{8}_\d+$")

TARGET_TABLE_TITLES = {
    "бухгалтерский баланс": "Бухгалтерский баланс",
    "отчет о финансовых результатах": "Отчет о финансовых результатах",
    "отчет о движении денежных средств": "Отчет о движении денежных средств",
}

META_KEY_ALIASES = {
    "отчетный период": "Отчетный период",
    "единица измерения": "Единица измерения",
    "источник данных": "Источник данных",
    "инн": "ИНН",
    "оквэд": "ОКВЭД",
    "форма №1 по окуд": "ОКУД",
    "форма n1 по окуд": "ОКУД",
    "форма №2 по окуд": "ОКУД",
    "форма №4 по окуд": "ОКУД",
}

CORE_METRIC_COLUMNS = [
    "assets_total",
    "ppe",
    "cash",
    "equity",
    "debt_lt",
    "debt_st",
    "revenue",
    "interest_expense",
    "net_income",
    "ebit",
    "cfo",
    "capex",
    "investing_payments",
    "financing_inflows",
    "loans_received",
    "bonds_issued",
    "debt_repayment",
    "interest_paid",
]

CORE_METRIC_CODE_MAP = {
    "Бухгалтерский баланс": {
        "1600": "assets_total",
        "1150": "ppe",
        "1250": "cash",
        "1300": "equity",
        "1410": "debt_lt",
        "1510": "debt_st",
    },
    "Отчет о финансовых результатах": {
        "2110": "revenue",
        "2330": "interest_expense",
        "2400": "net_income",
        "2300": "profit_before_tax",
        "2320": "interest_income",
        "2200": "operating_profit",
    },
    "Отчет о движении денежных средств": {
        "4100": "cfo",
        "4221": "capex",
        "4220": "investing_payments",
        "4310": "financing_inflows",
        "4311": "loans_received",
        "4314": "bonds_issued",
        "4323": "debt_repayment",
        "4123": "interest_paid",
        "4224": "interest_paid",
    },
}

CORE_METRIC_NAME_PATTERNS = [
    (re.compile(r"^баланс \(актив\)$", re.I), "assets_total"),
    (re.compile(r"^основные средства$", re.I), "ppe"),
    (
        re.compile(r"^денежные средства и денежные эквиваленты$", re.I),
        "cash",
    ),
    (re.compile(r"^итого по разделу iii$", re.I), "equity"),
    (re.compile(r"^выручка$", re.I), "revenue"),
    (re.compile(r"^проценты к уплате$", re.I), "interest_expense"),
    (re.compile(r"^чистая прибыль \(убыток\)$", re.I), "net_income"),
    (
        re.compile(r"^сальдо денежных потоков от текущих операций$", re.I),
        "cfo",
    ),
    (
        re.compile(
            r"в связи с приобретением, созданием, модернизацией, реконструкцией",
            re.I,
        ),
        "capex",
    ),
]

SUPPORTED_SPARK_SUFFIXES = {".docx"}


@dataclass
class ParseResult:
    raw_long: pd.DataFrame
    parse_log: pd.DataFrame


def clean_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").replace("\u2009", " ").replace("\u202f", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(value: object) -> str:
    return clean_text(value).lower().replace("ё", "е")


def normalize_code(value: object) -> str:
    code = re.sub(r"\D+", "", clean_text(value))
    return code


def join_unique(values: pd.Series) -> str:
    seen: set[str] = set()
    result: list[str] = []
    for value in values.dropna().astype(str):
        text = clean_text(value)
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return ", ".join(result)


def first_notna(values: pd.Series) -> object:
    for value in values:
        if pd.notna(value):
            return value
    return np.nan


def is_number_like(value: object) -> bool:
    text = clean_text(value)
    if not text:
        return False
    return bool(re.fullmatch(r"\(?-?\d[\d\s,\.]*\)?", text))


def to_number(value: object) -> float:
    text = clean_text(value)
    if text in {"", "-", "—", "–"}:
        return np.nan

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]

    text = text.replace(" ", "").replace(",", ".")
    num = pd.to_numeric(text, errors="coerce")
    if pd.isna(num):
        return np.nan
    if negative:
        return -float(num)
    return float(num)


def parse_filename(path: Path) -> tuple[str, str]:
    match = SPARK_REPORT_STEM_RE.match(path.stem)
    if not match:
        return path.stem, ""
    company = match.group(1).replace("_", " ").strip()
    inn = match.group(2)
    return company, inn


def build_input_inventory(input_dir: Path) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for path in sorted(input_dir.iterdir()):
        if not path.is_file():
            continue
        if not SPARK_REPORT_STEM_RE.match(path.stem):
            continue
        company, inn = parse_filename(path)
        rows.append(
            {
                "source_file": path.name,
                "source_path": str(path.resolve()),
                "source_ext": path.suffix.lower(),
                "supported_flag": path.suffix.lower() in SUPPORTED_SPARK_SUFFIXES,
                "company": company,
                "inn": inn,
            }
        )
    return pd.DataFrame(rows)


def parse_quarter_token(value: str) -> tuple[int, str]:
    text = clean_text(value).upper()
    match = re.fullmatch(r"(\d{4})\s*Q([1-4])", text)
    if not match:
        raise ValueError(f"Некорректный квартальный токен: {value}")
    return int(match.group(1)), f"Q{match.group(2)}"


def quarter_rank(period_type: str) -> int:
    return {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}.get(clean_text(period_type), 0)


def iter_quarters(start_period: str, end_period: str) -> list[tuple[int, str]]:
    start_year, start_quarter = parse_quarter_token(start_period)
    end_year, end_quarter = parse_quarter_token(end_period)
    current_year = start_year
    current_rank = quarter_rank(start_quarter)
    end_rank = quarter_rank(end_quarter)
    periods: list[tuple[int, str]] = []

    while (current_year < end_year) or (current_year == end_year and current_rank <= end_rank):
        period_type = f"Q{current_rank}"
        periods.append((current_year, period_type))
        current_rank += 1
        if current_rank > 4:
            current_rank = 1
            current_year += 1

    return periods


def quarter_label(year: int, period_type: str) -> str:
    return f"{int(year)}{clean_text(period_type)}"


def parse_period_label(label: str) -> tuple[object, object]:
    text = clean_text(label)
    if not text:
        return np.nan, np.nan

    year_match = re.search(r"(20\d{2}|19\d{2})", text)
    year = int(year_match.group(1)) if year_match else np.nan

    norm = normalize_text(text)
    quarter_patterns = [
        (r"\bi\b\s*кв", "Q1"),
        (r"\bii\b\s*кв", "Q2"),
        (r"\biii\b\s*кв", "Q3"),
        (r"\biv\b\s*кв", "Q4"),
        (r"\b1\b\s*кв", "Q1"),
        (r"\b2\b\s*кв", "Q2"),
        (r"\b3\b\s*кв", "Q3"),
        (r"\b4\b\s*кв", "Q4"),
        (r"полугоди", "Q2"),
        (r"\b6\s*мес", "Q2"),
        (r"\b9\s*мес", "Q3"),
        (r"\b1\s*квартал", "Q1"),
        (r"\b2\s*квартал", "Q2"),
        (r"\b3\s*квартал", "Q3"),
        (r"\b4\s*квартал", "Q4"),
    ]
    for pattern, period_type in quarter_patterns:
        if re.search(pattern, norm):
            return year, period_type

    return year, "FY" if year_match else np.nan


def parse_section_heading(heading: str) -> dict[str, str]:
    text = clean_text(heading)
    if not text:
        return {"section_heading": "", "heading_period": "", "heading_source": ""}

    parts = [clean_text(part) for part in text.split(",")]
    result = {
        "section_heading": text,
        "heading_period": "",
        "heading_source": "",
    }
    if parts and normalize_text(parts[0]).startswith("бухгалтерская отчетность"):
        if len(parts) >= 2:
            result["heading_period"] = parts[1]
        if len(parts) >= 3:
            result["heading_source"] = parts[2]
    return result


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def element_text(element: ET.Element, include_nested_tables: bool = True) -> str:
    parts: list[str] = []

    def walk(node: ET.Element) -> None:
        if not include_nested_tables and local_name(node.tag) == "tbl":
            return
        if local_name(node.tag) == "t" and node.text:
            parts.append(node.text)
        for child in list(node):
            walk(child)

    if include_nested_tables:
        walk(element)
    else:
        for child in list(element):
            if local_name(child.tag) == "tbl":
                continue
            walk(child)
    return clean_text("".join(parts))


def iter_body_blocks(body: ET.Element) -> Iterable[dict[str, object]]:
    for block_index, child in enumerate(list(body)):
        tag = local_name(child.tag)
        if tag == "p":
            text = element_text(child)
            if text:
                yield {"block_index": block_index, "kind": "paragraph", "text": text}
        elif tag == "tbl":
            yield {"block_index": block_index, "kind": "table", "element": child}


def table_to_rows(tbl: ET.Element, include_nested_tables: bool = False) -> list[list[str]]:
    rows: list[list[str]] = []
    for tr in tbl.findall("./w:tr", W_NS):
        row: list[str] = []
        for tc in tr.findall("./w:tc", W_NS):
            row.append(element_text(tc, include_nested_tables=include_nested_tables))
        rows.append(row)
    return rows


def pad_rows(rows: list[list[str]]) -> tuple[list[list[str]], int]:
    width = max((len(row) for row in rows), default=0)
    padded = [row + [""] * (width - len(row)) for row in rows]
    return padded, width


def direct_nested_tables(tbl: ET.Element) -> list[ET.Element]:
    nested: list[ET.Element] = []
    for tr in tbl.findall("./w:tr", W_NS):
        for tc in tr.findall("./w:tc", W_NS):
            nested.extend(tc.findall("./w:tbl", W_NS))
    return nested


def canonical_table_title(raw_title: str) -> str:
    return TARGET_TABLE_TITLES.get(normalize_text(raw_title), "")


def detect_outer_table_title(tbl: ET.Element) -> str:
    rows = table_to_rows(tbl, include_nested_tables=False)
    if not rows or not rows[0]:
        return ""
    return canonical_table_title(rows[0][0])


def is_meta_table(rows: list[list[str]]) -> bool:
    flat = " | ".join(
        normalize_text(cell)
        for row in rows[:4]
        for cell in row
        if clean_text(cell)
    )
    hints = [
        "отчетный период",
        "единица измерения",
        "источник данных",
        "инн",
        "оквэд",
    ]
    return sum(hint in flat for hint in hints) >= 2


def canonical_meta_key(key: str) -> str:
    norm = normalize_text(key)
    return META_KEY_ALIASES.get(norm, clean_text(key))


def extract_meta_from_rows(rows: list[list[str]]) -> dict[str, str]:
    meta: dict[str, str] = {}
    for row in rows:
        for idx in range(0, len(row) - 1, 2):
            key = canonical_meta_key(row[idx])
            value = clean_text(row[idx + 1])
            if key and value:
                meta[key] = value
    return meta


def is_header_row(row: list[str]) -> bool:
    if len(row) < 3:
        return False

    norm_cells = [normalize_text(cell) for cell in row]
    joined = " | ".join(norm_cells)
    first_cell = norm_cells[0]
    second_cell = norm_cells[1] if len(norm_cells) > 1 else ""
    value_headers = [cell for cell in norm_cells[2:] if cell]
    has_value_columns = bool(value_headers)
    if not has_value_columns:
        return False

    has_code_axis = any("код" in cell for cell in norm_cells[:2])
    has_metric_axis = "наименование показателя" in joined
    is_code_less_metric_header = first_cell == "наименование показателя" and not second_cell
    balance_header_hints = ("за отчетный период", "на отчетную дату", "на 31 декабря")
    is_balance_value_header = first_cell in {"актив", "пассив"} and any(
        any(hint in header for hint in balance_header_hints) for header in value_headers
    )

    return (has_metric_axis and (has_code_axis or is_code_less_metric_header)) or is_balance_value_header


def looks_like_data_table(rows: list[list[str]]) -> bool:
    if not rows:
        return False
    if len(rows) >= 2 and is_header_row(rows[1]):
        return True
    return any(is_header_row(row) for row in rows[:4])


def find_header_row_idx(rows: list[list[str]]) -> int | None:
    for idx, row in enumerate(rows[:5]):
        if is_header_row(row):
            return idx
    return None


def detect_statement_section(metric_name: str, metric_code: str, value_cells: list[str]) -> bool:
    if metric_code:
        return False

    if sum(is_number_like(value) for value in value_cells if clean_text(value)) == 0:
        return True

    norm = normalize_text(metric_name)
    return norm in {"актив", "пассив", "справочно", "справочно:"}


def standardize_metric(table_title: str, metric_code: str, metric_name: str) -> str | None:
    if table_title in CORE_METRIC_CODE_MAP:
        metric_std = CORE_METRIC_CODE_MAP[table_title].get(metric_code)
        if metric_std:
            return metric_std

    clean_name = clean_text(metric_name)
    for pattern, metric_std in CORE_METRIC_NAME_PATTERNS:
        if pattern.search(clean_name):
            return metric_std
    return None


def build_log_row(
    source_file: str,
    level: str,
    event: str,
    message: str,
    **extra: object,
) -> dict[str, object]:
    row = {
        "source_file": source_file,
        "level": level,
        "event": event,
        "message": message,
    }
    row.update(extra)
    return row


class SparkDocxParser:
    def __init__(self) -> None:
        self.allowed_titles = set(TARGET_TABLE_TITLES.values())

    def load_body(self, path: Path) -> ET.Element:
        with ZipFile(path) as docx_zip:
            xml_bytes = docx_zip.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        body = root.find("w:body", W_NS)
        if body is None:
            raise ValueError("Не найден word/document.xml/w:body")
        return body

    def parse_data_table(
        self,
        source_file: str,
        company: str,
        inn_from_filename: str,
        table_title: str,
        section_heading: str,
        block_index: int,
        meta: dict[str, str],
        data_tbl: ET.Element,
    ) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
        logs: list[dict[str, object]] = []
        rows = table_to_rows(data_tbl)
        padded_rows, width = pad_rows(rows)

        header_idx = find_header_row_idx(padded_rows)
        if header_idx is None:
            logs.append(
                build_log_row(
                    source_file,
                    "WARNING",
                    "header_not_found",
                    "Не удалось определить строку заголовка в таблице данных",
                    block_index=block_index,
                    table_title=table_title,
                    section_heading=section_heading,
                )
            )
            return [], logs

        header_row = padded_rows[header_idx]
        value_headers = [clean_text(col) for col in header_row[2:width]]
        report_period = clean_text(meta.get("Отчетный период", ""))
        if not report_period:
            heading_meta = parse_section_heading(section_heading)
            report_period = clean_text(heading_meta["heading_period"])

        report_year, period_type = parse_period_label(report_period)
        data_source = clean_text(meta.get("Источник данных", ""))
        unit = clean_text(meta.get("Единица измерения", ""))
        okved = clean_text(meta.get("ОКВЭД", ""))
        inn = clean_text(meta.get("ИНН", "")) or inn_from_filename

        out_rows: list[dict[str, object]] = []
        current_section = ""

        for row in padded_rows[header_idx + 1 :]:
            metric_name = clean_text(row[0]) if width >= 1 else ""
            metric_code = normalize_code(row[1]) if width >= 2 else ""
            value_cells = [clean_text(cell) for cell in row[2:width]]

            if not metric_name and not metric_code:
                continue

            if detect_statement_section(metric_name, metric_code, value_cells):
                current_section = metric_name
                continue

            if not metric_name:
                continue

            for value_idx, value_column in enumerate(value_headers):
                raw_value = value_cells[value_idx] if value_idx < len(value_cells) else ""
                out_rows.append(
                    {
                        "source_file": source_file,
                        "company": company,
                        "inn": inn,
                        "report_period": report_period,
                        "report_year": report_year,
                        "period_type": period_type,
                        "table_title": table_title,
                        "section_heading": section_heading,
                        "statement_section": current_section,
                        "data_source": data_source,
                        "unit": unit,
                        "okved": okved,
                        "metric_name": metric_name,
                        "metric_code": metric_code,
                        "value_column": value_column or f"value_col_{value_idx}",
                        "value_col_index": value_idx,
                        "value_raw": raw_value,
                        "value_num": to_number(raw_value),
                    }
                )

        logs.append(
            build_log_row(
                source_file,
                "INFO",
                "statement_parsed",
                "Таблица успешно распознана",
                block_index=block_index,
                table_title=table_title,
                section_heading=section_heading,
                report_period=report_period,
                rows_extracted=len(out_rows),
                data_source=data_source,
            )
        )
        return out_rows, logs

    def parse_file(self, path: Path) -> ParseResult:
        body = self.load_body(path)
        company, inn_from_filename = parse_filename(path)
        source_file = path.name

        raw_rows: list[dict[str, object]] = []
        logs: list[dict[str, object]] = [
            build_log_row(
                source_file,
                "INFO",
                "file_opened",
                "Файл открыт",
                company=company,
                inn=inn_from_filename,
            )
        ]

        last_heading = ""
        parsed_title_counter: Counter[str] = Counter()

        for block in iter_body_blocks(body):
            if block["kind"] == "paragraph":
                last_heading = str(block["text"])
                continue

            tbl = block["element"]
            assert isinstance(tbl, ET.Element)
            block_index = int(block["block_index"])

            table_title = detect_outer_table_title(tbl)
            if table_title not in self.allowed_titles:
                continue

            nested_tables = direct_nested_tables(tbl)
            candidates = nested_tables if nested_tables else [tbl]

            meta: dict[str, str] = {}
            data_tables: list[ET.Element] = []

            for candidate in candidates:
                candidate_rows = table_to_rows(candidate)
                if is_meta_table(candidate_rows):
                    meta.update(extract_meta_from_rows(candidate_rows))
                elif looks_like_data_table(candidate_rows):
                    data_tables.append(candidate)

            if not data_tables and looks_like_data_table(table_to_rows(tbl)):
                data_tables = [tbl]

            if not data_tables:
                logs.append(
                    build_log_row(
                        source_file,
                        "WARNING",
                        "data_table_missing",
                        "Внешняя таблица найдена, но data-table не распознана",
                        block_index=block_index,
                        table_title=table_title,
                        section_heading=last_heading,
                    )
                )
                continue

            if "Отчетный период" not in meta and last_heading:
                heading_meta = parse_section_heading(last_heading)
                if heading_meta["heading_period"]:
                    meta["Отчетный период"] = heading_meta["heading_period"]
                if heading_meta["heading_source"] and "Источник данных" not in meta:
                    meta["Источник данных"] = heading_meta["heading_source"]

            parsed_title_counter[table_title] += 1

            for data_tbl in data_tables:
                parsed_rows, parsed_logs = self.parse_data_table(
                    source_file=source_file,
                    company=company,
                    inn_from_filename=inn_from_filename,
                    table_title=table_title,
                    section_heading=last_heading,
                    block_index=block_index,
                    meta=meta,
                    data_tbl=data_tbl,
                )
                raw_rows.extend(parsed_rows)
                logs.extend(parsed_logs)

        for required_title in self.allowed_titles:
            if parsed_title_counter[required_title] == 0:
                logs.append(
                    build_log_row(
                        source_file,
                        "INFO",
                        "statement_absent",
                        "Форма отсутствует в файле или не представлена для доступных периодов",
                        table_title=required_title,
                    )
                )

        logs.append(
            build_log_row(
                source_file,
                "INFO",
                "file_summary",
                "Файл обработан",
                company=company,
                inn=inn_from_filename,
                raw_rows=len(raw_rows),
                statements_parsed=sum(parsed_title_counter.values()),
                balance_blocks=parsed_title_counter["Бухгалтерский баланс"],
                pnl_blocks=parsed_title_counter["Отчет о финансовых результатах"],
                cashflow_blocks=parsed_title_counter["Отчет о движении денежных средств"],
            )
        )

        raw_long = pd.DataFrame(raw_rows)
        parse_log = pd.DataFrame(logs)
        return ParseResult(raw_long=raw_long, parse_log=parse_log)


def ensure_raw_long_columns(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "table_title",
        "section_heading",
        "statement_section",
        "data_source",
        "unit",
        "okved",
        "metric_name",
        "metric_code",
        "value_column",
        "value_col_index",
        "value_raw",
        "value_num",
    ]
    for column in required_columns:
        if column not in df.columns:
            df[column] = np.nan
    return df[required_columns]


def build_panel_core(raw_long: pd.DataFrame) -> pd.DataFrame:
    if raw_long.empty:
        columns = [
            "source_file",
            "company",
            "inn",
            "report_period",
            "report_year",
            "period_type",
            "data_source",
            "unit",
        ] + CORE_METRIC_COLUMNS
        return pd.DataFrame(columns=columns)

    current_values = raw_long.loc[raw_long["value_col_index"] == 0].copy()
    current_values["metric_std"] = current_values.apply(
        lambda row: standardize_metric(
            table_title=clean_text(row["table_title"]),
            metric_code=normalize_code(row["metric_code"]),
            metric_name=clean_text(row["metric_name"]),
        ),
        axis=1,
    )

    panel_source = current_values.loc[current_values["metric_std"].notna()].copy()
    index_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "data_source",
        "unit",
    ]

    if panel_source.empty:
        return pd.DataFrame(columns=index_columns + CORE_METRIC_COLUMNS)

    panel = (
        panel_source.pivot_table(
            index=index_columns,
            columns="metric_std",
            values="value_num",
            aggfunc="first",
        )
        .reset_index()
    )
    panel.columns.name = None

    for column in ["interest_expense", "interest_income", "profit_before_tax", "operating_profit"]:
        if column not in panel.columns:
            panel[column] = np.nan

    ebit_from_pbt = panel["profit_before_tax"]
    if "interest_expense" in panel.columns:
        ebit_from_pbt = ebit_from_pbt + panel["interest_expense"].fillna(0)
    if "interest_income" in panel.columns:
        ebit_from_pbt = ebit_from_pbt - panel["interest_income"].fillna(0)

    panel["ebit"] = panel.get("ebit", pd.Series(np.nan, index=panel.index))
    panel["ebit"] = panel["ebit"].combine_first(ebit_from_pbt)
    panel["ebit"] = panel["ebit"].combine_first(panel["operating_profit"])

    for column in CORE_METRIC_COLUMNS:
        if column not in panel.columns:
            panel[column] = np.nan

    return panel[index_columns + CORE_METRIC_COLUMNS].sort_values(
        ["company", "report_year", "period_type", "source_file"],
        kind="stable",
    )


def build_coverage(raw_long: pd.DataFrame) -> pd.DataFrame:
    if raw_long.empty:
        return pd.DataFrame(
            columns=[
                "source_file",
                "company",
                "inn",
                "report_period",
                "report_year",
                "period_type",
                "data_source",
                "has_balance",
                "has_pnl",
                "has_cash_flow",
                "n_tables",
                "n_metrics_current",
                "n_values_total",
            ]
        )

    index_columns = [
        "source_file",
        "company",
        "inn",
        "report_period",
        "report_year",
        "period_type",
        "data_source",
    ]

    current_values = raw_long.loc[raw_long["value_col_index"] == 0].copy()
    form_presence = (
        current_values.assign(present=1)
        .pivot_table(
            index=index_columns,
            columns="table_title",
            values="present",
            aggfunc="max",
            fill_value=0,
        )
        .reset_index()
    )
    form_presence.columns.name = None

    rename_map = {
        "Бухгалтерский баланс": "has_balance",
        "Отчет о финансовых результатах": "has_pnl",
        "Отчет о движении денежных средств": "has_cash_flow",
    }
    form_presence = form_presence.rename(columns=rename_map)
    for column in rename_map.values():
        if column not in form_presence.columns:
            form_presence[column] = 0

    counts = (
        raw_long.groupby(index_columns, dropna=False)
        .agg(
            n_values_total=("metric_name", "size"),
            n_tables=("table_title", "nunique"),
        )
        .reset_index()
    )
    current_counts = (
        current_values.groupby(index_columns, dropna=False)
        .agg(n_metrics_current=("metric_name", "size"))
        .reset_index()
    )

    coverage = counts.merge(current_counts, on=index_columns, how="left").merge(
        form_presence,
        on=index_columns,
        how="left",
    )
    coverage["n_metrics_current"] = coverage["n_metrics_current"].fillna(0).astype(int)
    for column in rename_map.values():
        coverage[column] = coverage[column].fillna(0).astype(int)

    return coverage.sort_values(
        ["company", "report_year", "period_type", "source_file"],
        kind="stable",
    )


def normalize_panel_to_quarters(panel_core: pd.DataFrame) -> pd.DataFrame:
    if panel_core.empty:
        columns = [
            "company",
            "inn",
            "report_year",
            "period_type",
            "quarter_label",
            "source_file",
            "source_report_period",
            "source_period_type",
            "data_source",
            "unit",
        ] + CORE_METRIC_COLUMNS
        return pd.DataFrame(columns=columns)

    panel = panel_core.copy()
    panel = panel.loc[panel["report_year"].notna()].copy()
    panel["report_year"] = panel["report_year"].astype(int)
    panel["period_type_quarter"] = panel["period_type"].replace({"FY": "Q4"})
    panel = panel.loc[panel["period_type_quarter"].isin(["Q1", "Q2", "Q3", "Q4"])].copy()
    if panel.empty:
        return pd.DataFrame()

    panel["period_priority"] = np.where(panel["period_type"].eq(panel["period_type_quarter"]), 0, 1)
    panel = panel.sort_values(
        ["company", "report_year", "period_type_quarter", "period_priority", "source_file"],
        kind="stable",
    )

    group_columns = ["company", "inn", "report_year", "period_type_quarter"]
    aggregated = (
        panel.groupby(group_columns, dropna=False)
        .agg(
            source_file=("source_file", join_unique),
            source_report_period=("report_period", join_unique),
            source_period_type=("period_type", join_unique),
            data_source=("data_source", join_unique),
            unit=("unit", join_unique),
            **{column: (column, first_notna) for column in CORE_METRIC_COLUMNS},
        )
        .reset_index()
        .rename(columns={"period_type_quarter": "period_type"})
    )
    aggregated["quarter_label"] = aggregated.apply(
        lambda row: quarter_label(int(row["report_year"]), clean_text(row["period_type"])),
        axis=1,
    )
    return aggregated


def normalize_coverage_to_quarters(coverage: pd.DataFrame) -> pd.DataFrame:
    if coverage.empty:
        columns = [
            "company",
            "inn",
            "report_year",
            "period_type",
            "coverage_source_files",
            "n_tables",
            "n_metrics_current",
            "n_values_total",
            "has_balance",
            "has_pnl",
            "has_cash_flow",
        ]
        return pd.DataFrame(columns=columns)

    coverage_quarter = coverage.copy()
    coverage_quarter = coverage_quarter.loc[coverage_quarter["report_year"].notna()].copy()
    coverage_quarter["report_year"] = coverage_quarter["report_year"].astype(int)
    coverage_quarter["period_type_quarter"] = coverage_quarter["period_type"].replace({"FY": "Q4"})
    coverage_quarter = coverage_quarter.loc[
        coverage_quarter["period_type_quarter"].isin(["Q1", "Q2", "Q3", "Q4"])
    ].copy()
    if coverage_quarter.empty:
        return pd.DataFrame()

    aggregated = (
        coverage_quarter.groupby(["company", "inn", "report_year", "period_type_quarter"], dropna=False)
        .agg(
            coverage_source_files=("source_file", join_unique),
            n_tables=("n_tables", "max"),
            n_metrics_current=("n_metrics_current", "max"),
            n_values_total=("n_values_total", "max"),
            has_balance=("has_balance", "max"),
            has_pnl=("has_pnl", "max"),
            has_cash_flow=("has_cash_flow", "max"),
        )
        .reset_index()
        .rename(columns={"period_type_quarter": "period_type"})
    )
    return aggregated


def build_quarterly_panel(
    panel_core: pd.DataFrame,
    coverage: pd.DataFrame,
    input_inventory: pd.DataFrame,
    start_period: str,
    end_period: str,
) -> pd.DataFrame:
    if input_inventory.empty:
        company_universe = panel_core[["company", "inn"]].drop_duplicates().copy()
    else:
        inventory_for_grid = input_inventory.copy()
        if "supported_flag" in inventory_for_grid.columns:
            supported_inventory = inventory_for_grid.loc[inventory_for_grid["supported_flag"]].copy()
            if not supported_inventory.empty:
                inventory_for_grid = supported_inventory
        company_universe = inventory_for_grid[["company", "inn"]].drop_duplicates().copy()

    if company_universe.empty:
        columns = [
            "company",
            "inn",
            "report_year",
            "period_type",
            "quarter_label",
            "source_file",
            "source_report_period",
            "source_period_type",
            "data_source",
            "unit",
            "coverage_source_files",
            "n_tables",
            "n_metrics_current",
            "n_values_total",
            "has_balance",
            "has_pnl",
            "has_cash_flow",
        ] + CORE_METRIC_COLUMNS
        return pd.DataFrame(columns=columns)

    periods = iter_quarters(start_period=start_period, end_period=end_period)
    grid_rows: list[dict[str, object]] = []
    for _, company_row in company_universe.iterrows():
        for report_year, period_type in periods:
            grid_rows.append(
                {
                    "company": company_row["company"],
                    "inn": company_row["inn"],
                    "report_year": report_year,
                    "period_type": period_type,
                    "quarter_label": quarter_label(report_year, period_type),
                }
            )
    grid = pd.DataFrame(grid_rows)

    normalized_panel = normalize_panel_to_quarters(panel_core)
    normalized_coverage = normalize_coverage_to_quarters(coverage)

    output = grid.merge(
        normalized_panel,
        on=["company", "inn", "report_year", "period_type", "quarter_label"],
        how="left",
    ).merge(
        normalized_coverage,
        on=["company", "inn", "report_year", "period_type"],
        how="left",
    )

    for column in ["has_balance", "has_pnl", "has_cash_flow", "n_tables", "n_metrics_current", "n_values_total"]:
        if column not in output.columns:
            output[column] = np.nan
    for column in ["has_balance", "has_pnl", "has_cash_flow", "n_tables", "n_metrics_current", "n_values_total"]:
        output[column] = output[column].fillna(0).astype(int)
    for column in ["source_file", "source_report_period", "source_period_type", "data_source", "unit", "coverage_source_files"]:
        if column not in output.columns:
            output[column] = ""
        output[column] = output[column].fillna("")

    return output.sort_values(
        ["company", "report_year", "period_type"],
        key=lambda s: s.map(quarter_rank) if s.name == "period_type" else s,
        kind="stable",
    ).reset_index(drop=True)


def build_combined_report(
    docx_folder: Path,
    output_file: Path,
    pattern: str = "СПАРК-Отчет_*.docx",
    quarter_grid_start: str = "2014Q3",
    quarter_grid_end: str = "2025Q4",
) -> dict[str, pd.DataFrame]:
    parser = SparkDocxParser()
    input_inventory = build_input_inventory(docx_folder)
    if pattern in {"СПАРК-Отчет_*.docx", "СПАРК*Отчет_*.docx"}:
        files = sorted(
            path
            for path in docx_folder.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".docx"
            and SPARK_REPORT_STEM_RE.match(path.stem)
        )
    else:
        files = sorted(docx_folder.glob(pattern))
    if not files:
        raise FileNotFoundError(f"Не найдено файлов по шаблону: {pattern}")

    raw_frames: list[pd.DataFrame] = []
    log_frames: list[pd.DataFrame] = []

    for file_path in files:
        try:
            result = parser.parse_file(file_path)
            raw_frames.append(result.raw_long)
            log_frames.append(result.parse_log)
        except Exception as exc:  # pragma: no cover - нужен для batch-run
            log_frames.append(
                pd.DataFrame(
                    [
                        build_log_row(
                            file_path.name,
                            "ERROR",
                            "file_error",
                            str(exc),
                        )
                    ]
                )
            )

    raw_long = pd.concat(raw_frames, ignore_index=True) if raw_frames else pd.DataFrame()
    raw_long = ensure_raw_long_columns(raw_long)
    parse_log = pd.concat(log_frames, ignore_index=True) if log_frames else pd.DataFrame()

    if not input_inventory.empty:
        unsupported = input_inventory.loc[~input_inventory["supported_flag"]].copy()
        if not unsupported.empty:
            unsupported_logs = unsupported.apply(
                lambda row: build_log_row(
                    source_file=row["source_file"],
                    level="INFO",
                    event="file_skipped_unsupported",
                    message="Файл пропущен: формат пока не поддерживается текущим парсером",
                    company=row["company"],
                    inn=row["inn"],
                    source_ext=row["source_ext"],
                ),
                axis=1,
            ).tolist()
            parse_log = pd.concat([parse_log, pd.DataFrame(unsupported_logs)], ignore_index=True)

    panel_core = build_panel_core(raw_long)
    coverage = build_coverage(raw_long)
    panel_quarterly = build_quarterly_panel(
        panel_core=panel_core,
        coverage=coverage,
        input_inventory=input_inventory,
        start_period=quarter_grid_start,
        end_period=quarter_grid_end,
    )

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        input_inventory.to_excel(writer, sheet_name="input_inventory", index=False)
        raw_long.to_excel(writer, sheet_name="raw_long", index=False)
        panel_core.to_excel(writer, sheet_name="panel_core", index=False)
        panel_quarterly.to_excel(writer, sheet_name="panel_quarterly", index=False)
        coverage.to_excel(writer, sheet_name="coverage", index=False)
        parse_log.to_excel(writer, sheet_name="parse_log", index=False)

    return {
        "input_inventory": input_inventory,
        "raw_long": raw_long,
        "panel_core": panel_core,
        "panel_quarterly": panel_quarterly,
        "coverage": coverage,
        "parse_log": parse_log,
    }


def main(argv: list[str] | None = None) -> None:
    cli = argparse.ArgumentParser(
        description="Парсер DOCX-отчетов СПАРК в объединенный Excel.",
    )
    cli.add_argument(
        "--input-dir",
        type=Path,
        default=Path("."),
        help="Папка с DOCX-файлами СПАРК.",
    )
    cli.add_argument(
        "--pattern",
        default="СПАРК-Отчет_*.docx",
        help="Шаблон имен файлов.",
    )
    cli.add_argument(
        "--output",
        type=Path,
        default=Path("spark_combined_report.xlsx"),
        help="Путь к выходному Excel-файлу.",
    )
    cli.add_argument(
        "--quarter-grid-start",
        default="2014Q3",
        help="Начало квартальной сетки для panel_quarterly, например 2014Q3.",
    )
    cli.add_argument(
        "--quarter-grid-end",
        default="2025Q4",
        help="Конец квартальной сетки для panel_quarterly, например 2025Q4.",
    )
    if argv is None and "ipykernel" in sys.modules:
        argv = []

    args = cli.parse_args(argv)

    results = build_combined_report(
        docx_folder=args.input_dir,
        output_file=args.output,
        pattern=args.pattern,
        quarter_grid_start=args.quarter_grid_start,
        quarter_grid_end=args.quarter_grid_end,
    )
    print(f"Готово: {args.output}")
    print(f"Строк input_inventory: {len(results['input_inventory'])}")
    if not results["input_inventory"].empty:
        docx_count = int(results["input_inventory"]["source_ext"].eq(".docx").sum())
        print(f"Файлов docx: {docx_count}")
    print(f"Строк raw_long: {len(results['raw_long'])}")
    print(f"Строк panel_core: {len(results['panel_core'])}")
    print(f"Строк panel_quarterly: {len(results['panel_quarterly'])}")
    print(f"Строк coverage: {len(results['coverage'])}")
    print(f"Строк parse_log: {len(results['parse_log'])}")


if __name__ == "__main__":
    main()
