#!/usr/bin/env python3
"""Builds browser-ready JS data from local Excel sources."""

from __future__ import annotations

import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
OUT_FILES = (
    ROOT / "mortgage_calculator_data_v3.js",
    ROOT / "mortgage_calculator_data.js",
)
RATES_FILE_NAME = "Ставки 3.xlsx"
DATA_PREFIX = "window.MORTGAGE_CALC_DATA = "
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "doc_rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

ASSORTMENT_DETECT_HEADERS = {"МФЦ", "Объект", "Район", "Номер"}
ASSORTMENT_REQUIRED_HEADERS = [
    "GUID",
    "Район",
    "Объект",
    "Номер",
    "МФЦ",
]
RATES_REQUIRED_HEADERS = [
    "БАНК",
    "Особые условия",
    "Программа",
    "Срок кредита",
    "Срок действия дисконта",
    "Размер дисконта к процентной ставке (в процентных пунктах)",
    "Базовая ставка",
    "Ставка для клиента",
    "Размер ПВ",
]


def to_float(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        txt = value.strip().replace(" ", "").replace(",", ".")
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def to_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", to_str(value))


def parse_min_pv(raw: str) -> float | None:
    match = re.search(r"(\d+(?:[.,]\d+)?)", raw)
    if not match:
        return None
    try:
        return float(match.group(1).replace(",", "."))
    except ValueError:
        return None


def parse_credit_limit(special: str) -> dict[str, float | None]:
    txt = to_str(special).lower().replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)\s*млн", txt)
    if not match:
        return {"loan_min": None, "loan_max": None}
    try:
        amount = float(match.group(1)) * 1_000_000.0
    except ValueError:
        return {"loan_min": None, "loan_max": None}

    if "до" in txt:
        return {"loan_min": None, "loan_max": amount}
    if "более" in txt or "свыше" in txt or "от" in txt:
        return {"loan_min": amount, "loan_max": None}
    return {"loan_min": None, "loan_max": None}


def cell_ref_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    if not letters:
        return 0
    idx = 0
    for char in letters:
        idx = idx * 26 + (ord(char) - 64)
    return idx - 1


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in archive.namelist():
        return []

    root = ET.fromstring(archive.read(path))
    text_tag = f"{{{NS['main']}}}t"
    strings: list[str] = []
    for item in root.findall(f".//{{{NS['main']}}}si"):
        strings.append("".join(node.text or "" for node in item.iter(text_tag)))
    return strings


def resolve_first_sheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    first_sheet = workbook.find(f".//{{{NS['main']}}}sheet")
    if first_sheet is None:
        raise RuntimeError("Workbook does not contain sheets.")

    rel_id = first_sheet.attrib.get(f"{{{NS['doc_rel']}}}id")
    if not rel_id:
        raise RuntimeError("Could not resolve first sheet relationship.")

    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.attrib.get("Id") != rel_id:
            continue
        target = rel.attrib.get("Target", "").lstrip("/")
        if not target:
            break
        return target if target.startswith("xl/") else f"xl/{target}"

    raise RuntimeError("Could not locate first worksheet in workbook.")


def parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{{{NS['main']}}}v")

    if cell_type == "inlineStr":
        text_tag = f"{{{NS['main']}}}t"
        inline_node = cell.find(f"{{{NS['main']}}}is")
        if inline_node is None:
            return ""
        return "".join(node.text or "" for node in inline_node.iter(text_tag))

    if value_node is None:
        return None

    raw = value_node.text or ""
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return ""
    if cell_type == "b":
        return raw == "1"
    return raw


def read_first_sheet_rows(path: Path) -> list[list[Any]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        sheet_path = resolve_first_sheet_path(archive)
        sheet = ET.fromstring(archive.read(sheet_path))

    rows: list[list[Any]] = []
    for row_node in sheet.findall(f".//{{{NS['main']}}}sheetData/{{{NS['main']}}}row"):
        row_values: list[Any] = []
        for cell in row_node.findall(f"{{{NS['main']}}}c"):
            col_idx = cell_ref_to_index(cell.attrib.get("r", ""))
            while len(row_values) < col_idx:
                row_values.append(None)
            row_values.append(parse_cell_value(cell, shared_strings))
        rows.append(row_values)
    return rows


def build_header_index(header_row: list[Any]) -> dict[str, int]:
    idx: dict[str, int] = {}
    for pos, value in enumerate(header_row):
        key = normalize_header(value)
        if key and key not in idx:
            idx[key] = pos
    return idx


def get_cell(row: list[Any], idx: dict[str, int], column: str) -> Any:
    pos = idx.get(column)
    if pos is None or pos >= len(row):
        return None
    return row[pos]


def load_existing_payload() -> dict[str, Any]:
    for out_file in OUT_FILES:
        if not out_file.exists():
            continue
        text = out_file.read_text(encoding="utf-8")
        if not text.startswith(DATA_PREFIX):
            continue
        raw = text[len(DATA_PREFIX):].strip()
        if raw.endswith(";"):
            raw = raw[:-1]
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict):
            return data
    return {}


def pick_workbooks() -> tuple[Path | None, Path]:
    rates = ROOT / RATES_FILE_NAME
    if not rates.exists():
        raise RuntimeError(
            f"Rates file '{RATES_FILE_NAME}' not found. "
            "Calculator is configured to use only this file."
        )

    assortment: Path | None = None
    for xlsx in sorted(ROOT.glob("*.xlsx")):
        if xlsx.resolve() == rates.resolve():
            continue
        rows = read_first_sheet_rows(xlsx)
        if not rows:
            continue
        header_set = {normalize_header(cell) for cell in rows[0] if normalize_header(cell)}
        if ASSORTMENT_DETECT_HEADERS.issubset(header_set):
            assortment = xlsx
            break

    return assortment, rates


def format_apartment(value: Any) -> str:
    number = to_float(value)
    if number is not None and number.is_integer():
        return str(int(number))
    return to_str(value)


def build_assortment(path: Path) -> list[dict[str, Any]]:
    rows = read_first_sheet_rows(path)
    if not rows:
        raise RuntimeError("Assortment workbook is empty.")

    idx = build_header_index(rows[0])
    missing = [col for col in ASSORTMENT_REQUIRED_HEADERS if col not in idx]
    if missing:
        raise RuntimeError(f"Assortment file is missing required columns: {missing}")

    result: list[dict[str, Any]] = []
    for row in rows[1:]:
        mfc = to_float(get_cell(row, idx, "МФЦ"))
        if mfc is None or mfc <= 0:
            continue

        public_price = to_float(get_cell(row, idx, "Публичная цена"))
        if public_price is None or public_price <= 0:
            public_price = mfc

        record = {
            "id": to_str(get_cell(row, idx, "GUID")) or f"row-{len(result) + 1}",
            "status": to_str(get_cell(row, idx, "Статус")),
            "district": to_str(get_cell(row, idx, "Район")),
            "object": to_str(get_cell(row, idx, "Объект")),
            "section": to_str(get_cell(row, idx, "Подъезд (секция)")),
            "layout": to_str(get_cell(row, idx, "Тип планировки")),
            "floor": to_float(get_cell(row, idx, "Этаж")),
            "apartment": format_apartment(get_cell(row, idx, "Номер")),
            "area": to_float(get_cell(row, idx, "Площадь")),
            "mfc": round(mfc, 2),
            "public_price": round(public_price, 2),
        }
        result.append(record)

    result.sort(
        key=lambda item: (
            item["district"],
            item["object"],
            item["apartment"],
            item["id"],
        )
    )
    return result


def build_rates(path: Path) -> list[dict[str, Any]]:
    rows = read_first_sheet_rows(path)
    if not rows:
        raise RuntimeError("Rates workbook is empty.")

    idx = build_header_index(rows[0])
    commission_col = next((name for name in idx if name.startswith("Комиссия")), None)
    if not commission_col:
        raise RuntimeError("Rates file is missing commission column.")

    missing = [col for col in RATES_REQUIRED_HEADERS if col not in idx]
    if missing:
        raise RuntimeError(f"Rates file is missing required columns: {missing}")

    result: list[dict[str, Any]] = []
    for row in rows[1:]:
        commission = to_float(get_cell(row, idx, commission_col))
        client_rate = to_float(get_cell(row, idx, "Ставка для клиента"))
        if commission is None or client_rate is None:
            continue

        special = to_str(get_cell(row, idx, "Особые условия"))
        pv_text = to_str(get_cell(row, idx, "Размер ПВ"))
        limits = parse_credit_limit(special)

        record = {
            "bank": to_str(get_cell(row, idx, "БАНК")),
            "special": special,
            "program": to_str(get_cell(row, idx, "Программа")),
            "loan_term_rule": to_str(get_cell(row, idx, "Срок кредита")),
            "discount_term": to_str(get_cell(row, idx, "Срок действия дисконта")),
            "discount_delta": to_float(
                get_cell(
                    row,
                    idx,
                    "Размер дисконта к процентной ставке (в процентных пунктах)",
                )
            ),
            "base_rate": to_float(get_cell(row, idx, "Базовая ставка")),
            "client_rate": client_rate,
            "commission": commission,
            "pv_text": pv_text,
            "min_pv": parse_min_pv(pv_text),
            "loan_min": limits["loan_min"],
            "loan_max": limits["loan_max"],
        }
        if not record["bank"] or not record["program"] or not record["discount_term"]:
            continue
        result.append(record)

    return result


def build_payload() -> dict[str, Any]:
    assortment_file, rates_file = pick_workbooks()
    existing_payload = load_existing_payload()
    existing_sources = existing_payload.get("sources", {})
    if not isinstance(existing_sources, dict):
        existing_sources = {}

    if assortment_file is None:
        assortment = existing_payload.get("assortment", [])
        if not isinstance(assortment, list):
            assortment = []
        assortment_source = existing_sources.get("assortment_file")
    else:
        assortment = build_assortment(assortment_file)
        assortment_source = assortment_file.name

    rates = build_rates(rates_file)
    return {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "sources": {
            "assortment_file": assortment_source,
            "rates_file": rates_file.name,
        },
        "counts": {
            "apartments": len(assortment),
            "rates": len(rates),
        },
        "assortment": assortment,
        "rates": rates,
    }


def write_payload(payload: dict[str, Any]) -> None:
    content = DATA_PREFIX + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n"
    for out_file in OUT_FILES:
        out_file.write_text(content, encoding="utf-8")


def main() -> None:
    payload = build_payload()
    write_payload(payload)

    print("Updated files:")
    for out_file in OUT_FILES:
        print(f"- {out_file.name}")
    print(f"Apartments: {payload['counts']['apartments']}")
    print(f"Rates: {payload['counts']['rates']}")
    if payload["sources"]["assortment_file"]:
        print(f"Assortment source: {payload['sources']['assortment_file']}")
    else:
        print("Assortment source: not found, kept existing data or left empty")
    print(f"Rates source: {payload['sources']['rates_file']}")


if __name__ == "__main__":
    main()
