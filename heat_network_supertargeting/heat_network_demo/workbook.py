from __future__ import annotations

import math
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def excel_col_to_index(col: str) -> int:
    index = 0
    for ch in col:
        index = index * 26 + (ord(ch.upper()) - ord("A") + 1)
    return index - 1


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("main:si", NS):
        text_parts = []
        for node in si.iterfind(".//main:t", NS):
            text_parts.append(node.text or "")
        values.append("".join(text_parts))
    return values


def parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> object:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text = cell.findtext("main:is/main:t", default="", namespaces=NS)
        return text

    value_text = cell.findtext("main:v", default=None, namespaces=NS)
    if value_text is None:
        return None

    if cell_type == "s":
        return shared_strings[int(value_text)]
    if cell_type == "b":
        return value_text == "1"

    try:
        value = float(value_text)
    except ValueError:
        return value_text
    if math.isclose(value, round(value), abs_tol=1e-12):
        return int(round(value))
    return value


def parse_xlsx_rows(path: str | Path, sheet_name: str | None = None) -> list[list[object]]:
    workbook_path = Path(path)
    with zipfile.ZipFile(workbook_path) as zf:
        shared_strings = load_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall("pkgrel:Relationship", NS)
        }

        sheets = workbook.find("main:sheets", NS)
        if sheets is None:
            raise RuntimeError("No sheets found in workbook.")

        target_sheet = None
        for sheet in sheets.findall("main:sheet", NS):
            if sheet_name is None or sheet.attrib.get("name") == sheet_name:
                target_sheet = sheet
                break
        if target_sheet is None:
            raise RuntimeError(f"Sheet not found: {sheet_name}")

        rel_id = target_sheet.attrib.get(f"{{{NS['rel']}}}id")
        if rel_id is None or rel_id not in rel_map:
            raise RuntimeError("Worksheet relationship not found.")
        sheet_target = rel_map[rel_id]
        normalized_target = sheet_target.lstrip("/")
        sheet_path = normalized_target if normalized_target.startswith("xl/") else f"xl/{normalized_target}"
        sheet_xml = ET.fromstring(zf.read(sheet_path))

    rows: list[list[object]] = []
    sheet_data = sheet_xml.find("main:sheetData", NS)
    if sheet_data is None:
        return rows

    for row in sheet_data.findall("main:row", NS):
        current: list[object] = []
        next_col = 0
        for cell in row.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            col_index = excel_col_to_index(match.group(1)) if match else next_col
            while next_col < col_index:
                current.append(None)
                next_col += 1
            current.append(parse_cell_value(cell, shared_strings))
            next_col += 1
        rows.append(current)
    return rows


def stream_rows_from_workbook(rows: list[list[object]]) -> list[list[object]]:
    if not rows:
        return []
    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    header_map = {header.lower(): idx for idx, header in enumerate(headers)}

    def idx(*names: str) -> int:
        for name in names:
            key = name.lower()
            if key in header_map:
                return header_map[key]
        raise KeyError(f"Missing expected columns: {names}")

    idx_name = idx("Stream Information", "Stream", "Name")
    idx_ts = idx("Supply Temperture (°C)", "Supply Temperature (°C)", "Ts")
    idx_tt = idx("Target Temperature (°C)", "Tt")
    idx_q = idx("Heat Load (kW)", "Q (kW)", "Heat Load")
    idx_u = idx("U (KW/m2.K)", "U (kW/m2.K)", "U")
    idx_cp = idx("CP(kW/K)", "CP (kW/K)", "FCp")

    data_rows: list[list[object]] = []
    for raw in rows[1:]:
        if not raw or all(cell in (None, "") for cell in raw):
            continue
        padded = list(raw)
        data_rows.append(
            [
                padded[idx_name] if idx_name < len(padded) else None,
                padded[idx_ts] if idx_ts < len(padded) else None,
                padded[idx_tt] if idx_tt < len(padded) else None,
                padded[idx_q] if idx_q < len(padded) else None,
                padded[idx_u] if idx_u < len(padded) else None,
                padded[idx_cp] if idx_cp < len(padded) else None,
            ]
        )
    return data_rows
