from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


TARGET_COLUMNS = [
    "Direction",
    "Airline",
    "Month",
    "Fortnight",
    "Source File",
    "Flight No",
    "Flight Date",
    "AWB No",
    "Sfx",
    "Origin",
    "Dest",
    "Pcs",
    "Gross Wt",
    "Chg Wt",
    "Billing SHC",
    "Handling Amt",
    "X-Ray Amt",
    "Demurrage Amt",
    "Total Amt (Incl GST)",
]

SUMMARY_COLUMNS = [
    "Airline",
    "Direction",
    "Month",
    "AWB Count",
    "Total Pcs",
    "Total Chg Wt",
    "Total Amt (Incl GST)",
]

RECON_COLUMNS = [
    "Direction",
    "Airline",
    "Month",
    "Fortnight",
    "Source File",
    "Parsed Rows",
    "Line Total",
    "SC Total",
    "Variance",
    "Status",
]

HEADER_ALIASES = {
    "flight_no": ["FLT NO", "FLT NO.", "FLIGHT NO", "FLIGHT No.", "Flight No", "EXP TP SEG FLIGHT No"],
    "flight_date": ["FLT DATE", "FLIGHT DATE", "FLT DT", "Date", "EXP TP FLIGHT DATE"],
    "awb_no": ["AWB No", "AWB NO", "AWB.No", "AWB. NO.", "MAWB No", "MAWB NO", "MAWB", "AWBNo"],
    "sfx": ["AWB Sfx", "SFX", "AWB SFX"],
    "origin": ["ORIGIN", "ORG", "ORI"],
    "dest": ["DEST", "DES"],
    "pcs": ["PCS", "PKGS", "Rec_PCS"],
    "gross_wt": ["GRS Wg", "GRS WGT", "GROSS WT", "Gross wgt", "Received wgt"],
    "chg_wt": ["CHG WT.", "CHG WGT", "CH. WGT", "CH.WEIGHT", "CH.WT", "Received_Chg_Wgt", "CHARGEABLE WGHT"],
    "billing_shc": ["Billing SHC", "BILLING SHC", "SHC"],
}

MONTHS = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}


@dataclass
class ParsedFile:
    rows: list[dict[str, Any]]
    recon: dict[str, Any]
    log: dict[str, Any]


def norm(value: Any) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"[^A-Z0-9]+", "", text.upper())
    return text


ALIAS_LOOKUP = {
    norm(alias): canonical
    for canonical, aliases in HEADER_ALIASES.items()
    for alias in aliases
}


def to_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if text.startswith("="):
        return default
    try:
        return float(text)
    except ValueError:
        return default


def to_int(value: Any) -> int:
    return int(round(to_float(value)))


def excel_date(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return value


def split_awb_suffix(awb: Any, sfx: Any = None) -> tuple[str, str]:
    awb_text = "" if awb is None else str(awb).strip()
    sfx_text = "" if sfx is None else str(sfx).strip()
    if not sfx_text and re.match(r"^\d{3}-\d{8}[A-Z]$", awb_text):
        return awb_text[:-1], awb_text[-1]
    return awb_text, sfx_text


def direction_from_path(root: Path, path: Path) -> str:
    parts = [p.upper() for p in path.relative_to(root).parts]
    if "EXPORT" in parts:
        return "EXP"
    if "IMPORT" in parts:
        return "IMP"
    return ""


def airline_from_path(root: Path, path: Path) -> str:
    rel = path.relative_to(root)
    parts = rel.parts
    if len(parts) >= 3 and parts[0] in {"Export", "Import"}:
        return parts[1].upper()
    match = re.match(r"([A-Z0-9]{2,3})\b", path.name.upper())
    return match.group(1) if match else ""


def month_from_text(text: str) -> str:
    match = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'?(\d{4})\b", text.upper())
    if match:
        return f"{match.group(1)}'{match.group(2)}"
    return ""


def month_from_date(value: Any) -> str:
    value = excel_date(value)
    if isinstance(value, date):
        abbr = list(MONTHS.keys())[value.month - 1]
        return f"{abbr}'{value.year}"
    return ""


def fortnight_from_text(text: str) -> str:
    match = re.search(r"(01\s*-\s*15|16\s*-\s*(?:28|29|30|31))", text.upper())
    return match.group(1).replace(" ", "") if match else "Monthly"


def map_headers(values: list[Any]) -> dict[str, int]:
    result: dict[str, int] = {}
    for idx, value in enumerate(values):
        key = ALIAS_LOOKUP.get(norm(value))
        if key and key not in result:
            result[key] = idx
    return result


def find_header(ws) -> tuple[int, dict[str, int]] | tuple[None, dict[str, int]]:
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        mapping = map_headers(list(row))
        required = {"flight_no", "flight_date", "awb_no", "pcs", "chg_wt"}
        if len(required.intersection(mapping)) >= 4 and "awb_no" in mapping:
            return row_idx, mapping
    return None, {}


def detail_sheet_score(name: str, header_row: int | None, mapping: dict[str, int]) -> int:
    if not header_row:
        return -100
    upper = name.upper()
    if any(x in upper for x in ["SC", "CHALLAN", "DEM", "RATE", "CERT", "SCREEN", "OFFLOAD"]):
        return -50
    score = len(mapping) * 10
    if "DATA_INSERT" in upper:
        score += 80
    if "EXPORT TP_SC" in upper:
        score += 75
    if upper.startswith("ANNEX") or upper == "ANX" or "ANNEXURE" in upper:
        score += 70
    if upper == "SUMMARY":
        score += 35
    if "SUMMARY" in upper and upper != "SUMMARY":
        score -= 20
    return score


def selected_detail_sheets(wb) -> list[tuple[Any, int, dict[str, int]]]:
    candidates = []
    for ws in wb.worksheets:
        header_row, mapping = find_header(ws)
        score = detail_sheet_score(ws.title, header_row, mapping)
        if score > 0:
            candidates.append((score, ws, header_row, mapping))
    if not candidates:
        return []
    candidates.sort(key=lambda item: item[0], reverse=True)
    best_score = candidates[0][0]
    selected = []
    for score, ws, header_row, mapping in candidates:
        title = ws.title.upper()
        if title.startswith("ANNEX ") or score >= best_score - 5 or ("DATA_INSERT" in title and score >= best_score - 20):
            selected.append((ws, header_row, mapping))
    return selected


def parse_rate_card(wb) -> dict[str, Any]:
    rates = {
        "handling": defaultdict(lambda: 0.0),
        "xray": 0.0,
        "import_break_bulk": 1.2,
    }
    for ws in wb.worksheets:
        if "RATE" not in ws.title.upper():
            continue
        rows = list(ws.iter_rows(values_only=True))
        service_row_idx = None
        shc_cols: dict[str, int] = {}
        for i, row in enumerate(rows):
            if norm(row[0] if row else "") == "SERVICEELEMENT":
                service_row_idx = i
                for c, value in enumerate(row):
                    label = str(value).upper().strip() if value is not None else ""
                    if "GEN" == label:
                        shc_cols["GEN"] = c
                    elif "SPL" == label:
                        shc_cols["SPL"] = c
                    elif "TRM" in label:
                        shc_cols["TRM"] = c
                    elif "XRAY" in label or "X-RAY" in label:
                        shc_cols["XRAY"] = c
                break
        if service_row_idx is None:
            continue
        for row in rows[service_row_idx + 1 : service_row_idx + 8]:
            service = str(row[0]).upper() if row and row[0] is not None else ""
            if "CARGO HANDLING" in service:
                for shc, col in shc_cols.items():
                    if shc != "XRAY":
                        rates["handling"][shc] = to_float(row[col])
            elif "X-RAY" in service or "XRAY" in service:
                col = shc_cols.get("XRAY") or shc_cols.get("GEN", 1)
                rates["xray"] = to_float(row[col])
    return rates


def evaluate_amount_formula(formula: Any, row_values: dict[int, Any]) -> float:
    if not isinstance(formula, str) or not formula.startswith("="):
        return to_float(formula)
    expr = formula[1:]
    for col_letters in set(re.findall(r"\b([A-Z]+)\d+\b", expr)):
        col_idx = 0
        for char in col_letters:
            col_idx = col_idx * 26 + ord(char) - 64
        expr = re.sub(rf"\b{col_letters}\d+\b", str(to_float(row_values.get(col_idx))), expr)
    if not re.fullmatch(r"[0-9+\-*/(). ]+", expr):
        return 0.0
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return 0.0


def evaluate_cell(ws, cell, stack: set[str] | None = None) -> float:
    stack = stack or set()
    if cell.coordinate in stack:
        return 0.0
    value = cell.value
    if not isinstance(value, str) or not value.startswith("="):
        return to_float(value)
    stack.add(cell.coordinate)
    expr = value[1:]

    def replace_ref(match: re.Match) -> str:
        col_letters, row_number = match.group(1), int(match.group(2))
        ref_cell = ws.cell(row=row_number, column=column_index_from_string(col_letters))
        return str(evaluate_cell(ws, ref_cell, stack))

    expr = re.sub(r"\b([A-Z]+)(\d+)\b", replace_ref, expr)
    if not re.fullmatch(r"[0-9+\-*/(). ]+", expr):
        return 0.0
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return 0.0


def parse_challan_total(wb) -> float:
    total = 0.0
    for ws in wb.worksheets:
        upper = ws.title.upper()
        if not any(x in upper for x in ["SC", "CHALLAN", "HANDLING_SC", "SUMMRY", "SUMMARY"]):
            continue
        if any(x in upper for x in ["EXPORTS_SUMMARY"]):
            continue
        for row in ws.iter_rows():
            labels = " ".join(str(cell.value).upper() for cell in row if cell.value is not None)
            if "TOTAL" in labels and len([cell for cell in row if cell.value is not None]) <= 4:
                for cell in reversed(row):
                    amount = evaluate_cell(ws, cell)
                    if amount:
                        total += amount
                        break
                continue
            for cell in row:
                header = ws.cell(row=3, column=cell.column).value
                if header and "AMOUNT" in str(header).upper():
                    total += evaluate_cell(ws, cell)
                    break
    return round(total, 2)


def classify_service(description: str) -> str | None:
    text = description.upper()
    if "TOTAL" in text:
        return None
    if "DEMURRAGE" in text or "AIR SIDE DEM" in text:
        return "Demurrage Amt"
    if "X-RAY" in text or "XRAY" in text or "CERTIFICATION" in text:
        return "X-Ray Amt"
    if "HANDLING" in text or "BREAK BULK" in text or "PALLETIZATION" in text or "CARGO HANDLING" in text:
        return "Handling Amt"
    return None


def parse_challan_components(wb) -> dict[str, float]:
    components = {"Handling Amt": 0.0, "X-Ray Amt": 0.0, "Demurrage Amt": 0.0}
    accepted_titles = ["SC", "CHALLAN", "HANDLING_SC", "SUMMRY", "TRNSFR"]
    for ws in wb.worksheets:
        upper = ws.title.upper()
        if not any(token in upper for token in accepted_titles):
            continue
        if "EXPORTS_SUMMARY" in upper:
            continue
        header_row = None
        desc_col = None
        amount_col = None
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8)):
            for cell in row:
                text = str(cell.value).upper() if cell.value is not None else ""
                if "DESCRIPTION" in text or "SERVICE ELEMENT" in text:
                    desc_col = cell.column
                    header_row = cell.row
                if "AMOUNT" in text or "AMT" in text:
                    amount_col = cell.column
                    header_row = header_row or cell.row
            if desc_col and amount_col:
                break
        if not (header_row and desc_col and amount_col):
            continue
        for row_num in range(header_row + 1, ws.max_row + 1):
            description = ws.cell(row=row_num, column=desc_col).value
            service = classify_service(str(description or ""))
            if not service:
                continue
            amount = evaluate_cell(ws, ws.cell(row=row_num, column=amount_col))
            components[service] += amount
    return {key: round(value, 2) for key, value in components.items()}


def allocate_components_to_rows(rows: list[dict[str, Any]], components: dict[str, float]) -> None:
    if not rows:
        return
    total_weight = sum(max(row["Chg Wt"], 0.0) for row in rows) or float(len(rows))
    for component, total in components.items():
        allocated_sum = 0.0
        for idx, row in enumerate(rows):
            if idx == len(rows) - 1:
                amount = round(total - allocated_sum, 2)
            else:
                basis = max(row["Chg Wt"], 0.0) or (total_weight / len(rows))
                amount = round(total * basis / total_weight, 2)
                allocated_sum += amount
            row[component] = amount
    for row in rows:
        row["Total Amt (Incl GST)"] = round(row["Handling Amt"] + row["X-Ray Amt"] + row["Demurrage Amt"], 2)


def demurrage_by_awb(wb) -> dict[str, float]:
    dem: dict[str, float] = defaultdict(float)
    for ws in wb.worksheets:
        if "DEM" not in ws.title.upper():
            continue
        header_row, mapping = find_header(ws)
        if not header_row or "awb_no" not in mapping:
            continue
        headers = [cell.value for cell in ws[header_row]]
        amount_idx = None
        for idx, header in enumerate(headers):
            text = str(header).upper() if header is not None else ""
            if "INCL" in text or "INCLUDING" in text or "INC GST" in text:
                amount_idx = idx
        for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
            awb, _ = split_awb_suffix(row[mapping["awb_no"]].value)
            if not awb or "TOTAL" in awb.upper():
                continue
            if amount_idx is not None:
                dem[awb] += evaluate_cell(ws, row[amount_idx])
    return dem


def handling_rate_for(shc: str, rates: dict[str, Any]) -> float:
    text = (shc or "GEN").upper()
    if "TRM" in text or text in {"CDI"}:
        return rates["handling"]["TRM"] or rates["handling"]["GEN"]
    if text in rates["handling"]:
        return rates["handling"][text]
    if text in {"DGR", "VAL", "AVI", "PER", "VUN"}:
        return rates["handling"]["SPL"] or rates["handling"]["GEN"]
    return rates["handling"]["GEN"]


def parse_workbook(root: Path, path: Path) -> ParsedFile:
    wb = load_workbook(path, data_only=False)
    direction = direction_from_path(root, path)
    airline = airline_from_path(root, path)
    rates = parse_rate_card(wb)
    components = parse_challan_components(wb)
    dem = demurrage_by_awb(wb)
    source_file = path.name
    base_month = month_from_text(source_file)
    base_fortnight = fortnight_from_text(source_file)
    rows: list[dict[str, Any]] = []
    selected = selected_detail_sheets(wb)

    for ws, header_row, mapping in selected:
        sheet_fortnight = fortnight_from_text(ws.title)
        fortnight = sheet_fortnight if sheet_fortnight != "Monthly" else base_fortnight
        for cells in ws.iter_rows(min_row=header_row + 1, values_only=True):
            awb_raw = cells[mapping["awb_no"]] if mapping.get("awb_no") is not None else None
            awb, sfx = split_awb_suffix(awb_raw, cells[mapping["sfx"]] if "sfx" in mapping else None)
            if not awb or awb.upper().startswith("TOTAL") or awb == "… (your consolidated rows here) …":
                continue
            flight_date = excel_date(cells[mapping["flight_date"]]) if "flight_date" in mapping else ""
            month = base_month or month_from_date(flight_date)
            chg_wt = to_float(cells[mapping["chg_wt"]] if "chg_wt" in mapping else 0)
            billing_shc = str(cells[mapping["billing_shc"]]).strip() if "billing_shc" in mapping and cells[mapping["billing_shc"]] is not None else "GEN"
            if billing_shc == "CDI":
                billing_shc = "TRM"
            handling_rate = handling_rate_for(billing_shc, rates)
            xray_rate = rates["xray"] if direction == "EXP" else 0.0
            handling = chg_wt * handling_rate * 1.18
            if direction == "IMP":
                handling += chg_wt * rates["import_break_bulk"] * 1.18
            xray = chg_wt * xray_rate * 1.18
            dem_amt = dem.get(awb, 0.0)
            row = {
                "Direction": direction,
                "Airline": airline,
                "Month": month,
                "Fortnight": fortnight,
                "Source File": source_file,
                "Flight No": cells[mapping["flight_no"]] if "flight_no" in mapping else "",
                "Flight Date": flight_date,
                "AWB No": awb,
                "Sfx": sfx,
                "Origin": cells[mapping["origin"]] if "origin" in mapping else "",
                "Dest": cells[mapping["dest"]] if "dest" in mapping else "",
                "Pcs": to_int(cells[mapping["pcs"]] if "pcs" in mapping else 0),
                "Gross Wt": round(to_float(cells[mapping["gross_wt"]] if "gross_wt" in mapping else 0), 2),
                "Chg Wt": round(chg_wt, 2),
                "Billing SHC": billing_shc,
                "Handling Amt": round(handling, 2),
                "X-Ray Amt": round(xray, 2),
                "Demurrage Amt": round(dem_amt, 2),
            }
            row["Total Amt (Incl GST)"] = round(row["Handling Amt"] + row["X-Ray Amt"] + row["Demurrage Amt"], 2)
            rows.append(row)

    seen = set()
    deduped = []
    for row in rows:
        key = (row["Direction"], row["Airline"], row["Month"], row["Fortnight"], row["AWB No"], row["Sfx"])
        if key not in seen:
            seen.add(key)
            deduped.append(row)

    if any(components.values()):
        allocate_components_to_rows(deduped, components)
    line_total = round(sum(row["Total Amt (Incl GST)"] for row in deduped), 2)
    sc_total = round(sum(components.values()), 2) if any(components.values()) else parse_challan_total(wb)
    variance = round(line_total - sc_total, 2)
    recon = {
        "Direction": direction,
        "Airline": airline,
        "Month": base_month or (deduped[0]["Month"] if deduped else ""),
        "Fortnight": base_fortnight,
        "Source File": source_file,
        "Parsed Rows": len(deduped),
        "Line Total": line_total,
        "SC Total": sc_total,
        "Variance": variance,
        "Status": "OK" if abs(variance) <= 1.0 else "CHECK",
    }
    log = {
        "source": str(path.relative_to(root)),
        "selected_sheets": [ws.title for ws, _, _ in selected],
        "rows": len(deduped),
        "recon_status": recon["Status"],
        "variance": variance,
    }
    return ParsedFile(deduped, recon, log)


def autosize(ws, max_width: int = 42) -> None:
    for column_cells in ws.columns:
        letter = get_column_letter(column_cells[0].column)
        width = 10
        for cell in column_cells:
            value = cell.value
            if value is not None:
                width = max(width, min(max_width, len(str(value)) + 2))
        ws.column_dimensions[letter].width = width


def add_table(ws, name: str) -> None:
    if ws.max_row < 2 or ws.max_column < 1:
        return
    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)


def style_sheet(ws, money_columns: set[str], number_columns: set[str]) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    headers = [cell.value for cell in ws[1]]
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(bottom=thin)
            header = headers[cell.column - 1]
            if header in money_columns:
                cell.number_format = '#,##0.00'
            elif header in number_columns:
                cell.number_format = '#,##0.00'
            elif header == "Flight Date":
                cell.number_format = "dd-mmm-yyyy"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    autosize(ws)


def write_workbook(export_rows: list[dict[str, Any]], import_rows: list[dict[str, Any]], recon_rows: list[dict[str, Any]], output: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, rows in [("Export", export_rows), ("Import", import_rows)]:
        ws = wb.create_sheet(sheet_name)
        ws.append(TARGET_COLUMNS)
        for row in rows:
            ws.append([row.get(col, "") for col in TARGET_COLUMNS])
        style_sheet(
            ws,
            {"Handling Amt", "X-Ray Amt", "Demurrage Amt", "Total Amt (Incl GST)"},
            {"Pcs", "Gross Wt", "Chg Wt"},
        )
        add_table(ws, f"{sheet_name}Table")

    summary = defaultdict(lambda: {"count": 0, "pcs": 0, "chg": 0.0, "total": 0.0})
    for row in export_rows + import_rows:
        key = (row["Airline"], "EXPORT" if row["Direction"] == "EXP" else "IMPORT", row["Month"])
        summary[key]["count"] += 1
        summary[key]["pcs"] += row["Pcs"]
        summary[key]["chg"] += row["Chg Wt"]
        summary[key]["total"] += row["Total Amt (Incl GST)"]

    ws = wb.create_sheet("Month_Summary")
    ws.append(SUMMARY_COLUMNS)
    for key in sorted(summary):
        item = summary[key]
        ws.append([key[0], key[1], key[2], item["count"], item["pcs"], round(item["chg"], 2), round(item["total"], 2)])
    style_sheet(ws, {"Total Amt (Incl GST)"}, {"AWB Count", "Total Pcs", "Total Chg Wt"})
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = "#,##0"
        row[4].number_format = "#,##0"
        row[5].number_format = "#,##0.00"
    add_table(ws, "MonthSummaryTable")

    ws = wb.create_sheet("Reconciliation")
    ws.append(RECON_COLUMNS)
    for row in recon_rows:
        ws.append([row.get(col, "") for col in RECON_COLUMNS])
    style_sheet(ws, {"Line Total", "SC Total", "Variance"}, {"Parsed Rows"})
    for row in ws.iter_rows(min_row=2):
        row[5].number_format = "#,##0"
    for row in ws.iter_rows(min_row=2):
        status = row[9].value
        if status == "CHECK":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="FCE4D6")
        elif status == "OK":
            row[9].fill = PatternFill("solid", fgColor="E2F0D9")
    add_table(ws, "ReconciliationTable")

    readme = wb.create_sheet("README", 0)
    readme["A1"] = "DCSC MIS Consolidation Output"
    readme["A1"].font = Font(bold=True, size=14, color="1F4E78")
    notes = [
        "Generated by dcsc_mis_consolidator.py.",
        "Export and Import sheets are AWB-line level in the requested common schema.",
        "Month_Summary aggregates Airline x Month x Direction.",
        "Reconciliation compares AWB-line calculated charges to source Service Challan / SC totals.",
        "A CHECK status usually indicates source challan rounding/grouping or a layout requiring an airline-specific adapter.",
    ]
    for idx, note in enumerate(notes, start=3):
        readme.cell(idx, 1).value = note
    autosize(readme)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def consolidate(input_root: Path, output: Path, log_path: Path) -> dict[str, Any]:
    files = [
        path
        for path in sorted(input_root.rglob("*.xlsx"))
        if path.name != "MIS_Consolidated_TEMPLATE.xlsx" and not path.name.startswith("~$")
    ]
    export_rows: list[dict[str, Any]] = []
    import_rows: list[dict[str, Any]] = []
    recon_rows: list[dict[str, Any]] = []
    logs: list[dict[str, Any]] = []
    for path in files:
        parsed = parse_workbook(input_root, path)
        if direction_from_path(input_root, path) == "EXP":
            export_rows.extend(parsed.rows)
        elif direction_from_path(input_root, path) == "IMP":
            import_rows.extend(parsed.rows)
        recon_rows.append(parsed.recon)
        logs.append(parsed.log)

    export_rows.sort(key=lambda r: (r["Airline"], r["Month"], r["Fortnight"], str(r["Flight Date"]), r["AWB No"], r["Sfx"]))
    import_rows.sort(key=lambda r: (r["Airline"], r["Month"], r["Fortnight"], str(r["Flight Date"]), r["AWB No"], r["Sfx"]))
    write_workbook(export_rows, import_rows, recon_rows, output)
    summary = {
        "input_root": str(input_root),
        "files_scanned": len(files),
        "export_rows": len(export_rows),
        "import_rows": len(import_rows),
        "reconciliation_checks": sum(1 for r in recon_rows if r["Status"] == "CHECK"),
        "files": logs,
    }
    log_path.write_text(json.dumps(summary, indent=2, default=str), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Consolidate DCSC airline MIS workbooks.")
    parser.add_argument("--input", required=True, type=Path, help="Folder containing Export/ and Import/ folders")
    parser.add_argument("--output", required=True, type=Path, help="Output .xlsx path")
    parser.add_argument("--log", type=Path, default=Path("run_log.json"), help="Run log JSON path")
    args = parser.parse_args()
    summary = consolidate(args.input, args.output, args.log)
    print(json.dumps({k: v for k, v in summary.items() if k != "files"}, indent=2))


if __name__ == "__main__":
    main()
