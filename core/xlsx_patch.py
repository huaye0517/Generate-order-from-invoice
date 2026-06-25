# -*- coding: utf-8 -*-
"""快速修改 xlsx 单元格（避免整本加载大文件）"""
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict
from xml.etree import ElementTree as ET

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ET.register_namespace("", NS["main"])


def _col_row(cell_ref: str):
    m = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper())
    if not m:
        raise ValueError(f"无效单元格: {cell_ref}")
    return m.group(1), int(m.group(2))


def _find_sheet_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = wb.find("main:sheets", NS)
    if sheets is None:
        raise ValueError("workbook.xml 缺少 sheets")

    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map: Dict[str, str] = {}
    for rel in rels:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rel_map[rid] = target if target.startswith("xl/") else "xl/" + target.lstrip("/")

    for sheet in sheets.findall("main:sheet", NS):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if name == sheet_name and rid in rel_map:
            return rel_map[rid]
    raise ValueError(f"未找到工作表: {sheet_name}")


def _patch_sheet_xml(sheet_xml: bytes, cell_ref: str, value: str) -> bytes:
    _, row_num = _col_row(cell_ref)
    value = "" if value is None else str(value)
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find("main:sheetData", NS)
    if sheet_data is None:
        sheet_data = ET.SubElement(root, f"{{{NS['main']}}}sheetData")

    target_row = None
    for row in sheet_data.findall("main:row", NS):
        if int(row.attrib.get("r", "0")) == row_num:
            target_row = row
            break
    if target_row is None:
        target_row = ET.SubElement(sheet_data, f"{{{NS['main']}}}row", {"r": str(row_num)})

    target_cell = None
    for cell in target_row.findall("main:c", NS):
        if cell.attrib.get("r", "").upper() == cell_ref.upper():
            target_cell = cell
            break

    if target_cell is None:
        target_cell = ET.SubElement(
            target_row,
            f"{{{NS['main']}}}c",
            {"r": cell_ref.upper(), "t": "inlineStr"},
        )
    else:
        target_cell.attrib["t"] = "inlineStr"
        for child in list(target_cell):
            target_cell.remove(child)

    is_elem = ET.SubElement(target_cell, f"{{{NS['main']}}}is")
    t_elem = ET.SubElement(is_elem, f"{{{NS['main']}}}t")
    t_elem.text = value
    if value != value.strip() or "  " in value:
        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def set_cell_inline_str(xlsx_path: str, sheet_name: str, cell_ref: str, value: str) -> None:
    path = Path(xlsx_path)
    tmp_path = path.with_suffix(".tmp.xlsx")

    with zipfile.ZipFile(path, "r") as zin:
        sheet_target = _find_sheet_path(zin, sheet_name)
        new_sheet = _patch_sheet_xml(zin.read(sheet_target), cell_ref, value)
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_sheet if item.filename == sheet_target else zin.read(item.filename)
                zout.writestr(item, data)

    tmp_path.replace(path)


def _patch_hide_rows(sheet_xml: bytes, row_start: int, row_end: int) -> bytes:
    """隐藏指定行区间，减少产品表空白留白"""
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find("main:sheetData", NS)
    if sheet_data is None:
        return sheet_xml

    existing_rows = {
        int(row.attrib.get("r", "0")): row
        for row in sheet_data.findall("main:row", NS)
        if row.attrib.get("r", "").isdigit()
    }

    for row_num in range(row_start, row_end + 1):
        row = existing_rows.get(row_num)
        if row is None:
            row = ET.SubElement(sheet_data, f"{{{NS['main']}}}row", {"r": str(row_num)})
            existing_rows[row_num] = row
        row.set("hidden", "1")
        row.set("customHeight", "1")
        row.set("ht", "0")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def hide_sheet_rows(xlsx_path: str, sheet_name: str, row_start: int, row_end: int) -> None:
    """隐藏工作表中指定行"""
    if row_start > row_end:
        return
    path = Path(xlsx_path)
    tmp_path = path.with_suffix(".tmp.xlsx")

    with zipfile.ZipFile(path, "r") as zin:
        sheet_target = _find_sheet_path(zin, sheet_name)
        new_sheet = _patch_hide_rows(zin.read(sheet_target), row_start, row_end)
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_sheet if item.filename == sheet_target else zin.read(item.filename)
                zout.writestr(item, data)

    tmp_path.replace(path)


def copy_and_set_e2(
    template_path: str,
    dest_path: str,
    e2_value: str,
    hide_rows: tuple[int, int] | None = None,
) -> None:
    """复制模板、写入 E2，并可选隐藏产品区空行"""
    shutil.copy2(template_path, dest_path)
    set_cell_inline_str(dest_path, "模板", "E2", e2_value)
    if hide_rows:
        hide_sheet_rows(dest_path, "模板", hide_rows[0], hide_rows[1])
