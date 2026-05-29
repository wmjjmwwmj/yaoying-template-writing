#!/usr/bin/env python3
"""
Write a SinoUnited Health policy DOCX from a fixed Word layout.

The script edits only whitelisted content areas in word/document.xml and copies
all other OOXML package parts from the bundled layout unchanged.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "layout" / "yaoying-2026-policy-layout.docx"

NS = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o": "urn:schemas-microsoft-com:office:office",
    "oel": "http://schemas.microsoft.com/office/2019/extlst",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

W = NS["w"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def wval(el: ET.Element, name: str = "val") -> str | None:
    return el.attrib.get(w(name))


def text_of(el: ET.Element, *, include_instr: bool = False) -> str:
    out: list[str] = []
    for node in el.iter():
        if node.tag in (w("t"), w("delText")) and node.text:
            out.append(node.text)
        elif include_instr and node.tag == w("instrText") and node.text:
            out.append(node.text)
        elif node.tag == w("tab"):
            out.append("\t")
        elif node.tag == w("br"):
            out.append("\n")
    return "".join(out)


def as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        return "\n".join(as_text(item) for item in value if as_text(item))
    if isinstance(value, dict):
        lines = []
        for key, val in value.items():
            val_text = as_text(val)
            if val_text:
                lines.append(f"{key}：{val_text}")
        return "\n".join(lines)
    return str(value).strip()


def as_lines(value: Any) -> list[str]:
    text = as_text(value)
    if not text:
        return [""]
    return [line.strip() for line in text.splitlines()]


def first_run(container: ET.Element) -> ET.Element | None:
    return container.find(".//w:r", NS)


def make_run(text: str, template_run: ET.Element | None = None) -> ET.Element:
    run = ET.Element(w("r"))
    if template_run is not None:
        rpr = template_run.find("w:rPr", NS)
        if rpr is not None:
            run.append(copy.deepcopy(rpr))

    parts = text.split("\n")
    for line_index, line in enumerate(parts):
        if line_index:
            run.append(ET.Element(w("br")))
        tab_parts = line.split("\t")
        for tab_index, tab_part in enumerate(tab_parts):
            if tab_index:
                run.append(ET.Element(w("tab")))
            if tab_part or (not line and not tab_index):
                t = ET.Element(w("t"))
                if tab_part != tab_part.strip() or "  " in tab_part:
                    t.set(XML_SPACE, "preserve")
                t.text = tab_part
                run.append(t)
    return run


def set_paragraph_text(p: ET.Element, text: str, template_run: ET.Element | None = None) -> None:
    if template_run is None:
        template_run = first_run(p)
    for child in list(p):
        if child.tag != w("pPr"):
            p.remove(child)
    p.append(make_run(text, template_run))


def remove_bold_formatting(el: ET.Element) -> None:
    for node in list(el.iter()):
        for child in list(node):
            if child.tag in (w("b"), w("bCs")):
                node.remove(child)


def set_sdt_text(sdts: list[ET.Element], one_based_index: int, text: str) -> None:
    sdt = sdts[one_based_index - 1]
    content = sdt.find("w:sdtContent", NS)
    if content is None:
        raise ValueError(f"sdt {one_based_index} has no content")
    template_run = first_run(content)
    for child in list(content):
        content.remove(child)
    content.append(make_run(text, template_run))


def clear_sdt(sdts: list[ET.Element], one_based_index: int) -> None:
    sdt = sdts[one_based_index - 1]
    content = sdt.find("w:sdtContent", NS)
    if content is None:
        return
    template_run = first_run(content)
    for child in list(content):
        content.remove(child)
    content.append(make_run("", template_run))


def para_ilvl(p: ET.Element) -> str | None:
    node = p.find("./w:pPr/w:numPr/w:ilvl", NS)
    return wval(node) if node is not None else None


def set_content_tab_stop(p: ET.Element, level: str) -> None:
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        return
    tabs = ppr.find("w:tabs", NS)
    if tabs is None:
        tabs = ET.Element(w("tabs"))
        insert_at = 0
        for index, child in enumerate(list(ppr)):
            if child.tag in (w("pStyle"), w("numPr")):
                insert_at = index + 1
        ppr.insert(insert_at, tabs)

    for child in list(tabs):
        tabs.remove(child)

    tab = ET.Element(w("tab"))
    tab.set(w("val"), "left")
    tab.set(w("pos"), "1080" if level == "2" else "900")
    tabs.append(tab)


def set_numbered_content(sdts: list[ET.Element], items: Any) -> None:
    sdt = sdts[17 - 1]
    content = sdt.find("w:sdtContent", NS)
    if content is None:
        raise ValueError("content block has no sdtContent")
    templates = content.findall("w:p", NS)
    by_level = {para_ilvl(p): p for p in templates}
    fallback = templates[0] if templates else ET.Element(w("p"))

    normalized: list[dict[str, Any]] = []
    if isinstance(items, str):
        normalized = [{"level": 1, "text": line} for line in items.splitlines() if line.strip()]
    elif isinstance(items, list):
        for item in items:
            if isinstance(item, dict):
                normalized.append(item)
            else:
                normalized.append({"level": 1, "text": str(item)})

    for child in list(content):
        content.remove(child)

    if not normalized:
        normalized = [{"level": 1, "text": ""}]

    for item in normalized[:12]:
        level = str(int(item.get("level", 1)))
        template = by_level.get(level)
        if template is None:
            template = by_level.get("1")
        if template is None:
            template = fallback
        p = copy.deepcopy(template)
        set_content_tab_stop(p, level)
        text = as_text(item.get("text", ""))
        set_paragraph_text(p, text)
        content.append(p)


def tables(root: ET.Element) -> list[ET.Element]:
    return root.findall(".//w:tbl", NS)


def cell_at(table: ET.Element, row: int, col: int) -> ET.Element:
    rows = table.findall("w:tr", NS)
    if row < 1 or row > len(rows):
        raise IndexError(f"table row {row} out of range")
    cells = rows[row - 1].findall("w:tc", NS)
    if col < 1 or col > len(cells):
        raise IndexError(f"table column {col} out of range")
    return cells[col - 1]


def set_cell_lines(cell: ET.Element, lines: list[str]) -> None:
    paragraphs = cell.findall("w:p", NS)
    template_p = paragraphs[0] if paragraphs else ET.Element(w("p"))
    template_run = first_run(template_p)

    for child in list(cell):
        if child.tag != w("tcPr"):
            cell.remove(child)

    for line in lines:
        p = ET.Element(w("p"))
        ppr = template_p.find("w:pPr", NS)
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        p.append(make_run(line, template_run))
        cell.append(p)


def set_label_value(table: ET.Element, row: int, col: int, label: str, value: str) -> None:
    cell = cell_at(table, row, col)
    paragraphs = cell.findall("w:p", NS)
    if not paragraphs:
        paragraphs = [ET.SubElement(cell, w("p"))]
    p = paragraphs[0]
    label_run = first_run(p)
    value_run = None
    sdt = p.find(".//w:sdt", NS)
    if sdt is not None:
        value_run = first_run(sdt)
    if value_run is None:
        value_run = label_run

    for child in list(p):
        if child.tag != w("pPr"):
            p.remove(child)
    p.append(make_run(label, label_run))
    p.append(make_run(" " + value if value else " ", value_run))


def set_page_count(table: ET.Element, count: str) -> None:
    cell = cell_at(table, 5, 2)
    p = cell.find("w:p", NS)
    if p is None:
        p = ET.SubElement(cell, w("p"))
    template_run = first_run(p)
    for child in list(p):
        if child.tag != w("pPr"):
            p.remove(child)
    p.append(make_run("文件页数：共 ", template_run))
    p.append(make_run(str(count), template_run))
    p.append(make_run("  页", template_run))


def checkbox_nodes(p: ET.Element) -> list[ET.Element]:
    nodes = []
    for run in p.findall("w:r", NS):
        cb = run.find(".//w:checkBox", NS)
        if cb is not None:
            nodes.append(cb)
    return nodes


def set_checkbox(cb: ET.Element, checked: bool) -> None:
    node = cb.find("w:checked", NS)
    if node is None:
        node = ET.SubElement(cb, w("checked"))
    node.set(w("val"), "1" if checked else "0")


def set_all_checkboxes_in_paragraph(p: ET.Element, checked: bool) -> None:
    for cb in checkbox_nodes(p):
        set_checkbox(cb, checked)


def normalize_option(value: Any) -> str:
    return "".join(as_text(value).split()).lower()


def selected_options(value: Any) -> set[str]:
    if isinstance(value, list):
        return {normalize_option(item) for item in value if normalize_option(item)}
    text = as_text(value)
    if not text:
        return set()
    return {normalize_option(part) for part in re.split(r"[、,，;\n]+", text) if normalize_option(part)}


def set_cell_option_checkboxes(cell: ET.Element, selected: Any, options: list[str]) -> None:
    selected_set = selected_options(selected)
    checkboxes = cell.findall(".//w:checkBox", NS)
    for index, option in enumerate(options):
        if index >= len(checkboxes):
            break
        set_checkbox(checkboxes[index], normalize_option(option) in selected_set)


def set_document_nature(main_table: ET.Element, selected: str) -> None:
    if not selected:
        return
    allowed = ["新增", "修订", "废止", "回顾"]
    selected_set = {item.strip() for item in selected.replace(",", "，").split("，") if item.strip()}
    p = cell_at(main_table, 5, 1).find("w:p", NS)
    if p is None:
        return
    runs = p.findall("w:r", NS)
    begin_indexes = [i for i, run in enumerate(runs) if run.find(".//w:checkBox", NS) is not None]
    for pos, run_index in enumerate(begin_indexes):
        label = allowed[pos] if pos < len(allowed) else ""
        cb = runs[run_index].find(".//w:checkBox", NS)
        if cb is not None:
            set_checkbox(cb, label in selected_set)


def parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def normalized_heading_text(p: ET.Element) -> str:
    return "".join(text_of(p).split())


def find_heading_paragraph(root: ET.Element, heading: str) -> ET.Element | None:
    normalized = "".join(heading.split())
    for p in root.findall(".//w:p", NS):
        if normalized_heading_text(p) == normalized:
            return p
    return None


def set_or_insert_after_heading(root: ET.Element, heading: str, value: Any) -> None:
    text = as_text(value)
    if not text:
        return
    heading_p = find_heading_paragraph(root, heading)
    if heading_p is None:
        return

    parents = parent_map(root)
    parent = parents.get(heading_p)
    if parent is None:
        return
    siblings = list(parent)
    index = siblings.index(heading_p)
    target = siblings[index + 1] if index + 1 < len(siblings) and siblings[index + 1].tag == w("p") else None
    if target is not None and not text_of(target).strip():
        set_paragraph_text(target, text)
    else:
        new_p = copy.deepcopy(heading_p)
        num_pr = new_p.find("./w:pPr/w:numPr", NS)
        if num_pr is not None:
            new_p.find("./w:pPr", NS).remove(num_pr)
        set_paragraph_text(new_p, text)
        parent.insert(index + 1, new_p)


def insert_plain_after_heading(root: ET.Element, heading: str, value: Any) -> None:
    text = as_text(value)
    if not text:
        return
    heading_p = find_heading_paragraph(root, heading)
    if heading_p is None:
        return

    parents = parent_map(root)
    parent = parents.get(heading_p)
    if parent is None:
        return

    new_p = copy.deepcopy(heading_p)
    ppr = new_p.find("w:pPr", NS)
    if ppr is not None:
        num_pr = ppr.find("w:numPr", NS)
        if num_pr is not None:
            ppr.remove(num_pr)
    remove_bold_formatting(new_p)
    set_paragraph_text(new_p, text)
    parent.insert(list(parent).index(heading_p) + 1, new_p)


def set_heading_checkbox(root: ET.Element, heading: str, checked: bool) -> None:
    p = find_heading_paragraph(root, heading)
    if p is not None:
        set_all_checkboxes_in_paragraph(p, checked)


def metadata(data: dict[str, Any]) -> dict[str, Any]:
    return data.get("metadata", {}) or {}


def sections(data: dict[str, Any]) -> dict[str, Any]:
    return data.get("sections", {}) or {}


def resource_lookup(value: Any) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    if isinstance(value, dict):
        for key, row in value.items():
            if isinstance(row, dict):
                out[key] = {
                    "quantity": as_text(row.get("quantity", row.get("数量", ""))),
                    "description": as_text(row.get("description", row.get("用途说明", ""))),
                }
            else:
                out[key] = {"quantity": "", "description": as_text(row)}
    elif isinstance(value, list):
        for row in value:
            if isinstance(row, dict):
                name = as_text(row.get("name", row.get("名称", "")))
                if name:
                    out[name] = {
                        "quantity": as_text(row.get("quantity", row.get("数量", ""))),
                        "description": as_text(row.get("description", row.get("用途说明", ""))),
                    }
    return out


def apply_resources(all_tables: list[ET.Element], value: Any) -> None:
    if len(all_tables) < 2:
        return
    table = all_tables[1]
    fixed_names = ["1.人力资源", "2.财务预算", "3.设施空间", "4.资材物料", "5.科技信息"]
    lookup = resource_lookup(value)
    for row_index, fixed_name in enumerate(fixed_names, start=2):
        short_name = fixed_name.split(".", 1)[-1]
        row_data = lookup.get(fixed_name) or lookup.get(short_name) or {}
        set_cell_lines(cell_at(table, row_index, 2), [row_data.get("quantity", "")])
        set_cell_lines(cell_at(table, row_index, 3), [row_data.get("description", "")])


def apply_training(all_tables: list[ET.Element], value: Any) -> None:
    if len(all_tables) < 3 or not isinstance(value, list):
        return
    table = all_tables[2]
    cycle_options = ["1次或岗前", "1月", "1季度", "1年"]
    method_options = ["自学", "在线学习", "授课", "开会"]
    for offset, row_data in enumerate(value[:2], start=2):
        if not isinstance(row_data, dict):
            continue
        object_value = as_text(row_data.get("object", row_data.get("对象", "")))
        object_cell = cell_at(table, offset, 1)
        object_boxes = object_cell.findall(".//w:checkBox", NS)
        if object_boxes:
            row_label = "全员" if offset == 2 else "岗位"
            set_checkbox(object_boxes[0], normalize_option(object_value) == normalize_option(row_label))
        set_cell_option_checkboxes(cell_at(table, offset, 2), row_data.get("cycles", row_data.get("培训周期", "")), cycle_options)
        set_cell_option_checkboxes(cell_at(table, offset, 3), row_data.get("methods", row_data.get("培训方法", "")), method_options)
        set_cell_lines(cell_at(table, offset, 4), as_lines(row_data.get("remarks", row_data.get("备注", ""))))


def apply_revision_history(all_tables: list[ET.Element], value: Any) -> None:
    if len(all_tables) < 4 or not isinstance(value, list):
        return
    table = all_tables[3]
    for offset, row_data in enumerate(value[:3], start=2):
        if not isinstance(row_data, dict):
            continue
        set_cell_lines(cell_at(table, offset, 1), [as_text(row_data.get("version", row_data.get("版本", "")))])
        set_cell_lines(cell_at(table, offset, 2), [as_text(row_data.get("summary", row_data.get("概要", "")))])
        set_cell_lines(cell_at(table, offset, 3), [as_text(row_data.get("date", row_data.get("日期", "")))])


def apply_data(root: ET.Element, data: dict[str, Any]) -> None:
    sdts = root.findall(".//w:sdt", NS)
    all_tables = tables(root)
    main_table = all_tables[0]
    meta = metadata(data)
    secs = sections(data)

    set_sdt_text(sdts, 1, as_text(meta.get("file_name", meta.get("文件名", "文件名")) or "文件名"))
    set_label_value(main_table, 2, 1, "文件类别：", as_text(meta.get("file_category", meta.get("文件类别", ""))))
    set_label_value(main_table, 2, 2, "文件持有部门：", as_text(meta.get("department", meta.get("文件持有部门", ""))))
    set_label_value(main_table, 2, 3, "文件编号：", as_text(meta.get("document_number", meta.get("文件编号", ""))))
    set_label_value(main_table, 3, 1, "版次：", as_text(meta.get("version", meta.get("版次", ""))))
    set_label_value(main_table, 3, 2, "编写者：", as_text(meta.get("writer", meta.get("编写者", ""))))
    set_label_value(main_table, 3, 3, "审核者：", as_text(meta.get("reviewer", meta.get("审核者", ""))))
    set_label_value(main_table, 4, 1, "修订日期：", as_text(meta.get("revision_date", meta.get("修订日期", ""))))
    set_label_value(main_table, 4, 2, "审核日期：", as_text(meta.get("review_date", meta.get("审核日期", ""))))
    set_label_value(main_table, 4, 3, "执行日期：", as_text(meta.get("effective_date", meta.get("执行日期", ""))))
    set_document_nature(main_table, as_text(meta.get("document_nature", meta.get("文件性质", ""))))
    if meta.get("page_count") or meta.get("文件页数"):
        set_page_count(main_table, as_text(meta.get("page_count", meta.get("文件页数", ""))))

    set_sdt_text(sdts, 12, as_text(secs.get("purpose", secs.get("目的", ""))))
    set_sdt_text(sdts, 13, as_text(secs.get("scope", secs.get("范围", ""))))
    set_sdt_text(sdts, 15, as_text(secs.get("definitions", secs.get("定义", ""))))
    clear_sdt(sdts, 16)
    insert_plain_after_heading(root, "权责", secs.get("responsibilities", secs.get("权责", "")))
    set_numbered_content(sdts, secs.get("content", secs.get("内容", [])))
    set_or_insert_after_heading(root, "流程", secs.get("process", secs.get("流程", "")))
    set_or_insert_after_heading(root, "质量管理", secs.get("quality_management", secs.get("质量管理", "")))
    set_or_insert_after_heading(root, "风险管理", secs.get("risk_management", secs.get("风险管理", "")))
    set_or_insert_after_heading(root, "表单附件", secs.get("attachments", secs.get("表单附件", "")))
    set_or_insert_after_heading(root, "参考文献", secs.get("references", secs.get("参考文献", "")))

    set_heading_checkbox(root, "质量管理", bool(as_text(secs.get("quality_management", secs.get("质量管理", "")))))
    set_heading_checkbox(root, "风险管理", bool(as_text(secs.get("risk_management", secs.get("风险管理", "")))))
    set_heading_checkbox(root, "表单附件", bool(as_text(secs.get("attachments", secs.get("表单附件", "")))))
    set_heading_checkbox(root, "参考文献", bool(as_text(secs.get("references", secs.get("参考文献", "")))))
    set_heading_checkbox(root, "复审与修订历史记录", bool(data.get("revision_history")))

    apply_resources(all_tables, data.get("resources", data.get("资源分配", {})))
    apply_training(all_tables, data.get("training_plan", data.get("培训计划", [])))
    apply_revision_history(all_tables, data.get("revision_history", data.get("复审与修订历史记录", [])))


def unwrap_content_controls(root: ET.Element) -> None:
    while True:
        parents = parent_map(root)
        sdt = root.find(".//w:sdt", NS)
        if sdt is None:
            return

        parent = parents.get(sdt)
        content = sdt.find("w:sdtContent", NS)
        if parent is None or content is None:
            return

        index = list(parent).index(sdt)
        replacement_children = list(content)
        parent.remove(sdt)
        for offset, child in enumerate(replacement_children):
            parent.insert(index + offset, child)


def preserve_root_namespace_declarations(xml_bytes: bytes) -> bytes:
    """Keep namespace declarations required by mc:Ignorable after ElementTree serialization."""
    xml = xml_bytes.decode("utf-8")
    match = re.search(r"<w:document\b[^>]*>", xml)
    if not match:
        return xml_bytes

    start_tag = match.group(0)
    missing_attrs = []
    for prefix, uri in NS.items():
        if f"xmlns:{prefix}=" not in start_tag:
            missing_attrs.append(f' xmlns:{prefix}="{uri}"')

    if not missing_attrs:
        return xml_bytes

    replacement = start_tag[:-1] + "".join(missing_attrs) + ">"
    xml = xml[: match.start()] + replacement + xml[match.end() :]
    return xml.encode("utf-8")


def write_docx(template: Path, data_path: Path, output_path: Path) -> None:
    data = json.loads(data_path.read_text(encoding="utf-8"))
    with zipfile.ZipFile(template, "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        apply_data(root, data)
        unwrap_content_controls(root)
        document_xml = preserve_root_namespace_declarations(
            ET.tostring(root, encoding="utf-8", xml_declaration=True)
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w") as zout:
            for info in zin.infolist():
                payload = document_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                zout.writestr(info, payload)


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill the bundled policy DOCX layout from JSON.")
    parser.add_argument("input_json", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    args = parser.parse_args()

    if not args.template.exists():
        print(f"Template not found: {args.template}", file=sys.stderr)
        return 2
    write_docx(args.template, args.input_json, args.output_docx)
    print(args.output_docx)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
