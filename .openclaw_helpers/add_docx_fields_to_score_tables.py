#!/usr/bin/env python3
"""Add JSON/DOCX answer fields to generated ai4sci score workbooks.

The score table exporter writes simple two-column XLSX files without external
dependencies.  This script keeps that contract: it reads Q2 research fields
from the structured answer JSON, extracts Q47 open answer text from the
formatted answer DOCX, adds selected question evidence beside matching score
rows, then edits the matching score workbook in-place using only the Python
standard library.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Sequence
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHEET = {"m": SHEET_NS}

ET.register_namespace("", SHEET_NS)
ET.register_namespace("r", REL_NS)

FIELD_OPTIONS = [
    ("A", "数学与系统科学"),
    ("B", "物理"),
    ("C", "化学"),
    ("D", "空间科学"),
    ("E", "生命科学"),
    ("F", "地球科学"),
    ("G", "信息科技"),
    ("H", "材料科学"),
    ("I", "能源科学"),
    ("J", "海洋科学"),
    ("K", "环境与生态"),
    ("L", "力学"),
    ("M", "精密仪器与装备"),
    ("N", "土木水利"),
    ("O", "大科学工程"),
    ("P", "交叉科学"),
]
LETTER_TO_FIELD = dict(FIELD_OPTIONS)

ADDED_LABELS = {
    "第一部分 受访者基本信息",
    "主要研究领域（至多三项）",
    "最后一个问题回答（Q47）",
}

EVIDENCE_TARGETS = (
    (13, "自主研发AI科研工具或平台情况"),
    (14, "领域专用AI模型（垂域模型）研发情况"),
    (15, "AI支撑平台使用情况"),
    (17, "AI主要应用环节"),
    (25, "科研组织机制调整内容"),
    (35, "AI带来的核心价值"),
    (36, "AI科研应用发展核心瓶颈"),
    (39, "数据资源改革需求"),
    (40, "计算资源改革需求"),
    (41, "专属AI平台建设需求"),
    (42, "组织机制优化方向"),
    (43, "考核评价机制改革需求"),
    (44, "知识产权归属与保护改革需求"),
    (45, "AI4S人才培养与引进改革需求"),
    (46, "开放合作推进需求"),
)
EVIDENCE_LABELS = {label for _, label in EVIDENCE_TARGETS}
MULTI_EVIDENCE_QUESTIONS = {15, 17, 25, 35, 36, 39, 40, 41, 42, 43, 44, 45, 46}


def docx_lines(path: Path) -> list[str]:
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    lines: list[str] = []
    for paragraph in root.findall(".//w:p", WORD_NS):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", WORD_NS)).strip()
        if text:
            lines.append(text)
    return lines


def is_question_marker(line: str, number: int | None = None) -> bool:
    stripped = line.strip()
    if number is not None:
        return bool(re.fullmatch(rf"Q\s*0*{number}", stripped, flags=re.I))
    return bool(re.fullmatch(r"Q\s*(?:[1-9]|[1-3][0-9]|4[0-7])", stripped, flags=re.I))


def question_segment(lines: Sequence[str], number: int) -> list[str]:
    start = next((index for index, line in enumerate(lines) if is_question_marker(line, number)), None)
    if start is None:
        raise ValueError(f"missing Q{number}")

    end = len(lines)
    for index in range(start + 1, len(lines)):
        if is_question_marker(lines[index]):
            end = index
            break
    return list(lines[start:end])


def answer_text(segment: Sequence[str]) -> str:
    answer_lines: list[str] = []
    collecting = False
    stop_prefixes = (
        "【证据】",
        "【置信度】",
        "【原题】",
        "【原选项】",
        "【填答要求】",
        "【作答与依据】",
        "【可选维度】",
        "【原题补充】",
    )

    for raw_line in segment:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("【答案】"):
            collecting = True
            remainder = line.removeprefix("【答案】").strip()
            if remainder:
                answer_lines.append(remainder)
            continue
        if not collecting:
            continue
        if line.startswith(stop_prefixes) or line.startswith("──"):
            break
        if is_question_marker(line):
            break
        answer_lines.append(line)

    return "\n".join(answer_lines).strip()


def clean_answer_line(line: str) -> str:
    text = line.strip()
    text = text.replace("**", "").replace("__", "").replace("`", "")
    text = re.sub(r"^\s*[-*]\s+", "", text)
    return text.strip()


def answer_and_evidence_text(segment: Sequence[str]) -> str:
    answer_lines: list[str] = []
    evidence_lines: list[str] = []
    mode: str | None = None

    for raw_line in segment:
        line = clean_answer_line(raw_line)
        if not line:
            continue
        if is_question_marker(line):
            continue

        marker = re.match(r"^【([^】]+)】\s*(.*)$", line)
        if marker:
            name = marker.group(1).strip()
            remainder = marker.group(2).strip()
            if name == "答案":
                mode = "answer"
                if remainder:
                    answer_lines.append(remainder)
                continue
            if name.startswith("证据"):
                mode = "evidence"
                if remainder:
                    evidence_lines.append(remainder)
                continue
            if name == "置信度":
                break
            if name in {"原题", "原选项", "填答要求", "作答与依据", "可选维度", "原题补充"}:
                mode = None
                continue

        if mode == "answer":
            answer_lines.append(line)
        elif mode == "evidence":
            evidence_lines.append(line)

    parts: list[str] = []
    if answer_lines:
        parts.append("答案：" + "\n".join(answer_lines))
    if evidence_lines:
        parts.append("证据：" + "\n".join(evidence_lines))
    return "\n".join(parts).strip()


def evidence_lines(segment: Sequence[str]) -> list[str]:
    lines: list[str] = []
    collecting = False

    for raw_line in segment:
        line = clean_answer_line(raw_line)
        if not line:
            continue
        if is_question_marker(line):
            continue

        marker = re.match(r"^【([^】]+)】\s*(.*)$", line)
        if marker:
            name = marker.group(1).strip()
            remainder = marker.group(2).strip()
            if name.startswith("证据"):
                collecting = True
                if remainder:
                    lines.append(remainder)
                continue
            if name == "置信度":
                break
            if name in {"答案", "原题", "原选项", "填答要求", "作答与依据", "可选维度", "原题补充"}:
                collecting = False
                continue

        if collecting:
            lines.append(line)

    return lines


def selected_letters_from_json(data: dict[str, Any], question_number: int) -> list[str]:
    value = data.get(str(question_number))
    selected = value.get("selected") if isinstance(value, dict) else None
    if not isinstance(selected, list):
        return []

    letters: list[str] = []
    for item in selected:
        letter = str(item).strip().upper()
        if re.fullmatch(r"[A-Z]", letter) and letter not in letters:
            letters.append(letter)
    return letters


def option_evidence_cells(segment: Sequence[str], selected_letters: Sequence[str]) -> tuple[list[str], list[str]]:
    selected = [letter.upper() for letter in selected_letters]
    selected_set = set(selected)
    collected: dict[str, list[str]] = {letter: [] for letter in selected}
    warnings: list[str] = []
    current_letter: str | None = None

    option_line_pattern = re.compile(
        r"^(?:选|选择)?\s*([A-Z])\s*(?:[（(][^）)]*[）)])?\s*(?:[：:：\-—]+)?\s*(.*)$"
    )

    for line in evidence_lines(segment):
        if re.match(r"^(?:不选|未选|不选择)\s*[A-Z]", line):
            current_letter = None
            continue
        if line.startswith(("评分计算", "评分：", "评分依据", "综上", "因此")):
            current_letter = None
            continue

        match = option_line_pattern.match(line)
        if match:
            letter = match.group(1).upper()
            if letter in selected_set:
                current_letter = letter
                collected[letter].append(line)
            else:
                current_letter = None
            continue

        if current_letter:
            collected[current_letter].append(line)

    fallback = answer_and_evidence_text(segment)
    cells: list[str] = []
    for letter in selected:
        option_text = "\n".join(collected.get(letter) or []).strip()
        if option_text:
            cells.append(option_text)
        elif fallback:
            cells.append(f"选{letter}：未在证据段落中拆出逐项证据，以下为该题综合依据。\n{fallback}")
        else:
            cells.append(f"选{letter}：未提取到证据")
            warnings.append(f"Q option {letter}: no evidence extracted")
    return cells, warnings


def canonical_field(text: str) -> str | None:
    normalized = re.sub(r"\s+", "", text)
    for _, field in FIELD_OPTIONS:
        if field in normalized:
            return field
    return None


def unique_limited(values: Sequence[str], limit: int = 3) -> list[str]:
    result: list[str] = []
    for value in values:
        cleaned = value.strip()
        if cleaned and cleaned not in result:
            result.append(cleaned)
        if len(result) >= limit:
            break
    return result


def institute_from_docx(path: Path) -> str:
    stem = path.stem
    prefix = "问卷_格式整理_"
    return stem[len(prefix) :] if stem.startswith(prefix) else stem


def institute_from_xlsx(path: Path) -> str:
    stem = path.stem
    prefix = "评分表_"
    return stem[len(prefix) :] if stem.startswith(prefix) else stem


def research_fields_from_json(json_path: Path) -> list[str]:
    if not json_path.is_file():
        return []
    with json_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    value = data.get("2") if isinstance(data, dict) else None
    selected = value.get("selected") if isinstance(value, dict) else None
    if not isinstance(selected, list):
        return []
    fields: list[str] = []
    for item in selected:
        text = str(item).strip()
        field = canonical_field(text)
        if field is None and re.fullmatch(r"[A-Pa-p]", text):
            field = LETTER_TO_FIELD[text.upper()]
        if field is not None:
            fields.append(field)
    return unique_limited(fields)


def answer_json_data(json_path: Path) -> dict[str, Any]:
    if not json_path.is_file():
        return {}
    with json_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else {}


def extract_fields(docx_path: Path, json_path: Path) -> tuple[str, str, dict[str, list[str]], list[str]]:
    lines = docx_lines(docx_path)
    q47_answer = answer_text(question_segment(lines, 47))
    answer_data = answer_json_data(json_path)

    warnings: list[str] = []
    fields = research_fields_from_json(json_path)
    if not fields:
        warnings.append(f"{json_path.name}: no Q2 research fields extracted from JSON")
    if not q47_answer:
        warnings.append(f"{docx_path.name}: no Q47 answer extracted")

    evidence_by_label: dict[str, list[str]] = {}
    for question_number, label in EVIDENCE_TARGETS:
        try:
            segment = question_segment(lines, question_number)
        except ValueError as error:
            warnings.append(f"{docx_path.name}: {error}")
            continue
        if question_number in MULTI_EVIDENCE_QUESTIONS:
            selected_letters = selected_letters_from_json(answer_data, question_number)
            if not selected_letters:
                warnings.append(f"{json_path.name}: no selected letters found for Q{question_number}")
                evidence_text = answer_and_evidence_text(segment)
                if evidence_text:
                    evidence_by_label[label] = [evidence_text]
                continue
            cells, option_warnings = option_evidence_cells(segment, selected_letters)
            evidence_by_label[label] = cells
            warnings.extend(f"{docx_path.name}: Q{question_number} {warning}" for warning in option_warnings)
        else:
            evidence_text = answer_and_evidence_text(segment)
            if evidence_text:
                evidence_by_label[label] = [evidence_text]
            else:
                warnings.append(f"{docx_path.name}: no answer/evidence extracted for Q{question_number}")

    return "、".join(fields), q47_answer, evidence_by_label, warnings


def column_number_from_ref(reference: str) -> int:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", reference)
    if not match:
        return 0
    number = 0
    for letter in match.group(1):
        number = number * 26 + ord(letter) - 64
    return number


def cell_column_number(cell: ET.Element) -> int:
    return column_number_from_ref(cell.get("r", ""))


def qname(name: str) -> str:
    return f"{{{SHEET_NS}}}{name}"


def cell_ref(row_number: int, column_number: int) -> str:
    name = ""
    number = column_number
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return f"{name}{row_number}"


def inline_string_cell(row_number: int, column_number: int, value: str, style_id: int) -> ET.Element:
    cell = ET.Element(
        qname("c"),
        {"r": cell_ref(row_number, column_number), "s": str(style_id), "t": "inlineStr"},
    )
    inline = ET.SubElement(cell, qname("is"))
    text = ET.SubElement(inline, qname("t"))
    text.text = value
    return cell


def added_row(row_number: int, kind: str, label: str, value: str = "") -> ET.Element:
    if kind == "part":
        label_style = 1
        value_style = 5
        height = 22
    elif label == "最后一个问题回答（Q47）":
        label_style = 3
        value_style = 3
        height = 92
    else:
        label_style = 3
        value_style = 3
        height = 24

    row = ET.Element(
        qname("row"),
        {"r": str(row_number), "ht": str(height), "customHeight": "1"},
    )
    row.append(inline_string_cell(row_number, 1, label, label_style))
    row.append(inline_string_cell(row_number, 2, value, value_style))
    return row


def row_first_text(row: ET.Element) -> str:
    first_cell = row.find("m:c", SHEET)
    if first_cell is None:
        return ""
    text_node = first_cell.find("m:is/m:t", SHEET)
    return text_node.text or "" if text_node is not None else ""


def remove_cell(row: ET.Element, column_number: int) -> None:
    for cell in list(row.findall("m:c", SHEET)):
        if cell_column_number(cell) == column_number:
            row.remove(cell)


def remove_cells_from(row: ET.Element, min_column_number: int) -> None:
    for cell in list(row.findall("m:c", SHEET)):
        if cell_column_number(cell) >= min_column_number:
            row.remove(cell)


def set_inline_string(row: ET.Element, row_number: int, column_number: int, value: str, style_id: int) -> None:
    remove_cell(row, column_number)
    row.append(inline_string_cell(row_number, column_number, value, style_id))
    cells = list(row.findall("m:c", SHEET))
    for cell in cells:
        row.remove(cell)
    for cell in sorted(cells, key=cell_column_number):
        row.append(cell)


def evidence_row_height(text: str) -> int:
    explicit_lines = max(1, len(text.splitlines()))
    wrapped_lines = max(explicit_lines, len(text) // 48)
    return min(150, max(40, 22 + wrapped_lines * 14))


def set_min_row_height(row: ET.Element, height: int) -> None:
    current = float(row.get("ht", "0") or 0)
    if current < height:
        row.set("ht", str(height))
        row.set("customHeight", "1")


def shift_row(row: ET.Element, row_number: int) -> None:
    row.set("r", str(row_number))
    for cell in row.findall("m:c", SHEET):
        reference = cell.get("r", "")
        match = re.fullmatch(r"([A-Z]+)([0-9]+)", reference)
        if match:
            cell.set("r", f"{match.group(1)}{row_number}")


def widen_value_column(root: ET.Element) -> None:
    cols = root.find("m:cols", SHEET)
    if cols is None:
        cols = ET.Element(qname("cols"))
        sheet_data = root.find("m:sheetData", SHEET)
        insert_at = list(root).index(sheet_data) if sheet_data is not None else 0
        root.insert(insert_at, cols)

    found_b = False
    found_c = False
    for col in cols.findall("m:col", SHEET):
        min_col = int(col.get("min", "0"))
        max_col = int(col.get("max", "0"))
        if min_col <= 2 <= max_col:
            col.set("width", "86")
            col.set("customWidth", "1")
            found_b = True
        if min_col <= 3 <= max_col:
            col.set("width", "96")
            col.set("customWidth", "1")
            found_c = True
    if not found_b:
        cols.append(ET.Element(qname("col"), {"min": "2", "max": "2", "width": "86", "customWidth": "1"}))
    if not found_c:
        cols.append(ET.Element(qname("col"), {"min": "3", "max": "3", "width": "96", "customWidth": "1"}))


def widen_evidence_columns(root: ET.Element, max_evidence_cells: int) -> None:
    if max_evidence_cells <= 0:
        return
    cols = root.find("m:cols", SHEET)
    if cols is None:
        cols = ET.Element(qname("cols"))
        sheet_data = root.find("m:sheetData", SHEET)
        insert_at = list(root).index(sheet_data) if sheet_data is not None else 0
        root.insert(insert_at, cols)

    start_col = 3
    end_col = start_col + max_evidence_cells - 1
    existing = {
        (int(col.get("min", "0")), int(col.get("max", "0")))
        for col in cols.findall("m:col", SHEET)
    }
    for column_number in range(start_col, end_col + 1):
        if any(min_col <= column_number <= max_col for min_col, max_col in existing):
            continue
        cols.append(
            ET.Element(
                qname("col"),
                {"min": str(column_number), "max": str(column_number), "width": "58", "customWidth": "1"},
            )
        )


def enrich_xlsx(xlsx_path: Path, research_fields: str, q47_answer: str, evidence_by_label: dict[str, list[str]]) -> None:
    with ZipFile(xlsx_path, "r") as source:
        root = ET.fromstring(source.read("xl/worksheets/sheet1.xml"))
        sheet_data = root.find("m:sheetData", SHEET)
        if sheet_data is None:
            raise ValueError(f"{xlsx_path}: missing sheetData")

        original_rows = list(sheet_data.findall("m:row", SHEET))
        clean_rows = [row for row in original_rows if row_first_text(row) not in ADDED_LABELS]
        new_rows = [
            added_row(1, "part", "第一部分 受访者基本信息"),
            added_row(2, "item", "主要研究领域（至多三项）", research_fields),
            added_row(3, "item", "最后一个问题回答（Q47）", q47_answer),
        ]

        sheet_data.clear()
        for row in new_rows:
            sheet_data.append(row)
        for offset, row in enumerate(clean_rows, start=4):
            shift_row(row, offset)
            label = row_first_text(row)
            remove_cells_from(row, 3)
            if label in EVIDENCE_LABELS:
                evidence_cells = evidence_by_label.get(label, [])
                max_height = 0
                for cell_offset, evidence_text in enumerate(evidence_cells):
                    if not evidence_text:
                        continue
                    set_inline_string(row, offset, 3 + cell_offset, evidence_text, 3)
                    max_height = max(max_height, evidence_row_height(evidence_text))
                if max_height:
                    set_min_row_height(row, max_height)
            sheet_data.append(row)

        widen_value_column(root)
        widen_evidence_columns(root, max((len(values) for values in evidence_by_label.values()), default=0))
        worksheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as handle:
            tmp_path = Path(handle.name)
        with ZipFile(tmp_path, "w", ZIP_DEFLATED) as target:
            for item in source.infolist():
                payload = worksheet_xml if item.filename == "xl/worksheets/sheet1.xml" else source.read(item.filename)
                target.writestr(item, payload)

    shutil.move(str(tmp_path), xlsx_path)


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Add Q2, Q47, and selected question evidence from answer DOCX/JSON to matching score XLSX files.",
    )
    parser.add_argument("--docx-dir", type=Path, default=Path("result"))
    parser.add_argument("--xlsx-dir", type=Path, default=Path("score_tables"))
    parser.add_argument("--json-dir", type=Path, default=Path("result"))
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    docx_files = sorted(args.docx_dir.glob("问卷_格式整理_*.docx"))
    if not docx_files:
        raise SystemExit(f"no formatted answer DOCX files found in: {args.docx_dir}")

    xlsx_by_institute = {
        institute_from_xlsx(path): path
        for path in sorted(args.xlsx_dir.glob("评分表_*.xlsx"))
    }
    if not xlsx_by_institute:
        raise SystemExit(f"no score XLSX files found in: {args.xlsx_dir}")

    updated = 0
    warnings: list[str] = []
    for docx_path in docx_files:
        institute = institute_from_docx(docx_path)
        xlsx_path = xlsx_by_institute.get(institute)
        if xlsx_path is None:
            warnings.append(f"{docx_path.name}: matching score workbook not found")
            continue

        json_path = args.json_dir / f"单选题答案_{institute}.json"
        research_fields, q47_answer, evidence_by_label, docx_warnings = extract_fields(docx_path, json_path)
        warnings.extend(docx_warnings)
        enrich_xlsx(xlsx_path, research_fields, q47_answer, evidence_by_label)
        updated += 1
        print(f"updated {xlsx_path}")

    for warning in warnings:
        print(f"WARNING: {warning}", file=sys.stderr)

    if updated != len(xlsx_by_institute):
        print(
            f"WARNING: updated {updated} workbook(s), found {len(xlsx_by_institute)} workbook(s)",
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
