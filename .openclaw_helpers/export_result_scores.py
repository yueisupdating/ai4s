#!/usr/bin/env python3
"""Export ai4sci JSON answers to compact score/selection tables.

Output is one `.xlsx` or `.txt` per JSON file.  The Excel writer uses only the
Python standard library, so it can run on the openclaw host without openpyxl.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


@dataclass(frozen=True)
class Item:
    key: str
    label: str
    mode: str = "score"  # score, selected, scored_selected, urgency


@dataclass(frozen=True)
class Section:
    title: str
    items: tuple[Item, ...]


@dataclass(frozen=True)
class Part:
    title: str
    sections: tuple[Section, ...]
    items: tuple[Item, ...] = ()


@dataclass(frozen=True)
class OutputRow:
    kind: str
    label: str
    value: int | float | str | None


BASIC_INFO_PART = Part(
    "第一部分 受访者基本信息",
    (),
    (
        Item("2", "主要研究领域", "selected_text"),
    ),
)

READY_PART = Part(
    "第二部分 AI赋能科研就绪度调研",
    (
        Section(
            "数据资源就绪度",
            (
                Item("3", "科研数据整体情况"),
                Item("4", "数据管理平台建设情况"),
                Item("5", "科研数据可访问性"),
                Item("6", "内部数据共享机制完善度"),
                Item("7", "外部数据共享合作参与度"),
            ),
        ),
        Section(
            "计算资源就绪度",
            (
                Item("8", "计算资源来源构成情况"),
                Item("9", "AI4S计算资源科研需求满足情况"),
                Item("10", "专用AI计算硬件（如GPU集群）情况"),
                Item("11", "云计算资源便捷使用情况"),
                Item("12", "计算资源分配机制合理性"),
            ),
        ),
        Section(
            "模型与平台就绪度",
            (
                Item("13", "自主研发AI科研工具或平台情况"),
                Item("14", "领域专用AI模型（垂域模型）研发情况"),
                Item("15", "AI支撑平台使用情况", "scored_selected"),
            ),
        ),
        Section(
            "科研创新就绪度",
            (
                Item("16", "科研人员AI使用情况"),
                Item("17", "AI主要应用环节", "selected"),
                Item("18", "AI4S相关项目数量和质量"),
                Item("19", "AI驱动重大科研成果情况"),
                Item("20", "AI科研成果转化能力"),
            ),
        ),
        Section(
            "人才储备就绪度",
            (
                Item("21", "复合型人才储备情况"),
                Item("22", "AI相关培训情况"),
                Item("23", "AI4S人才激励机制情况"),
            ),
        ),
        Section(
            "组织机制就绪度",
            (
                Item("24", "科研组织机制调整情况"),
                Item("25", "科研组织机制调整内容", "selected"),
                Item("26", "AI4S跨学科团队建设支持情况"),
                Item("27", "分布式科研团队组建情况", "scored_selected"),
                Item("28", "AI4S科研绩效考核适配情况"),
                Item("29", "高风险创新项目支持情况"),
                Item("30", "AI4S相关成果知识产权保护情况"),
            ),
        ),
        Section(
            "开放合作就绪度",
            (
                Item("31", "AI领域外部合作情况"),
                Item("32", "国际AI4S科研合作情况"),
                Item("33", "AI4S开源社区或开放科学平台参与情况"),
                Item("34", "内部合作文化"),
            ),
        ),
        Section(
            "整体就绪度",
            (
                Item("35", "AI带来的核心价值", "selected"),
                Item("36", "AI科研应用发展核心瓶颈", "selected"),
                Item("37", "整体就绪度"),
            ),
        ),
    ),
)

DEMAND_PART = Part(
    "第三部分 AI赋能科研需求调研",
    (),
    (
        Item("38", "AI时代科研组织模式与资源配置改革紧迫性", "urgency"),
        Item("39", "数据资源改革需求", "selected"),
        Item("40", "计算资源改革需求", "selected"),
        Item("41", "专属AI平台建设需求", "selected"),
        Item("42", "组织机制优化方向", "selected"),
        Item("43", "考核评价机制改革需求", "selected"),
        Item("44", "知识产权归属与保护改革需求", "selected"),
        Item("45", "AI4S人才培养与引进改革需求", "selected"),
        Item("46", "开放合作推进需求", "selected"),
    ),
)

EXPECTED_KEYS = {str(number) for number in range(2, 47)}
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sorted_selected_letters(selected: Any) -> list[str]:
    if isinstance(selected, str):
        letters = [letter for letter in selected.upper() if letter.isalpha()]
    elif isinstance(selected, list):
        letters = [str(item).strip().upper() for item in selected]
    else:
        return []
    return [letter for letter in letters if re.fullmatch(r"[A-Z]", letter)]


def get_score(value: Any) -> int | None:
    if type(value) is int and value in {1, 2, 3, 4}:
        return value
    if isinstance(value, dict):
        score = value.get("score")
        if type(score) is int and score in {1, 2, 3, 4}:
            return score
    return None


def get_selected(value: Any) -> list[str]:
    if isinstance(value, dict):
        return sorted_selected_letters(value.get("selected", []))
    return []


def get_selected_texts(value: Any) -> list[str]:
    if not isinstance(value, dict):
        return []
    selected = value.get("selected", [])
    if not isinstance(selected, list):
        return []
    return [str(item).strip() for item in selected if str(item).strip()]


def selected_display(value: Any) -> str:
    letters = get_selected(value)
    if letters:
        return "".join(letters)
    if type(value) is int:
        return f"旧格式缺少选项（{value}）"
    return ""


def item_display_value(item: Item, answers: dict[str, Any]) -> int | float | str | None:
    value = answers.get(item.key)
    if item.mode == "score":
        return get_score(value)
    if item.mode == "selected":
        return selected_display(value)
    if item.mode == "selected_text":
        return "、".join(get_selected_texts(value))
    if item.mode == "scored_selected":
        return get_score(value)
    if item.mode == "urgency":
        score = get_score(value)
        letters = get_selected(value)
        if letters and score is not None:
            return f"{letters[0]}（{score}分）"
        if score is not None:
            return f"{score}分"
        return selected_display(value)
    raise ValueError(f"unknown item mode: {item.mode}")


def item_average_score(item: Item, answers: dict[str, Any]) -> int | None:
    if item.mode not in {"score", "scored_selected"}:
        return None
    return get_score(answers.get(item.key))


def average(values: Iterable[int]) -> float | None:
    numeric = list(values)
    if not numeric:
        return None
    return sum(numeric) / len(numeric)


def build_rows(answers: dict[str, Any], include_basic_info: bool) -> list[OutputRow]:
    parts: list[Part] = []
    if include_basic_info:
        parts.append(BASIC_INFO_PART)
    parts.extend((READY_PART, DEMAND_PART))

    rows: list[OutputRow] = []
    for part in parts:
        rows.append(OutputRow("part", part.title, None))
        for item in part.items:
            rows.append(OutputRow("item", item.label, item_display_value(item, answers)))
        for section in part.sections:
            section_scores = [
                score
                for item in section.items
                if (score := item_average_score(item, answers)) is not None
            ]
            rows.append(OutputRow("section", section.title, average(section_scores)))
            for item in section.items:
                rows.append(OutputRow("item", item.label, item_display_value(item, answers)))
    return rows


def read_answer_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"{path}: JSON must be an object")

    missing = sorted(EXPECTED_KEYS - set(data), key=int)
    extra = sorted(
        set(data) - EXPECTED_KEYS,
        key=lambda item: (not item.isdigit(), int(item) if item.isdigit() else item),
    )
    if missing or extra:
        raise ValueError(f"{path}: answer keys must be 2..46; missing={missing}, extra={extra}")
    return data


def institute_name_from_json(path: Path) -> str:
    stem = path.stem
    for prefix in ("单选题答案_", "答案_", "score_", "scores_"):
        if stem.startswith(prefix):
            stem = stem[len(prefix) :]
            break
    return stem


def safe_filename(name: str) -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", name).strip()
    return cleaned or "score_table"


def value_to_text(value: int | float | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        return format(value, ".15g")
    return str(value)


def write_txt(path: Path, rows: Sequence[OutputRow]) -> None:
    lines = [f"{row.label}\t{value_to_text(row.value)}" for row in rows]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")


def xml_text(value: str) -> str:
    return escape(value, {"\n": "&#10;"})


def cell_ref(row_number: int, column_number: int) -> str:
    name = ""
    number = column_number
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return f"{name}{row_number}"


def inline_string_cell(row_number: int, column_number: int, value: str, style_id: int) -> str:
    reference = cell_ref(row_number, column_number)
    return (
        f'<c r="{reference}" s="{style_id}" t="inlineStr">'
        f"<is><t>{xml_text(value)}</t></is>"
        "</c>"
    )


def value_cell(row_number: int, column_number: int, value: int | float | str | None, style_id: int) -> str:
    reference = cell_ref(row_number, column_number)
    if value is None:
        return f'<c r="{reference}" s="{style_id}"/>'
    if isinstance(value, int):
        return f'<c r="{reference}" s="{style_id}"><v>{value}</v></c>'
    if isinstance(value, float):
        return f'<c r="{reference}" s="{style_id}"><v>{format(value, ".15g")}</v></c>'
    return inline_string_cell(row_number, column_number, value, style_id)


def row_xml(row_number: int, row: OutputRow) -> str:
    if row.kind == "part":
        label_style = 1
        value_style = 5
        height = 22
    elif row.kind == "section":
        label_style = 2
        value_style = 6
        height = 21
    else:
        label_style = 3
        value_style = 4
        height = 20

    cells = [
        inline_string_cell(row_number, 1, row.label, label_style),
        value_cell(row_number, 2, row.value, value_style),
    ]
    return (
        f'<row r="{row_number}" ht="{height}" customHeight="1">'
        f"{''.join(cells)}"
        "</row>"
    )


def worksheet_xml(rows: Sequence[OutputRow]) -> str:
    rows_body = "\n".join(row_xml(index, row) for index, row in enumerate(rows, start=1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>
    <col min="1" max="1" width="42" customWidth="1"/>
    <col min="2" max="2" width="16" customWidth="1"/>
  </cols>
  <sheetData>
{rows_body}
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="SimSun"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="SimSun"/><family val="2"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAF2F8"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7FBFF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def workbook_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="评分表" sheetId="1" r:id="rId1"/></sheets>
</workbook>
"""


def workbook_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def core_props_xml() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>ai4sci</dc:creator>
  <cp:lastModifiedBy>ai4sci</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""


def app_props_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>ai4sci export_result_scores.py</Application>
</Properties>
"""


def write_xlsx(path: Path, rows: Sequence[OutputRow]) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml())
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml(rows))
        archive.writestr("xl/styles.xml", styles_xml())
        archive.writestr("docProps/core.xml", core_props_xml())
        archive.writestr("docProps/app.xml", app_props_xml())


def export_one(
    json_path: Path,
    output_dir: Path,
    output_format: str,
    include_basic_info: bool,
) -> list[Path]:
    answers = read_answer_json(json_path)
    rows = build_rows(answers, include_basic_info=include_basic_info)
    institute = institute_name_from_json(json_path)
    output_stem = safe_filename(f"评分表_{institute}")

    created: list[Path] = []
    if output_format in {"xlsx", "both"}:
        output_path = output_dir / f"{output_stem}.xlsx"
        write_xlsx(output_path, rows)
        created.append(output_path)
    if output_format in {"txt", "both"}:
        output_path = output_dir / f"{output_stem}.txt"
        write_txt(output_path, rows)
        created.append(output_path)
    return created


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export result/*.json answers to compact xlsx or txt tables.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("result"),
        help="Directory containing answer JSON files. Default: result",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("score_tables"),
        help="Directory for generated files. Default: score_tables",
    )
    parser.add_argument(
        "--format",
        choices=("xlsx", "txt", "both"),
        default="xlsx",
        help="Output format. Default: xlsx",
    )
    parser.add_argument(
        "--include-basic-info",
        action="store_true",
        help="Also include key 2 under 第一部分. By default output starts at 第二部分.",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    input_dir: Path = args.input_dir
    output_dir: Path = args.output_dir

    if not input_dir.is_dir():
        raise SystemExit(f"input directory not found: {input_dir}")

    json_files = sorted(input_dir.glob("*.json"))
    if not json_files:
        raise SystemExit(f"no JSON files found in: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    all_outputs: list[Path] = []
    for json_path in json_files:
        all_outputs.extend(
            export_one(
                json_path=json_path,
                output_dir=output_dir,
                output_format=args.format,
                include_basic_info=args.include_basic_info,
            )
        )

    for output_path in all_outputs:
        print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
