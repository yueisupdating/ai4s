#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


EXPECTED_SINGLE_CHOICE_KEYS = {
    "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
    "16", "18", "19", "20", "21", "22", "23", "24", "26", "28", "29",
    "30", "31", "32", "33", "34", "37", "38",
}
EXPECTED_MULTI_CHOICE_SCORE_KEYS = {
    "2", "15", "17", "25", "27", "35", "36", "39", "40", "41", "42",
    "43", "44", "45", "46",
}
EXPECTED_SCORE_KEYS = {str(number) for number in range(2, 47)}

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SECTION_TITLES = {
    "受访者基本信息",
    "AI赋能科研就绪度调研",
    "数据资源就绪度",
    "计算资源就绪度",
    "模型与平台就绪度",
    "科研创新就绪度",
    "人才储备就绪度",
    "组织机制就绪度",
    "开放合作就绪度",
    "整体就绪度",
    "AI赋能科研需求调研",
}


def docx_lines(path: Path) -> list[str]:
    if not path.is_file() or path.stat().st_size == 0:
        raise SystemExit(f"DOCX is missing or empty: {path}")

    with ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    root = ET.fromstring(xml)

    lines: list[str] = []
    for paragraph in root.findall(".//w:p", WORD_NS):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", WORD_NS)).strip()
        if text:
            lines.append(text)
    return lines


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text.replace("☐", "").replace("⬜", ""))


def is_source_question_line(text: str) -> bool:
    if not text or text in SECTION_TITLES:
        return False
    if text.startswith("院内科研单位AI4S") or text.startswith("注意："):
        return False
    if text.startswith("您所在的机构名称") or text.startswith("您所在机构/团队"):
        return True
    if "？" in text:
        return True
    if text.endswith("：") and (text.startswith("您所在机构") or text.startswith("AI为")):
        return True
    if text.startswith("所在领域"):
        return True
    if text.startswith("从整体上看"):
        return True
    if text.startswith("在") and "认为" in text:
        return True
    if text.startswith("请从以下可选维度"):
        return True
    return False


def find_source_questionnaire(path: Path) -> Path | None:
    search_dirs = [Path.cwd(), path.parent, path.parent.parent]
    seen: set[Path] = set()
    for directory in search_dirs:
        directory = directory.resolve()
        if directory in seen or not directory.is_dir():
            continue
        seen.add(directory)
        candidates = [
            candidate
            for candidate in directory.glob("院内科研单位AI4S*.docx")
            if candidate.resolve() != path.resolve()
        ]
        if candidates:
            return candidates[0]
    return None


def source_question_blocks(source_path: Path) -> list[dict[str, object]]:
    blocks: list[dict[str, object]] = []
    current: dict[str, object] | None = None

    for text in docx_lines(source_path):
        if text in SECTION_TITLES or text.startswith("注意："):
            continue
        if is_source_question_line(text):
            if current is not None:
                blocks.append(current)
            current = {
                "number": len(blocks) + 1,
                "question": text,
                "options": [],
            }
            continue
        if current is not None:
            current["options"].append(text)

    if current is not None:
        blocks.append(current)
    return blocks


def validate_docx_source_form_text(path: Path, source_path: Path | None = None) -> None:
    source_path = source_path or find_source_questionnaire(path)
    if source_path is None:
        raise SystemExit(f"source questionnaire DOCX was not found near {path}")

    source_blocks = source_question_blocks(source_path)
    if len(source_blocks) != 47:
        raise SystemExit(
            f"source questionnaire parsing failed: expected 47 questions, "
            f"found={len(source_blocks)}, source={source_path}"
        )

    body = normalize_text("\n".join(docx_lines(path)))
    missing: list[str] = []

    for block in source_blocks:
        number = int(block["number"])
        required_lines = [str(block["question"]), *[str(option) for option in block["options"]]]
        for required in required_lines:
            normalized = normalize_text(required)
            if normalized and normalized not in body:
                missing.append(f"Q{number}: {required[:80]}")
                if len(missing) >= 12:
                    break
        if len(missing) >= 12:
            break

    if missing:
        raise SystemExit(
            "DOCX must retain original source questionnaire question text and option text; "
            f"missing_examples={missing}"
        )


def validate_score_json(path: Path) -> None:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)

    if not isinstance(data, dict):
        raise SystemExit('score JSON must be an object like {"2": 3, "3": 1}')

    bad = {
        key: value
        for key, value in data.items()
        if type(value) is not int or value not in {1, 2, 3, 4}
    }
    if bad:
        raise SystemExit(f"score JSON values must be JSON numbers 1/2/3/4 only: {bad}")

    actual_keys = set(data)
    missing = sorted(EXPECTED_SCORE_KEYS - actual_keys, key=int)
    extra = sorted(actual_keys - EXPECTED_SCORE_KEYS, key=lambda item: (not item.isdigit(), int(item) if item.isdigit() else item))
    if missing or extra:
        raise SystemExit(
            "score JSON keys must exactly match expected 47-question form keys 2 through 46; "
            f"missing={missing}, extra={extra}"
        )

    with path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def validate_docx_question_coverage(path: Path) -> None:
    body = "\n".join(docx_lines(path))
    found: set[int] = set()
    patterns = (
        re.compile(r"(?<![A-Za-z0-9])Q\s*0*([1-9]|[1-3][0-9]|4[0-7])(?![0-9])", re.I),
        re.compile(r"(?m)^\s*(?:问题\s*)?0*([1-9]|[1-3][0-9]|4[0-7])\s*[、.．:：]"),
        re.compile(r"(?m)^\s*第\s*0*([1-9]|[1-3][0-9]|4[0-7])\s*[题问]"),
    )
    for pattern in patterns:
        found.update(int(match.group(1)) for match in pattern.finditer(body))

    missing = [number for number in range(1, 48) if number not in found]
    if missing:
        raise SystemExit(
            f"DOCX does not cover all 47 questions: {path}; "
            f"found={len(found)}, missing={missing}"
        )


def validate_docx_q47(path: Path) -> None:
    lines = docx_lines(path)

    start = None
    for index, line in enumerate(lines):
        stripped = line.strip()
        if (
            stripped.upper().startswith("Q47")
            or stripped.startswith("47")
            or "最核心的科研组织管理方面创新诉求" in stripped
            or "请从以下可选维度中选出" in stripped
        ):
            start = index
            break

    if start is None:
        raise SystemExit(f"DOCX is missing Q47 open-ended demand ranking: {path}")

    segment = "\n".join(lines[start:])
    answer_positions = [
        position
        for marker in ("答案", "【答案】", "**答案")
        if (position := segment.find(marker)) >= 0
    ]
    if answer_positions:
        segment = segment[min(answer_positions):]

    allowed_dimensions = [
        "建立扁平化组织架构",
        "建立跨学科分布式团队",
        "改革考核评价机制",
        "明确知识产权归属",
        "优化资源配置",
        "引进与培养复合型人才",
        "培育开放合作与协同文化",
    ]
    ranks = ["第一", "第二", "第三", "第四"]
    items: list[tuple[str, str]] = []
    for rank in ranks:
        marker_index = segment.find(rank)
        if marker_index < 0:
            continue
        following = [
            candidate
            for next_rank in ranks
            if (candidate := segment.find(next_rank, marker_index + len(rank))) >= 0
        ]
        end = min(following) if following else len(segment)
        items.append((rank, segment[marker_index:end]))

    if len(items) < 2:
        raise SystemExit(
            f"Q47 must rank at least two demands using 第一/第二/...: {path}; "
            f"found={len(items)}"
        )

    selected: list[str] = []
    bad: list[str] = []
    for rank, item in items:
        matches = [dimension for dimension in allowed_dimensions if dimension in item]
        if not matches:
            bad.append(f"{rank}: {item[:120]}")
        else:
            selected.append(matches[0])

    if bad:
        raise SystemExit(
            "Q47 ranked items must each use one of the source questionnaire dimensions "
            f"{allowed_dimensions}; bad_items={bad}"
        )

    if len(set(selected)) != len(selected):
        raise SystemExit(f"Q47 ranked dimensions must not repeat: {path}; selected={selected}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate ai4sci generated questionnaire outputs.")
    parser.add_argument("check", choices=["score-json", "single-choice", "docx-coverage", "docx-q47", "docx-source-text"])
    parser.add_argument("paths", nargs="+", type=Path)
    args = parser.parse_args()

    if args.check in {"score-json", "single-choice"}:
        validate_score_json(args.paths[0])
    elif args.check == "docx-coverage":
        validate_docx_question_coverage(args.paths[0])
    elif args.check == "docx-q47":
        validate_docx_q47(args.paths[0])
    elif args.check == "docx-source-text":
        if len(args.paths) == 1:
            validate_docx_source_form_text(args.paths[0])
        elif len(args.paths) == 2:
            validate_docx_source_form_text(args.paths[1], args.paths[0])
        else:
            raise SystemExit("docx-source-text expects generated DOCX, or source DOCX plus generated DOCX")
    else:
        raise AssertionError(args.check)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
