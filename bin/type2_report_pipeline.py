#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from process_excel_to_word import (
    create_gemini_client,
    get_interaction_output_text,
    load_gemini_api_key,
    update_word_fields,
)


for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")


SUPPORTED_EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
SUMMARY_SHEET = "汇总-资产基础法"
CLASSIFICATION_SHEET = "2分类汇总"
DEFAULT_MODEL = "gemini-3.5-flash"

RESPONSE_SCHEMA: Dict[str, Any] = {
    "type": "object",
    "properties": {
        "report": {
            "type": "object",
            "properties": {
                "scope_intro": {"type": "array", "items": {"type": "string"}},
                "asset_descriptions": {
                    "type": "array",
                    "items": {"type": "string"},
                },
                "method_sections": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string"},
                            "paragraphs": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                        },
                        "required": ["subject", "paragraphs"],
                    },
                },
                "conclusion_intro": {
                    "type": "array",
                    "items": {"type": "string"},
                },
                "final_conclusion": {
                    "type": "array",
                    "items": {"type": "string"},
                },
            },
            "required": [
                "scope_intro",
                "asset_descriptions",
                "method_sections",
                "conclusion_intro",
                "final_conclusion",
            ],
        },
        "explanation": {
            "type": "object",
            "properties": {
                "scope_intro": {"type": "array", "items": {"type": "string"}},
                "physical_assets": {
                    "type": "array",
                    "items": {"type": "string"},
                },
                "technical_sections": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string"},
                            "paragraphs": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                        },
                        "required": ["subject", "paragraphs"],
                    },
                },
            },
            "required": [
                "scope_intro",
                "physical_assets",
                "technical_sections",
            ],
        },
        "warnings": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["report", "explanation", "warnings"],
}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_name(value: Any) -> str:
    return re.sub(r"[\s　:：()（）\-—_]+", "", clean_text(value)).lower()


def normalize_sheet_name(value: Any) -> str:
    return re.sub(r"\s+", "", clean_text(value))


def normalize_matrix(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [[value]]
    if value and not isinstance(value[0], tuple):
        return [list(value)]
    return [list(row) for row in value]


def numeric_value(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(",", "").replace("%", "")
    if not text or text in {"-", "—", "/"}:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if "%" in clean_text(value):
        number /= 100
    return number


def format_number(value: Any, decimals: int = 2) -> str:
    number = numeric_value(value)
    if number is None:
        return "-"
    return f"{number:,.{decimals}f}"


def format_rate(value: Any) -> str:
    number = numeric_value(value)
    if number is None:
        return "-"
    return f"{number * 100:.2f}%"


def sanitize_filename_part(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", clean_text(value))
    return cleaned.rstrip(". ") or "未命名"


def discover_excel_files(input_dir: Path) -> List[Path]:
    files = sorted(
        path
        for path in input_dir.rglob("*")
        if (
            path.is_file()
            and not path.name.startswith("~$")
            and path.suffix.lower() in SUPPORTED_EXCEL_SUFFIXES
        )
    )
    if not files:
        raise FileNotFoundError(f"项目目录中未找到 Excel 文件: {input_dir}")
    return files


@contextmanager
def excel_application() -> Iterator[Any]:
    if not sys.platform.startswith("win"):
        raise RuntimeError("类型 2 的 .xls 读取目前需要 Windows 和 Microsoft Excel")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass
        yield excel
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


@contextmanager
def open_workbook(excel: Any, path: Path) -> Iterator[Any]:
    workbook = excel.Workbooks.Open(
        str(path),
        UpdateLinks=0,
        ReadOnly=True,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
        Notify=False,
    )
    try:
        yield workbook
    finally:
        workbook.Close(False)


def visible_worksheets(workbook: Any) -> List[Any]:
    return [
        workbook.Worksheets.Item(index)
        for index in range(1, workbook.Worksheets.Count + 1)
        if int(workbook.Worksheets.Item(index).Visible) == -1
    ]


def find_worksheet(workbook: Any, expected_name: str) -> Optional[Any]:
    expected = normalize_sheet_name(expected_name)
    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets.Item(index)
        if normalize_sheet_name(sheet.Name) == expected:
            return sheet
    return None


def read_range(
    sheet: Any,
    *,
    max_rows: Optional[int] = None,
    max_columns: Optional[int] = None,
) -> List[List[Any]]:
    used = sheet.UsedRange
    first_row = int(used.Row)
    first_column = int(used.Column)
    row_count = int(used.Rows.Count)
    column_count = int(used.Columns.Count)
    if max_rows is not None:
        row_count = min(row_count, max_rows)
    if max_columns is not None:
        column_count = min(column_count, max_columns)
    if row_count < 1 or column_count < 1:
        return []
    value = sheet.Range(
        sheet.Cells(first_row, first_column),
        sheet.Cells(first_row + row_count - 1, first_column + column_count - 1),
    ).Value2
    return normalize_matrix(value)


def find_prefixed_value(matrix: Sequence[Sequence[Any]], label: str) -> str:
    normalized_label = normalize_name(label)
    for row in matrix[:12]:
        for cell in row:
            text = clean_text(cell)
            compact = normalize_name(text)
            if normalized_label not in compact:
                continue
            for separator in ("：", ":"):
                if separator in text:
                    return text.split(separator, 1)[1].strip()
    return ""


def find_header_row(
    matrix: Sequence[Sequence[Any]],
    required_headers: Sequence[str],
) -> int:
    required = [normalize_name(header) for header in required_headers]
    for index, row in enumerate(matrix[:20]):
        cells = [normalize_name(cell) for cell in row]
        if all(any(header in cell for cell in cells) for header in required):
            return index
    return -1


def header_column(row: Sequence[Any], expected: str) -> int:
    target = normalize_name(expected)
    for index, cell in enumerate(row):
        if target in normalize_name(cell):
            return index
    return -1


def row_value(row: Sequence[Any], index: int) -> Any:
    return row[index] if 0 <= index < len(row) else None


def parse_summary_matrix(matrix: Sequence[Sequence[Any]]) -> Dict[str, Any]:
    entity = find_prefixed_value(matrix, "被评估单位")
    benchmark_date = find_prefixed_value(matrix, "评估基准日")
    header_index = find_header_row(matrix, ("账面价值", "评估价值", "增减值"))
    if header_index < 0:
        raise ValueError(f"{SUMMARY_SHEET} 未找到汇总表头")
    header = matrix[header_index]
    book_col = header_column(header, "账面价值")
    assessed_col = header_column(header, "评估价值")
    change_col = header_column(header, "增减值")
    rate_col = header_column(header, "增值率")
    rows = []
    for source_row in matrix[header_index + 1 :]:
        sequence = clean_text(row_value(source_row, 0))
        project = clean_text(row_value(source_row, 1))
        if not project and source_row:
            candidate = clean_text(source_row[0])
            if candidate and not re.fullmatch(r"\d+", candidate):
                project = candidate
        if not project or project in {"A", "B", "C=B-A", "D=C/A×100%"}:
            continue
        book = numeric_value(row_value(source_row, book_col))
        assessed = numeric_value(row_value(source_row, assessed_col))
        change = numeric_value(row_value(source_row, change_col))
        rate = numeric_value(row_value(source_row, rate_col))
        if all(value is None for value in (book, assessed, change, rate)):
            continue
        rows.append(
            {
                "序号": sequence,
                "项目": project,
                "账面价值": book,
                "评估价值": assessed,
                "增减值": change,
                "增值率": rate,
            }
        )
    if not entity:
        raise ValueError(f"{SUMMARY_SHEET} 未找到被评估单位")
    return {
        "sheet": SUMMARY_SHEET,
        "unit": "人民币万元",
        "entity": entity,
        "benchmark_date": benchmark_date,
        "rows": rows,
    }


def parse_classification_matrix(
    matrix: Sequence[Sequence[Any]],
) -> Dict[str, Any]:
    header_index = find_header_row(matrix, ("科目名称", "账面价值", "评估价值"))
    if header_index < 0:
        raise ValueError(f"{CLASSIFICATION_SHEET} 未找到分类汇总表头")
    header = matrix[header_index]
    subject_col = header_column(header, "科目名称")
    book_col = header_column(header, "账面价值")
    assessed_col = header_column(header, "评估价值")
    change_col = header_column(header, "增减值")
    rate_col = header_column(header, "增值率")
    rows = []
    category = ""
    for source_row in matrix[header_index + 1 :]:
        sequence = clean_text(row_value(source_row, 0))
        category_marker = clean_text(row_value(source_row, 1))
        subject = clean_text(row_value(source_row, subject_col))
        if not subject:
            continue
        if category_marker:
            category = subject.replace("合计", "").strip()
        rows.append(
            {
                "序号": sequence,
                "类别": category,
                "科目名称": subject.replace("合计", "").strip(),
                "账面价值": numeric_value(row_value(source_row, book_col)),
                "评估价值": numeric_value(row_value(source_row, assessed_col)),
                "增减值": numeric_value(row_value(source_row, change_col)),
                "增值率": numeric_value(row_value(source_row, rate_col)),
            }
        )
    return {
        "sheet": CLASSIFICATION_SHEET,
        "unit": "人民币元",
        "rows": rows,
    }


def read_detail_candidate(excel: Any, path: Path) -> Dict[str, Any]:
    with open_workbook(excel, path) as workbook:
        summary_sheet = find_worksheet(workbook, SUMMARY_SHEET)
        classification_sheet = find_worksheet(workbook, CLASSIFICATION_SHEET)
        if summary_sheet is None or classification_sheet is None:
            raise ValueError("缺少汇总-资产基础法或2分类汇总")
        summary = parse_summary_matrix(
            read_range(summary_sheet, max_rows=120, max_columns=12)
        )
        classification = parse_classification_matrix(
            read_range(classification_sheet, max_rows=180, max_columns=12)
        )
        return {
            "file": str(path),
            "entity": summary["entity"],
            "benchmark_date": summary["benchmark_date"],
            "summary": summary,
            "classification": classification,
        }


def selection_score(input_dir: Path, candidate: Dict[str, Any]) -> Tuple[int, str]:
    path = Path(candidate["file"])
    entity = clean_text(candidate["entity"])
    root_name = normalize_name(input_dir.name)
    entity_name = normalize_name(entity)
    filename = normalize_name(path.stem)
    score = 0
    reasons = []
    if entity_name and (entity_name in root_name or root_name in entity_name):
        score += 100
        reasons.append("被评估单位与项目目录名称一致")
    if re.match(r"^3(?:-|_)", path.name):
        score += 40
        reasons.append("文件名以3-开头")
    elif re.match(r"^3\.1(?:-|_)", path.name):
        score += 35
        reasons.append("文件名以3.1-开头")
    if entity_name and any(
        token and token in filename
        for token in re.split(r"公司|股份|有限|科技|集团", entity_name)
    ):
        score += 15
        reasons.append("文件名包含被评估单位简称")
    return score, "；".join(reasons) or "按文件名顺序候选"


def choose_primary_workbook(
    input_dir: Path,
    candidates: Sequence[Dict[str, Any]],
) -> Dict[str, Any]:
    if not candidates:
        raise ValueError("未找到同时包含汇总-资产基础法和2分类汇总的评估明细表")
    ranked = []
    for candidate in candidates:
        score, reason = selection_score(input_dir, candidate)
        ranked.append((score, reason, candidate))
    ranked.sort(key=lambda item: (-item[0], item[2]["file"]))
    if len(ranked) > 1 and ranked[0][0] == ranked[1][0]:
        first = Path(ranked[0][2]["file"]).name
        second = Path(ranked[1][2]["file"]).name
        raise ValueError(
            "无法唯一确定母公司评估明细表，请调整文件名或目录名: "
            f"{first}, {second}"
        )
    primary = dict(ranked[0][2])
    primary["selection_reason"] = ranked[0][1]
    return primary


def company_aliases(candidate: Dict[str, Any]) -> List[str]:
    aliases = []
    entity = clean_text(candidate.get("entity"))
    entity_base = re.sub(
        r"(股份有限公司|有限责任公司|有限公司)$",
        "",
        entity,
    )
    for suffix in ("科技", "智慧科技", "无线科技", "投资", "集团"):
        if entity_base.endswith(suffix):
            entity_base = entity_base[: -len(suffix)]
            break
    if len(entity_base) >= 2:
        aliases.append(entity_base)

    stem = Path(clean_text(candidate.get("file"))).stem
    if "评估明细表-" in stem:
        stem = stem.split("评估明细表-", 1)[1]
    stem = re.sub(r"\d{4}(?:[.\-]\d+)*.*$", "", stem)
    stem = re.sub(r"(反馈修改稿|报告日.*)$", "", stem)
    stem = stem.strip("-_ ")
    if len(stem) >= 2 and stem not in aliases:
        aliases.append(stem)
    return aliases


def subject_from_sheet_name(name: str, aliases: Dict[str, str]) -> str:
    subject = re.sub(r"^\d+(?:\.\d+)*[.、\-]?", "", clean_text(name))
    subject = re.sub(r"(汇总表?|评估明细表)$", "", subject).strip()
    for alias, canonical in aliases.items():
        if normalize_name(alias) == normalize_name(subject):
            return canonical
    return subject


def combined_headers(
    matrix: Sequence[Sequence[Any]],
    header_index: int,
) -> List[str]:
    width = max((len(row) for row in matrix[: header_index + 2]), default=0)
    headers = []
    for column in range(width):
        parts = []
        for row_index in (header_index, header_index + 1):
            if row_index >= len(matrix):
                continue
            text = clean_text(row_value(matrix[row_index], column))
            if text and text not in parts:
                parts.append(text)
        headers.append("/".join(parts))
    return headers


def matching_column(headers: Sequence[str], candidates: Sequence[str]) -> int:
    for candidate in candidates:
        target = normalize_name(candidate)
        for index, header in enumerate(headers):
            if target in normalize_name(header):
                return index
    return -1


def summarize_subject_sheet(
    sheet: Any,
    aliases: Dict[str, str],
) -> Dict[str, Any]:
    matrix = read_range(sheet, max_rows=10000, max_columns=50)
    subject = subject_from_sheet_name(clean_text(sheet.Name), aliases)
    header_index = find_header_row(matrix, ("序号",))
    if header_index < 0:
        return {
            "subject": subject,
            "sheet": clean_text(sheet.Name),
            "item_count": None,
            "detail_headers": [],
            "content_summary": [],
            "location_summary": [],
            "representative_rows": [],
            "warnings": ["未找到序号表头"],
        }
    headers = combined_headers(matrix, header_index)
    sequence_col = matching_column(headers, ("序号",))
    content_col = matching_column(
        headers,
        (
            "业务内容",
            "设备名称",
            "名称及规格型号",
            "名称",
            "证书名称及证书号",
            "欠款单位名称",
        ),
    )
    location_col = matching_column(
        headers,
        ("存放地点", "坐落", "位置", "使用地点"),
    )
    age_col = matching_column(headers, ("账龄", "库龄"))
    content_counts: Counter[str] = Counter()
    location_counts: Counter[str] = Counter()
    age_counts: Counter[str] = Counter()
    max_sequence = 0
    representative_rows = []
    for row in matrix[header_index + 1 :]:
        sequence_text = clean_text(row_value(row, sequence_col))
        match = re.fullmatch(r"(\d+)(?:\.0+)?", sequence_text)
        if not match:
            continue
        sequence = int(match.group(1))
        max_sequence = max(max_sequence, sequence)
        content = clean_text(row_value(row, content_col))
        location = clean_text(row_value(row, location_col))
        age = clean_text(row_value(row, age_col))
        if content:
            content_counts[content] += 1
        if location:
            location_counts[location] += 1
        if age:
            age_counts[age] += 1
        if len(representative_rows) < 8:
            representative_rows.append(
                {
                    headers[index] or f"列{index + 1}": clean_text(value)
                    for index, value in enumerate(row[: min(len(row), len(headers))])
                    if clean_text(value)
                }
            )
    return {
        "subject": subject,
        "sheet": clean_text(sheet.Name),
        "item_count": max_sequence or None,
        "detail_headers": [header for header in headers if header],
        "content_summary": [
            {"value": value, "count": count}
            for value, count in content_counts.most_common(8)
        ],
        "location_summary": [
            {"value": value, "count": count}
            for value, count in location_counts.most_common(6)
        ],
        "main_age_bucket": age_counts.most_common(1)[0][0] if age_counts else "",
        "representative_rows": representative_rows,
        "warnings": [],
    }


def summarize_primary_workbook(
    excel: Any,
    primary: Dict[str, Any],
    rules: Dict[str, Any],
) -> Dict[str, Any]:
    path = Path(primary["file"])
    aliases = rules.get("subject_aliases", {})
    classification_by_subject = {
        normalize_name(row["科目名称"]): row
        for row in primary["classification"]["rows"]
    }
    with open_workbook(excel, path) as workbook:
        visible = visible_worksheets(workbook)
        hidden_names = [
            clean_text(workbook.Worksheets.Item(index).Name)
            for index in range(1, workbook.Worksheets.Count + 1)
            if int(workbook.Worksheets.Item(index).Visible) != -1
        ]
        subjects = []
        for sheet in visible:
            sheet_name = normalize_sheet_name(sheet.Name)
            if sheet_name in {
                normalize_sheet_name(SUMMARY_SHEET),
                normalize_sheet_name(CLASSIFICATION_SHEET),
            }:
                continue
            if "汇总" in sheet_name or sheet_name == "基本情况":
                continue
            summary = summarize_subject_sheet(sheet, aliases)
            classification = classification_by_subject.get(
                normalize_name(summary["subject"]),
                {},
            )
            summary.update(
                {
                    "category": classification.get("类别", ""),
                    "unit": "人民币元",
                    "book_value": classification.get("账面价值"),
                    "assessed_value": classification.get("评估价值"),
                    "increase_decrease": classification.get("增减值"),
                    "increase_rate": classification.get("增值率"),
                }
            )
            subjects.append(summary)
        return {
            "file": str(path),
            "entity": primary["entity"],
            "selection_reason": primary["selection_reason"],
            "visible_sheets": [clean_text(sheet.Name) for sheet in visible],
            "hidden_sheets_ignored": hidden_names,
            "subjects": subjects,
        }


def detect_method(texts: Iterable[str]) -> str:
    joined = "\n".join(texts)
    methods = [
        method
        for method in (
            "假设开发法",
            "上市公司比较法",
            "交易案例比较法",
            "市场法",
            "收益法",
            "重置成本法",
            "成本法",
            "资产基础法",
        )
        if method in joined
    ]
    return "、".join(methods)


def summarize_workpaper(excel: Any, path: Path) -> Dict[str, Any]:
    with open_workbook(excel, path) as workbook:
        sheets = visible_worksheets(workbook)
        key_text = []
        sheet_names = []
        for sheet in sheets:
            sheet_names.append(clean_text(sheet.Name))
            matrix = read_range(sheet, max_rows=15, max_columns=20)
            for row in matrix:
                for cell in row:
                    text = clean_text(cell)
                    if text and text not in key_text:
                        key_text.append(text)
                    if len(key_text) >= 80:
                        break
                if len(key_text) >= 80:
                    break
            if len(key_text) >= 80:
                break
        return {
            "file": str(path),
            "kind": detect_method([path.name, *sheet_names, *key_text]),
            "visible_sheets": sheet_names,
            "valuation_method": detect_method(key_text),
            "key_text": key_text[:40],
            "warnings": [],
        }


def build_input_data(
    input_dir: Path,
    excel_files: Sequence[Path],
    rules: Dict[str, Any],
) -> Dict[str, Any]:
    detail_paths = [path for path in excel_files if "评估明细表" in path.name]
    candidates = []
    candidate_errors = []
    with excel_application() as excel:
        for index, path in enumerate(detail_paths, start=1):
            print(f"[明细表 {index}/{len(detail_paths)}] {path.name}")
            try:
                candidates.append(read_detail_candidate(excel, path))
            except Exception as exc:
                candidate_errors.append(f"{path.name}: {exc}")
        primary = choose_primary_workbook(input_dir, candidates)
        print(f"母公司明细表: {Path(primary['file']).name}")
        print(f"被评估单位: {primary['entity']}")
        primary_details = summarize_primary_workbook(excel, primary, rules)

        related = [
            {
                "file": candidate["file"],
                "entity": candidate["entity"],
                "relationship": "subsidiary_or_related_entity",
                "used_for_parent_scope": False,
            }
            for candidate in candidates
            if candidate["file"] != primary["file"]
        ]
        related_candidates = [
            candidate
            for candidate in candidates
            if candidate["file"] != primary["file"]
        ]
        related_aliases = [
            alias
            for candidate in related_candidates
            for alias in company_aliases(candidate)
        ]
        all_workpaper_paths = [
            path
            for path in excel_files
            if path not in detail_paths
        ]
        workpaper_paths = [
            path
            for path in all_workpaper_paths
            if not any(alias in path.stem for alias in related_aliases)
        ]
        excluded_related_workpapers = [
            str(path)
            for path in all_workpaper_paths
            if path not in workpaper_paths
        ]
        workpapers = []
        for index, path in enumerate(workpaper_paths, start=1):
            print(f"[工作底稿 {index}/{len(workpaper_paths)}] {path.name}")
            try:
                workpapers.append(summarize_workpaper(excel, path))
            except Exception as exc:
                workpapers.append(
                    {
                        "file": str(path),
                        "kind": "",
                        "visible_sheets": [],
                        "valuation_method": "",
                        "key_text": [],
                        "warnings": [str(exc)],
                    }
                )

    classification_rows = primary["classification"]["rows"]
    excluded_subjects = set(rules.get("summary_categories", []))
    excluded_subjects.update({"负债合计", "净资产", "所有者权益"})
    subject_order = [
        row["科目名称"]
        for row in classification_rows
        if (
            numeric_value(row.get("账面价值")) not in (None, 0)
            or numeric_value(row.get("评估价值")) not in (None, 0)
        )
        and row["科目名称"] not in excluded_subjects
    ]
    physical_subjects = set(rules.get("physical_asset_subjects", []))
    physical_assets = [
        {
            "subject": item["subject"],
            "item_count": item["item_count"],
            "main_names": item["content_summary"],
            "locations": item["location_summary"],
            "source_sheets": [item["sheet"]],
            "source_workpapers": [],
        }
        for item in primary_details["subjects"]
        if item["subject"] in physical_subjects
        or any(
            marker in item["subject"]
            for marker in (
                "存货",
                "原材料",
                "产成品",
                "在产品",
                "周转材料",
                "发出商品",
                "委托加工物资",
                "房屋",
                "构筑物",
                "设备",
                "车辆",
                "土地",
                "在建工程",
            )
        )
    ]
    return {
        "schema_version": "1.0",
        "document_type": "2",
        "source_type": "valuation_project_folder",
        "source_root": str(input_dir),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "project": {
            "被评估单位": primary["entity"],
            "评估基准日": primary["benchmark_date"],
            "评估对象": f"{primary['entity']}于评估基准日的股东全部权益价值",
            "最终评估方法": "",
        },
        "primary_workbook": primary_details,
        "related_workbooks": related,
        "asset_basis_summary": primary["summary"],
        "classification_summary": primary["classification"],
        "subject_order": subject_order,
        "subjects": primary_details["subjects"],
        "physical_assets": physical_assets,
        "off_balance_assets": [],
        "workpapers": workpapers,
        "excluded_related_workpapers": excluded_related_workpapers,
        "warnings": candidate_errors,
        "errors": [],
    }


def select_method_library(
    library: Dict[str, Any],
    subject_order: Sequence[str],
) -> List[Dict[str, Any]]:
    aliases = library.get("aliases", {})
    entries = library.get("entries", [])
    selected = []
    used = set()
    for subject in subject_order:
        target = aliases.get(subject, subject)
        target_names = [normalize_name(target)]
        if any(marker in subject for marker in ("房屋", "构筑物")):
            target_names.append(normalize_name("房屋建（构）筑物"))
        if "技术类无形资产" in subject:
            target_names.extend(
                (
                    normalize_name("专利权"),
                    normalize_name("商标权"),
                )
            )
        matches = [
            entry
            for entry in entries
            if any(
                target_name in normalize_name(entry.get("subject"))
                or normalize_name(entry.get("subject")) in target_name
                for target_name in target_names
            )
        ]
        for entry in matches:
            key = entry.get("subject")
            if key in used:
                continue
            used.add(key)
            selected.append(
                {
                    "subject": key,
                    "plain_text": clean_text(entry.get("plain_text"))[:16000],
                }
            )
    return selected


def compact_ai_input(data: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "project": data["project"],
        "asset_basis_summary": data["asset_basis_summary"],
        "classification_summary": data["classification_summary"],
        "subject_order": data["subject_order"],
        "subjects": [
            {
                key: item.get(key)
                for key in (
                    "category",
                    "subject",
                    "sheet",
                    "unit",
                    "book_value",
                    "assessed_value",
                    "increase_decrease",
                    "increase_rate",
                    "item_count",
                    "content_summary",
                    "location_summary",
                    "main_age_bucket",
                    "representative_rows",
                    "warnings",
                )
            }
            for item in data["subjects"]
        ],
        "physical_assets": data["physical_assets"],
        "workpapers": data["workpapers"],
        "method_library": data.get("method_library", []),
        "warnings": data["warnings"],
    }


def call_gemini(
    *,
    api_key: str,
    model: str,
    thinking_level: str,
    prompt_text: str,
    data: Dict[str, Any],
    store: bool,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    runtime_prompt = (
        f"{prompt_text}\n\n"
        "## 本次输出结构补充要求\n"
        "只返回 JSON。scope_intro 只写范围表之前的段落；"
        "asset_descriptions 写范围表之后的委估资产负债、表外资产及引用报告段落；"
        "method_sections 和 technical_sections 按 subject_order 顺序输出。"
        "不要在段落中重复生成表格，汇总表由程序根据源 Excel 写入。"
        "final_conclusion.method 为空时不得选择最终评估方法。\n\n"
        "## 本次结构化数据\n"
        + json.dumps(compact_ai_input(data), ensure_ascii=False, indent=2)
    )
    client = create_gemini_client(api_key)
    interaction = client.create(
        model=model,
        input=runtime_prompt,
        response_format={
            "type": "text",
            "mime_type": "application/json",
            "schema": RESPONSE_SCHEMA,
        },
        generation_config={"thinking_level": thinking_level},
        store=store,
    )
    if interaction.get("status") == "failed":
        raise RuntimeError(
            f"Gemini interaction 失败: {interaction.get('error') or interaction}"
        )
    output_text = get_interaction_output_text(interaction)
    if not output_text:
        raise RuntimeError("Gemini 返回了空响应")
    try:
        payload = json.loads(output_text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Gemini 返回的结构化 JSON 无法解析: {exc}") from exc
    metadata = {
        "interaction_id": clean_text(interaction.get("id")),
        "model": clean_text(interaction.get("model")) or model,
        "usage": interaction.get("usage")
        if isinstance(interaction.get("usage"), dict)
        else {},
    }
    return payload, metadata


def paragraph_before(anchor: Paragraph, text: str, style: Optional[str] = None) -> Paragraph:
    element = OxmlElement("w:p")
    anchor._p.addprevious(element)
    paragraph = Paragraph(element, anchor._parent)
    if style:
        try:
            paragraph.style = style
        except KeyError:
            pass
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_after = Pt(0)
    if style not in {"报告章标题", "报告节标题"}:
        paragraph.paragraph_format.first_line_indent = Cm(0.74)
    run = paragraph.add_run(clean_text(text))
    run.font.name = "宋体" if style not in {"报告章标题", "报告节标题"} else "黑体"
    run._element.rPr.rFonts.set(
        qn("w:eastAsia"),
        "宋体" if style not in {"报告章标题", "报告节标题"} else "黑体",
    )
    run.font.size = Pt(12 if style not in {"报告章标题", "报告节标题"} else 14)
    run.font.bold = style in {"报告章标题", "报告节标题"}
    return paragraph


def table_before(
    document: Document,
    anchor: Paragraph,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
) -> Table:
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = clean_text(value)
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "宋体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    run.font.size = Pt(9)
    anchor._p.addprevious(table._tbl)
    return table


def find_placeholder(document: Document, placeholder: str) -> Paragraph:
    expected = "{{" + placeholder + "}}"
    for paragraph in document.paragraphs:
        if paragraph.text.strip() == expected:
            return paragraph
    raise ValueError(f"模板中未找到占位符: {expected}")


def remove_paragraph(paragraph: Paragraph) -> None:
    paragraph._element.getparent().remove(paragraph._element)


def replace_header_placeholder(document: Document, name: str, value: str) -> None:
    token = "{{" + name + "}}"
    for section in document.sections:
        for paragraph in section.header.paragraphs:
            if token in paragraph.text:
                for run in paragraph.runs:
                    if token in run.text:
                        run.text = run.text.replace(token, value)


def scope_table_rows(summary_rows: Sequence[Dict[str, Any]]) -> List[List[str]]:
    fixed = {
        "流动资产",
        "非流动资产",
        "资产总计",
        "流动负债",
        "非流动负债",
        "负债总计",
        "股东全部权益价值",
        "股东全部权益价值（单体）",
        "股东全部权益价值（合并）",
        "净资产（所有者权益）",
    }
    rows = []
    for row in summary_rows:
        project = clean_text(row.get("项目"))
        book = numeric_value(row.get("账面价值"))
        if project not in fixed and book in (None, 0):
            continue
        rows.append([project, format_number(book)])
    return rows


def conclusion_table_rows(
    summary_rows: Sequence[Dict[str, Any]],
) -> List[List[str]]:
    rows = []
    for row in summary_rows:
        book = numeric_value(row.get("账面价值"))
        assessed = numeric_value(row.get("评估价值"))
        if book in (None, 0) and assessed in (None, 0):
            continue
        rows.append(
            [
                clean_text(row.get("项目")),
                format_number(book),
                format_number(assessed),
                format_number(row.get("增减值")),
                format_rate(row.get("增值率")),
            ]
        )
    return rows


def add_paragraphs(anchor: Paragraph, paragraphs: Any) -> None:
    if not isinstance(paragraphs, list):
        return
    for text in paragraphs:
        if clean_text(text):
            paragraph_before(anchor, clean_text(text))


def add_section_objects(anchor: Paragraph, sections: Any) -> None:
    if not isinstance(sections, list):
        return
    for section in sections:
        if not isinstance(section, dict):
            continue
        subject = clean_text(section.get("subject"))
        if subject:
            paragraph_before(anchor, subject, "报告节标题")
        add_paragraphs(anchor, section.get("paragraphs"))


def render_report(
    template_file: Path,
    output_file: Path,
    data: Dict[str, Any],
    generated: Dict[str, Any],
) -> None:
    document = Document(template_file)
    summary_rows = data["asset_basis_summary"]["rows"]
    report = generated["report"]

    scope_anchor = find_placeholder(document, "报告_评估对象和评估范围")
    add_paragraphs(scope_anchor, report.get("scope_intro"))
    paragraph_before(scope_anchor, "金额单位：人民币万元")
    table_before(
        document,
        scope_anchor,
        ("项目", "账面价值"),
        scope_table_rows(summary_rows),
    )
    add_paragraphs(scope_anchor, report.get("asset_descriptions"))
    remove_paragraph(scope_anchor)

    method_anchor = find_placeholder(
        document,
        "报告_资产基础法具体评估方法介绍",
    )
    add_section_objects(method_anchor, report.get("method_sections"))
    remove_paragraph(method_anchor)

    conclusion_anchor = find_placeholder(document, "报告_评估结论")
    add_paragraphs(conclusion_anchor, report.get("conclusion_intro"))
    paragraph_before(conclusion_anchor, "金额单位：人民币万元")
    table_before(
        document,
        conclusion_anchor,
        ("项目", "账面价值", "评估价值", "增减值", "增值率"),
        conclusion_table_rows(summary_rows),
    )
    add_paragraphs(conclusion_anchor, report.get("final_conclusion"))
    remove_paragraph(conclusion_anchor)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_file)


def render_explanation(
    template_file: Path,
    output_file: Path,
    data: Dict[str, Any],
    generated: Dict[str, Any],
) -> None:
    document = Document(template_file)
    explanation = generated["explanation"]
    summary_rows = data["asset_basis_summary"]["rows"]
    entity = clean_text(data["project"]["被评估单位"])
    replace_header_placeholder(document, "项目名称", entity)

    scope_anchor = find_placeholder(document, "说明_评估对象与评估范围说明")
    add_paragraphs(scope_anchor, explanation.get("scope_intro"))
    paragraph_before(scope_anchor, "金额单位：人民币万元")
    table_before(
        document,
        scope_anchor,
        ("项目", "账面价值"),
        scope_table_rows(summary_rows),
    )
    remove_paragraph(scope_anchor)

    physical_anchor = find_placeholder(
        document,
        "说明_实物资产分布情况及特点",
    )
    add_paragraphs(physical_anchor, explanation.get("physical_assets"))
    remove_paragraph(physical_anchor)

    technical_anchor = find_placeholder(
        document,
        "说明_资产基础法评估技术说明",
    )
    add_section_objects(technical_anchor, explanation.get("technical_sections"))
    remove_paragraph(technical_anchor)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_file)


def build_parser(project_root: Path) -> argparse.ArgumentParser:
    resource_root = (
        Path(getattr(sys, "_MEIPASS", "")) / "resources" / "2"
        if bool(getattr(sys, "frozen", False))
        else project_root / "bin" / "template" / "2"
    )
    parser = argparse.ArgumentParser(
        description="Generate type 2 enterprise valuation report documents."
    )
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--word-dir", default=str(project_root / "word"))
    parser.add_argument(
        "--records-file",
        default=str(project_root / "bin" / "json" / "类型2输入数据.json"),
    )
    parser.add_argument(
        "--report-template",
        default=str(resource_root / "企业价值评估报告-自动生成基底模板.docx"),
    )
    parser.add_argument(
        "--explanation-template",
        default=str(resource_root / "企业价值评估说明-自动生成基底模板.docx"),
    )
    parser.add_argument(
        "--prompt-file",
        default=str(resource_root / "企业价值评估报告自动生成-Prompt.md"),
    )
    parser.add_argument(
        "--rules-file",
        default=str(resource_root / "企业价值评估报告生成规则.json"),
    )
    parser.add_argument(
        "--method-library",
        default=str(resource_root / "资产基础法评估方法库.json"),
    )
    parser.add_argument("--model", default=None)
    parser.add_argument(
        "--api-key-file",
        default=str(project_root / "gemini_api.txt"),
    )
    parser.add_argument("--extract-only", action="store_true")
    parser.add_argument("--no-word-field-update", action="store_true")
    return parser


def get_project_root() -> Path:
    if not bool(getattr(sys, "frozen", False)):
        return Path(__file__).resolve().parent.parent
    executable_dir = Path(sys.executable).resolve().parent
    if (
        executable_dir.name.lower() == "dist"
        and (executable_dir.parent / "bin").is_dir()
    ):
        return executable_dir.parent
    return executable_dir


def main(argv: Optional[Sequence[str]] = None) -> int:
    project_root = get_project_root()
    args = build_parser(project_root).parse_args(argv)
    input_dir = Path(args.input_dir).expanduser().resolve()
    word_dir = Path(args.word_dir).expanduser().resolve()
    records_file = Path(args.records_file).expanduser().resolve()
    report_template = Path(args.report_template).expanduser().resolve()
    explanation_template = Path(args.explanation_template).expanduser().resolve()
    prompt_file = Path(args.prompt_file).expanduser().resolve()
    rules_file = Path(args.rules_file).expanduser().resolve()
    method_library_file = Path(args.method_library).expanduser().resolve()
    api_key_file = Path(args.api_key_file).expanduser().resolve()

    if not input_dir.is_dir():
        raise FileNotFoundError(f"项目目录不存在: {input_dir}")
    for label, path in (
        ("评估报告模板", report_template),
        ("评估说明模板", explanation_template),
        ("Prompt", prompt_file),
        ("生成规则", rules_file),
        ("评估方法库", method_library_file),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label}不存在: {path}")

    rules = json.loads(rules_file.read_text(encoding="utf-8-sig"))
    prompt_text = prompt_file.read_text(encoding="utf-8-sig").strip()
    method_library = json.loads(
        method_library_file.read_text(encoding="utf-8-sig")
    )
    excel_files = discover_excel_files(input_dir)
    print(f"Excel files found: {len(excel_files)}")
    data = build_input_data(input_dir, excel_files, rules)
    method_subjects = [
        *data["subject_order"],
        *(item["subject"] for item in data["subjects"]),
    ]
    data["method_library"] = select_method_library(
        method_library,
        method_subjects,
    )
    records_file.parent.mkdir(parents=True, exist_ok=True)
    records_file.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"类型2 JSON: {records_file}")
    if args.extract_only:
        print("Extract-only mode: Word generation skipped.")
        return 0

    ai_config = rules.get("ai", {})
    model = clean_text(args.model) or clean_text(ai_config.get("model")) or DEFAULT_MODEL
    thinking_level = clean_text(ai_config.get("thinking_level")) or "medium"
    store = bool(ai_config.get("store_interactions", False))
    api_key = load_gemini_api_key(api_key_file)
    print(f"Calling Gemini: {model}")
    generated, interaction = call_gemini(
        api_key=api_key,
        model=model,
        thinking_level=thinking_level,
        prompt_text=prompt_text,
        data=data,
        store=store,
    )
    data["generated_sections"] = generated
    data["gemini"] = interaction
    records_file.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    entity = sanitize_filename_part(data["project"]["被评估单位"])
    benchmark = sanitize_filename_part(data["project"]["评估基准日"])
    report_file = word_dir / f"1-评估报告-{entity}-{benchmark}.docx"
    explanation_file = word_dir / f"2-评估说明-{entity}-{benchmark}.docx"
    render_report(report_template, report_file, data, generated)
    render_explanation(explanation_template, explanation_file, data, generated)

    field_warnings = []
    if not args.no_word_field_update:
        for output_file in (report_file, explanation_file):
            warning = update_word_fields(output_file)
            if warning:
                field_warnings.append(f"{output_file.name}: {warning}")

    manifest = {
        "document_type": "2",
        "source_type": "valuation_project_folder",
        "source_root": str(input_dir),
        "records_file": str(records_file),
        "reports": [
            {"kind": "report", "filename": report_file.name},
            {"kind": "explanation", "filename": explanation_file.name},
        ],
        "gemini": interaction,
        "warnings": [*data.get("warnings", []), *generated.get("warnings", []), *field_warnings],
        "errors": [],
    }
    word_dir.mkdir(parents=True, exist_ok=True)
    manifest_file = word_dir / "generation_manifest.json"
    manifest_file.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Saved: {report_file}")
    print(f"Saved: {explanation_file}")
    print(f"Manifest: {manifest_file}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
