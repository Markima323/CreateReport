#!/usr/bin/env python3
"""
Generate one Word value-analysis report per non-empty Sheet1 row.

Inputs default to:
1) bin/template/价值分析报告-自动生成基底模板.docx
2) bin/template/价值分析报告自动生成-Prompt.md
3) bin/template/价值分析报告生成规则.json
4) excel/数据表.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape as xml_escape

from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

try:
    import openpyxl
except Exception:  # pragma: no cover - import guard
    print(
        "Missing dependency: openpyxl. Install with: pip install -r bin/requirements.txt",
        file=sys.stderr,
    )
    raise

try:
    from docx import Document
except Exception:  # pragma: no cover - import guard
    print(
        "Missing dependency: python-docx. Install with: pip install -r bin/requirements.txt",
        file=sys.stderr,
    )
    raise


DEFAULT_PROJECT_VALUES: Dict[str, Any] = {
    "项目年": 2025,
    "报告年": 2025,
    "报告批次": "1017",
    "分析基准日": "2025年6月30日",
    "报告日期": "2025年10月16日",
    "报告日期大写": "二〇二五年十月十六日",
    "有效期截止日": "2025年12月30日",
    "折现率": 0.03,
}

MONEY_TOLERANCE = Decimal("1")
COEFFICIENT_TOLERANCE = Decimal("0.0001")
PERCENT_RATIO_TOLERANCE = Decimal("0.0001")

FIELD_COLUMNS: Dict[str, int] = {
    "序号": 1,
    "支行": 2,
    "姓名": 3,
    "基准日本金": 4,
    "基准日利息": 5,
    "基准日本息": 6,
    "地址": 7,
    "面积": 8,
    "第一顺位": 9,
    "登记类型": 10,
    "性质": 11,
    "权利价值": 12,
    "预计回收": 13,
    "评估单价": 14,
    "评估价值": 15,
    "快速变现系数": 16,
    "快速变现价值": 17,
    "清收完成时间": 18,
    "折现期限": 19,
    "折现系数": 20,
    "折现价值": 21,
    "处置费用": 22,
    "安置费用": 23,
    "可回收金额": 24,
    "债权评估值": 25,
    "偿债率": 26,
    "权证编号": 29,
    "证载权利人": 30,
    "保证人": 51,
}

DEBTOR_COLUMNS: List[Tuple[str, int, int, int, int, int]] = [
    ("债务人1", 31, 32, 33, 34, 35),
    ("债务人2", 36, 37, 38, 39, 40),
    ("债务人3", 41, 42, 43, 44, 45),
    ("债务人4", 46, 47, 48, 49, 50),
]

ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:\*\?"<>\|]')


@dataclass
class Debtor:
    source_index: int
    label: str
    name: str
    gender: str
    ethnicity: str
    id_number: str
    address: str


@dataclass
class PreparedReport:
    row_number: int
    sequence: str
    branch: str
    name: str
    report_number: str
    filename: str
    debtor_count: int
    guarantor_included: bool
    placement_cost_included: bool
    minimum_item: str
    replacements: Dict[str, str]
    conclusion_text: str
    guarantor_extra_lines: List[str]
    warnings: List[str]


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return format_chinese_date(value)
    if isinstance(value, date):
        return format_chinese_date(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_chinese_date(value: date) -> str:
    return f"{value.year}年{value.month}月{value.day}日"


def to_decimal(value: Any, field_name: str) -> Decimal:
    if is_blank(value):
        raise ValueError(f"{field_name}为空")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        return Decimal(text[:-1]) / Decimal("100")
    try:
        return Decimal(text)
    except InvalidOperation:
        raise ValueError(f"{field_name}不是有效数字: {value}")


def optional_decimal(value: Any, default: Decimal = Decimal("0")) -> Decimal:
    if is_blank(value):
        return default
    return to_decimal(value, "数值")


def round_half_up(value: Decimal, places: int = 0) -> Decimal:
    quant = Decimal("1") if places == 0 else Decimal("1").scaleb(-places)
    return value.quantize(quant, rounding=ROUND_HALF_UP)


def fmt_yuan(value: Decimal) -> str:
    return format(round_half_up(value, 2), ",.2f")


def fmt_wanyuan(value: Decimal) -> str:
    return format(round_half_up(value / Decimal("10000"), 2), ",.2f")


def fmt_area(value: Decimal) -> str:
    return format(round_half_up(value, 2), ",.2f")


def fmt_decimal(value: Decimal, places: int) -> str:
    return format(round_half_up(value, places), f",.{places}f")


def fmt_plain_number(value: Decimal) -> str:
    quantized = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    text = format(quantized.normalize(), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def fmt_percent(value: Decimal) -> str:
    return f"{fmt_decimal(value * Decimal('100'), 2)}%"


def fmt_percent_compact(value: Decimal) -> str:
    return f"{fmt_plain_number(value * Decimal('100'))}%"


def sanitize_filename_part(value: str) -> Tuple[str, Optional[str]]:
    sanitized = ILLEGAL_FILENAME_CHARS.sub("-", value).strip()
    if sanitized != value:
        return sanitized, f"文件名非法字符已替换: {value} -> {sanitized}"
    return sanitized, None


def normalize_certificate_number(value: Any) -> str:
    return clean_text(value).replace("(", "（").replace(")", "）")


def id_card_checksum(first17: str) -> str:
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    checks = "10X98765432"
    total = sum(int(ch) * weight for ch, weight in zip(first17, weights))
    return checks[total % 11]


def validate_or_repair_id_card(value: Any) -> Tuple[str, Optional[str]]:
    text = clean_text(value).replace(" ", "").upper()
    if re.fullmatch(r"\d{17}", text):
        repaired = text + id_card_checksum(text)
        return repaired, f"身份证号码17位已按校验位补齐: {text} -> {repaired}"
    if not re.fullmatch(r"\d{17}[\dX]", text):
        raise ValueError(f"身份证号码格式错误: {text}")
    birth = text[6:14]
    try:
        datetime.strptime(birth, "%Y%m%d")
    except ValueError:
        raise ValueError(f"身份证号码出生日期无效: {text}")
    expected = id_card_checksum(text[:17])
    if text[-1] != expected:
        raise ValueError(f"身份证号码校验位错误: {text}，应为 {expected}")
    return text, None


def normalize_ethnicity(value: Any) -> str:
    text = clean_text(value)
    if text and not text.endswith("族"):
        return text + "族"
    return text


def extract_floor(address: str) -> str:
    marker = address.rfind("室")
    if marker < 0:
        raise ValueError(f"地址无法提取楼层，未找到“室”: {address}")
    before_room = address[:marker].strip()
    match = re.search(r"(\d+(?:-\d+)?)\s*$", before_room)
    if not match:
        raise ValueError(f"地址无法提取楼层，室号不明确: {address}")
    room_token = match.group(1)
    room = room_token.split("-")[-1]
    if len(room) < 3:
        raise ValueError(f"地址无法提取楼层，房号位数不足: {address}")
    floor = int(room) // 100
    if floor <= 0:
        raise ValueError(f"地址无法提取楼层，房号异常: {address}")
    return f"{floor}层"


def choose_cached_or_computed(
    values: Dict[str, Any],
    field_name: str,
    computed: Decimal,
    tolerance: Decimal,
    warnings: List[str],
) -> Decimal:
    raw = values.get(field_name)
    if is_blank(raw):
        warnings.append(f"{field_name}为空，已由规则补算为 {computed}")
        return computed
    cached = to_decimal(raw, field_name)
    if abs(cached - computed) > tolerance:
        warnings.append(f"{field_name}与规则复核值不一致: Excel={cached}, 复核={computed}")
    return cached


def get_row_values(ws: Any, row_number: int) -> Dict[str, Any]:
    values: Dict[str, Any] = {}
    for field, col in FIELD_COLUMNS.items():
        values[field] = ws.cell(row_number, col).value
    for i, (_, name_col, gender_col, ethnicity_col, id_col, addr_col) in enumerate(DEBTOR_COLUMNS, start=1):
        values[f"债务人{i}"] = ws.cell(row_number, name_col).value
        values[f"债务人{i}性别"] = ws.cell(row_number, gender_col).value
        values[f"债务人{i}民族"] = ws.cell(row_number, ethnicity_col).value
        values[f"债务人{i}身份证号码"] = ws.cell(row_number, id_col).value
        values[f"债务人{i}住址"] = ws.cell(row_number, addr_col).value
    return values


def build_debtors(values: Dict[str, Any], warnings: List[str]) -> List[Debtor]:
    raw_debtors: List[Tuple[int, str, str, str, str, str]] = []
    for i in range(1, 5):
        name = clean_text(values.get(f"债务人{i}"))
        if not name:
            continue
        gender = clean_text(values.get(f"债务人{i}性别"))
        ethnicity = normalize_ethnicity(values.get(f"债务人{i}民族"))
        address = clean_text(values.get(f"债务人{i}住址"))
        missing = [
            label
            for label, item in [
                ("性别", gender),
                ("民族", ethnicity),
                ("身份证号码", values.get(f"债务人{i}身份证号码")),
                ("住址", address),
            ]
            if is_blank(item)
        ]
        if missing:
            raise ValueError(f"债务人{i}{name}缺少字段: {', '.join(missing)}")
        id_number, warning = validate_or_repair_id_card(values.get(f"债务人{i}身份证号码"))
        if warning:
            warnings.append(f"债务人{i}{name}: {warning}")
        raw_debtors.append((i, name, gender, ethnicity, id_number, address))

    if not raw_debtors:
        raise ValueError("未找到债务人")

    labels = ["债务人"] if len(raw_debtors) == 1 else ["债务人一", "债务人二", "债务人三", "债务人四"]
    debtors: List[Debtor] = []
    for index, item in enumerate(raw_debtors):
        source_index, name, gender, ethnicity, id_number, address = item
        debtors.append(
            Debtor(
                source_index=source_index,
                label=labels[index],
                name=name,
                gender=gender,
                ethnicity=ethnicity,
                id_number=id_number,
                address=address,
            )
        )
    return debtors


def build_prepared_report(
    row_number: int,
    values: Dict[str, Any],
    project: Dict[str, Any],
    known_guarantors: Dict[str, Dict[str, Any]],
) -> PreparedReport:
    warnings: List[str] = []
    required = ["序号", "支行", "姓名", "地址", "面积", "权利价值", "预计回收", "评估单价", "权证编号", "证载权利人"]
    missing = [field for field in required if is_blank(values.get(field))]
    if missing:
        raise ValueError(f"必要字段为空: {', '.join(missing)}")

    sequence = clean_text(values["序号"])
    branch = clean_text(values["支行"])
    name = clean_text(values["姓名"])
    address = clean_text(values["地址"])
    floor = extract_floor(address)
    debtors = build_debtors(values, warnings)

    principal = to_decimal(values["基准日本金"], "基准日本金")
    interest = to_decimal(values["基准日利息"], "基准日利息")
    basis = choose_cached_or_computed(
        values,
        "基准日本息",
        principal + interest,
        MONEY_TOLERANCE,
        warnings,
    )
    if basis == 0:
        raise ValueError("基准日本息为0，无法计算偿债率")

    area = to_decimal(values["面积"], "面积")
    unit_price = to_decimal(values["评估单价"], "评估单价")
    market_value = choose_cached_or_computed(
        values,
        "评估价值",
        round_half_up(area * unit_price, 0),
        MONEY_TOLERANCE,
        warnings,
    )
    quick_coeff = to_decimal(values["快速变现系数"], "快速变现系数")
    quick_value = choose_cached_or_computed(
        values,
        "快速变现价值",
        round_half_up(market_value * quick_coeff, 0),
        MONEY_TOLERANCE,
        warnings,
    )
    discount_term = to_decimal(values["折现期限"], "折现期限")
    discount_rate = to_decimal(project["折现率"], "折现率")
    discount_coeff_calc = round_half_up(
        Decimal(str(float((Decimal("1") + discount_rate)) ** float(-discount_term))),
        4,
    )
    discount_coeff = choose_cached_or_computed(
        values,
        "折现系数",
        discount_coeff_calc,
        COEFFICIENT_TOLERANCE,
        warnings,
    )
    discounted_value = choose_cached_or_computed(
        values,
        "折现价值",
        round_half_up(quick_value * discount_coeff, 0),
        MONEY_TOLERANCE,
        warnings,
    )
    disposal_cost = optional_decimal(values.get("处置费用"))
    placement_cost = optional_decimal(values.get("安置费用"))
    recover_value = choose_cached_or_computed(
        values,
        "可回收金额",
        discounted_value - disposal_cost - placement_cost,
        MONEY_TOLERANCE,
        warnings,
    )
    right_value = to_decimal(values["权利价值"], "权利价值")
    debt_value_calc = min(basis, right_value, recover_value)
    debt_value = choose_cached_or_computed(
        values,
        "债权评估值",
        debt_value_calc,
        MONEY_TOLERANCE,
        warnings,
    )
    repayment_ratio = choose_cached_or_computed(
        values,
        "偿债率",
        debt_value / basis,
        PERCENT_RATIO_TOLERANCE,
        warnings,
    )

    property_nature = clean_text(values.get("性质"))
    has_placement_cost = placement_cost > 0
    if has_placement_cost:
        placement_title_suffix = "及安置费用"
        placement_sentence = (
            "同时，由于抵押物为债务人的唯一住房，需考虑安置费用，"
            f"根据委托人提供的资料，本次考虑安置费用为{fmt_yuan(placement_cost)}元。"
        )
        recover_formula = "抵押物折现价值-抵押物处置费用-安置费用"
        recover_calculation = f"{fmt_yuan(discounted_value)}-{fmt_yuan(disposal_cost)}-{fmt_yuan(placement_cost)}"
    else:
        placement_title_suffix = ""
        placement_sentence = "同时，由于抵押物并非债务人的唯一住房，无需考虑安置费用。"
        recover_formula = "抵押物折现价值-抵押物处置费用"
        recover_calculation = f"{fmt_yuan(discounted_value)}-{fmt_yuan(disposal_cost)}"

    if "非唯一" in property_nature and has_placement_cost:
        warnings.append("性质为非唯一住房，但安置费用大于0，已按安置费用生成")
    elif "唯一" in property_nature and not has_placement_cost:
        warnings.append("性质为唯一住房，但安置费用为空或0，已按安置费用生成")

    minimum_candidates = [
        ("抵押物可回收金额", recover_value),
        ("资产抵押债权金额", basis),
        ("抵押权利价值", right_value),
    ]
    minimum_value = min(item[1] for item in minimum_candidates)
    minimum_names = [label for label, amount in minimum_candidates if abs(amount - minimum_value) <= MONEY_TOLERANCE]
    if len(minimum_names) > 1:
        minimum_item = "与".join(minimum_names)
        warnings.append(f"{minimum_item}并列为三者最低值")
    else:
        minimum_item = minimum_names[0]

    conclusion = (
        "由于抵押物可偿债金额为抵押物可回收金额、资产抵押债权金额、抵押权利价值三者中孰低者，"
        f"抵押物可回收金额为{fmt_yuan(recover_value)}元，资产抵押债权金额为{fmt_yuan(basis)}元，"
        f"抵押权利价值为{fmt_yuan(right_value)}元，{minimum_item}"
    )
    if len(minimum_names) > 1:
        conclusion += f"并列为三者最低值，所以{name}户债权资产分析价值为{minimum_item}，即{fmt_yuan(debt_value)}元。"
    else:
        conclusion += f"为三者最低值，所以{name}户债权资产分析价值为{minimum_item}，即{fmt_yuan(debt_value)}元。"

    branch_for_filename, branch_warning = sanitize_filename_part(branch)
    name_for_filename, name_warning = sanitize_filename_part(name)
    for warning in [branch_warning, name_warning]:
        if warning:
            warnings.append(warning)

    report_year = clean_text(project["报告年"])
    report_batch = clean_text(project["报告批次"])
    report_number = f"天地恒安[{report_year}]资评咨字第{report_batch}-{sequence}号"
    filename = f"价值分析报告-工行个贷不良资产-{branch_for_filename}-{name_for_filename}.docx"

    guarantor = clean_text(values.get("保证人"))
    debtor_names_overview = "及".join(debtor.name for debtor in debtors)
    if guarantor:
        debtor_names_overview = f"{debtor_names_overview}，债务责任关联方为{guarantor}"
    debtor_names_table = "、".join(debtor.name for debtor in debtors)

    guarantor_details = known_guarantors.get(guarantor, {}) if guarantor else {}
    guarantor_extra_lines = [
        clean_text(line) for line in guarantor_details.get("经营范围附加行", []) if clean_text(line)
    ]

    replacements: Dict[str, str] = {
        "保证人名称": guarantor,
        "保证人成立日期": clean_text(guarantor_details.get("成立日期")),
        "保证人法定代表人": clean_text(guarantor_details.get("法定代表人")),
        "保证人类型": clean_text(guarantor_details.get("类型")),
        "保证人经营范围": clean_text(guarantor_details.get("经营范围")),
        "保证人统一社会信用代码": clean_text(guarantor_details.get("统一社会信用代码")),
        "保证人营业场所": clean_text(guarantor_details.get("营业场所")),
        "债务人": debtor_names_table,
        "债务人姓名_及": debtor_names_overview,
        "债权评估值": fmt_yuan(debt_value),
        "偿债率百分比": fmt_percent(repayment_ratio),
        "分析基准日": clean_text(project["分析基准日"]),
        "利息万元": fmt_wanyuan(interest),
        "单价": fmt_yuan(unit_price),
        "可回收公式": recover_formula,
        "可回收计算式": recover_calculation,
        "可回收金额": fmt_yuan(recover_value),
        "坐落": address,
        "基准日本息": fmt_yuan(basis),
        "处置费用": fmt_yuan(disposal_cost),
        "姓名": name,
        "安置费用标题后缀": placement_title_suffix,
        "安置费用说明": placement_sentence,
        "层": floor,
        "市值": fmt_yuan(market_value),
        "序号": sequence,
        "快速变现价值": fmt_yuan(quick_value),
        "快速变现系数": fmt_plain_number(quick_coeff),
        "折现价值": fmt_yuan(discounted_value),
        "折现期限": fmt_plain_number(discount_term),
        "折现率百分比": fmt_percent_compact(discount_rate),
        "折现系数": fmt_decimal(discount_coeff, 4),
        "报告年": report_year,
        "报告批次": report_batch,
        "报告日期": clean_text(project["报告日期"]),
        "报告日期大写": clean_text(project["报告日期大写"]),
        "最低值名称": minimum_item,
        "有效期截止日": clean_text(project["有效期截止日"]),
        "本息万元": fmt_wanyuan(basis),
        "本金万元": fmt_wanyuan(principal),
        "权利人": clean_text(values["证载权利人"]),
        "权利价值": fmt_yuan(right_value),
        "权利价值万元": fmt_wanyuan(right_value),
        "权证号": normalize_certificate_number(values["权证编号"]),
        "评估值万元": fmt_wanyuan(debt_value),
        "面积": fmt_area(area),
        "项目年": clean_text(project["项目年"]),
        "预计回收": clean_text(values["预计回收"]),
    }

    for i in range(1, 5):
        if i <= len(debtors):
            debtor = debtors[i - 1]
            replacements.update(
                {
                    f"债务人{i}": debtor.name,
                    f"债务人{i}住址": debtor.address,
                    f"债务人{i}性别": debtor.gender,
                    f"债务人{i}民族显示": debtor.ethnicity,
                    f"债务人{i}称谓": debtor.label,
                    f"债务人{i}身份证号码": debtor.id_number,
                }
            )
        else:
            replacements.update(
                {
                    f"债务人{i}": "",
                    f"债务人{i}住址": "",
                    f"债务人{i}性别": "",
                    f"债务人{i}民族显示": "",
                    f"债务人{i}称谓": "",
                    f"债务人{i}身份证号码": "",
                }
            )

    return PreparedReport(
        row_number=row_number,
        sequence=sequence,
        branch=branch,
        name=name,
        report_number=report_number,
        filename=filename,
        debtor_count=len(debtors),
        guarantor_included=bool(guarantor),
        placement_cost_included=has_placement_cost,
        minimum_item=minimum_item,
        replacements=replacements,
        conclusion_text=conclusion,
        guarantor_extra_lines=guarantor_extra_lines,
        warnings=warnings,
    )


def remove_paragraph(paragraph: Any) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def set_paragraph_text(paragraph: Any, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        paragraph.runs[0].font.highlight_color = None
        for run in paragraph.runs[1:]:
            run.text = ""
            run.font.highlight_color = None
    else:
        paragraph.add_run(text)


def insert_paragraph_after(paragraph: Any, text: str) -> Any:
    new_element = OxmlElement("w:p")
    if paragraph._p.pPr is not None:
        new_element.append(deepcopy(paragraph._p.pPr))
    paragraph._p.addnext(new_element)
    new_paragraph = Paragraph(new_element, paragraph._parent)
    run = new_paragraph.add_run(text)
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        run._r.insert(0, deepcopy(paragraph.runs[0]._r.rPr))
    return new_paragraph


def iter_paragraphs(container: Any) -> Iterable[Any]:
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)


def iter_all_paragraphs(doc: Any) -> Iterable[Any]:
    yield from iter_paragraphs(doc)
    for section in doc.sections:
        for container in [
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ]:
            yield from iter_paragraphs(container)


def clear_highlights(doc: Any) -> None:
    for paragraph in iter_all_paragraphs(doc):
        for run in paragraph.runs:
            run.font.highlight_color = None


def apply_literal_adjustments(doc: Any, has_guarantor: bool) -> None:
    if has_guarantor:
        return
    for paragraph in iter_all_paragraphs(doc):
        text = paragraph.text
        if "保证人的偿债能力" not in text and "保证合同" not in text:
            continue
        adjusted = text.replace("、保证人的偿债能力", "")
        adjusted = adjusted.replace("贷款合同、借款凭据、抵押合同、保证合同等资料", "贷款合同、借款凭据、抵押合同等资料")
        adjusted = adjusted.replace("等五个方面", "等四个方面")
        if adjusted != text:
            set_paragraph_text(paragraph, adjusted)


def apply_static_toc(doc: Any, report: PreparedReport) -> None:
    if report.guarantor_included:
        toc_lines = {
            "声": "声  明\t1",
            "价值分析报告摘要": "价值分析报告摘要\t2",
            "价值分析报告": "价值分析报告\t4",
            "一、": "一、委托人、债务人及债务责任关联方简介\t4",
            "二、": "二、价值分析目的\t5",
            "三、": "三、价值分析对象和价值分析范围\t5",
            "四、": "四、价值类型\t6",
            "五、": "五、价值分析基准日\t6",
            "六、": "六、价值分析思路和方法\t6",
            "七、": "七、价值分析程序实施过程和情况\t9",
            "八、": "八、价值分析假设\t11",
            "九、": "九、价值分析结论\t11",
            "十、": "十、特别事项说明\t12",
            "十一、": "十一、价值分析报告使用限制说明\t13",
            "十二、": "十二、价值分析报告日\t13",
            "十三、": "十三、资产评估机构印章\t14",
            "价值分析报告附件": "价值分析报告附件\t15",
        }
    elif report.debtor_count == 1:
        toc_lines = {
            "声": "声  明\t1",
            "价值分析报告摘要": "价值分析报告摘要\t2",
            "价值分析报告": "价值分析报告\t4",
            "一、": "一、委托人、债务人及债务责任关联方简介\t4",
            "二、": "二、价值分析目的\t5",
            "三、": "三、价值分析对象和价值分析范围\t5",
            "四、": "四、价值类型\t5",
            "五、": "五、价值分析基准日\t6",
            "六、": "六、价值分析思路和方法\t6",
            "七、": "七、价值分析程序实施过程和情况\t9",
            "八、": "八、价值分析假设\t10",
            "九、": "九、价值分析结论\t11",
            "十、": "十、特别事项说明\t11",
            "十一、": "十一、价值分析报告使用限制说明\t12",
            "十二、": "十二、价值分析报告日\t13",
            "十三、": "十三、资产评估机构印章\t13",
            "价值分析报告附件": "价值分析报告附件\t14",
        }
    else:
        toc_lines = {
            "声": "声  明\t1",
            "价值分析报告摘要": "价值分析报告摘要\t2",
            "价值分析报告": "价值分析报告\t4",
            "一、": "一、委托人、债务人及债务责任关联方简介\t4",
            "二、": "二、价值分析目的\t5",
            "三、": "三、价值分析对象和价值分析范围\t5",
            "四、": "四、价值类型\t5",
            "五、": "五、价值分析基准日\t6",
            "六、": "六、价值分析思路和方法\t6",
            "七、": "七、价值分析程序实施过程和情况\t9",
            "八、": "八、价值分析假设\t10",
            "九、": "九、价值分析结论\t11",
            "十、": "十、特别事项说明\t12",
            "十一、": "十一、价值分析报告使用限制说明\t13",
            "十二、": "十二、价值分析报告日\t13",
            "十三、": "十三、资产评估机构印章\t13",
            "价值分析报告附件": "价值分析报告附件\t14",
        }

    in_toc = False
    ordered_toc_lines = sorted(toc_lines.items(), key=lambda item: len(item[0]), reverse=True)
    for paragraph in doc.paragraphs:
        normalized = re.sub(r"\s+", "", paragraph.text)
        if normalized == "目录":
            in_toc = True
            continue
        if in_toc and normalized == "声明":
            break
        if not in_toc:
            continue
        for prefix, text in ordered_toc_lines:
            if paragraph.text.strip().startswith(prefix):
                set_paragraph_text(paragraph, text)
                break


def delete_optional_blocks(doc: Any, report: PreparedReport) -> None:
    for paragraph in list(doc.paragraphs):
        text = paragraph.text.strip()
        for debtor_index in range(report.debtor_count + 1, 5):
            if f"{{{{债务人{debtor_index}称谓}}}}" in text:
                remove_paragraph(paragraph)
                break

    if report.guarantor_included:
        optional_guarantor_fields = [
            ("统一社会信用代码：{{保证人统一社会信用代码}}", "保证人统一社会信用代码"),
            ("类    型：{{保证人类型}}", "保证人类型"),
            ("法定代表人：{{保证人法定代表人}}", "保证人法定代表人"),
            ("成立日期：{{保证人成立日期}}", "保证人成立日期"),
            ("营业场所：{{保证人营业场所}}", "保证人营业场所"),
            ("经营范围：{{保证人经营范围}}", "保证人经营范围"),
        ]
        removable_prefixes = [
            prefix for prefix, key in optional_guarantor_fields if not report.replacements.get(key)
        ]
    else:
        removable_prefixes = [
            "债务责任关联方简介",
            "保证人：{{保证人名称}}",
            "统一社会信用代码：{{保证人统一社会信用代码}}",
            "类    型：{{保证人类型}}",
            "法定代表人：{{保证人法定代表人}}",
            "成立日期：{{保证人成立日期}}",
            "营业场所：{{保证人营业场所}}",
            "经营范围：{{保证人经营范围}}",
        ]

    for paragraph in list(doc.paragraphs):
        text = paragraph.text.strip()
        if any(text.startswith(prefix) for prefix in removable_prefixes):
            remove_paragraph(paragraph)


def replace_placeholders_in_paragraph(paragraph: Any, replacements: Dict[str, str]) -> None:
    original = paragraph.text
    for run in paragraph.runs:
        run.font.highlight_color = None
        text = run.text
        if "{{" not in text:
            continue
        for key, value in replacements.items():
            text = text.replace("{{" + key + "}}", value)
        run.text = text

    if "{{" in paragraph.text:
        replaced = original
        for key, value in replacements.items():
            replaced = replaced.replace("{{" + key + "}}", value)
        if replaced != original:
            set_paragraph_text(paragraph, replaced)


def replace_placeholders(doc: Any, report: PreparedReport) -> None:
    for paragraph in iter_all_paragraphs(doc):
        if "由于抵押物可偿债金额为抵押物可回收金额" in paragraph.text:
            set_paragraph_text(paragraph, report.conclusion_text)
            continue
        replace_placeholders_in_paragraph(paragraph, report.replacements)


def add_guarantor_extra_lines(doc: Any, report: PreparedReport) -> None:
    if not report.guarantor_extra_lines:
        return
    in_guarantor_block = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text.startswith("保证人："):
            in_guarantor_block = True
            continue
        if in_guarantor_block and text.startswith("委托合同约定的价值分析报告使用人"):
            return
        if in_guarantor_block and text.startswith("经营范围："):
            current = paragraph
            for line in report.guarantor_extra_lines:
                current = insert_paragraph_after(current, line)
            return


def postprocess_docx_xml(path: Path, replacements: Dict[str, str]) -> None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as fp:
        temp_path = Path(fp.name)
    try:
        with ZipFile(path, "r") as source, ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    text = data.decode("utf-8", errors="ignore")
                    for key, value in replacements.items():
                        text = text.replace("{{" + key + "}}", xml_escape(value))
                    text = re.sub(r"<w:highlight\b[^>]*/>", "", text)
                    text = re.sub(r"<w:updateFields\b[^>]*/>", "", text)
                    data = text.encode("utf-8")
                target.writestr(info, data)
        shutil.move(str(temp_path), str(path))
    finally:
        if temp_path.exists():
            temp_path.unlink()


def scan_docx(path: Path) -> List[str]:
    issues: List[str] = []
    try:
        Document(path)
    except Exception as exc:
        issues.append(f"文件无法用 python-docx 打开: {exc}")
    with ZipFile(path, "r") as zf:
        for name in zf.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            text = zf.read(name).decode("utf-8", errors="ignore")
            if "{{" in text or "}}" in text:
                issues.append(f"{name} 中仍有占位符")
            if "<w:highlight" in text:
                issues.append(f"{name} 中仍有高亮")
    return issues


def render_report(template_file: Path, output_file: Path, report: PreparedReport) -> List[str]:
    doc = Document(template_file)
    delete_optional_blocks(doc, report)
    apply_static_toc(doc, report)
    replace_placeholders(doc, report)
    add_guarantor_extra_lines(doc, report)
    clear_highlights(doc)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_file))
    postprocess_docx_xml(output_file, report.replacements)
    return scan_docx(output_file)


def find_single_file(directory: Path, pattern: str, description: str) -> Path:
    files = sorted(directory.glob(pattern))
    if not files:
        raise FileNotFoundError(f"未找到{description}: {directory / pattern}")
    if len(files) > 1:
        raise ValueError(f"找到多个{description}，请通过参数指定: {', '.join(str(p) for p in files)}")
    return files[0]


def resolve_defaults(project_root: Path) -> Dict[str, Path]:
    template_dir = project_root / "bin" / "template"
    excel_dir = project_root / "excel"
    return {
        "rules_file": find_single_file(template_dir, "*规则.json", "规则JSON"),
        "template_file": find_single_file(template_dir, "*基底模板.docx", "Word模板"),
        "prompt_file": find_single_file(template_dir, "*Prompt.md", "Prompt文件"),
        "workbook": excel_dir / "数据表.xlsx",
    }


def load_rules(rules_file: Path) -> Dict[str, Any]:
    return json.loads(rules_file.read_text(encoding="utf-8-sig"))


def load_project_values(rules: Dict[str, Any]) -> Dict[str, Any]:
    project = dict(DEFAULT_PROJECT_VALUES)
    project.update(rules.get("project_defaults", {}))
    return project


def build_parser(project_root: Path) -> argparse.ArgumentParser:
    defaults = resolve_defaults(project_root)
    parser = argparse.ArgumentParser(description="Generate one value-analysis Word report per Sheet1 data row.")
    parser.add_argument("--template-file", default=str(defaults["template_file"]))
    parser.add_argument("--prompt-file", default=str(defaults["prompt_file"]))
    parser.add_argument("--rules-file", default=str(defaults["rules_file"]))
    parser.add_argument("--workbook", default=str(defaults["workbook"]))
    parser.add_argument("--sheet", default="Sheet1")
    parser.add_argument("--word-dir", default=str(project_root / "word"))
    return parser


def main() -> int:
    project_root = Path(__file__).resolve().parent.parent
    args = build_parser(project_root).parse_args()

    template_file = Path(args.template_file).expanduser().resolve()
    prompt_file = Path(args.prompt_file).expanduser().resolve()
    rules_file = Path(args.rules_file).expanduser().resolve()
    workbook_file = Path(args.workbook).expanduser().resolve()
    word_dir = Path(args.word_dir).expanduser().resolve()

    for label, path in [
        ("Word模板", template_file),
        ("Prompt文件", prompt_file),
        ("规则JSON", rules_file),
        ("Excel数据表", workbook_file),
    ]:
        if not path.is_file():
            raise FileNotFoundError(f"{label}不存在: {path}")

    # The prompt is part of the configured generation package; reading it here
    # catches missing/encoding issues before any report is written.
    prompt_text = prompt_file.read_text(encoding="utf-8-sig").strip()
    if not prompt_text:
        raise ValueError(f"Prompt文件为空: {prompt_file}")

    rules = load_rules(rules_file)
    project = load_project_values(rules)
    known_guarantors = rules.get("known_guarantors", {})
    workbook = openpyxl.load_workbook(workbook_file, data_only=True)
    if args.sheet not in workbook.sheetnames:
        raise ValueError(f"Excel中不存在工作表 {args.sheet}; 可用工作表: {', '.join(workbook.sheetnames)}")
    ws = workbook[args.sheet]

    manifest: Dict[str, Any] = {
        "template": str(template_file),
        "prompt": str(prompt_file),
        "workbook": str(workbook_file),
        "sheet": args.sheet,
        "reports": [],
        "errors": [],
    }

    total_records = 0
    generated = 0
    print(f"Using template: {template_file}")
    print(f"Using prompt: {prompt_file}")
    print(f"Using workbook: {workbook_file}")

    for row_number in range(2, ws.max_row + 1):
        if is_blank(ws.cell(row_number, FIELD_COLUMNS["序号"]).value):
            continue
        total_records += 1
        values = get_row_values(ws, row_number)
        sequence = clean_text(values.get("序号"))
        name = clean_text(values.get("姓名"))
        print(f"Processing row {row_number} | 序号 {sequence} | {name}")
        try:
            report = build_prepared_report(row_number, values, project, known_guarantors)
            output_file = word_dir / report.filename
            validation_issues = render_report(template_file, output_file, report)
            all_warnings = report.warnings + validation_issues
            manifest["reports"].append(
                {
                    "row_number": report.row_number,
                    "sequence": report.sequence,
                    "branch": report.branch,
                    "name": report.name,
                    "filename": report.filename,
                    "report_number": report.report_number,
                    "debtor_count": report.debtor_count,
                    "guarantor_included": report.guarantor_included,
                    "placement_cost_included": report.placement_cost_included,
                    "minimum_item": report.minimum_item,
                    "output_status": "success" if not validation_issues else "warning",
                    "warnings": all_warnings,
                }
            )
            generated += 1
            print(f"  Saved: {output_file}")
            for warning in all_warnings:
                print(f"  WARNING: {warning}")
        except Exception as exc:
            error = {
                "row_number": row_number,
                "sequence": sequence,
                "name": name,
                "output_status": "error",
                "error": str(exc),
            }
            manifest["errors"].append(error)
            print(f"  ERROR: {exc}")

    word_dir.mkdir(parents=True, exist_ok=True)
    manifest_file = word_dir / "generation_manifest.json"
    manifest_file.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nRows selected: {total_records}")
    print(f"Reports generated: {generated}")
    print(f"Errors: {len(manifest['errors'])}")
    print(f"Manifest: {manifest_file}")
    return 0 if not manifest["errors"] else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
