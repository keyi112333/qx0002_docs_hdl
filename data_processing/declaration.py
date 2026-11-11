#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口【详情导出】Excel 解析（开发文档 6.3 / 7.5 / 9.3 / 9.5）

详情导出结构（示例数据验证）：
- 工作表“基本信息&商品信息”：表头区域为“标签/值”成对排列，第 22 行为明细表头
  （项号、商品编号、成交数量、单价、总价、币制...），其后为明细行；
- 工作表“随附单证信息”：表头为“随附单证代码/单证编号”，代码 a 表示保税核注清单。

表头行、标签与列名可能随单一窗口版本变化，解析不使用固定行列：
- 明细表头按“项号 + 成交数量 + 总价”在整表中定位，明细列按列名（含别名）取索引；
- 表头字段在明细表头之前的区域内按标签任意位置查找，值只取标签右侧紧邻的第一格，
  该格为空或本身是标签时视为无值（不跨列跳格、不取下行的同名单元格）；
- 随附单证页在整表中查找表头行与代码/编号列，列顺序可互换。

解析失败时抛出 ParseError，由调用方记“失败-解析失败”并转人工。
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from log.loggingConfig import logger

BONDED_DOC_CODE = "a"
# 币种小数位：默认 2 位，日元等无小数币种取 0 位
CURRENCY_DECIMALS = {"JPY": 0, "KRW": 0, "VND": 0}


class ParseError(Exception):
    """报关单详情无法解析或关键字段缺失"""
    pass


class ManualError(Exception):
    """需要人工判断的业务异常（无合同号、核注清单编号不唯一等）"""
    pass


# 标签别名 -> Declaration 字段名（标签统一去掉空白后比较）
HEADER_ALIASES: Dict[str, Tuple[str, ...]] = {
    "unified_no": ("统一编号",),
    "no": ("报关单号", "海关编号"),
    "contract_no": ("合同协议号",),
    "export_date": ("进/出口日期", "进出口日期", "出口日期"),
    "domestic_consignor": ("境内收发货人名称",),
    "overseas_consignor": ("境外收发货人名称（外文）", "境外收发货人名称(外文)", "境外收发货人名称"),
    "overseas_consignor_cn": ("境外收发货人名称（中文）", "境外收发货人名称(中文)"),
    "supervision_code": ("监管方式",),
    "trade_mode_code": ("成交方式",),
    "freight_flag": ("运费标识",),
    "insurance_flag": ("保费标识",),
    "record_no": ("备案号",),
    "ie_flag": ("进出口标识",),
}

DETAIL_HEADER_KEYS = ("项号", "成交数量", "总价")
DETAIL_ALIASES: Dict[str, Tuple[str, ...]] = {
    "item_no": ("项号",),
    "hs_code": ("商品编号",),
    "name": ("商品名称",),
    "qty": ("成交数量",),
    "unit": ("成交计量单位",),
    "price": ("单价",),
    "amount": ("总价",),
    "currency": ("币制",),
}


def _normalize_label(value) -> str:
    """标签归一化：去空白、统一全角括号与冒号，便于跨版本匹配"""
    text = "" if value is None else str(value)
    text = re.sub(r"\s+", "", text)
    return text.replace("（", "(").replace("）", ")").replace("：", ":")


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _decimal(value) -> Optional[Decimal]:
    text = _text(value).replace(",", "")
    if not text or text == "-":
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _parse_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _text(value)
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


@dataclass
class DeclarationLine:
    """报关单明细行"""

    item_no: str
    hs_code: str = ""
    name: str = ""
    qty: Optional[Decimal] = None
    unit: str = ""
    price: Optional[Decimal] = None
    amount: Optional[Decimal] = None
    currency: str = ""


@dataclass
class Declaration:
    """报关单主数据（BIP 录入与文件命名的数据源）"""

    no: str
    unified_no: str = ""
    contract_no: str = ""
    export_date: Optional[date] = None
    domestic_consignor: str = ""
    overseas_consignor: str = ""
    supervision_code: str = ""
    supervision_mode: str = ""          # 中文监管方式，来自列表导出
    trade_mode_code: str = ""
    freight_flag: str = ""
    insurance_flag: str = ""
    record_no: str = ""
    ie_flag: str = ""
    lines: List[DeclarationLine] = field(default_factory=list)
    attached_docs: List[Tuple[str, str]] = field(default_factory=list)
    bonded_nos: List[str] = field(default_factory=list)
    source_file: str = ""
    export_date_from_list = False       # 出口日期是否取自列表导出

    @property
    def currency(self) -> str:
        for line in self.lines:
            if line.currency:
                return line.currency
        return ""

    @property
    def bonded_no(self) -> str:
        return self.bonded_nos[0] if self.bonded_nos else ""

    @property
    def total_qty(self) -> Decimal:
        return sum((line.qty for line in self.lines if line.qty is not None), Decimal("0"))

    @property
    def total_amount(self) -> Decimal:
        return sum((line.amount for line in self.lines if line.amount is not None), Decimal("0"))

    def currency_decimals(self) -> int:
        return CURRENCY_DECIMALS.get(self.currency.upper(), 2)

    def unit_price(self) -> Optional[Decimal]:
        """报关单价 = 汇总总价 / 汇总数量，四舍五入到币种小数位（9.3）"""
        qty = self.total_qty
        if not qty:
            return None
        digits = self.currency_decimals()
        quant = Decimal(1).scaleb(-digits)
        return (self.total_amount / qty).quantize(quant, rounding=ROUND_HALF_UP)

    def summary(self) -> str:
        export_text = self.export_date.strftime("%Y-%m-%d") if self.export_date else "缺失"
        return (
            f"{self.no} 合同号={self.contract_no or '无'} 出口日期={export_text} "
            f"监管方式={self.supervision_mode or self.supervision_code} 明细{len(self.lines)}行"
        )


def _find_detail_header(rows: List[tuple]) -> int:
    """在整表中定位明细表头行（项号 + 成交数量 + 总价），不依赖固定行号"""
    for index, row in enumerate(rows):
        labels = {_normalize_label(value) for value in row}
        if all(key in labels for key in DETAIL_HEADER_KEYS):
            return index
    raise ParseError("详情导出中未找到明细表头（项号/成交数量/总价）")


def _read_header_fields(rows: List[tuple], detail_index: int) -> Dict[str, str]:
    """在明细表头之前按标签取值，不依赖固定行列

    标签位置动态查找；值只取标签右侧紧邻的第一格（右边第一格），
    该格为空或本身是标签时视为无值，不跨列跳格、不取下行的单元格。
    """
    alias_map = {
        _normalize_label(alias): name
        for name, aliases in HEADER_ALIASES.items()
        for alias in aliases
    }
    values: Dict[str, str] = {}
    for row_index in range(detail_index):
        cells = list(rows[row_index])
        for col, cell in enumerate(cells):
            label = _normalize_label(cell)
            name = alias_map.get(label)
            if not name or name in values:
                continue
            value = _value_at_right(cells, col, alias_map)
            if value is not None:
                values[name] = value
    return values


def _value_at_right(cells: List, col: int, alias_map: Dict[str, str]) -> Optional[str]:
    """取标签右侧紧邻的第一格；为空或本身是另一个标签时视为无值"""
    index = col + 1
    if index >= len(cells):
        return None
    normalized = _normalize_label(cells[index])
    if not normalized or normalized in alias_map:
        return None
    return _text(cells[index])


def _read_lines(rows: List[tuple], header_index: int) -> List[DeclarationLine]:
    header = [_normalize_label(value) for value in rows[header_index]]
    columns: Dict[str, Optional[int]] = {}
    for name, aliases in DETAIL_ALIASES.items():
        columns[name] = next(
            (header.index(_normalize_label(alias)) for alias in aliases
             if _normalize_label(alias) in header),
            None,
        )
    for required in ("qty", "amount"):
        if columns.get(required) is None:
            raise ParseError(f"详情导出明细缺少列: {DETAIL_ALIASES[required][0]}")

    def cell(values, key: str):
        index = columns.get(key)
        if index is None or index >= len(values):
            return None
        return values[index]

    lines: List[DeclarationLine] = []
    for row in rows[header_index + 1:]:
        item_no = _text(cell(row, "item_no"))
        qty = _decimal(cell(row, "qty"))
        amount = _decimal(cell(row, "amount"))
        if not item_no and qty is None and amount is None:
            continue
        lines.append(DeclarationLine(
            item_no=item_no,
            hs_code=_text(cell(row, "hs_code")),
            name=_text(cell(row, "name")),
            qty=qty,
            unit=_text(cell(row, "unit")),
            price=_decimal(cell(row, "price")),
            amount=amount,
            currency=_text(cell(row, "currency")),
        ))
    return lines


def _find_sheet(workbook, keyword: str):
    for name in workbook.sheetnames:
        if keyword in name:
            return workbook[name]
    return None


def _read_attached_docs(workbook) -> List[Tuple[str, str]]:
    """读取“随附单证”页的（代码, 编号）列表，表头行与列顺序均动态识别"""
    sheet = _find_sheet(workbook, "随附单证")
    if sheet is None:
        return []

    rows = list(sheet.iter_rows(values_only=True))
    header_index, code_col, no_col = None, None, None
    code_key, no_key = _normalize_label("随附单证代码"), _normalize_label("单证编号")
    for index, row in enumerate(rows):
        labels = [_normalize_label(value) for value in row]
        if code_key in labels and no_key in labels:
            header_index = index
            code_col = labels.index(code_key)
            no_col = labels.index(no_key)
            break
    if header_index is None:
        return []

    docs: List[Tuple[str, str]] = []
    for row in rows[header_index + 1:]:
        values = list(row)
        code = _text(values[code_col]) if code_col < len(values) else ""
        no = _text(values[no_col]) if no_col < len(values) else ""
        if code and no:
            docs.append((code, no))
    return docs


def parse_detail(path, supervision_mode: str = "", export_date_fallback: Optional[date] = None,
                 unified_no: str = "") -> Declaration:
    """解析报关单详情导出 Excel

    Args:
        path: 详情导出文件路径
        supervision_mode: 列表导出中的中文监管方式（判断“进料”用中文名，详情里只有代码）
        export_date_fallback: 详情“进/出口日期”为空时的兜底日期（列表导出的进出口日期）
        unified_no: 列表导出的统一编号（详情中缺失时补齐）
    """
    path = Path(path)
    if not path.exists():
        raise ParseError(f"报关单详情文件不存在: {path}")

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = _find_sheet(workbook, "基本信息")
        if sheet is None and workbook.worksheets:
            # 页签被改名时退回第一个工作表，仍按明细表头定位
            sheet = workbook.worksheets[0]
            logger.warning(f"详情导出未找到“基本信息&商品信息”页，改用首个工作表: {sheet.title}")
        if sheet is None:
            raise ParseError(f"详情导出没有可解析的工作表: {path.name}")

        rows = list(sheet.iter_rows(values_only=True))
        header_index = _find_detail_header(rows)
        header = _read_header_fields(rows, header_index)
        lines = _read_lines(rows, header_index)
        attached_docs = _read_attached_docs(workbook)
    finally:
        workbook.close()

    no = header.get("no", "")
    if not no:
        raise ParseError(f"详情导出缺少报关单号: {path.name}")
    if not lines:
        raise ParseError(f"详情导出没有商品明细: {path.name}")

    export_date = _parse_date(header.get("export_date"))
    from_list = False
    if export_date is None and export_date_fallback is not None:
        export_date = export_date_fallback
        from_list = True
    if export_date is None:
        raise ParseError(f"详情导出出口日期为空且列表导出无进出口日期: {no}")

    declaration = Declaration(
        no=no,
        unified_no=header.get("unified_no") or unified_no,
        contract_no=header.get("contract_no", ""),
        export_date=export_date,
        domestic_consignor=header.get("domestic_consignor", ""),
        overseas_consignor=header.get("overseas_consignor", "") or header.get("overseas_consignor_cn", ""),
        supervision_code=header.get("supervision_code", ""),
        supervision_mode=supervision_mode,
        trade_mode_code=header.get("trade_mode_code", ""),
        freight_flag=header.get("freight_flag", ""),
        insurance_flag=header.get("insurance_flag", ""),
        record_no=header.get("record_no", ""),
        ie_flag=header.get("ie_flag", ""),
        lines=lines,
        attached_docs=attached_docs,
        source_file=str(path),
    )
    declaration.export_date_from_list = from_list

    bonded_codes = {code.strip().lower() for code, _ in attached_docs}
    declaration.bonded_nos = sorted({
        number for code, number in attached_docs if code.strip().lower() == BONDED_DOC_CODE
    })
    if not declaration.currency:
        raise ParseError(f"详情导出明细缺少币制: {no}")

    logger.info(f"详情解析完成: {declaration.summary()} 随附单证代码={sorted(bonded_codes)}")
    return declaration
