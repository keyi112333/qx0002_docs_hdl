#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报关单列表导出文件解析与分段合并（开发文档 8.4 / 4.3）

单一窗口【列表导出】实际可能是三种格式：
- HTML 表格伪装成 .xls（最常见）
- 二进制 .xls（需要 xlrd）
- .xlsx
读取时按文件头自动识别，调用方无需关心格式。
"""
from dataclasses import dataclass
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from log.loggingConfig import logger

# 列表导出的关键列（别名按单一窗口版本可能变化，取第一个命中的列）
COL_UNIFIED_NO = ("统一编号", "统一编号/报关单号")
COL_CUSTOMS_NO = ("海关编号", "报关单号")
COL_DOMESTIC = ("境内收发货人", "境内收发货人名称")
COL_EXPORT_DATE = ("进出口日期", "出口日期")
COL_SUPERVISION = ("监管方式",)
COL_CONTRACT = ("合同协议号", "合同号")
COL_STATUS = ("报关状态",)
COL_IE_FLAG = ("进出口标志",)


class ListFormatError(Exception):
    """列表导出文件格式无法识别"""
    pass


def _read_text(path: Path) -> str:
    """按 UTF-8 / GB18030 依次尝试解码"""
    raw = path.read_bytes()
    for encoding in ("utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


class _TableRowParser(HTMLParser):
    """解析 HTML 表格，按行返回单元格文本"""

    def __init__(self):
        super().__init__()
        self.rows: List[List[str]] = []
        self._row: Optional[List[str]] = None
        self._cell: Optional[List[str]] = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def iter_excel_rows(path) -> Iterable[list]:
    """按实际内容逐行读取导出文件（xlsx / 二进制 xls / HTML 表格 xls）"""
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == ".xls":
        with path.open("rb") as file:
            head = file.read(4)
        if head.startswith(b"<"):
            parser = _TableRowParser()
            parser.feed(_read_text(path))
            yield from parser.rows
            return
        try:
            import xlrd
        except ImportError:
            raise ListFormatError("导出文件为二进制 .xls，请先安装 xlrd: pip install xlrd") from None
        sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
        for index in range(sheet.nrows):
            yield sheet.row_values(index)
        return

    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for row in workbook.active.iter_rows(values_only=True):
            yield list(row)
    finally:
        workbook.close()


def _find_header(rows: List[list]) -> int:
    """定位表头行：同时包含“海关编号”与“统一编号”"""
    for index, row in enumerate(rows[:20]):
        values = [str(value or "").strip() for value in row]
        if any(alias in values for alias in COL_CUSTOMS_NO) and any(
            alias in values for alias in COL_UNIFIED_NO
        ):
            return index
    raise ListFormatError("导出文件中未找到“统一编号/海关编号”表头，请确认是报关单列表导出文件")


def _column_index(headers: List[str], aliases: Iterable[str]) -> Optional[int]:
    for alias in aliases:
        if alias in headers:
            return headers.index(alias)
    return None


def _cell(values: List[str], index: Optional[int]) -> str:
    if index is None or index >= len(values):
        return ""
    return str(values[index] or "").strip()


def _parse_date(value: str) -> Optional[date]:
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


@dataclass
class ListDeclaration:
    """报关单列表导出中的一行（8.4 的候选集合）"""

    unified_no: str                     # 统一编号：单一窗口逐单查询用
    no: str                             # 海关编号（报关单号）：业务主键
    domestic_consignor: str = ""        # 境内收发货人
    export_date: Optional[date] = None  # 进出口日期
    supervision_mode: str = ""          # 监管方式（中文，如“进料对口”）
    contract_no: str = ""
    customs_status: str = ""            # 报关状态
    ie_flag: str = ""                   # 进出口标志
    update_window: Optional[Tuple[date, date]] = None   # 命中的“最近操作时间”分段

    def window_text(self) -> str:
        """最近操作时间段文本：YYYY-MM-DD~YYYY-MM-DD"""
        if not self.update_window:
            return ""
        start, end = self.update_window
        return f"{start:%Y-%m-%d}~{end:%Y-%m-%d}"

    def is_export_closed(self) -> bool:
        """是否出口且已结关（查询条件已限定，导出结果再校验一次）"""
        if self.ie_flag and "出口" not in self.ie_flag:
            return False
        if self.customs_status and "结关" not in self.customs_status:
            return False
        return True


def read_list_file(path) -> List[ListDeclaration]:
    """读取单个列表导出文件，返回报关单行（跳过无海关编号的合计/空行）"""
    path = Path(path)
    rows = list(iter_excel_rows(path))
    if not rows:
        raise ListFormatError(f"列表导出文件为空: {path}")

    header_index = _find_header(rows)
    headers = [str(value or "").strip() for value in rows[header_index]]
    col_unified = _column_index(headers, COL_UNIFIED_NO)
    col_no = _column_index(headers, COL_CUSTOMS_NO)
    col_domestic = _column_index(headers, COL_DOMESTIC)
    col_date = _column_index(headers, COL_EXPORT_DATE)
    col_supervision = _column_index(headers, COL_SUPERVISION)
    col_contract = _column_index(headers, COL_CONTRACT)
    col_status = _column_index(headers, COL_STATUS)
    col_ie = _column_index(headers, COL_IE_FLAG)

    result: List[ListDeclaration] = []
    for row in rows[header_index + 1:]:
        values = [str(value or "").strip() for value in row]
        no = _cell(values, col_no)
        if not no or no == "-":
            continue
        item = ListDeclaration(
            unified_no=_cell(values, col_unified),
            no=no,
            domestic_consignor=_cell(values, col_domestic),
            export_date=_parse_date(_cell(values, col_date)),
            supervision_mode=_cell(values, col_supervision),
            contract_no=_cell(values, col_contract),
            customs_status=_cell(values, col_status),
            ie_flag=_cell(values, col_ie),
        )
        result.append(item)

    logger.info(f"列表导出解析完成: {path.name} -> {len(result)} 条")
    return result


def read_list_files(files: Iterable) -> List[ListDeclaration]:
    """读取多个分段导出文件，按报关单号去重（保留最新一段的记录）"""
    deduped: Dict[str, ListDeclaration] = {}
    for file in files:
        for item in read_list_file(file):
            deduped[item.no] = item
    return list(deduped.values())


def merge_exports(files: Iterable, save_path) -> Path:
    """合并分段导出结果，按报关单号去重后另存为 xlsx（8.4 第 4 步）"""
    files = [Path(file) for file in files]
    items = read_list_files(files)
    if not items:
        raise ListFormatError("各分段导出均无数据，无法生成合并文件")

    headers = ["统一编号", "海关编号", "境内收发货人", "进出口日期", "监管方式", "合同协议号", "报关状态", "进出口标志"]
    import openpyxl

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "报关单列表"
    sheet.append(headers)
    for item in items:
        sheet.append([
            item.unified_no,
            item.no,
            item.domestic_consignor,
            item.export_date.strftime("%Y-%m-%d") if item.export_date else "",
            item.supervision_mode,
            item.contract_no,
            item.customs_status,
            item.ie_flag,
        ])
    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = f"A1:H{sheet.max_row}"
    for index in range(1, len(headers) + 1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = 22

    save_path = Path(save_path)
    save_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(save_path)
    workbook.close()
    logger.info(f"报关单列表合并完成: {len(items)} 条 -> {save_path}")
    return save_path
