#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
运行结果表读写（开发文档 6.5 / 8.5 / 8.9 / 8.11）

- 文件名：{公司}-报关单和放行单等单据处理流程_运行结果表-{YYYY}.xlsx
- 存放位置：{日志目录}/{公司}/运行结果表/
- 工作表：绑定失败清单，冻结窗格 C2，共 33 列 + 末尾“最近操作时间段”（第 34 列）
- 本程序只写第 1～29 列与“最近操作时间段”，OA 四列（30～33）只读不写
- 按报关单号 upsert，不整表重写；保存用“临时文件 + 替换”避免损坏
- 早期 33 列的结果表在打开时自动补“最近操作时间段”列，不影响原有数据
"""
import os
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import openpyxl
from openpyxl.utils import get_column_letter

from log.loggingConfig import logger

SHEET_NAME = "绑定失败清单"
COL_BILL_NO = "报关单号"
COL_STATUS = "运行状态"
COL_REASON = "异常原因"
COL_START = "开始时间"
COL_END = "结束时间"
# 记录该单在单一窗口列表导出中命中的“最近操作时间”分段，逐单查询时重设时间条件
COL_UPDATE_WINDOW = "最近操作时间段"

RESULT_COLUMNS = [
    "销售组织", "报关单号", "贸易种类", "报关客户", "报关币种", "合同号", "报关数量", "报关单价",
    "总价", "运输费", "保险费", "出口日期", "手册号", "贸易方式", "报关汇率", "监管方式", "成交方式",
    "保税核注清单编号", "放行通知单路径", "电子委托协议路径", "发票路径", "装箱单路径",
    "报税核注清单路径", "报关单路径", "绑定状态", "运行状态", "异常原因", "开始时间", "结束时间",
    "OA处理状态", "OA处理异常原因", "OA处理开始时间", "OA处理结束时间",
    COL_UPDATE_WINDOW,
]
WRITABLE_COLUMNS = RESULT_COLUMNS[:29] + [COL_UPDATE_WINDOW]
# 允许写空值（用于清除上次的异常原因）
ALWAYS_WRITE_COLUMNS = {COL_STATUS, COL_REASON, COL_START, COL_END}


class ResultTableError(Exception):
    """运行结果表读写失败"""
    pass


def build_result_file_name(company: str, year: int) -> str:
    return f"{company}-报关单和放行单等单据处理流程_运行结果表-{year}.xlsx"


def find_template(config_dir) -> Optional[Path]:
    """在 01_Config 中查找运行结果表模板"""
    config_dir = Path(config_dir)
    if not config_dir.exists():
        return None
    candidates = sorted(config_dir.glob("*运行结果表*.xlsx"))
    return candidates[0] if candidates else None


class ResultTable:
    """单个公司、单个出口年度的运行结果表"""

    def __init__(self, path, template: Optional[Path] = None):
        self.path = Path(path)
        self.template = Path(template) if template else None
        self.workbook = None
        self.sheet = None
        self.columns: Dict[str, int] = {}
        self._row_index: Dict[str, int] = {}
        self._open()

    def _open(self) -> None:
        if not self.path.exists():
            if self.template and self.template.exists():
                self._create_from_template()
            else:
                logger.warning(f"未找到运行结果表模板，按 33 列标准表头新建: {self.path}")
                self._create_empty()

        self.workbook = openpyxl.load_workbook(self.path)
        if SHEET_NAME in self.workbook.sheetnames:
            self.sheet = self.workbook[SHEET_NAME]
        else:
            self.sheet = self.workbook.worksheets[0]

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in self.sheet[1]]
        self._append_missing_columns(headers)
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in self.sheet[1]]
        self.columns = {name: headers.index(name) + 1 for name in RESULT_COLUMNS}
        if COL_BILL_NO not in headers:
            raise ResultTableError(f"运行结果表缺少“报关单号”列: {self.path}")

        for row in range(2, self.sheet.max_row + 1):
            bill_no = self._text(self.sheet.cell(row=row, column=self.columns[COL_BILL_NO]).value)
            if bill_no:
                self._row_index[bill_no] = row

    def _append_missing_columns(self, headers: List[str]) -> None:
        """老版本结果表缺少新列时在末尾补表头（保存时随数据一起写入）"""
        missing = [name for name in RESULT_COLUMNS if name not in headers]
        if not missing:
            return
        logger.warning(f"运行结果表缺少列，已在末尾补充: {missing} -> {self.path}")
        for name in missing:
            column = max(self.sheet.max_column, len(headers)) + 1
            self.sheet.cell(row=1, column=column, value=name)
            self.sheet.column_dimensions[get_column_letter(column)].width = 20
            headers.append(name)

    def _create_from_template(self) -> None:
        """复制模板并清空示例数据行，保留表头与格式"""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        import shutil

        shutil.copy2(self.template, self.path)
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook[SHEET_NAME] if SHEET_NAME in workbook.sheetnames else workbook.worksheets[0]
        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row - 1)
        workbook.save(self.path)
        workbook.close()
        logger.info(f"按模板创建运行结果表: {self.path}")

    def _create_empty(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = SHEET_NAME
        sheet.append(RESULT_COLUMNS)
        sheet.freeze_panes = "C2"
        workbook.save(self.path)
        workbook.close()

    @staticmethod
    def _text(value) -> str:
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return str(value).strip()

    def rows(self) -> List[dict]:
        result = []
        for row in range(2, self.sheet.max_row + 1):
            record = {
                name: self.sheet.cell(row=row, column=column).value
                for name, column in self.columns.items()
            }
            if self._text(record.get(COL_BILL_NO)):
                result.append(record)
        return result

    def get(self, bill_no: str) -> Optional[dict]:
        row = self._row_index.get(bill_no)
        if row is None:
            return None
        return {
            name: self.sheet.cell(row=row, column=column).value
            for name, column in self.columns.items()
        }

    def status(self, bill_no: str) -> str:
        record = self.get(bill_no)
        return self._text(record.get(COL_STATUS)) if record else ""

    def bills_with_status(self, statuses: Iterable[str]) -> set:
        wanted = set(statuses)
        return {
            bill_no
            for bill_no, row in self._row_index.items()
            if self._text(self.sheet.cell(row=row, column=self.columns[COL_STATUS]).value) in wanted
        }

    def upsert(self, bill_no: str, values: dict) -> str:
        """按报关单号更新（不存在则追加），返回 'insert' 或 'update'"""
        bill_no = str(bill_no or "").strip()
        if not bill_no:
            raise ResultTableError("报关单号为空，无法写入运行结果表")

        row = self._row_index.get(bill_no)
        action = "update"
        if row is None:
            row = self.sheet.max_row + 1
            self.sheet.cell(row=row, column=self.columns[COL_BILL_NO], value=bill_no)
            self._row_index[bill_no] = row
            action = "insert"

        for name, value in values.items():
            if name not in WRITABLE_COLUMNS or name not in self.columns:
                continue
            if value is None or (isinstance(value, str) and value == ""):
                if name not in ALWAYS_WRITE_COLUMNS:
                    continue
                value = None
            self.sheet.cell(row=row, column=self.columns[name], value=value)
        return action

    def save(self) -> None:
        """临时文件 + 原子替换（9.7）"""
        temp_path = self.path.with_name(self.path.stem + ".tmp.xlsx")
        self.workbook.save(temp_path)
        try:
            os.replace(temp_path, self.path)
        except PermissionError:
            temp_path.unlink(missing_ok=True)
            raise ResultTableError(f"运行结果表被占用，请复制到本地后重试: {self.path}") from None
        logger.info(f"运行结果表已保存: {self.path}")

    def close(self) -> None:
        if self.workbook:
            self.workbook.close()


def open_result_table(log_dir, company: str, year: int, config_dir=None) -> ResultTable:
    """打开（必要时创建）某公司某出口年度的运行结果表（存放在日志目录）"""
    path = Path(log_dir) / company / "运行结果表" / build_result_file_name(company, year)
    return ResultTable(path, template=find_template(config_dir) if config_dir else None)


def year_of(value) -> int:
    """取出口日期所属年度"""
    if isinstance(value, datetime):
        return value.year
    if isinstance(value, date):
        return value.year
    text = str(value or "").strip()
    return int(text[:4]) if len(text) >= 4 and text[:4].isdigit() else date.today().year
