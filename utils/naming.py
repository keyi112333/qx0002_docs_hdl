#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单据文件命名与归档路径（开发文档 4.3 / 9.6）

统一规则：海关编号（报关单号）-合同编号-出口日期-单据名称.pdf
按报关单“出口日期”所属年月（YYYYMM）归档。
"""
import re
from datetime import date, datetime
from pathlib import Path

# 文件名非法字符与长度限制
ILLEGAL_PATTERN = re.compile(r'[<>:"/\\|?*]')
MAX_NAME_LENGTH = 180
# 报关单没有合同协议号时的占位符
CONTRACT_PLACEHOLDER = "无合同号"
DECLARATION_DOC_TYPE = "报关单"


def parse_export_date(value) -> date:
    """把出口日期统一成 date，支持 20260108 / 2026-01-08 / 2026/01/08"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"出口日期格式无法识别: {value}，支持 YYYYMMDD / YYYY-MM-DD")


def build_document_name(entry_no: str, contract_no: str, export_date, doc_type: str, ext: str = ".pdf") -> str:
    """生成单据文件名：报关单号-合同编号-出口日期-单据名称.扩展名"""
    entry_no = str(entry_no or "").strip()
    contract = str(contract_no or "").strip() or CONTRACT_PLACEHOLDER
    date_str = parse_export_date(export_date).strftime("%Y%m%d")

    name = ILLEGAL_PATTERN.sub("_", f"{entry_no}-{contract}-{date_str}-{doc_type}")
    name = re.sub(r"\s+", " ", name).strip()

    # 超长时截断合同编号，保留报关单号与出口日期（长度含扩展名）
    limit = MAX_NAME_LENGTH - len(ext)
    if len(name) > limit:
        head, tail = f"{entry_no}-", f"-{date_str}-{doc_type}"
        keep = limit - len(head) - len(tail)
        name = f"{head}{contract[:max(keep, 0)]}{tail}"
    return name + ext


def build_document_path(base_dir, entry_no: str, contract_no: str, export_date, doc_type: str,
                        ext: str = ".pdf") -> Path:
    """生成单据归档路径：base_dir/YYYYMM/报关单号-合同编号-出口日期-单据名称.扩展名"""
    yyyymm = parse_export_date(export_date).strftime("%Y%m")
    return Path(base_dir) / yyyymm / build_document_name(entry_no, contract_no, export_date, doc_type, ext)
