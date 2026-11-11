#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BIP 已上传附件数据（开发文档 6.4 / 8.5 / 9.2）

数据来自 BIP 顾问提供的 SQL（视图 RPA_xsbgdfjrz：销售报关单表头 left join 附件日志），
由本程序直连 Oracle 查询后落盘为
03_Output/{公司}/BIP已上传附件表/{公司}-BIP已上传附件表-YYYYMMDD HHMMSS.xlsx，
作为本次“待处理清单”比对的依据保留。

列含义：
- VBILLCODE：报关单号（与单一窗口海关编号一致）
- FILENAME/OP_TIME：附件名称与操作时间
- CZMC：操作名称（上传/修改/删除）；为空表示该报关单没有任何附件操作记录
"""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set

from log.loggingConfig import logger

DEFAULT_DSN = "10.98.7.31:1522/orcl"
DEFAULT_USER = "yonbiptest"
DEFAULT_PASSWORD = "yonbiptest"

# 与 tests/test_rpa_xsbgdfjrz_view.py 中的视图 select 一致
DEFAULT_SQL = """
select bgh.pk_iota_h,
       bgh.vbillcode,
       substr(fjrz.filepath,1,20) as djid,
       fjrz.filename,
       fjrz.op_time,
       case fjrz.op_type when '1' then '上传' when '2' then '修改' when '3' then '删除' when '4' then '移动'
            else '' end as czmc
from iot_iota_h bgh
left join sm_filelog fjrz
       on bgh.pk_iota_h = substr(fjrz.filepath,1,20) and fjrz.dr=0
where bgh.dr=0
"""

EXPORT_COLUMNS = ("PK_IOTA_H", "VBILLCODE", "DJID", "FILENAME", "OP_TIME", "CZMC")
COL_BILL_ALIASES = ("VBILLCODE", "报关单号")
COL_FILE_ALIASES = ("FILENAME", "附件名称")
COL_TIME_ALIASES = ("OP_TIME", "操作时间")
COL_ACTION_ALIASES = ("CZMC", "操作名称")
# 只有“上传/修改”才算存在附件；“删除”视为已移除
VALID_ACTIONS = ("上传", "修改")


class BipQueryError(Exception):
    """BIP 附件数据查询失败"""
    pass


@dataclass
class BipAttachmentData:
    """BIP 已上传附件数据（按报关单号分组）"""

    rows: List[dict] = field(default_factory=list)
    by_bill: Dict[str, List[dict]] = field(default_factory=dict)
    source_file: str = ""

    def has_attachment(self, bill_no: str) -> bool:
        for row in self.by_bill.get(bill_no, []):
            if str(row.get("action") or "").strip() in VALID_ACTIONS and str(row.get("filename") or "").strip():
                return True
        return False

    def attached_bills(self) -> Set[str]:
        return {no for no in self.by_bill if self.has_attachment(no)}

    def bills_without_attachment(self) -> Set[str]:
        """无附件记录的报关单号（CZMC 为空，或仅有“删除”记录）"""
        return {no for no in self.by_bill if not self.has_attachment(no)}

    def summary(self) -> str:
        return (
            f"BIP附件数据: 记录{len(self.rows)}行，单号{len(self.by_bill)}个，"
            f"有附件{len(self.attached_bills())}个，无附件{len(self.bills_without_attachment())}个"
        )


def fetch_attachment_export(config, save_path) -> Path:
    """直连 Oracle 查询 BIP 附件数据并落盘为 Excel"""
    dsn = str(config.get("查询配置.BIP数据库连接", DEFAULT_DSN) or DEFAULT_DSN).strip()
    user = str(config.get("查询配置.BIP数据库用户", DEFAULT_USER) or DEFAULT_USER).strip()
    password = str(config.get("查询配置.BIP数据库口令", DEFAULT_PASSWORD) or DEFAULT_PASSWORD).strip()
    sql = str(config.get("查询配置.BIP已上传附件SQL", DEFAULT_SQL) or DEFAULT_SQL).strip()

    try:
        import oracledb
    except ImportError:
        raise BipQueryError("需要安装 oracledb: pip install oracledb") from None

    logger.info(f"查询 BIP 已上传附件数据: {user}@{dsn}")
    try:
        connection = oracledb.connect(user=user, password=password, dsn=dsn, tcp_connect_timeout=10)
    except Exception as e:
        raise BipQueryError(f"BIP 数据库连接失败: {dsn} - {e}") from None

    try:
        cursor = connection.cursor()
        cursor.execute(sql)
        headers = [str(description[0]).upper() for description in cursor.description]
        rows = cursor.fetchall()
    except Exception as e:
        raise BipQueryError(f"BIP 已上传附件查询失败: {e}") from None
    finally:
        connection.close()

    save_path = Path(save_path)
    save_path.parent.mkdir(parents=True, exist_ok=True)
    _write_export(save_path, headers, rows, sql)
    logger.info(f"BIP 已上传附件数据已导出: {save_path}（{len(rows)} 行）")
    return save_path


def _write_export(path: Path, headers: List[str], rows: List[tuple], sql: str) -> None:
    import openpyxl
    from openpyxl.utils import get_column_letter

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "查询结果"
    sheet.append(headers)
    for row in rows:
        sheet.append([
            value.read() if hasattr(value, "read") else value
            for value in row
        ])
    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{sheet.max_row}"
    for index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = max(len(header) + 6, 22)

    info = workbook.create_sheet("执行信息")
    info.append(["项目", "内容"])
    info.append(["查询时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    info.append(["返回行数", len(rows)])
    info.append(["执行SQL", sql])
    info.column_dimensions["A"].width = 14
    info.column_dimensions["B"].width = 100

    temp_path = path.with_suffix(".tmp.xlsx")
    workbook.save(temp_path)
    workbook.close()
    temp_path.replace(path)


def load_attachment_export(path) -> BipAttachmentData:
    """读取附件数据导出文件（用于离线验证与重跑比对）"""
    import openpyxl

    path = Path(path)
    if not path.exists():
        raise BipQueryError(f"BIP 已上传附件文件不存在: {path}")

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.worksheets[0]
        rows = list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        raise BipQueryError(f"BIP 已上传附件文件为空: {path}")

    headers = [str(value or "").strip().upper() for value in rows[0]]

    def column(aliases):
        for alias in aliases:
            alias_upper = alias.upper()
            if alias_upper in headers:
                return headers.index(alias_upper)
        return None

    col_bill = column(COL_BILL_ALIASES)
    col_file = column(COL_FILE_ALIASES)
    col_time = column(COL_TIME_ALIASES)
    col_action = column(COL_ACTION_ALIASES)
    if col_bill is None:
        raise BipQueryError(f"BIP 已上传附件文件缺少报关单号列: {path}")

    data = BipAttachmentData(source_file=str(path))
    for row in rows[1:]:
        values = list(row)
        bill_no = _cell(values, col_bill)
        if not bill_no:
            continue
        record = {
            "bill_no": bill_no,
            "filename": _cell(values, col_file),
            "op_time": _cell(values, col_time),
            "action": _cell(values, col_action),
        }
        data.rows.append(record)
        data.by_bill.setdefault(bill_no, []).append(record)
    logger.info(data.summary())
    return data


def _cell(values: List, index: Optional[int]) -> str:
    if index is None or index >= len(values):
        return ""
    value = values[index]
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value).strip()
