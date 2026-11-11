#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
销售报关单 + 附件日志视图（RPA_xsbgdfjrz）SQL 独立测试脚本

被测 SQL：
create or replace view RPA_xsbgdfjrz as
select bgh.pk_iota_h, bgh.vbillcode,
       substr(fjrz.filepath,1,20) as djid,
       fjrz.filename, fjrz.op_time,
       case fjrz.op_type when '1' then '上传' ... end as czmc
from iot_iota_h bgh
left join sm_filelog fjrz
       on bgh.pk_iota_h = substr(fjrz.filepath,1,20) and fjrz.dr=0

默认只执行视图里的 select 语句（不改动数据库），查询结果写入 Excel；
加 --create-view 会先在测试库执行 create or replace view，再查询该视图。

用法：
    python tests/test_rpa_xsbgdfjrz_view.py
    python tests/test_rpa_xsbgdfjrz_view.py --bill-code 531620240160772613
    python tests/test_rpa_xsbgdfjrz_view.py --limit 100
    python tests/test_rpa_xsbgdfjrz_view.py --create-view
    python tests/test_rpa_xsbgdfjrz_view.py --output d:/tmp/result.xlsx
"""
import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import oracledb
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

DB_DSN = "10.98.7.31:1522/orcl"
DB_USER = "yonbiptest"
DB_PASSWORD = "yonbiptest"

VIEW_NAME = "RPA_xsbgdfjrz"

# 视图内的 select 语句（与原始 SQL 一致）
VIEW_SELECT_SQL = """
select bgh.pk_iota_h/*销售报关单.表头主键*/,
       bgh.vbillcode/*销售报关单.单据号*/,
       substr(fjrz.filepath,1,20) /*附件日志.单据表头主键*/ as djid,
       fjrz.filename/*附件日志.附件名称*/,
       fjrz.op_time/*附件日志.操作时间*/,
       case fjrz.op_type/*附件日志.操作名称*/
            when '1' then '上传' when '2' then '修改' when '3' then '删除' when '4' then '移动'
            else '' end as czmc
from iot_iota_h/*销售报关单表头*/ bgh
left join sm_filelog/*附件日志*/ fjrz
       on bgh.pk_iota_h = substr(fjrz.filepath,1,20) and fjrz.dr=0
"""

DEFAULT_OUTPUT_DIR = ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "03_Output"


def build_query(bill_code: str = "", limit: int = 0) -> str:
    """拼接待执行的查询 SQL（可选按单据号过滤、限制行数）"""
    sql = VIEW_SELECT_SQL
    if bill_code:
        sql += "\nwhere bgh.dr=0 and bgh.vbillcode = :bill_code"
    if limit > 0:
        sql = f"select * from (\n{sql}\n) where rownum <= :row_limit"
    return sql


def normalize(value):
    """LOB / 特殊类型转换为可写入 Excel 的值"""
    if isinstance(value, oracledb.LOB):
        return value.read()
    return value


def write_excel(output_path: Path, headers, rows, exec_sql: str, elapsed: float, mode: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "查询结果"

    sheet.append(list(headers))
    for row in rows:
        sheet.append([normalize(v) for v in row])

    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{sheet.max_row}"

    for idx, header in enumerate(headers, start=1):
        width = max(len(str(header)) + 4, 18)
        if rows:
            sample = max((len(str(normalize(row[idx - 1]))) for row in rows[:50] if row[idx - 1] is not None), default=0)
            width = min(max(width, sample + 2), 60)
        sheet.column_dimensions[get_column_letter(idx)].width = width

    info = workbook.create_sheet("执行信息")
    info.append(["项目", "内容"])
    info.append(["数据库", f"{DB_USER}@{DB_DSN}"])
    info.append(["测试方式", mode])
    info.append(["执行时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    info.append(["耗时(秒)", round(elapsed, 3)])
    info.append(["返回行数", len(rows)])
    info.append(["执行SQL", exec_sql.strip()])
    info.column_dimensions["A"].width = 14
    info.column_dimensions["B"].width = 100

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="测试 RPA_xsbgdfjrz 视图 SQL 并导出 Excel")
    parser.add_argument("--bill-code", default="", help="按销售报关单号过滤（默认不过滤）")
    parser.add_argument("--limit", type=int, default=0, help="最多查询行数，0 表示不限制（默认 0）")
    parser.add_argument("--output", default="", help="Excel 输出路径，默认写入 03_Output")
    parser.add_argument("--create-view", action="store_true", help="在测试库执行 create or replace view 后再查询视图")
    args = parser.parse_args()

    if args.output:
        output_path = Path(args.output)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = DEFAULT_OUTPUT_DIR / f"{VIEW_NAME}_{stamp}.xlsx"

    query = build_query(args.bill_code, args.limit)
    params = {}
    if args.bill_code:
        params["bill_code"] = args.bill_code
    if args.limit > 0:
        params["row_limit"] = args.limit

    print(f"连接数据库 {DB_USER}@{DB_DSN} ...")
    try:
        connection = oracledb.connect(
            user=DB_USER, password=DB_PASSWORD, dsn=DB_DSN, tcp_connect_timeout=10
        )
    except oracledb.OperationalError as exc:
        print(f"数据库连接失败：{exc}\n请确认已连接内网/VPN，且 {DB_DSN} 可达。")
        return 1

    try:
        cursor = connection.cursor()

        if args.create_view:
            print(f"执行 create or replace view {VIEW_NAME} ...")
            cursor.execute(f"create or replace view {VIEW_NAME} as {VIEW_SELECT_SQL}")
            print("视图创建成功")

        print(f"执行查询（bill_code={args.bill_code or '全部'}, limit={args.limit or '不限'}）...")
        start = time.perf_counter()
        cursor.execute(query, params)
        headers = [d[0] for d in cursor.description]
        rows = cursor.fetchall()
        elapsed = time.perf_counter() - start

        print(f"查询完成：{len(rows)} 行，耗时 {elapsed:.3f} 秒")
        print("列：" + " | ".join(headers))
        for row in rows[:5]:
            print("   ", " | ".join("" if v is None else str(normalize(v)) for v in row))
        if len(rows) > 5:
            print(f"    ... 其余 {len(rows) - 5} 行见 Excel")

        mode = "创建视图后查询视图" if args.create_view else "直接执行视图 select"
        write_excel(output_path, headers, rows, query, elapsed, mode)
        print(f"结果已导出：{output_path}")
    finally:
        connection.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
