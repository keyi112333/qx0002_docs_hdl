#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
销售发票汇总视图（RPA_xsfphz）SQL 独立测试脚本

被测 SQL（多个出库单合并一单发票，出库单号/报关单号/合同号用逗号拼接且去重）：
select fph.pk_group, fph.pk_org_v, fph.vbillcode, fph.vdef3,
       LISTAGG(DISTINCT fpb.vsrccode, ',')  ... as 销售出库单号,
       LISTAGG(DISTINCT bg1.vbillcode, ',') ... as 销售报关单号,
       LISTAGG(DISTINCT bg1.contractno, ',') ... as 销售报关单的合同号,
       LISTAGG(DISTINCT fpb.vbdef8, ',')    ... as 销售发票的合同号
from so_saleinvoice fph
left join so_saleinvoice_b fpb on fph.csaleinvoiceid = fpb.csaleinvoiceid and fpb.dr=0
left join bd_busitype lc on lc.pk_busitype = fph.cbiztypeid and lc.dr=0
left join (销售报关单关联销售出库单) bg1 on bg1.outvbillcode = fpb.vsrccode
where fph.dr=0 and lc.businame in ('出口销售','出口销售（合约）')
group by fph.pk_group, fph.pk_org_v, fph.vbillcode, fph.vdef3

注意：原始 SQL 本身不是视图定义，脚本中的 --create-view 只是把这条查询包成
测试视图（RPA_xsfphz）方便对比查询结果。

默认只执行 select 语句（不改动数据库），查询结果写入 Excel。

用法：
    python tests/test_rpa_xsfphz_view.py
    python tests/test_rpa_xsfphz_view.py --bill-code ZHHSI25070100049495
    python tests/test_rpa_xsfphz_view.py --limit 100
    python tests/test_rpa_xsfphz_view.py --create-view
    python tests/test_rpa_xsfphz_view.py --output d:/tmp/result.xlsx
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

VIEW_NAME = "RPA_xsfphz"

# 视图内的 select 语句（与原始 SQL 一致，末行 __BILL_FILTER__ 为可选单据号过滤的插入点）
VIEW_SELECT_SQL = """
select fph.pk_group/*销售发票.集团*/ as 集团,
       fph.pk_org_v/*销售发票.组织*/ as 组织,
       fph.vbillcode/*销售发票.单据号*/ as BIP发票单据号,
       fph.vdef3/*销售发票.税务发票号*/ as 税务发票号,
       LISTAGG(DISTINCT fpb.vsrccode, ',')
                WITHIN GROUP (ORDER BY fpb.vsrccode) as 销售出库单号,
       LISTAGG(DISTINCT bg1.vbillcode, ',')
                WITHIN GROUP (ORDER BY bg1.vbillcode) as 销售报关单号,
       LISTAGG(DISTINCT bg1.contractno, ',')
                WITHIN GROUP (ORDER BY bg1.contractno) as 销售报关单的合同号,
       LISTAGG(DISTINCT fpb.vbdef8, ',')
                WITHIN GROUP (ORDER BY fpb.vbdef8) as 销售发票的合同号
from so_saleinvoice/*销售发票表头*/ fph
left join so_saleinvoice_b/*销售发票表体*/ fpb
       on fph.csaleinvoiceid = fpb.csaleinvoiceid and fpb.dr=0
left join bd_busitype/*业务流程*/ lc
       on lc.pk_busitype = fph.cbiztypeid and lc.dr=0
left join (
    /*销售报关单关联销售出库单*/
    select bgh.vbillcode/*销售报关单表头.报关单号*/,
           bgh.contractno/*销售报关单表头.合同号*/,
           bgb.outvbillcode/*销售报关单表体.出库单号*/
    from iot_iota_h/*销售报关单表头*/ bgh
    left join iot_iota_b/*销售报关单表体*/ bgb
           on bgh.pk_iota_h = bgb.pk_iota_h and bgb.dr=0
    where bgh.dr=0
) bg1 on bg1.outvbillcode = fpb.vsrccode
where fph.dr=0 and lc.businame in ('出口销售','出口销售（合约）') /*__BILL_FILTER__*/
group by fph.pk_group, fph.pk_org_v, fph.vbillcode, fph.vdef3
"""

BILL_FILTER_MARK = "/*__BILL_FILTER__*/"

DEFAULT_OUTPUT_DIR = ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "03_Output"


def build_query(bill_code: str = "", limit: int = 0) -> str:
    """拼接待执行的查询 SQL（可选按单据号过滤、限制行数）"""
    sql = VIEW_SELECT_SQL
    if bill_code:
        sql = sql.replace(BILL_FILTER_MARK, "and fph.vbillcode = :bill_code")
    else:
        sql = sql.replace(BILL_FILTER_MARK, "")
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
    parser = argparse.ArgumentParser(description="测试 RPA_xsfphz 视图 SQL 并导出 Excel")
    parser.add_argument("--bill-code", default="", help="按销售发票号过滤（默认不过滤）")
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
            view_sql = VIEW_SELECT_SQL.replace(BILL_FILTER_MARK, "")
            cursor.execute(f"create or replace view {VIEW_NAME} as {view_sql}")
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
