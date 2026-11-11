#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式流程直接测试：真实登录单一窗口/电子口岸、真实查询导出、真实取 BIP 附件数据、真实下载
并回写运行结果表（不使用任何假数据）。

与 main.py 的差别：
- 不执行 initEnv（不会关闭 Excel 等进程）；
- 可指定公司、查询天数、本次最多处理几个单号，便于小批量试跑；
- 全程打印待处理清单、下载结果与异常原因，便于人工核对。

用法：
    python tests/test_qx0002_pipeline_e2e.py --company 珠海 --days 7 --max-bills 1
    python tests/test_qx0002_pipeline_e2e.py --company 珠海 --days 7 --dry-run
    python tests/test_qx0002_pipeline_e2e.py --company 珠海 --mode rerun --max-bills 2
    # 不查 BIP 数据库，直接用已导出的附件表做比对
    python tests/test_qx0002_pipeline_e2e.py --company 珠海 ^
        --bip-file "D:\RPA CODE\qx0002_Docs_Hdl\data\测试工作目录\QX0002_报关单和放行通知单等单据的处理\03_Output\珠海\BIP已上传附件表\珠海-BIP已上传附件表-20261111 153812.xlsx"

前置条件：
- 对应公司的 U 盾端口已开启（RPUSI），Windows 凭据中有该公司的卡口令；
- 01_Config/config.xlsx 中系统配置、查询配置（BIP 库）、公司盾配置已按实施环境填写；
- 运行期间不要打开 03_Output 下待写入的文件。
"""
import argparse
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config import Config
from core.context import TaskContext
from log.loggingConfig import logger
from utils.helpers import format_time
from workflows.qx0002_docs_hdl.subtask_declaration import DeclarationSubtask
from workflows.qx0002_docs_hdl.subtask_download import DownloadSubtask
from workflows.qx0002_docs_hdl.subtask_login import SingleWindowLoginSubtask

USER_CONFIG = ROOT / "data" / "UserConfig.xlsx"


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


class _ConfigProxy:
    """在真实配置外层覆盖查询天数，避免为试跑修改 config.xlsx"""

    def __init__(self, base, days: int = 0):
        self._base = base
        self._days = days

    def get(self, key, default=None):
        if self._days and key == "业务参数.查询最近操作天数":
            return self._days
        return self._base.get(key, default)

    def get_path(self, *args, **kwargs):
        return self._base.get_path(*args, **kwargs)

    def read_table(self, *args, **kwargs):
        return self._base.read_table(*args, **kwargs)

    @property
    def project_name(self):
        return getattr(self._base, "project_name", "QX0002")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="QX0002 正式流程直接测试")
    parser.add_argument("--company", default=None, help="公司名称，缺省取第一家启用的公司")
    parser.add_argument("--days", type=int, default=0,
                        help="查询最近操作天数，覆盖 config 的业务参数（单一窗口按 7 天分段）")
    parser.add_argument("--max-bills", type=int, default=0,
                        help="本次最多下载几个单号，0 表示全部；仅影响下载阶段，清单会完整生成")
    parser.add_argument("--mode", default="normal", choices=("normal", "rerun"),
                        help="normal/rerun；rerun 时同名文件按 skip 复用")
    parser.add_argument("--dry-run", action="store_true",
                        help="只生成待处理清单并写入结果表，不下载单据")
    parser.add_argument("--bip-file", default="",
                        help="直接从该 Excel 读取 BIP 已上传附件数据，不查询 BIP 数据库（测试/重跑用）")
    return parser.parse_args()


def load_config(days: int):
    if not USER_CONFIG.exists():
        raise SystemExit(f"配置文件不存在: {USER_CONFIG}")
    os.chdir(ROOT)
    config = Config(str(USER_CONFIG))
    # 创建 01_Config~06_Temp 与 03_Output/{公司} 子目录
    config.ensure_directories()
    config.load_business_config()
    return _ConfigProxy(config, days)


def select_company(config, name: str = None) -> dict:
    companies = [
        company for company in config.read_table("公司盾配置") if _as_bool(company.get("启用"))
    ]
    if not companies:
        raise SystemExit("公司盾配置中没有启用的公司")
    if not name:
        return companies[0]
    for company in companies:
        if str(company.get("公司名称") or "").strip() == name:
            return company
    raise SystemExit(f"公司盾配置中未找到公司: {name}")


def print_todo(todo) -> None:
    print(f"待处理报关单 {len(todo)} 条：")
    for item in todo:
        print(f"  - 统一编号={item.unified_no} 报关单号={item.no} 监管方式={item.supervision_mode} "
              f"合同号={item.contract_no or '无'} 进出口日期={item.export_date}")


def main() -> int:
    args = parse_args()
    started = time.time()
    config = load_config(args.days)
    company = select_company(config, args.company)
    company_name = str(company.get("公司名称") or "").strip()
    print(f"公司={company_name} 模式={args.mode} 查询天数={args.days or '配置值'} "
          f"最多下载={args.max_bills or '全部'}")

    context = TaskContext()
    context.config = config
    try:
        if not SingleWindowLoginSubtask(context, company).run():
            print("FAIL 单一窗口登录失败，请检查 U 盾端口与卡口令")
            return 1
        print("PASS 单一窗口登录")

        bip_file = Path(args.bip_file) if args.bip_file else None
        if bip_file:
            print(f"BIP 附件数据来源: 本地文件 {bip_file}（不查询数据库）")
        declaration = DeclarationSubtask(context, company, args.mode, bip_file=bip_file)
        if not declaration.run():
            print("PASS 未生成待处理清单（无数据或无交集），流程结束")
            return 0

        stats = context.get("candidate_stats") or {}
        print("PASS 待处理清单: " + "，".join(f"{key}={value}" for key, value in stats.items()))
        print(f"  合并列表: {context.get('declaration_list_file')}")
        print(f"  BIP附件表: {context.get('bip_attachment_file')}")
        missing_segments = context.get("missing_segments") or []
        if missing_segments:
            print(f"  WARNING 以下分段导出失败，结果可能缺失: {', '.join(missing_segments)}")

        todo = context.get("todo") or []
        print_todo(todo)
        if args.dry_run:
            print("PASS 干跑结束（未下载单据）")
            return 0

        if args.max_bills and len(todo) > args.max_bills:
            context.set("todo", todo[:args.max_bills])
            print(f"WARNING 本次只下载前 {args.max_bills} 条，其余单号已写入结果表但状态仍为“待处理”")

        if not DownloadSubtask(context, company, args.mode).run():
            print("FAIL 下载阶段异常，详见日志")
            return 1

        records = context.get("download_records") or []
        print(f"下载结果 {len(records)} 条：")
        for record in records:
            docs = "、".join(sorted(record["docs"])) or "无"
            reason = "；".join(record["errors"])
            print(f"  - {record['item'].no} 状态={record['status']} 已下载={docs}"
                  + (f" 异常={reason}" if reason else ""))
        print(f"PASS 流程结束，耗时 {format_time(time.time() - started)}")
        return 0
    except Exception as e:
        logger.exception("正式流程直接测试失败")
        print(f"FAIL {e}")
        return 1
    finally:
        context.close_all_browsers()


if __name__ == "__main__":
    sys.exit(main())
