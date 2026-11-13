#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国电子口岸导出报关单独立测试脚本

只执行“电子口岸登录 + 出口退税联网稽查 + 报关单查询下载”部分：
- 不加载 UserConfig，不执行 initEnv（不会关闭其他进程）
- 公司、凭据目标名、浏览器参数直接读取 config.xlsx
- 卡口令从 Windows 凭据管理器读取（与单一窗口共用同一张 U 盾）

导出的报关单直接保存到输出目录：电子口岸导出的才是正确版本；
单一窗口打印的报关单仅供提取信息，存储在临时目录（见 test_single_window.py）。

归档路径：03_Output/{公司}/下载单据/YYYYMM/{报关单号}-{合同编号}-{出口日期}-报关单.pdf
其中报关单号、合同编号、出口日期来自单一窗口报关单信息，详情 Excel 解析未实现前用命令行参数。

用法：
    python tests/test_chinaport.py
    python tests/test_chinaport.py --entry-no 531620260004863835
    python tests/test_chinaport.py --entry-no 531620260004863835 --contract-no HT2026-018 --export-date 20260108
    python tests/test_chinaport.py --company 珠海 --entry-no 531620260004863835 --pause 30
    python tests/test_chinaport.py --config "data/.../config.xlsx"
"""
import argparse
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config_loader import ConfigError, ConfigLoader
from log.loggingConfig import logger
from utils.credential import CredentialError, get_password
from utils.naming import DECLARATION_DOC_TYPE, build_document_path, parse_export_date
from web.browser import BrowserManager
from web.chinaport.declaration import ChinaPortDeclarationExport
from web.chinaport.login import DEFAULT_URL, ChinaPortLogin

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)
DEFAULT_OUTPUT_DIR = ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "03_Output"
# 录制脚本（recordings/chinaport_login_export）中使用的报关单号，可用 --entry-no 覆盖
DEFAULT_ENTRY_NO = "531620260004863835"


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="中国电子口岸导出报关单独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--company", default=None, help="公司名称，缺省取第一家启用的公司")
    parser.add_argument(
        "--entry-no", action="append", default=[], metavar="报关单号",
        help=f"要导出的报关单号，可重复指定；缺省用录制值 {DEFAULT_ENTRY_NO}",
    )
    parser.add_argument("--out-dir", default=str(DEFAULT_OUTPUT_DIR), help="导出目录，默认 data/.../03_Output")
    parser.add_argument(
        "--contract-no", default=None,
        help="合同编号（来自单一窗口报关单合同协议号）；不传用占位符“无合同号”，仅用于文件命名",
    )
    parser.add_argument(
        "--export-date", default=None,
        help="出口日期 YYYYMMDD（来自单一窗口报关单详情）；不传用当天，仅用于命名与归档年月",
    )
    parser.add_argument("--url", default=None, help=f"电子口岸地址，默认 {DEFAULT_URL}")
    parser.add_argument("--credential", default=None, help="卡口令凭据目标名，缺省取公司盾配置中的凭据目标名")
    parser.add_argument("--pause", type=int, default=30, help="导出后浏览器保持打开的秒数，0 表示立即关闭")
    return parser.parse_args()


def load_company(config_path: Path, company_name: str = None):
    """读取公司盾配置，返回 (配置加载器, 公司行)"""
    if not config_path.exists():
        raise ConfigError(f"config.xlsx 不存在: {config_path}")

    loader = ConfigLoader(str(config_path))
    companies = [
        company
        for company in loader.read_table("公司盾配置")
        if str(company.get("启用") or "").strip() in ("是", "true", "True", "1")
    ]
    if not companies:
        raise ConfigError("公司盾配置中没有启用的公司")

    if company_name:
        for company in companies:
            if str(company.get("公司名称") or "").strip() == company_name:
                return loader, company
        raise ConfigError(f"公司盾配置中未找到公司: {company_name}")

    return loader, companies[0]


def run(args: argparse.Namespace) -> None:
    loader, company = load_company(Path(args.config), args.company)
    company_name = str(company.get("公司名称") or "").strip()
    company_path_name = str(company.get("公司全称") or company_name).strip()
    credential_target = str(args.credential or company.get("凭据目标名") or "").strip()
    if not credential_target:
        raise ConfigError(f"公司未配置凭据目标名: {company_name}")

    entry_nos = list(args.entry_no) or [DEFAULT_ENTRY_NO]
    logger.info(f"测试公司: {company_name}，凭据目标名: {credential_target}")
    logger.info(f"待导出报关单: {', '.join(entry_nos)}")
    # 合同编号/出口日期取自单一窗口报关单信息（解析未实现），测试用命令行参数
    contract_no = (args.contract_no or "").strip()
    export_date = parse_export_date(args.export_date) if args.export_date else date.today()
    if not (args.contract_no and args.export_date):
        logger.warning(
            f"未提供合同编号/出口日期，按 {contract_no or '无合同号'}/{export_date:%Y%m%d} 命名，"
            "接入单一窗口报关单信息后应改为取合同协议号与出口日期"
        )
    password = get_password(credential_target)

    headless = _as_bool(loader.get("系统配置.无头模式", False))
    channel = loader.get("系统配置.浏览器通道", "msedge")
    page_timeout = int(loader.get("系统配置.页面超时(秒)", 30)) * 1000
    login_timeout = int(loader.get("系统配置.登录超时(秒)", 120)) * 1000
    download_timeout = int(loader.get("系统配置.下载超时(秒)", 60)) * 1000

    browser = BrowserManager(headless=headless, channel=channel)
    browser.start()
    try:
        login = ChinaPortLogin(browser, url=args.url or DEFAULT_URL, page_timeout=page_timeout)
        login.login(password, timeout=login_timeout)
        app = login.open_export_tax_refund()

        exporter = ChinaPortDeclarationExport(
            app, page_timeout=page_timeout, download_timeout=download_timeout
        )
        exported = []
        declaration_dir = Path(args.out_dir) / company_path_name / "下载单据"
        logger.info(f"报关单归档目录: {declaration_dir}")
        for entry_no in entry_nos:
            exporter.query(entry_no)
            # 报关单号/合同编号/出口日期来自单一窗口报关单信息，输出到 03_Output/公司/下载单据/YYYYMM
            path = build_document_path(
                declaration_dir, entry_no, contract_no, export_date, DECLARATION_DOC_TYPE
            )
            exported.append(exporter.download_declaration(entry_no, path))

        print(f"PASS: {company_name} 电子口岸导出报关单完成，共 {len(exported)} 份")
        for path in exported:
            print(f"导出文件: {path}")
        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认导出结果...")
            app.wait_for_timeout(args.pause * 1000)
    finally:
        browser.close()


def main() -> int:
    args = parse_args()
    try:
        run(args)
        print("测试通过")
        return 0
    except (ConfigError, CredentialError) as e:
        print(f"测试前置条件失败: {e}")
        return 1
    except Exception as e:
        logger.exception("中国电子口岸测试失败")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
