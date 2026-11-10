#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口独立测试脚本

只执行“单一窗口登录 + 报关数据查询”部分：
- 不加载 UserConfig，不执行 initEnv（不会关闭其他进程）
- 公司、凭据目标名、浏览器参数直接读取 config.xlsx
- 卡口令从 Windows 凭据管理器读取

用法：
    python tests/test_single_window.py
    python tests/test_single_window.py --company 珠海 --pause 60
    python tests/test_single_window.py --config "data/.../config.xlsx"
"""
import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config_loader import ConfigError, ConfigLoader
from log.loggingConfig import logger
from utils.credential import CredentialError, get_password
from web.browser import BrowserManager
from web.single_window.login import SingleWindowLogin
from web.single_window.query import DeclarationQuery

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="单一窗口登录+查询独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--company", default=None, help="公司名称，缺省取第一家启用的公司")
    parser.add_argument("--days", type=int, default=7, help="最近操作时间天数，默认 7 天")
    parser.add_argument("--pause", type=int, default=30, help="查询后浏览器保持打开的秒数，0 表示立即关闭")
    parser.add_argument("--no-submit", action="store_true", help="只填写查询条件、不点击查询，便于人工确认")
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
    credential_target = str(company.get("凭据目标名") or "").strip()
    if not credential_target:
        raise ConfigError(f"公司未配置凭据目标名: {company_name}")

    logger.info(f"测试公司: {company_name}，凭据目标名: {credential_target}")
    password = get_password(credential_target)

    headless = _as_bool(loader.get("系统配置.无头模式", False))
    channel = loader.get("系统配置.浏览器通道", "msedge")
    url = loader.get("系统配置.单一窗口地址", "https://www.singlewindow.cn/#/")
    page_timeout = int(loader.get("系统配置.页面超时(秒)", 30)) * 1000
    login_timeout = int(loader.get("系统配置.登录超时(秒)", 120)) * 1000

    browser = BrowserManager(headless=headless, channel=channel)
    browser.start()
    try:
        login = SingleWindowLogin(browser, url=url, page_timeout=page_timeout)
        login.login(password, timeout=login_timeout)
        goods_page = login.open_declaration_query()

        end_date = date.today()
        start_date = end_date - timedelta(days=args.days - 1)
        DeclarationQuery(goods_page).search(start_date, end_date, submit=not args.no_submit)

        action = "已填写查询条件（未提交）" if args.no_submit else "已提交查询"
        print(f"PASS: {company_name} {action}，最近操作时间 {start_date} ~ {end_date}")
        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认查询结果...")
            goods_page.wait_for_timeout(args.pause * 1000)
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
        logger.error(f"单一窗口测试失败: {e}")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
