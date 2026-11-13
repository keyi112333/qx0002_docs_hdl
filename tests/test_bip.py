#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BIP 登录 + 工作台搜索独立测试脚本

只执行“BIP 登录 + 工作台 Alt+K 搜索并打开菜单”部分：
- 不加载 UserConfig，不执行 initEnv（不会关闭其他进程）
- BIP 地址、浏览器参数、登录环境直接读取 config.xlsx 的系统配置
- 账号密码从 Windows 凭据管理器读取（系统配置.BIP凭据目标名）

默认搜索并打开【销售报关单】页面，可用 --menu 覆盖。

用法：
    python tests/test_bip.py
    python tests/test_bip.py --menu 销售报关单
    python tests/test_bip.py --credential QX0002_BIP --url https://bip.example.com
    python tests/test_bip.py --pause 30
    python tests/test_bip.py --config "data/.../config.xlsx"
"""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config_loader import ConfigError, ConfigLoader
from log.loggingConfig import logger
from web.bip.login import MENU_NAME, BipLogin

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BIP 登录 + 工作台搜索独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--menu", default=MENU_NAME, help=f"工作台搜索的菜单名，默认 {MENU_NAME}")
    parser.add_argument("--url", default=None, help="BIP 登录地址，缺省按系统配置.BIP地址拼接")
    parser.add_argument("--credential", default=None, help="凭据目标名，缺省取系统配置.BIP凭据目标名")
    parser.add_argument("--pause", type=int, default=30, help="搜索后浏览器保持打开的秒数，0 表示立即关闭")
    return parser.parse_args()


def run(args: argparse.Namespace) -> None:
    config_path = Path(args.config)
    if not config_path.exists():
        raise ConfigError(f"config.xlsx 不存在: {config_path}")
    loader = ConfigLoader(str(config_path))

    credential_name = str(args.credential or loader.get("系统配置.BIP凭据目标名") or "").strip()
    if not credential_name:
        raise ConfigError("未配置 BIP 凭据目标名: 系统配置.BIP凭据目标名（或用 --credential 指定）")

    logger.info(f"测试菜单: {args.menu}，凭据目标名: {credential_name}")
    bip = BipLogin(
        loader,
        login_url=args.url or "",
        credential_name=credential_name,
    )
    try:
        bip.open()
        logger.info("BIP 登录成功，开始工作台搜索")
        page = bip.search_menu(args.menu)
        print(f"PASS: 已打开【{args.menu}】")
        print(f"页面标题: {page.title()}")
        print(f"页面地址: {page.url}")
        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认页面...")
            page.wait_for_timeout(args.pause * 1000)
    finally:
        bip.close()


def main() -> int:
    args = parse_args()
    try:
        run(args)
        print("测试通过")
        return 0
    except ConfigError as e:
        print(f"测试前置条件失败: {e}")
        return 1
    except Exception as e:
        logger.exception("BIP 测试失败")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
