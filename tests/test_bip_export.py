#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BIP【外币汇率-全局】导出独立测试脚本

流程（对应 recordings/bip_export.py）：
1. 登录 BIP → 工作台 Alt+K 搜索【外币汇率-全局】并打开；
2. 搜索并选中汇率方案 E001 基准汇率方案-中间汇率（记账）；
3. 读取列表总条数，每页设为 5000 条；
4. 输出到 EXCEL 文件：
   - 总条数 ≤ 5000：一次导出；
   - 总条数 > 5000：导出第 1 页后再跳转到第 2 页导出，最后把两次结果合并。

超过 10000 条时两次导出只能覆盖前 10000 条，脚本会给出告警。

用法：
    python tests/test_bip_export.py --credential QX0002_BIP
    python tests/test_bip_export.py --credential QX0002_BIP --url http://218.14.150.108:8088
    python tests/test_bip_export.py --credential QX0002_BIP --pause 30
"""
import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl import Workbook

from config.config_loader import ConfigError, ConfigLoader
from log.loggingConfig import logger
from web.bip.login import LOGIN_PATH, BipLogin

MENU_NAME = "外币汇率-全局"
SCHEME_KEYWORD = "E001"
SCHEME_NAME = "E001 基准汇率方案-中间汇率（记账）"
PAGE_SIZE = 5000

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)
DEFAULT_OUTPUT_DIR = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "03_Output" / "BIP外币汇率"
)
# recordings/bip_export.py 使用的测试环境登录地址
DEFAULT_LOGIN_URL = "http://218.14.150.108:8088" + LOGIN_PATH

TOTAL_PATTERN = re.compile(r"共\s*(\d+)\s*条")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BIP 外币汇率-全局 导出独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--url", default=DEFAULT_LOGIN_URL, help="BIP 登录地址，可只传域名")
    parser.add_argument("--credential", default=None, help="凭据目标名，缺省取系统配置.BIP凭据目标名")
    parser.add_argument("--out-dir", default=str(DEFAULT_OUTPUT_DIR), help="导出目录")
    parser.add_argument("--pause", type=int, default=30, help="导出后浏览器保持打开的秒数，0 表示立即关闭")
    return parser.parse_args()


def _login_url(value: str) -> str:
    url = str(value or "").strip() or DEFAULT_LOGIN_URL
    return url if LOGIN_PATH in url else url.rstrip("/") + LOGIN_PATH


def _main_frame(page):
    return page.frame_locator("#forwardFrame").frame_locator("#mainiframe")


def _scopes(page):
    """业务弹窗的元素可能在主内容 iframe、forwardFrame 或顶层文档。"""
    return (_main_frame(page), page.frame_locator("#forwardFrame"), page)


def _locate(page, make_locator, timeout: int = 15000):
    """跨 scope 轮询找到第一个可见元素。"""
    deadline = time.monotonic() + timeout / 1000.0
    while True:
        for scope in _scopes(page):
            try:
                locator = make_locator(scope)
                if locator.count() > 0:
                    first = locator.first
                    if first.is_visible():
                        return first
            except Exception:
                continue
        if time.monotonic() >= deadline:
            raise RuntimeError("页面元素在超时时间内未出现")
        page.wait_for_timeout(300)


def read_total(page) -> int:
    """读取列表分页处的“共 N 条”，解析失败返回 0。"""
    for scope in _scopes(page):
        try:
            matches = scope.get_by_text(TOTAL_PATTERN)
            total = matches.count()
        except Exception:
            continue
        for index in range(min(total, 20)):
            try:
                text = matches.nth(index).inner_text(timeout=1000)
            except Exception:
                continue
            found = TOTAL_PATTERN.search(text or "")
            if found:
                return int(found.group(1))
    return 0


def select_scheme(page) -> None:
    """搜索并选中 E001 基准汇率方案-中间汇率（记账）。"""
    box = _locate(page, lambda scope: scope.get_by_role("textbox", name="请搜索"))
    box.click()
    box.fill(SCHEME_KEYWORD)
    page.wait_for_timeout(1000)

    # 录制里先点主内容区的搜索图标再点方案链接；图标定位失败时不影响后续查找。
    try:
        icon = _main_frame(page).locator("i").first
        if icon.is_visible():
            icon.click()
    except Exception as exc:
        logger.warning("BIP方案搜索图标未点击: %s", exc)

    link = _locate(
        page,
        lambda scope: scope.locator("a").filter(has_text=SCHEME_NAME),
        timeout=30000,
    )
    link.click()
    page.wait_for_timeout(1500)
    logger.info("BIP已选择汇率方案: %s", SCHEME_NAME)


def set_page_size(page, size: int = PAGE_SIZE) -> None:
    """把列表每页条数设为 size（rc_select_0）。"""
    select = _locate(page, lambda scope: scope.locator("#rc_select_0"), timeout=10000)
    select.click()
    select.fill(str(size))
    select.press("Enter")
    page.wait_for_timeout(2000)
    logger.info("BIP列表每页条数已设为 %d", size)


def jump_to_page(page, number: int) -> None:
    """跳转到指定页码（rc-pagination / antd 分页的“跳至页”输入框）。"""
    makers = (
        lambda scope: scope.locator(".rc-pagination-options-quick-jumper input"),
        lambda scope: scope.locator(".ant-pagination-options-quick-jumper input"),
    )
    last_error = None
    for maker in makers:
        try:
            box = _locate(page, maker, timeout=5000)
        except Exception as exc:
            last_error = exc
            continue
        box.click()
        box.fill(str(number))
        box.press("Enter")
        page.wait_for_timeout(2000)
        logger.info("BIP列表已跳转到第 %d 页", number)
        return
    raise RuntimeError(f"未找到分页“跳至页”输入框: {last_error}")


def _find_visible(page, make_locator, timeout: int = 2000):
    """找可见元素，超时返回 None（不抛异常）。"""
    try:
        return _locate(page, make_locator, timeout=timeout)
    except RuntimeError:
        return None


def open_output_entry(page) -> None:
    """展开工具栏“输出”入口。

    录制里点“输出”前先点了一个无文字图标按钮
    （get_by_role("button").filter(has_text=^$).nth(1)），工具栏菜单才会展开。
    """
    output_menu = lambda scope: scope.get_by_role("menuitem", name="输出")
    if _find_visible(page, output_menu, timeout=1000):
        return
    # 录制中的工具栏按钮挂在弹窗页本身，优先在顶层文档找。
    scopes = (page, _main_frame(page), page.frame_locator("#forwardFrame"))
    for scope in scopes:
        try:
            buttons = scope.get_by_role("button").filter(has_text=re.compile(r"^$"))
            total = buttons.count()
        except Exception:
            continue
        for index in (1, 0, 2):
            if index >= total:
                continue
            button = buttons.nth(index)
            try:
                if not button.is_visible():
                    continue
                button.click(timeout=5000)
            except Exception as exc:
                logger.warning("BIP工具栏图标按钮(nth=%d)点击失败: %s", index, exc)
                continue
            page.wait_for_timeout(800)
            if _find_visible(page, output_menu, timeout=2000):
                logger.info("BIP工具栏输出入口已展开(nth=%d)", index)
                return
    logger.warning("BIP工具栏输出入口未展开，继续查找“输出”菜单")


def export_excel(page, output: Path, timeout: int = 180000) -> Path:
    """辅助功能/工具栏“输出” → 输出到EXCEL文件 → 确定(Y) → 保存下载文件。"""
    open_output_entry(page)
    menu = _locate(page, lambda scope: scope.get_by_role("menuitem", name="输出"), timeout=30000)
    menu.click()
    page.wait_for_timeout(500)

    radio = _locate(
        page,
        lambda scope: scope.get_by_role("radio", name="输出到EXCEL文件"),
        timeout=15000,
    )
    radio.check()
    page.wait_for_timeout(500)

    with page.expect_download(timeout=timeout) as info:
        confirm = _locate(page, lambda scope: scope.get_by_role("button", name="确定(Y)"))
        confirm.click()
    download = info.value
    output.parent.mkdir(parents=True, exist_ok=True)
    download.save_as(str(output))
    logger.info("BIP导出文件已保存: %s", output)
    return output


def merge_xlsx(parts, output: Path) -> int:
    """按行合并多个导出的首张工作表，重复表头只保留一次；返回合并后的数据行数。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "合并"
    header = None
    rows = 0
    for path in parts:
        source = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            for values in source.worksheets[0].iter_rows(values_only=True):
                if all(value is None or str(value).strip() == "" for value in values):
                    continue
                if header is None:
                    header = tuple(values)
                    sheet.append(list(values))
                    continue
                if tuple(values) == header:
                    continue  # 跳过重复表头
                sheet.append(list(values))
                rows += 1
        finally:
            source.close()
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    logger.info("BIP导出结果已合并: %s（数据行 %d）", output, rows)
    return rows


def run(args: argparse.Namespace) -> None:
    config_path = Path(args.config)
    if not config_path.exists():
        raise ConfigError(f"config.xlsx 不存在: {config_path}")
    loader = ConfigLoader(str(config_path))

    credential_name = str(args.credential or loader.get("系统配置.BIP凭据目标名") or "").strip()
    if not credential_name:
        raise ConfigError("未配置 BIP 凭据目标名: 系统配置.BIP凭据目标名（或用 --credential 指定）")

    out_dir = Path(args.out_dir)
    bip = BipLogin(loader, login_url=_login_url(args.url), credential_name=credential_name)
    try:
        bip.open()
        page = bip.search_menu(MENU_NAME)
        select_scheme(page)

        total = read_total(page)
        if total:
            logger.info("BIP列表总条数: %d", total)
        else:
            logger.warning("未能读取列表总条数，按单次导出处理")

        set_page_size(page, PAGE_SIZE)
        stamp = datetime.now().strftime("%Y%m%d %H%M%S")

        if total > PAGE_SIZE:
            if total > PAGE_SIZE * 2:
                logger.warning("总条数 %d 超过两次导出上限 %d，仅导出前 %d 条",
                               total, PAGE_SIZE * 2, PAGE_SIZE * 2)
            first = export_excel(page, out_dir / f"{MENU_NAME}-第1批-{stamp}.xlsx")
            jump_to_page(page, 2)
            second = export_excel(page, out_dir / f"{MENU_NAME}-第2批-{stamp}.xlsx")
            merged = out_dir / f"{MENU_NAME}-合并-{stamp}.xlsx"
            rows = merge_xlsx([first, second], merged)
            print(f"PASS: 共 {total} 条，已分两批导出并合并")
            print(f"合并文件: {merged}（数据行 {rows}）")
        else:
            output = export_excel(page, out_dir / f"{MENU_NAME}-{stamp}.xlsx")
            print(f"PASS: 共 {total or '未知'} 条，已单次导出")
            print(f"导出文件: {output}")

        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认导出结果...")
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
        logger.exception("BIP 外币汇率导出测试失败")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
