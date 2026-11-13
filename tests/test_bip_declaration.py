#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BIP【销售报关单】新增写入独立测试脚本

数据来源（二选一）：
1. --detail-file：报关单详情导出 Excel，按 data_processing.declaration 解析，
   再经 data_processing.rules.result_row_from_declaration 得到与结果表一致的字段；
2. --result-file（可选 --bill-no）：运行结果表，直接取该单号已写入的结果行。

页面流程（对应销售报关单录入录制）：
登录 BIP → 工作台搜索【销售报关单】→ 新增 → 按录制选择器填写表头 →
保存（--dry-run 时只填不保存）。

用法：
    python tests/test_bip_declaration.py --credential QX0002_BIP \
        --detail-file "docs/示例数据/E20260001895493493-报关单详情-20261111 135639.xlsx"
    python tests/test_bip_declaration.py --credential QX0002_BIP \
        --result-file "data/.../运行结果表.xlsx" --bill-no 531620260004864277
    python tests/test_bip_declaration.py --credential QX0002_BIP --detail-file ... --dry-run
"""
import argparse
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config_loader import ConfigError, ConfigLoader
from data_processing.declaration import parse_detail
from data_processing.result_table import ResultTable
from data_processing.rules import result_row_from_declaration
from log.loggingConfig import logger
from web.bip.login import LOGIN_PATH, BipLogin

MENU_NAME = "销售报关单"
# 与运行结果表第 1～18 列一致，也是销售报关单表头需要的数据
BIP_FIELDS = [
    "销售组织", "报关单号", "贸易种类", "报关客户", "报关币种", "合同号",
    "报关数量", "报关单价", "总价", "运输费", "保险费", "出口日期", "手册号",
    "贸易方式", "报关汇率", "监管方式", "成交方式", "保税核注清单编号",
]

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)
DEFAULT_LOGIN_URL = "http://218.14.150.108:8088" + LOGIN_PATH

# 录制中的字段选择器（容器内的输入框，按顺序填写）
FIELD_SELECTORS = {
    "报关币种": (".pk_currency input", ".pk_currency .wui-input"),
    "合同号": (".contractno input", ".contractno .wui-input"),
    "报关数量": (".wui-input.md.base-input.number-formcontrol",),
    "报关单价": (".nqtorigtaxprice input", ".nqtorigtaxprice .wui-input"),
    "总价": (".norigtaxmny input", ".norigtaxmny .wui-input"),
    "运输费": (".shipping input", ".shipping .wui-input"),
    "保险费": (".insurance input", ".insurance .wui-input"),
    "贸易方式": (".def3 input", ".def3 .wui-input"),
    "手册号": (".def2 input", ".def2 .wui-input"),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BIP 销售报关单新增写入独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--url", default=DEFAULT_LOGIN_URL, help="BIP 登录地址，可只传域名")
    parser.add_argument("--credential", default=None, help="凭据目标名，缺省取系统配置.BIP凭据目标名")
    parser.add_argument("--detail-file", default=None, help="报关单详情导出 Excel（优先）")
    parser.add_argument("--result-file", default=None, help="运行结果表 Excel")
    parser.add_argument("--bill-no", default=None, help="结果表中的报关单号，缺省取最后一行")
    parser.add_argument("--supervision-mode", default="", help="详情解析用的中文监管方式（判“进料”用）")
    parser.add_argument("--dry-run", action="store_true", help="只填写不保存")
    parser.add_argument("--pause", type=int, default=30, help="填写/保存后浏览器保持打开的秒数")
    return parser.parse_args()


def _login_url(value: str) -> str:
    url = str(value or "").strip() or DEFAULT_LOGIN_URL
    return url if LOGIN_PATH in url else url.rstrip("/") + LOGIN_PATH


def _number_text(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def _date_text(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value or "").strip()
    if re.fullmatch(r"\d{8}", text):
        return f"{text[:4]}-{text[4:6]}-{text[6:]}"
    return text


def load_bip_data(args: argparse.Namespace) -> dict:
    """从详情导出或结果表取出写入 BIP 的字段。"""
    if args.detail_file:
        declaration = parse_detail(
            Path(args.detail_file),
            supervision_mode=str(args.supervision_mode or "").strip(),
        )
        logger.info("详情解析完成: %s", declaration.summary())
        data = result_row_from_declaration(declaration)
        # 结果表按报关单号单独 upsert，result_row_from_declaration 不含该列
        data["报关单号"] = declaration.no
        return data

    if args.result_file:
        table = ResultTable(Path(args.result_file))
        try:
            if args.bill_no:
                row = table.get(args.bill_no)
                if row is None:
                    raise ConfigError(f"结果表中未找到报关单号: {args.bill_no}")
            else:
                rows = table.rows()
                if not rows:
                    raise ConfigError(f"结果表没有数据行: {args.result_file}")
                row = rows[-1]
        finally:
            table.close()
        return {field: row.get(field) for field in BIP_FIELDS}

    raise ConfigError("请用 --detail-file（详情导出）或 --result-file（运行结果表）指定数据来源")


def _main_frame(page):
    return page.frame_locator("#forwardFrame").frame_locator("#mainiframe")


def _scopes(page):
    """表单元素可能在主内容 iframe、forwardFrame 或弹窗顶层文档。"""
    return (_main_frame(page), page.frame_locator("#forwardFrame"), page)


def _locate(page, make_locator, timeout: int = 15000, label: str = "页面元素"):
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
            raise RuntimeError(f"{label}在超时时间内未出现")
        page.wait_for_timeout(300)


def _find_visible(page, make_locator, timeout: int = 2000):
    try:
        return _locate(page, make_locator, timeout=timeout)
    except RuntimeError:
        return None


def _editable(page, selectors, label: str):
    """依次尝试候选选择器，取第一个可见且可编辑的输入框。"""
    for selector in selectors:
        locator = _find_visible(page, lambda scope, s=selector: scope.locator(s), timeout=3000)
        if locator is None:
            continue
        try:
            if locator.is_editable():
                return locator
        except Exception:
            continue
    raise RuntimeError(f"{label}输入框在超时时间内未出现")


def _textbox_by_index(page, index: int):
    """按录制中的 textbox 序号定位（nth）。"""
    for scope in _scopes(page):
        try:
            boxes = scope.get_by_role("textbox")
            if boxes.count() > index and boxes.nth(index).is_visible():
                return boxes.nth(index)
        except Exception:
            continue
    raise RuntimeError(f"未找到第 {index} 个可见输入框")


def _pick_option(page, value: str, timeout: int = 2000) -> bool:
    """引用/下拉字段输入后，若出现匹配选项则点选。"""
    text = str(value or "").strip()
    if not text:
        return False
    escaped = re.escape(text)
    makers = (
        lambda scope: scope.locator(".wui-select-item-option-content").filter(
            has_text=re.compile(escaped)),
        lambda scope: scope.locator(".wui-select-item").filter(has_text=re.compile(escaped)),
        lambda scope: scope.get_by_role("option", name=re.compile(escaped)),
    )
    for maker in makers:
        option = _find_visible(page, maker, timeout=timeout)
        if option is None:
            continue
        try:
            option.click(timeout=3000)
            page.wait_for_timeout(300)
            return True
        except Exception:
            continue
    return False


def _fill(locator, value, page, enter: bool = False) -> None:
    locator.click()
    # 结果表/详情里缺失的字段是 None，必须按空值填写，不能写成字符串 "None"
    locator.fill("" if value is None else str(value))
    if enter:
        locator.press("Enter")
    page.wait_for_timeout(300)


def click_new(page) -> None:
    button = _locate(page, lambda scope: scope.get_by_role("button", name="新增"),
                     timeout=30000, label="新增按钮")
    button.click()
    page.wait_for_timeout(3000)
    logger.info("BIP销售报关单已点击新增")


def fill_declaration(page, data: dict) -> None:
    """按录制顺序填写表头。"""
    # 销售组织：第一个输入框，输入公司后回车选中
    org = _textbox_by_index(page, 0)
    _fill(org, data.get("销售组织", ""), page, enter=True)
    page.wait_for_timeout(500)
    _pick_option(page, data.get("销售组织", ""))
    logger.info("BIP已填销售组织: %s", data.get("销售组织"))

    # 报关单号（录制为第 3 个输入框）
    code = _textbox_by_index(page, 2)
    _fill(code, data.get("报关单号", ""), page)
    logger.info("BIP已填报关单号: %s", data.get("报关单号"))

    # 贸易种类（录制为第 4 个输入框，值取监管方式）
    trade_kind = _textbox_by_index(page, 3)
    _fill(trade_kind, data.get("贸易种类", ""), page, enter=True)
    page.wait_for_timeout(500)
    _pick_option(page, data.get("贸易种类", ""))
    logger.info("BIP已填贸易种类: %s", data.get("贸易种类"))

    # 报关客户（录制为第 5 个输入框）
    customer = _textbox_by_index(page, 4)
    _fill(customer, data.get("报关客户", ""), page, enter=True)
    page.wait_for_timeout(500)
    _pick_option(page, data.get("报关客户", ""))
    logger.info("BIP已填报关客户: %s", data.get("报关客户"))

    # 报关币种
    currency = _editable(page, FIELD_SELECTORS["报关币种"], "报关币种")
    _fill(currency, data.get("报关币种", ""), page, enter=True)
    page.wait_for_timeout(500)
    _pick_option(page, data.get("报关币种", ""))
    logger.info("BIP已填报关币种: %s", data.get("报关币种"))

    # 合同号
    contract = _editable(page, FIELD_SELECTORS["合同号"], "合同号")
    _fill(contract, data.get("合同号", ""), page)
    logger.info("BIP已填合同号: %s", data.get("合同号"))

    # 报关数量
    qty = _editable(page, FIELD_SELECTORS["报关数量"], "报关数量")
    _fill(qty, _number_text(data.get("报关数量")), page)
    logger.info("BIP已填报关数量: %s", _number_text(data.get("报关数量")))

    # 报关单价
    price = _editable(page, FIELD_SELECTORS["报关单价"], "报关单价")
    _fill(price, _number_text(data.get("报关单价")), page)
    logger.info("BIP已填报关单价: %s", _number_text(data.get("报关单价")))

    # 总价：录制里只点击，由数量 × 单价自动计算
    total = _editable(page, FIELD_SELECTORS["总价"], "总价")
    total.click()
    page.wait_for_timeout(300)
    logger.info("BIP总价由页面自动计算，未手工覆盖（结果表值=%s）",
                _number_text(data.get("总价")))

    # 运输费 / 保险费
    for field in ("运输费", "保险费"):
        locator = _editable(page, FIELD_SELECTORS[field], field)
        _fill(locator, _number_text(data.get(field)), page)
        logger.info("BIP已填%s: %s", field, _number_text(data.get(field)))

    # 出口日期（录制为第一个 .single-input）
    export_date = _find_visible(page, lambda scope: scope.locator(".single-input"), timeout=5000)
    if export_date is None:
        export_date = _locate(
            page,
            lambda scope: scope.get_by_role("textbox", name=re.compile("出口日期")),
            timeout=8000,
            label="出口日期",
        )
    date_value = _date_text(data.get("出口日期"))
    _fill(export_date, date_value, page)
    logger.info("BIP已填出口日期: %s", date_value)

    # 贸易方式（固定值“出口”）
    trade_mode = _editable(page, FIELD_SELECTORS["贸易方式"], "贸易方式")
    _fill(trade_mode, data.get("贸易方式") or "出口", page, enter=True)
    page.wait_for_timeout(500)
    _pick_option(page, data.get("贸易方式") or "出口")
    logger.info("BIP已填贸易方式: %s", data.get("贸易方式") or "出口")

    # 手册号（备案号）
    record_no = _editable(page, FIELD_SELECTORS["手册号"], "手册号")
    _fill(record_no, data.get("手册号", ""), page)
    logger.info("BIP已填手册号: %s", data.get("手册号"))

    for field in ("监管方式", "成交方式", "保税核注清单编号", "报关汇率"):
        if data.get(field):
            logger.info("BIP字段 [%s] 录制中未固定填写，保持页面默认: %s", field, data[field])


def save_declaration(page) -> bool:
    """保存单据；返回是否看到保存成功提示。"""
    try:
        button = _locate(page, lambda scope: scope.get_by_role("button", name=re.compile(r"^保存")),
                         timeout=15000, label="保存按钮")
    except RuntimeError:
        raise RuntimeError("未找到“保存”按钮，可用 --dry-run 先检查页面填写结果") from None

    button.click()
    page.wait_for_timeout(3000)

    confirm = _find_visible(page, lambda scope: scope.get_by_role("button", name="确定(Y)"), 3000)
    if confirm is not None:
        confirm.click()
        page.wait_for_timeout(3000)

    if _find_visible(page, lambda scope: scope.get_by_text("保存成功"), 5000):
        logger.info("BIP销售报关单保存成功")
        return True
    logger.warning("未捕获“保存成功”提示，请人工确认单据是否已保存")
    return False


def run(args: argparse.Namespace) -> None:
    config_path = Path(args.config)
    if not config_path.exists():
        raise ConfigError(f"config.xlsx 不存在: {config_path}")
    loader = ConfigLoader(str(config_path))

    credential_name = str(args.credential or loader.get("系统配置.BIP凭据目标名") or "").strip()
    if not credential_name:
        raise ConfigError("未配置 BIP 凭据目标名: 系统配置.BIP凭据目标名（或用 --credential 指定）")

    data = load_bip_data(args)
    logger.info("待写入BIP销售报关单: %s", {k: data.get(k) for k in BIP_FIELDS if data.get(k)})

    bip = BipLogin(loader, login_url=_login_url(args.url), credential_name=credential_name)
    try:
        bip.open()
        page = bip.search_menu(MENU_NAME)
        click_new(page)
        fill_declaration(page, data)

        if args.dry_run:
            print(f"DRY-RUN: 已填写【{MENU_NAME}】未保存，报关单号={data.get('报关单号')}")
        else:
            saved = save_declaration(page)
            print(f"PASS: 已提交【{MENU_NAME}】保存，报关单号={data.get('报关单号')}"
                  f"（保存成功提示={saved}）")

        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认录入结果...")
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
        logger.exception("BIP 销售报关单写入测试失败")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
