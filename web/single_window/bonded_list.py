#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口-加工贸易账册-综合查询-保税核注清单查询与打印下载

入口：门户-业务应用-口岸执法申报-加工贸易账册（新窗口）-综合查询-保税核注清单查询。
企业类型/海关编码改为“点击输入框 → 空格 → 输入值 → 回车”确认，不再查找下拉项；
查不到数据时页面会弹“未查询到相关的清单数据”，需点确定关闭。
用完调用 close()：关闭账册窗口并回到门户首页，下一个模块重新进入。
"""
import re
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from web.base_page import BasePage

QUERY_ANCHOR = 'input[name="bondInvtNo"]'
ENTERPRISE_TYPE_PLACEHOLDER = "请选择要查询的企业类型"
CUSTOMS_CODE_PLACEHOLDER = (
    "请选择企业类型对应的海关编码；若选择授权企业，仅支持查询当前登录企业申报或录入的清单数据。"
)
START_DATE_INPUT = "#inputDateStart"
END_DATE_INPUT = "#inputDateEnd"
DEFAULT_HOME_URL = "https://www.singlewindow.cn/#/"

# 真实控件：文本显示域 + 隐藏编码域（日志实测）
ENTERPRISE_TYPE_HIDDEN = "#qryEntType"
ENTERPRISE_TYPE_NAME = "#qryEntTypeName"
IE_FLAG_HIDDEN = "#impExpMarkCd"
IE_FLAG_NAME = "#impExpMarkCdName"
IE_FLAG_EXPORT_CODE = "E"
IE_FLAG_EXPORT_NAME = "出口"

JS_SET_VALUE = """
(el, value) => {
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
}
"""


class BondedListQuery(BasePage):
    """保税核注清单查询与打印"""

    def __init__(self, portal_page, page_timeout: int = 30000, download_timeout: int = 60000,
                 home_url: str = DEFAULT_HOME_URL):
        super().__init__(portal_page)
        self.page_timeout = page_timeout
        self.download_timeout = download_timeout
        self.portal_page = portal_page
        self.home_url = home_url

    def open(self) -> None:
        """从门户打开保税核注清单查询页（加工贸易账册为新窗口）"""
        police = self._visible_locator(self.page.get_by_role("link", name="口岸执法申报"))
        if police is None:
            # 门户页头“业务应用”可能匹配到多个（含已展开的菜单项），点第一个
            self.page.get_by_text("业务应用", exact=True).first.click()
            police = self._wait_visible(self.page.get_by_role("link", name="口岸执法申报"))
        if police is None:
            raise RuntimeError("门户未找到【口岸执法申报】入口")
        police.click()

        with self.page.expect_popup() as popup_info:
            account = self._wait_visible(self.page.get_by_role("link", name="加工贸易账册"))
            if account is None:
                raise RuntimeError("未找到【加工贸易账册】入口")
            account.click()

        app = popup_info.value
        app.wait_for_load_state()
        app.set_default_timeout(self.page_timeout)
        self._handle_optional_dialogs(app)
        app.locator("a").filter(has_text="综合查询").first.click()
        app.get_by_text("保税核注清单查询").first.click()
        self.page = app
        logger.info("已进入保税核注清单查询")

    def query(self, bond_invt_no: str, start_date, end_date, enterprise_code: str = "4404260A2T") -> bool:
        """填写查询条件并查询

        - 企业类型：点击输入框 → 空格 → 输入“加工单位” → 回车；回车没选中时改点下拉项；
        - 企业海关编码：同法输入后回车；
        - 进出口标志：直接设置隐藏域/显示域（impExpMarkCd=E / 出口），避免下拉选成“进口”；
        返回 True 表示查到数据，False 表示页面提示“未查询到相关的清单数据”。
        """
        frame = self._find_frame()

        self._select_by_typing(frame, ENTERPRISE_TYPE_PLACEHOLDER, "加工单位", ENTERPRISE_TYPE_HIDDEN)
        self._select_by_typing(frame, CUSTOMS_CODE_PLACEHOLDER, enterprise_code)
        self._set_export_flag(frame)

        bond_input = frame.locator(QUERY_ANCHOR)
        bond_input.click()
        bond_input.fill(bond_invt_no)

        frame.locator(START_DATE_INPUT).fill(start_date.strftime("%Y%m%d"))
        frame.locator(END_DATE_INPUT).fill(end_date.strftime("%Y%m%d"))

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        return self._wait_query_result(bond_invt_no)

    def _select_by_typing(self, frame, placeholder: str, value: str,
                          hidden_selector: str = None) -> None:
        """点击 → 空格 → 输入 → 点选下拉中的同名项（不按回车，回车会选中下拉第一项）

        Args:
            hidden_selector: 该字段对应的隐藏编码域，传入时回读编码用于校验
        """
        locator = frame.get_by_role("textbox", name=placeholder)
        try:
            locator.click(timeout=5000)
            locator.press(" ", timeout=3000)
            self.page.wait_for_timeout(300)
            locator.press_sequentially(value, delay=50, timeout=10000)
            self.page.wait_for_timeout(500)

            if self._pick_item_by_text(frame, value):
                logger.info(f"已点选下拉项: {value} 实际={locator.input_value(timeout=3000)}")
                self._log_hidden_code(frame, hidden_selector, value)
                return

            logger.warning(
                f"下拉中未找到匹配项: {value}，当前输入值={locator.input_value(timeout=3000)}"
            )
            self._log_dropdown_items(frame)
        except PlaywrightTimeoutError as e:
            logger.warning(f"输入框填写超时，继续按清单号查询: {value} - {e}")

    @staticmethod
    def _log_dropdown_items(frame) -> None:
        """输出当前可见的下拉项文本，便于按实际选项调整匹配"""
        try:
            items = frame.evaluate(
                """() => Array.from(document.querySelectorAll('li, dd'))
                    .filter(el => el.offsetParent !== null && el.innerText && el.innerText.trim())
                    .map(el => `${el.tagName}:${el.innerText.trim().slice(0, 40)}`)
                    .slice(0, 20)"""
            )
            logger.warning(f"可见下拉项: {items}")
        except Exception as e:
            logger.warning(f"读取下拉项失败: {e}")

    @staticmethod
    def _log_hidden_code(frame, hidden_selector: str, value: str) -> None:
        """回读隐藏编码域，确认编码不是空值（避免显示正确但编码未写入）"""
        if not hidden_selector:
            return
        try:
            code = frame.locator(hidden_selector).input_value(timeout=3000)
            if code:
                logger.info(f"隐藏编码已写入: {value} -> {code}")
            else:
                logger.warning(f"隐藏编码为空，可能未真正选中: {value}")
        except Exception as e:
            logger.warning(f"读取隐藏编码失败: {e}")

    def _pick_item_by_text(self, frame, text: str, timeout: int = 3000) -> bool:
        """在已展开的下拉中点选匹配项（li/dd）：先精确匹配，再按包含匹配"""
        exact = re.compile(rf"^\s*{re.escape(text)}\s*$")
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            for tag in ("li", "dd"):
                items = frame.locator(tag)
                for index in range(items.count()):
                    item = items.nth(index)
                    try:
                        if not item.is_visible():
                            continue
                        content = item.inner_text().strip()
                        if len(content) > 60:
                            continue
                        if exact.match(content) or text in content:
                            item.click()
                            return True
                    except Exception:
                        continue
            self.page.wait_for_timeout(200)
        return False

    def _set_export_flag(self, frame) -> None:
        """进出口标志直接设为“出口”（隐藏域 E + 显示域 出口）"""
        try:
            frame.locator(IE_FLAG_HIDDEN).evaluate(JS_SET_VALUE, IE_FLAG_EXPORT_CODE)
            frame.locator(IE_FLAG_NAME).evaluate(JS_SET_VALUE, IE_FLAG_EXPORT_NAME)
            actual = frame.locator(IE_FLAG_NAME).input_value(timeout=3000)
            logger.info(f"进出口标志已设置: 期望={IE_FLAG_EXPORT_NAME} 实际={actual}")
        except Exception as e:
            logger.warning(f"设置进出口标志失败: {e}")

    def close(self) -> None:
        """关闭加工贸易账册窗口并回到门户首页，便于重新进入下一个模块"""
        app = self.page
        try:
            if app is not None and app is not self.portal_page and not app.is_closed():
                app.close()
                logger.info("已关闭加工贸易账册窗口")
        except Exception as e:
            logger.warning(f"关闭加工贸易账册窗口失败: {e}")

        self.page = self.portal_page
        try:
            if self.portal_page is not None and not self.portal_page.is_closed():
                self.portal_page.goto(self.home_url)
                self.portal_page.wait_for_timeout(500)
                logger.info("已返回单一窗口首页")
        except Exception as e:
            logger.warning(f"返回单一窗口首页失败: {e}")

    def _wait_query_result(self, bond_invt_no: str) -> bool:
        """等待结果出现；出现“未查询到相关的清单数据”时点确定并返回 False"""
        deadline = time.time() + self.page_timeout / 1000
        while time.time() < deadline:
            if self._find_text(bond_invt_no)[1] is not None:
                logger.info(f"保税核注清单查询完成: {bond_invt_no}")
                return True

            for frame in list(self.page.frames):
                if self._dismiss_no_data(frame):
                    logger.warning(f"保税核注清单未查询到数据: {bond_invt_no}")
                    return False

            self.page.wait_for_timeout(300)
        raise RuntimeError(f"保税核注清单查询超时: {bond_invt_no}")

    def _find_text(self, text: str):
        """在所有 frame 中查找可见文本，返回 (frame, locator)，找不到返回 (None, None)"""
        for frame in self.page.frames:
            try:
                item = self._visible_locator(frame.get_by_text(text))
                if item is not None:
                    return frame, item
            except Exception:
                continue
        return None, None

    @staticmethod
    def _dismiss_no_data(frame) -> bool:
        """关闭“未查询到相关的清单数据”提示"""
        message = frame.get_by_text(re.compile("未查询到"))
        try:
            if message.count() == 0 or not message.first.is_visible():
                return False
        except Exception:
            return False

        confirm = frame.locator("a.layui-layer-btn0")
        if confirm.count() == 0:
            confirm = frame.get_by_text("确定", exact=True)
        confirm.first.click()
        return True

    def print_list(self, bond_invt_no: str, save_path) -> Path:
        """点击清单号选中该行后点【打印】，保存打印出的 PDF"""
        frame, number_cell = None, None
        deadline = time.time() + self.page_timeout / 1000
        while time.time() < deadline and number_cell is None:
            frame, number_cell = self._find_text(bond_invt_no)
            if number_cell is None:
                self.page.wait_for_timeout(300)
        if number_cell is None:
            raise RuntimeError(f"结果列表中未找到清单号: {bond_invt_no}")
        number_cell.click()

        print_button = frame.get_by_role("button", name=re.compile("打印"))
        if print_button.count() == 0:
            frame = self._find_frame_with_button("打印")
            print_button = frame.get_by_role("button", name=re.compile("打印"))

        path = self.save_download(
            lambda: print_button.first.click(),
            save_path,
            pdf_preview=True,
        )
        logger.info(f"保税核注清单下载完成: {bond_invt_no} -> {path}")
        return path

    @staticmethod
    def _handle_optional_dialogs(page) -> None:
        """账册应用打开后可能出现提示/修改密码弹窗，出现则关闭"""
        for locator in (
            page.locator("#layui-layer3").get_by_text("关闭"),
            page.get_by_text("已了解，继续访问应用系统"),
        ):
            try:
                locator.first.wait_for(state="visible", timeout=2000)
            except PlaywrightTimeoutError:
                continue
            locator.first.click()

    @staticmethod
    def _visible_locator(locator):
        """返回第一个可见项，全部不可见时返回 None"""
        for index in range(locator.count()):
            item = locator.nth(index)
            try:
                if item.is_visible():
                    return item
            except Exception:
                continue
        return None

    def _wait_visible(self, locator, timeout: int = None):
        """等待出现可见项（门户菜单异步展开）"""
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            item = self._visible_locator(locator)
            if item is not None:
                return item
            self.page.wait_for_timeout(300)
        return None

    def _find_frame_with_button(self, button_name: str, timeout: int = None):
        """在所有 frame 中查找包含指定按钮的页面"""
        pattern = re.compile(button_name)
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    button = frame.get_by_role("button", name=pattern)
                    if button.count() > 0 and button.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(300)
        raise RuntimeError(f"未找到包含【{button_name}】按钮的页面")

    def _find_frame(self, timeout: int = None):
        """查找可见的保税核注清单查询 iframe（按清单号输入框定位）"""
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    anchor = frame.locator(QUERY_ANCHOR)
                    if anchor.count() > 0 and anchor.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(500)
        raise RuntimeError(f"未找到保税核注清单查询页（{QUERY_ANCHOR}）")
