#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口-加工贸易账册-综合查询-保税核注清单查询与打印下载

入口：门户-业务应用-口岸执法申报-加工贸易账册（新窗口）-综合查询-保税核注清单查询。
企业类型/海关编码按录制只输入空格，进出口标志填“出口”，实际以清单号和申报日期为准；
查不到数据时页面会弹“未查询到相关的清单数据”，需点确定关闭。
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
IE_FLAG_INPUT = 'input[name="impExpMarkCdName"]'
START_DATE_INPUT = "#inputDateStart"
END_DATE_INPUT = "#inputDateEnd"


class BondedListQuery(BasePage):
    """保税核注清单查询与打印"""

    def __init__(self, portal_page, page_timeout: int = 30000, download_timeout: int = 60000):
        super().__init__(portal_page)
        self.page_timeout = page_timeout
        self.download_timeout = download_timeout

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

        企业类型/海关编码需先点输入框、再输入空格唤出下拉，然后点下拉里的文本选项；
        返回 True 表示查到数据，False 表示页面提示“未查询到相关的清单数据”。
        """
        frame = self._find_frame()

        enterprise_type = frame.get_by_role("textbox", name=ENTERPRISE_TYPE_PLACEHOLDER)
        enterprise_type.click()
        enterprise_type.press(" ")
        self._pick_option("加工单位")

        customs_code = frame.get_by_role("textbox", name=CUSTOMS_CODE_PLACEHOLDER)
        customs_code.click()
        customs_code.press(" ")
        self._pick_option(enterprise_code)

        ie_flag = frame.locator(IE_FLAG_INPUT)
        ie_flag.click()
        ie_flag.fill("出口")

        bond_input = frame.locator(QUERY_ANCHOR)
        bond_input.click()
        bond_input.fill(bond_invt_no)

        frame.locator(START_DATE_INPUT).fill(start_date.strftime("%Y%m%d"))
        frame.locator(END_DATE_INPUT).fill(end_date.strftime("%Y%m%d"))

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        return self._wait_query_result(bond_invt_no)

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

    def _pick_option(self, text: str, timeout: int = 5000) -> bool:
        """点输入框后按空格键才会弹出自动完成下拉，点其中匹配文本才算选中"""
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    options = frame.locator(".autocompleter-list li, .autocompleter li")
                    for index in range(options.count()):
                        option = options.nth(index)
                        if option.is_visible() and text in option.inner_text():
                            option.click()
                            logger.info(f"已选择下拉项: {option.inner_text().strip()}")
                            return True
                except Exception:
                    continue
            self.page.wait_for_timeout(200)
        logger.warning(f"下拉中未找到匹配项: {text}")
        return False

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
