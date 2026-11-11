#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国电子口岸登录模块

录制脚本 recordings/chinaport_login_export 的登录部分：
电子口岸首页 → 登录 → 卡介质登录 → 卡口令 → 勾选协议 → 登录 → 首页
→ 出口退税联网稽查（新窗口）→ 报关单查询下载

电子口岸与单一窗口共用同一张 U 盾，卡口令仍从 Windows 凭据管理器读取。
"""
import re
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from web.base_page import BasePage
from web.browser import BrowserManager

DEFAULT_URL = "https://e.chinaport.gov.cn/"
DEFAULT_LOGIN_TIMEOUT = 120000

CARD_PASSWORD_INPUT = "卡口令"
AGREEMENT_CHECKBOX = "#agreementCardId > .custom-checkbox-wrapper > .custom-checkbox"
INTEL_CHECKBOX = "#checkboxIntel"
# 提交按钮文字中间带空格（“登 录”），用空白匹配避开页头的“登录”入口
SUBMIT_LOGIN_NAME = re.compile(r"登\s+录")
EXPORT_TAX_REFUND_LINK = re.compile("出口退税联网稽查")
DECLARATION_QUERY_ENTRY = "报关单查询下载"


class ChinaPortLogin(BasePage):
    """中国电子口岸登录

    使用卡介质（U盾）登录，只需输入卡口令。登录后从首页进入
    “出口退税联网稽查”应用，该应用在新窗口打开。
    """

    def __init__(self, browser: BrowserManager, url: str = DEFAULT_URL, page_timeout: int = 30000):
        super().__init__(browser.page)
        self.browser = browser
        self.url = url or DEFAULT_URL
        self.page_timeout = page_timeout
        self.page.set_default_timeout(page_timeout)

    def login(self, password: str, timeout: int = DEFAULT_LOGIN_TIMEOUT) -> bool:
        """卡介质登录并等待登录完成

        Args:
            password: 卡口令（从 Windows 凭据管理器读取，不落日志）
            timeout: 等待登录完成的超时时间（毫秒），U盾识别较慢时需放宽
        """
        logger.info(f"开始登录中国电子口岸（卡介质登录）: {self.url}")

        self.goto(self.url)
        self.page.get_by_role("link", name="登录").click()
        self.page.get_by_role("link", name="卡介质登录").click()

        # 直接使用 locator 填写，避免 BasePage.fill 将卡口令写入日志
        password_input = self.page.get_by_role("textbox", name=CARD_PASSWORD_INPUT)
        password_input.click()
        password_input.fill(password)

        self.page.locator(AGREEMENT_CHECKBOX).click()
        self.page.locator(INTEL_CHECKBOX).check()

        submit = self.page.get_by_role("link", name=SUBMIT_LOGIN_NAME)
        if submit.count() == 0:
            submit = self.page.get_by_role("button", name=SUBMIT_LOGIN_NAME)
        submit.first.click()

        self._wait_login_result(timeout)
        self.goto(self.url)
        logger.info("中国电子口岸登录成功")
        return True

    def _wait_login_result(self, timeout: int) -> None:
        """等待登录完成：登录表单消失后重新进入首页"""
        try:
            self.page.get_by_role("textbox", name=CARD_PASSWORD_INPUT).wait_for(
                state="detached", timeout=timeout
            )
        except PlaywrightTimeoutError:
            raise RuntimeError("中国电子口岸登录超时，请检查U盾是否就绪、卡口令是否正确") from None

    def open_export_tax_refund(self):
        """从首页进入出口退税联网稽查，返回应用新窗口的页面对象"""
        logger.info("进入中国电子口岸-出口退税联网稽查")

        link = self._wait_visible(self.page.get_by_role("link", name=EXPORT_TAX_REFUND_LINK))
        if link is None:
            raise RuntimeError("电子口岸首页未找到【出口退税联网稽查】入口，请确认登录是否成功")

        with self.page.expect_popup() as popup_info:
            link.click()

        app = popup_info.value
        app.wait_for_load_state()
        app.set_default_timeout(self.page_timeout)
        self._handle_optional_dialogs(app)
        app.get_by_text(DECLARATION_QUERY_ENTRY).first.click()
        logger.info("已进入报关单查询下载")
        return app

    @staticmethod
    def _handle_optional_dialogs(page) -> None:
        """处理进入应用后可能出现的提示/修改密码弹窗，出现则关闭"""
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
        """等待出现可见项（首页菜单异步渲染）"""
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            item = self._visible_locator(locator)
            if item is not None:
                return item
            self.page.wait_for_timeout(300)
        return None
