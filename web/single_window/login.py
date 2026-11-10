#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口登录模块
"""
import time

from log.loggingConfig import logger
from web.base_page import BasePage
from web.browser import BrowserManager

DEFAULT_URL = "https://www.singlewindow.cn/#/"
HOME_URL = "https://www.singlewindow.cn/#/"
DEFAULT_LOGIN_TIMEOUT = 120000
LOGGED_IN_TIMEOUT = 15000

# 登录页 #info 中的过程提示与真正的错误提示
LOGIN_PROGRESS_HINTS = ("请稍等", "登录中", "正在登录", "正在处理")
LOGIN_ERROR_HINTS = ("错误", "失败", "不正确", "无效", "不存在", "锁定", "异常", "过期", "拒绝")


class SingleWindowLogin(BasePage):
    """中国国际贸易单一窗口登录

    使用卡介质（U盾）登录，只需输入卡口令，无需账号。
    卡口令由调用方从 Windows 凭据管理器读取后传入。
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
            timeout: 等待登录跳转的超时时间（毫秒），U盾识别较慢时需放宽
        """
        logger.info("开始登录单一窗口（卡介质登录）")

        self.goto(self.url)
        self.page.get_by_role("link", name="登录").click()
        self.page.get_by_role("link", name="卡介质登录").click()

        # 直接使用 locator 填写，避免 BasePage.fill 将卡口令写入日志
        password_input = self.page.get_by_role("textbox", name="卡口令")
        password_input.click()
        password_input.fill(password)

        self.page.locator("#agreementCardId > .custom-checkbox-wrapper > .custom-checkbox").click()
        self.page.locator("#checkboxIntel").check()
        self.page.get_by_role("button", name="登    录").click()

        self._wait_login_result(timeout)
        self.goto(HOME_URL)
        self._wait_logged_in()
        logger.info("单一窗口登录成功")
        return True

    def _wait_login_result(self, timeout: int) -> None:
        """等待登录结果

        卡介质登录会先跳转 CAS 登录页，认证成功后跳回 login.swdo 换取会话，
        再进入门户首页；失败时停留在登录页，并在 #info 中给出错误提示。
        """
        deadline = time.time() + timeout / 1000
        last_message = ""
        while time.time() < deadline:
            try:
                url = self.page.url
                if "cas/login" not in url and "login.swdo" not in url:
                    logger.info(f"登录页已跳转: {url}")
                    return
                error = self.page.locator("#info").inner_text(timeout=1000).strip()
            except Exception:
                # 页面跳转过程中上下文可能失效，继续等待
                error = ""

            if error and error != last_message:
                logger.info(f"登录页提示: {error}")
                last_message = error

            if (
                error
                and not any(hint in error for hint in LOGIN_PROGRESS_HINTS)
                and any(hint in error for hint in LOGIN_ERROR_HINTS)
            ):
                raise RuntimeError(f"单一窗口登录失败: {error}")

            self.page.wait_for_timeout(500)

        raise RuntimeError("单一窗口登录超时，请检查U盾是否就绪、卡口令是否正确")

    def _wait_logged_in(self) -> None:
        """在门户首页确认登录态：页头“登录”入口消失后才算登录成功"""
        try:
            self.page.locator(".header-top-inside").wait_for(
                state="visible", timeout=LOGGED_IN_TIMEOUT
            )
            # 等待页头完成登录态渲染
            self.page.wait_for_timeout(1000)
        except Exception:
            pass

        if self.page.locator(".header-btn-login").is_visible():
            raise RuntimeError("单一窗口登录未成功，首页仍显示“登录”入口")
        logger.info("首页已进入登录状态")

    def open_goods_declaration(self):
        """从首页进入口岸执法申报-货物申报，返回货物申报页面"""
        logger.info("进入单一窗口-货物申报")

        self.page.get_by_text("业务应用", exact=True).click()
        self.page.get_by_role("link", name="口岸执法申报").click()
        with self.page.expect_popup() as popup_info:
            self.page.get_by_role("link", name="货物申报").click()

        goods_page = popup_info.value
        goods_page.wait_for_load_state()
        goods_page.set_default_timeout(self.page_timeout)
        self._handle_optional_dialogs(goods_page)
        return goods_page

    def open_declaration_query(self):
        """进入综合查询-报关数据查询，返回货物申报页面"""
        goods_page = self.open_goods_declaration()
        goods_page.locator("a").filter(has_text="综合查询").first.click()
        goods_page.get_by_text("报关数据查询").click()
        logger.info("已进入综合查询-报关数据查询")
        return goods_page

    def _handle_optional_dialogs(self, page) -> None:
        """处理登录后可能出现的提示弹窗

        两类弹窗均为“可能出现”：
        - 提示弹窗：出现则点击关闭
        - 修改密码提示：出现则点击“已了解，继续访问应用系统”
        """
        # 提示弹窗：出现则点击关闭
        self._click_if_visible(page, page.locator("#layui-layer3").get_by_text("关闭"))

        # 修改密码提示：出现则点击已了解，继续访问应用系统
        self._click_if_visible(page, page.get_by_text("修改密码："))
        self._click_if_visible(page, page.get_by_text("已了解，继续访问应用系统"))

    @staticmethod
    def _click_if_visible(page, locator, timeout: int = 2000) -> bool:
        """元素可见时点击，不存在或不可见时跳过"""
        try:
            locator.wait_for(state="visible", timeout=timeout)
        except Exception:
            return False
        locator.click()
        return True
