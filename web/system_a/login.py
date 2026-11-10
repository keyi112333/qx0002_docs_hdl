#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
系统A登录模块
"""
from typing import Optional
from log.loggingConfig import logger
from web.browser import BrowserManager
from web.base_page import BasePage


class SystemALogin(BasePage):
    """系统A登录模块

    继承BasePage，封装系统A特定的登录逻辑
    """

    def __init__(self, browser: BrowserManager):
        super().__init__(browser.page)
        self.browser = browser

    def login(self, url: str, username: str, password: str) -> bool:
        """执行登录

        Args:
            url: 登录页面URL
            username: 用户名
            password: 密码

        Returns:
            bool: 登录是否成功
        """
        logger.info(f"系统A登录: {username}")

        try:
            # 访问登录页
            self.goto(url)

            # 输入用户名和密码（根据实际系统修改选择器）
            self.fill("#username", username)
            self.fill("#password", password)

            # 点击登录按钮
            self.click("#login-btn")

            # 等待登录成功标志（根据实际系统修改）
            self.wait_for_selector("#dashboard", timeout=30000)

            logger.info("系统A登录成功")
            return True

        except Exception as e:
            logger.error(f"系统A登录失败: {e}")
            return False

    def is_logged_in(self) -> bool:
        """检查是否已登录"""
        try:
            return self.is_visible("#dashboard")
        except:
            return False

    def logout(self) -> bool:
        """退出登录"""
        try:
            self.click("#logout-btn")
            self.wait_for_selector("#login-btn")
            logger.info("系统A退出登录成功")
            return True
        except Exception as e:
            logger.error(f"系统A退出登录失败: {e}")
            return False