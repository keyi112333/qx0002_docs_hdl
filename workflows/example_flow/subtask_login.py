#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：登录系统A
"""
from core.task_base import TaskBase
from web.browser import BrowserManager
from web.system_a.login import SystemALogin


class LoginSubtask(TaskBase):
    """子任务：登录系统A

    职责：
    - 启动浏览器
    - 访问系统A登录页
    - 输入用户名密码登录
    - 将浏览器存入上下文供后续子任务使用
    """

    def __init__(self, context):
        super().__init__("登录系统A", context)

    def execute(self) -> bool:
        # 从上下文获取配置
        config = self.context.config
        url = config.get("系统A配置.网址", "https://example.com/login")
        username = config.get("系统A配置.用户名", "admin")
        password = config.get("系统A配置.密码", "123456")
        headless = config.get("系统A配置.无头模式", False)

        # 创建浏览器
        browser = BrowserManager(headless=headless)
        browser.start()

        # 存入上下文供后续子任务使用
        self.context.set_browser("system_a", browser)

        # 执行登录
        login_module = SystemALogin(browser)
        return login_module.login(url, username, password)