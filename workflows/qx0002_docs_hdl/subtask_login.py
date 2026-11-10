#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：登录单一窗口
"""
from datetime import date, timedelta

from core.task_base import TaskBase
from log.loggingConfig import logger
from utils.credential import get_password
from web.browser import BrowserManager
from web.single_window.login import SingleWindowLogin
from web.single_window.query import DeclarationQuery


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


class SingleWindowLoginSubtask(TaskBase):
    """子任务：登录单一窗口并进入报关数据查询

    公司信息来自 config.xlsx 的“公司盾配置”，卡口令按该公司的
    “凭据目标名”从 Windows 凭据管理器读取。
    """

    def __init__(self, context, company: dict):
        company_name = str(company.get("公司名称") or "").strip()
        super().__init__(f"登录单一窗口-{company_name}", context)
        self.company = company

    def execute(self) -> bool:
        config = self.context.config
        company_name = str(self.company.get("公司名称") or "").strip()
        credential_target = str(self.company.get("凭据目标名") or "").strip()
        if not company_name or not credential_target:
            raise ValueError(f"公司盾配置不完整（公司名称/凭据目标名）: {self.company}")

        logger.info(f"读取卡口令: {company_name} -> Windows凭据 {credential_target}")
        password = get_password(credential_target)

        headless = _as_bool(config.get("系统配置.无头模式", False))
        channel = config.get("系统配置.浏览器通道", "msedge")
        url = config.get("系统配置.单一窗口地址", "https://www.singlewindow.cn/#/")
        page_timeout = int(config.get("系统配置.页面超时(秒)", 30)) * 1000
        query_step = int(config.get("业务参数.查询分段天数", 7))

        browser = BrowserManager(headless=headless, channel=channel)
        browser.start()
        self.context.set_browser("single_window", browser)
        self.context.set("current_company", self.company)

        login = SingleWindowLogin(browser, url=url, page_timeout=page_timeout)
        login.login(password)
        goods_page = login.open_declaration_query()

        end_date = date.today()
        start_date = end_date - timedelta(days=query_step - 1)
        DeclarationQuery(goods_page).search(start_date, end_date)
        return True
