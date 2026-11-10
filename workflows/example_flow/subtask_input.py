#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：录入数据到系统
"""
from core.task_base import TaskBase
from web.base_page import BasePage


class InputSubtask(TaskBase):
    """子任务：录入数据到系统A

    职责：
    - 从上下文获取已读取的数据
    - 遍历数据逐条录入到系统A
    """

    def __init__(self, context):
        super().__init__("录入数据", context)

    def execute(self) -> bool:
        # 从上下文获取浏览器和数据
        browser = self.context.get_browser("system_a")
        data = self.context.get("input_data", [])

        if not browser or not browser.page:
            return False

        page = BasePage(browser.page)

        # 导航到录入页面（根据实际系统修改）
        page.click("#data-input-menu")
        page.wait_for_selector("#input-form")

        # 遍历数据录入
        success_count = 0
        for row in data:
            try:
                # 录入单条数据（根据实际系统字段修改）
                page.fill("#field1", str(row.get("字段1", "")))
                page.fill("#field2", str(row.get("字段2", "")))
                page.click("#submit-btn")
                page.wait_for_selector("#success-msg", timeout=5000)
                success_count += 1
            except Exception as e:
                # 单条失败继续下一条
                continue

        return success_count > 0