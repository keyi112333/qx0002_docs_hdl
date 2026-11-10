#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：下载报表
"""
from core.task_base import TaskBase
from web.system_a.download_report import DownloadReport


class DownloadSubtask(TaskBase):
    """子任务：下载报表

    职责：
    - 从系统A下载报表
    - 保存到输出目录
    """

    def __init__(self, context):
        super().__init__("下载报表", context)

    def execute(self) -> bool:
        # 从上下文获取浏览器和配置
        browser = self.context.get_browser("system_a")
        config = self.context.config

        if not browser or not browser.page:
            return False

        # 获取下载保存路径
        output_path = str(config.get_path('output', 'report.xlsx'))

        # 执行下载
        download_module = DownloadReport(browser)

        # 导航到报表页面
        if not download_module.navigate_to_report_page():
            return False

        # 下载报表
        return download_module.download_report(output_path)