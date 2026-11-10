#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
系统A下载报表模块
"""
from typing import Optional
from log.loggingConfig import logger
from web.browser import BrowserManager
from web.base_page import BasePage


class DownloadReport(BasePage):
    """系统A下载报表模块

    尧承BasePage，封装系统A特定的报表下载逻辑
    """

    def __init__(self, browser: BrowserManager):
        super().__init__(browser.page)
        self.browser = browser

    def navigate_to_report_page(self) -> bool:
        """导航到报表页面"""
        try:
            self.click("#reports-menu")
            self.wait_for_selector("#report-list")
            logger.info("导航到报表页面成功")
            return True
        except Exception as e:
            logger.error(f"导航到报表页面失败: {e}")
            return False

    def select_report_type(self, report_type: str) -> bool:
        """选择报表类型"""
        try:
            self.select_option("#report-type", report_type)
            logger.info(f"选择报表类型: {report_type}")
            return True
        except Exception as e:
            logger.error(f"选择报表类型失败: {e}")
            return False

    def set_date_range(self, start_date: str, end_date: str) -> bool:
        """设置日期范围"""
        try:
            self.fill("#start-date", start_date)
            self.fill("#end-date", end_date)
            logger.info(f"设置日期范围: {start_date} ~ {end_date}")
            return True
        except Exception as e:
            logger.error(f"设置日期范围失败: {e}")
            return False

    def download_report(self, save_path: Optional[str] = None) -> bool:
        """下载报表

        Args:
            save_path: 保存路径（可选，使用浏览器默认下载路径）

        Returns:
            bool: 下载是否成功
        """
        try:
            # 开始下载（等待下载事件）
            with self.page.expect_download() as download_info:
                self.click("#download-btn")

            download = download_info.value

            # 如果指定保存路径，则保存到指定位置
            if save_path:
                download.save_as(save_path)
                logger.info(f"报表已保存到: {save_path}")
            else:
                logger.info(f"报表已下载: {download.path()}")

            return True

        except Exception as e:
            logger.error(f"下载报表失败: {e}")
            return False