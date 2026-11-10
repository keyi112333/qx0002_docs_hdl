#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
页面操作基类
"""
import time
from pathlib import Path
from typing import Optional

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger


class BasePage:
    """页面操作基类

    提供通用的页面操作方法：点击、输入、等待、截图、下载保存等
    """

    def __init__(self, page):
        self.page = page

    def goto(self, url: str) -> None:
        """访问页面"""
        logger.info(f"访问页面: {url}")
        self.page.goto(url)

    def click(self, selector: str, timeout: int = 10000) -> None:
        """点击元素"""
        logger.info(f"点击元素: {selector}")
        self.page.click(selector, timeout=timeout)

    def fill(self, selector: str, value: str, timeout: int = 10000) -> None:
        """输入文本"""
        logger.info(f"输入文本: {selector} -> {value}")
        self.page.fill(selector, value, timeout=timeout)

    def type(self, selector: str, value: str, delay: int = 50, timeout: int = 10000) -> None:
        """逐字符输入（模拟真实输入）"""
        logger.info(f"逐字符输入: {selector}")
        self.page.type(selector, value, delay=delay, timeout=timeout)

    def wait_for_selector(self, selector: str, timeout: int = 30000) -> None:
        """等待元素出现"""
        logger.info(f"等待元素: {selector}")
        self.page.wait_for_selector(selector, timeout=timeout)

    def wait_for_url(self, url: str, timeout: int = 30000) -> None:
        """等待URL变化"""
        logger.info(f"等待URL: {url}")
        self.page.wait_for_url(url, timeout=timeout)

    def get_text(self, selector: str) -> str:
        """获取元素文本"""
        return self.page.locator(selector).text_content()

    def get_value(self, selector: str) -> str:
        """获取输入框值"""
        return self.page.input_value(selector)

    def screenshot(self, path: str) -> None:
        """截图"""
        logger.info(f"截图保存: {path}")
        self.page.screenshot(path=path)

    def is_visible(self, selector: str) -> bool:
        """判断元素是否可见"""
        return self.page.locator(selector).is_visible()

    def select_option(self, selector: str, value: str) -> None:
        """选择下拉框选项"""
        logger.info(f"选择选项: {selector} -> {value}")
        self.page.select_option(selector, value)

    def check(self, selector: str) -> None:
        """勾选复选框"""
        logger.info(f"勾选: {selector}")
        self.page.check(selector)

    def uncheck(self, selector: str) -> None:
        """取消勾选"""
        logger.info(f"取消勾选: {selector}")
        self.page.uncheck(selector)

    def wait(self, timeout: int = 1000) -> None:
        """等待指定时间"""
        self.page.wait_for_timeout(timeout)

    @staticmethod
    def check_box(locator) -> None:
        """勾选表格复选框

        bootstrap-table 的固定表头与页面侧边导航会拦截真实点击，
        改为直接派发 click，再校验勾选状态，避免静默失败。
        """
        if not locator.is_checked():
            locator.evaluate("el => el.click()")
        if not locator.is_checked():
            raise RuntimeError(f"复选框勾选未生效: {locator}")

    def save_download(self, click, save_path, pdf_preview: bool = False, timeout: int = None) -> Path:
        """执行 click 并保存结果文件

        详情导出会直接产生下载；打印/打印预览会在弹窗中打开浏览器 PDF 查看器，
        其工具栏位于扩展的 shadow DOM 中无法点击，改为按弹窗地址直接取 PDF。
        下载事件可能挂在主页面或弹窗页上，因此两处都监听。
        """
        timeout = timeout or getattr(self, "download_timeout", 60000)
        path = Path(save_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        downloads = []
        popups = []

        # Playwright 会给事件处理器设置属性，不能直接传 list.append；
        # 弹窗一创建就挂下载监听，避免弹窗内立即触发的下载漏接
        def collect_page_download(download) -> None:
            downloads.append(download)

        def collect_popup(popup) -> None:
            popups.append(popup)

            def collect_popup_download(download) -> None:
                downloads.append(download)

            popup.on("download", collect_popup_download)
            logger.info(f"已打开弹窗: {popup.url}")

        self.page.on("download", collect_page_download)
        self.page.on("popup", collect_popup)
        try:
            click()

            # 等待弹窗/下载事件登记；直接下载时无需等待弹窗
            deadline = time.time() + 10
            while time.time() < deadline and not popups and not downloads:
                self.page.wait_for_timeout(100)

            if pdf_preview:
                for popup in popups:
                    if self._fetch_popup_pdf(popup, path):
                        return path

            download = self._wait_download(downloads, timeout)
            suggested = Path(download.suggested_filename)
            if suggested.suffix and suggested.suffix.lower() != path.suffix.lower():
                path = path.with_suffix(suggested.suffix)
                logger.info(f"按实际文件格式保存: {suggested.name} -> {path.name}")
            download.save_as(path)
            return path
        finally:
            self.page.remove_listener("download", collect_page_download)
            self.page.remove_listener("popup", collect_popup)
            for popup in popups:
                try:
                    popup.close()
                except Exception:
                    pass

    def _fetch_popup_pdf(self, popup, path: Path) -> bool:
        """按预览弹窗地址取回 PDF

        浏览器 PDF 查看器的工具栏在扩展的 shadow DOM 中无法点击，
        这里用与浏览器共享 Cookie 的 APIRequestContext 直接请求该地址。
        """
        deadline = time.time() + 15
        while time.time() < deadline and not popup.url:
            self.page.wait_for_timeout(200)

        url = popup.url
        if not url:
            logger.warning("预览弹窗地址为空，改为等待下载事件")
            return False

        try:
            response = popup.context.request.get(url, headers={"Referer": self.page.url})
            body = response.body() if response.ok else b""
        except Exception as e:
            raise RuntimeError(f"按预览地址获取 PDF 失败: {e}") from None

        if not body.startswith(b"%PDF"):
            raise RuntimeError(f"预览地址未返回 PDF: HTTP {response.status} {url}")

        path.write_bytes(body)
        logger.info(f"已按预览地址保存 PDF: {url}")
        return True

    def _wait_download(self, downloads: list, timeout: int):
        """等待任一页面产生的下载事件"""
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            if downloads:
                return downloads[0]
            self.page.wait_for_timeout(200)
        raise RuntimeError(f"等待下载超时（{timeout / 1000:.0f} 秒）")