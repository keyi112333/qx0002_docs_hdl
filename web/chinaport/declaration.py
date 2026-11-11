#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国电子口岸-出口退税联网稽查-报关单查询下载

录制脚本 recordings/chinaport_login_export 的报关单导出部分：
报关单号查询 → 勾选结果行 → 批量预览/打印 → 保存 PDF。

电子口岸导出的报关单才是正确版本，按开发文档 4.3 归档到
03_Output/{公司}/下载单据/YYYYMM/{报关单号}-{合同编号}-{出口日期}-报关单.pdf（路径由调用方生成）；
单一窗口打印的报关单仅供提取信息，保存到 06_Temp
（见 DeclarationExport.download_documents 的 declaration_path 参数）。
"""
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from web.base_page import BasePage

# 报关单号输入框，也作为“报关单查询下载页面”的定位依据
ENTRY_ID_INPUT = "#entryId"
ROW_CHECK_BOX = 'input[name="btSelectItem"]'
# 按钮名用普通字符串做子串匹配；含“/”的正则会被 Playwright 拼成非法 CSS 选择器
QUERY_BUTTON_NAME = "查询"
BATCH_PRINT_BUTTON_NAME = "批量预览/打印"


class ChinaPortDeclarationExport(BasePage):
    """报关单查询下载页面操作

    查询页位于 iframe01 中，页面可能存在多个同名 iframe，
    按“可见的报关单号输入框”定位实际使用的 frame。
    """

    def __init__(self, app_page, page_timeout: int = 30000, download_timeout: int = 60000):
        super().__init__(app_page)
        self.page_timeout = page_timeout
        self.download_timeout = download_timeout

    def query(self, entry_id: str) -> None:
        """按报关单号查询并等待结果列表出现该单号"""
        frame = self._find_frame()
        field = frame.locator(ENTRY_ID_INPUT)
        field.click()
        field.fill(entry_id)
        if field.input_value() != entry_id:
            raise RuntimeError(f"报关单号未填写成功: {entry_id}")

        frame.get_by_role("button", name=QUERY_BUTTON_NAME).first.click()
        try:
            frame.get_by_text(entry_id).first.wait_for(timeout=self.page_timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"报关单查询无结果: {entry_id}") from None
        logger.info(f"报关单查询完成: {entry_id}")

    def download_declaration(self, entry_id: str, save_path) -> Path:
        """勾选该条报关单后点击【批量预览/打印】，保存打印出的文件

        打印会打开预览窗口，下载事件可能挂在主页面或预览页上，
        因此对整个浏览器上下文监听下载；没有下载事件时再按预览地址取 PDF。
        """
        frame = self._find_frame()
        self._select_entry(frame, entry_id)

        path = Path(save_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        saved = self._batch_print(frame, path)
        logger.info(f"报关单下载完成: {entry_id} -> {saved}")
        return saved

    def _select_entry(self, frame, entry_id: str) -> None:
        """勾选查询结果中的报关单

        查询已按报关单号过滤，取第一条结果的复选框即可（与录制脚本一致）；
        不能按 get_by_role("row") 定位，页面里存在同样包含单号但没有复选框的行。
        """
        checkbox = frame.locator(ROW_CHECK_BOX).first
        try:
            checkbox.wait_for(timeout=self.page_timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"报关单查询无结果: {entry_id}") from None
        self.check_box(checkbox)
        logger.info(f"已勾选报关单: {entry_id}")

    def _batch_print(self, frame, path: Path) -> Path:
        """点击【批量预览/打印】并捕获下载/预览文件"""
        context = self.page.context
        downloads, opened = [], []

        def on_download(download) -> None:
            downloads.append(download)

        def on_page(page) -> None:
            opened.append(page)
            page.on("download", on_download)

        context.on("page", on_page)
        for page in context.pages:
            page.on("download", on_download)

        try:
            button = frame.get_by_role("button", name=BATCH_PRINT_BUTTON_NAME)
            if button.count() == 0:
                raise RuntimeError("报关单查询下载页未找到【批量预览/打印】按钮")
            button.first.click()

            deadline = time.time() + self.download_timeout / 1000
            while time.time() < deadline and not downloads:
                self.page.wait_for_timeout(200)

            if downloads:
                download = downloads[0]
                failure = download.failure()
                if failure:
                    raise RuntimeError(f"报关单下载失败: {failure}")

                suggested = Path(download.suggested_filename)
                if suggested.suffix and suggested.suffix.lower() != path.suffix.lower():
                    path = path.with_suffix(suggested.suffix)
                    logger.info(f"按实际文件格式保存: {suggested.name} -> {path.name}")

                download.save_as(path)
                return path

            # 没有下载事件时，预览窗口地址可能就是 PDF 本身（与单一窗口打印一致）
            fetched = self._fetch_pdf(opened or [p for p in context.pages if p is not self.page], path)
            if fetched:
                return fetched
            raise RuntimeError("批量预览/打印未产生下载，且预览窗口中没有可取的 PDF")
        finally:
            context.remove_listener("page", on_page)
            for page in context.pages:
                try:
                    page.remove_listener("download", on_download)
                except Exception:
                    pass
            for page in opened:
                try:
                    page.close()
                except Exception:
                    pass

    def _fetch_pdf(self, pages, path: Path):
        """按预览窗口地址取回 PDF（与浏览器共享 Cookie）"""
        for page in reversed(pages):
            try:
                url = page.url
            except Exception:
                continue
            if not url or url.startswith(("about:", "blob:", "edge:", "chrome:", "devtools:")):
                continue

            try:
                response = page.context.request.get(url, headers={"Referer": self.page.url})
                body = response.body() if response.ok else b""
            except Exception as e:
                logger.warning(f"按预览地址获取 PDF 失败: {url} - {e}")
                continue

            if body.startswith(b"%PDF"):
                path.write_bytes(body)
                logger.info(f"已按预览地址保存 PDF: {url}")
                return path
            logger.warning(f"预览地址不是 PDF: HTTP {response.status} {url}")
        return None

    def _find_frame(self, timeout: int = None):
        """在所有 frame 中查找可见的报关单查询下载页"""
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    anchor = frame.locator(ENTRY_ID_INPUT)
                    if anchor.count() > 0 and anchor.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(500)

        candidates = []
        for frame in self.page.frames:
            try:
                count = frame.locator(ENTRY_ID_INPUT).count()
                if count:
                    candidates.append(f"{frame.name}|{frame.url}|count={count}")
            except Exception:
                continue
        raise RuntimeError(f"未找到可见的报关单查询下载页（{ENTRY_ID_INPUT}），候选: {candidates or '无'}")
