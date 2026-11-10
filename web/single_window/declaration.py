#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口-综合查询-报关数据结果列表：全选、列表导出、按统一编号查询、详情导出、单据打印下载

录制脚本 recordings/singlewwindow_dowload.py 中两处定位在此修正：
1. 全选：录屏命中不到表头复选框，实际应勾选 input[name="btSelectAll"]；
2. 统一编号查询框 #cusCiqNo1 与其它查询条件同因（jQuery 校验/联动），
   直接 fill 不生效，改为赋值后触发 input/change。

详情导出与打印预览都会先弹窗再下载，下载事件可能挂在弹窗页上，
因此统一按“主页面 + 弹窗”双监听处理。
"""
import re
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from web.base_page import BasePage

# 统一编号查询框：既用于填值，也作为“当前可见货物申报 frame”的判断依据
QUERY_ANCHOR = "#cusCiqNo1"
# 报关单结果表：同页还有报关状态、运输状态等表格，都带 btSelectAll/btSelectItem
RESULT_TABLE = "#declareTable"
SELECT_ALL_BOX = f'{RESULT_TABLE} input[name="btSelectAll"]'
ROW_CHECK_BOX = f'{RESULT_TABLE} input[name="btSelectItem"]'
LIST_EXPORT_BUTTON = "列表导出"
DETAIL_EXPORT_BUTTON = "详情导出"
PREVIEW_TEXT = "打印预览"

# 需要下载的三类单据：(单据名称, 打印按钮名称)
DOCUMENT_BUTTONS = (
    ("报关单", "打印报关单"),
    ("放行通知单", "打印通知书"),
    ("电子委托协议", "打印电子委托协议"),
)

JS_SET_VALUE = """
(el, value) => {
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
}
"""


class DeclarationExport(BasePage):
    """报关数据结果列表操作

    列表来源于上一步“报关数据查询”的结果；全选只作用于当前结果页，
    因此调用方需保证查询条件已提交（单次查询不超过 7 天）。
    页面可能存在多个同名 iframe，按“可见的统一编号查询框”定位实际使用的 frame。
    """

    def __init__(self, page, page_timeout: int = 30000, download_timeout: int = 60000):
        super().__init__(page)
        self.page_timeout = page_timeout
        self.download_timeout = download_timeout

    def select_all(self) -> int:
        """勾选表头全选复选框，返回勾选到的数据行数"""
        frame = self._find_frame()
        try:
            # 查询是异步渲染，先等第一行出现再全选
            frame.locator(ROW_CHECK_BOX).first.wait_for(timeout=self.page_timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError("结果列表为空，无法全选导出，请确认查询条件与时间段") from None

        # 带固定表头的表格会在 .fixed-table-header 里渲染一份可见的表头，
        # #declareTable 内的那份被其遮挡，优先点可见的那份
        select_all = frame.locator(
            f'.bootstrap-table:has({RESULT_TABLE}) .fixed-table-header input[name="btSelectAll"]'
        )
        if select_all.count() == 0:
            select_all = frame.locator(SELECT_ALL_BOX)
        self._check_box(select_all)

        count = frame.locator(f"{ROW_CHECK_BOX}:checked").count()
        if count == 0:
            raise RuntimeError("全选未生效：结果列表没有勾选到任何数据")
        logger.info(f"已全选报关数据: {count} 条")
        return count

    def export_list(self, save_path) -> Path:
        """全选后点击【列表导出】，保存全量列表"""
        self.select_all()
        frame = self._find_frame()
        path = Path(save_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with self.page.expect_download(timeout=self.download_timeout) as download_info:
            frame.get_by_role("button", name=re.compile(LIST_EXPORT_BUTTON)).click()

        download = download_info.value
        failure = download.failure()
        if failure:
            raise RuntimeError(f"列表导出下载失败: {failure}")

        # 导出文件的实际格式以浏览器建议文件名为准（可能是 .xls/.zip），
        # 强行改成 .xlsx 会导致文件打不开
        suggested = Path(download.suggested_filename)
        if suggested.suffix and suggested.suffix.lower() != path.suffix.lower():
            path = path.with_suffix(suggested.suffix)
            logger.info(f"按实际导出格式保存: {suggested.name} -> {path.name}")

        download.save_as(path)
        path = self._normalize_export(path)
        logger.info(f"列表导出完成: {path}")
        return path

    @staticmethod
    def _normalize_export(path: Path) -> Path:
        """校验并整理导出文件：压缩包解开其中的 Excel，非 Excel 直接报错"""
        size = path.stat().st_size
        with path.open("rb") as file:
            head = file.read(4)

        if head == b"PK\x03\x04":
            import zipfile

            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                if any(name.startswith("xl/") for name in names):
                    logger.info(f"导出文件校验通过: {path}（{size} 字节）")
                    return path
                excels = [n for n in names if n.lower().endswith((".xlsx", ".xls"))]
                data = archive.read(excels[0]) if len(excels) == 1 else None

            if data:
                target = path.with_suffix(Path(excels[0]).suffix)
                target.write_bytes(data)
                logger.info(f"导出文件为压缩包，已取出 {excels[0]}: {target}（{size} 字节）")
                return target
            raise RuntimeError(f"导出压缩包内未找到唯一的 Excel 文件: {path}")

        if head == b"\xd0\xcf\x11\xe0":
            logger.info(f"导出文件为 xls 格式: {path}（{size} 字节）")
            return path

        # 单一窗口的列表导出实际是 HTML 表格伪装成 .xls（文件名“单一窗口海关数据列表导出模板.xls”），
        # Excel 可正常打开，不能按二进制 Excel 校验
        if head.startswith(b"<"):
            with path.open("rb") as file:
                snippet = file.read(65536).lower()
            if b"<table" in snippet or b"<tr" in snippet:
                logger.info(f"导出文件为 xls(HTML 表格) 格式: {path}（{size} 字节）")
                return path
            raise RuntimeError(f"导出文件为 HTML 页面而非数据表: {path}，请检查列表导出是否成功")

        raise RuntimeError(
            f"导出文件不是 Excel: {path}（{size} 字节，头部 {head!r}），请检查列表导出是否成功"
        )

    def query_by_no(self, cus_ciq_no: str) -> None:
        """按统一编号查询，等待该编号出现在结果列表"""
        frame = self._find_frame()
        field = frame.locator(QUERY_ANCHOR)
        field.evaluate(JS_SET_VALUE, cus_ciq_no)
        if field.input_value() != cus_ciq_no:
            raise RuntimeError(f"统一编号未填写成功: {cus_ciq_no}")

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        try:
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).first.wait_for(timeout=self.page_timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"统一编号查询无结果: {cus_ciq_no}") from None
        logger.info(f"统一编号查询完成: {cus_ciq_no}")

    def export_detail(self, cus_ciq_no: str, save_path) -> Path:
        """勾选该条报关数据后点击【详情导出】，保存详情文件"""
        frame = self._find_frame()
        self._check_box(
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).get_by_role("checkbox")
        )
        path = self._save_download(
            lambda: frame.get_by_role("button", name=re.compile(DETAIL_EXPORT_BUTTON)).click(),
            save_path,
        )
        logger.info(f"详情导出完成: {cus_ciq_no} -> {path}")
        return path

    def download_documents(self, cus_ciq_no: str, save_dir) -> dict:
        """依次下载报关单、放行通知单、电子委托协议

        需先执行 query_by_no 查询到该条数据；详情导出后行勾选一般还在，
        这里重新勾选一次，避免列表刷新导致选择丢失。
        """
        frame = self._find_frame()
        self._check_box(
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).get_by_role("checkbox")
        )

        out_dir = Path(save_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        saved = {}
        for doc_name, button_name in DOCUMENT_BUTTONS:
            frame = self._find_frame()
            frame.get_by_role("button", name=re.compile(button_name)).click()
            path = self._save_download(
                lambda: frame.get_by_text(PREVIEW_TEXT).click(),
                out_dir / f"{cus_ciq_no}-{doc_name}.pdf",
                pdf_preview=True,
            )
            saved[doc_name] = path
            logger.info(f"单据下载完成: {doc_name} -> {path}")
        return saved

    @staticmethod
    def _check_box(locator) -> None:
        """勾选表格复选框

        bootstrap-table 的固定表头与页面侧边导航会拦截真实点击，
        改为直接派发 click，再校验勾选状态，避免静默失败。
        """
        if not locator.is_checked():
            locator.evaluate("el => el.click()")
        if not locator.is_checked():
            raise RuntimeError(f"复选框勾选未生效: {locator}")

    def _save_download(self, click, save_path, pdf_preview: bool = False) -> Path:
        """执行 click 并保存结果文件

        详情导出直接产生下载；打印预览会在弹窗中打开浏览器 PDF 查看器，
        其工具栏位于 shadow DOM，点不到“保存”，改为按弹窗地址直接取 PDF。
        下载事件可能挂在主页面或弹窗页上，因此两处都监听。
        """
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

            download = self._wait_download(downloads)
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

    def _find_frame(self, timeout: int = None):
        """在所有 frame 中查找可见的报关数据查询结果页"""
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    anchor = frame.locator(QUERY_ANCHOR)
                    if anchor.count() > 0 and anchor.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(500)

        candidates = []
        for frame in self.page.frames:
            try:
                count = frame.locator(QUERY_ANCHOR).count()
                if count:
                    candidates.append(f"{frame.name}|{frame.url}|count={count}")
            except Exception:
                continue
        raise RuntimeError(f"未找到可见的报关数据查询页（{QUERY_ANCHOR}），候选: {candidates or '无'}")

    def _wait_download(self, downloads: list):
        """等待任一页面产生的下载事件"""
        deadline = time.time() + self.download_timeout / 1000
        while time.time() < deadline:
            if downloads:
                return downloads[0]
            self.page.wait_for_timeout(200)
        raise RuntimeError(f"等待下载超时（{self.download_timeout / 1000:.0f} 秒）")
