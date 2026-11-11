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
ATTACHMENT_BUTTON = "随附单据"
ATTACHMENT_DOCS = ("发票", "装箱单")

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

    def __init__(self, page, page_timeout: int = 30000, download_timeout: int = 60000,
                 step_interval: int = 0):
        super().__init__(page)
        self.page_timeout = page_timeout
        self.download_timeout = download_timeout
        # 查询→全选→列表导出之间的停顿（毫秒），避免页面未刷新完就连续操作导致卡住
        self.step_interval = max(int(step_interval), 0)

    def _pause(self, factor: int = 1) -> None:
        if self.step_interval:
            self.page.wait_for_timeout(self.step_interval * factor)

    def wait_result_settled(self, timeout: int = None, interval: int = 300) -> int:
        """等待查询结果刷新完成：加载遮罩消失，且行数连续两次不变

        分段查询时上一段的结果仍在页面上，直接全选会把旧数据导出，
        因此每次点【查询】后先等列表稳定，再进入全选/导出。
        """
        frame = self._find_frame()
        # 点击查询后先给页面一个短暂的渲染窗口，再开始判断
        self.page.wait_for_timeout(600)

        deadline = time.time() + (timeout or self.page_timeout) / 1000
        last_count, stable = None, 0
        while time.time() < deadline:
            try:
                loading = frame.locator(".fixed-table-loading")
                if loading.count() > 0 and loading.first.is_visible():
                    last_count, stable = None, 0
                    self.page.wait_for_timeout(interval)
                    continue
            except Exception:
                pass

            count = frame.locator(ROW_CHECK_BOX).count()
            if count > 0 and count == last_count:
                stable += 1
                if stable >= 2:
                    logger.info(f"查询结果已刷新: {count} 条")
                    return count
            else:
                stable = 0
            last_count = count
            self.page.wait_for_timeout(interval)

        logger.warning(f"等待查询结果刷新超时，当前行数: {last_count or 0}")
        return last_count or 0

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
        self.check_box(select_all)

        count = frame.locator(f"{ROW_CHECK_BOX}:checked").count()
        if count == 0:
            raise RuntimeError("全选未生效：结果列表没有勾选到任何数据")
        logger.info(f"已全选报关数据: {count} 条")
        return count

    def export_list(self, save_path) -> Path:
        """全选后点击【列表导出】，保存全量列表（导出完成后停顿，供下一段重新填写时间）"""
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
        self._pause()
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

    def query_by_no(self, cus_ciq_no: str, start_date=None, end_date=None) -> None:
        """按统一编号查询，等待该编号出现在结果列表

        Args:
            cus_ciq_no: 统一编号
            start_date/end_date: 该单所属“最近操作时间”分段；不传则沿用页面上的时间条件，
                传入可避免列表查询后时间条件过期导致查不到早前分段的报关单
        """
        frame = self._find_frame()
        if start_date is not None and end_date is not None:
            frame.locator("#updateTime1").evaluate(JS_SET_VALUE, start_date.strftime("%Y-%m-%d"))
            frame.locator("#updateTimeEnd1").evaluate(JS_SET_VALUE, end_date.strftime("%Y-%m-%d"))
            logger.info(f"按单查询前重设最近操作时间: {start_date} ~ {end_date}")
            self._pause()

        field = frame.locator(QUERY_ANCHOR)
        field.evaluate(JS_SET_VALUE, cus_ciq_no)
        if field.input_value() != cus_ciq_no:
            raise RuntimeError(f"统一编号未填写成功: {cus_ciq_no}")

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        try:
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).first.wait_for(timeout=self.page_timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"统一编号查询无结果: {cus_ciq_no}") from None
        self._pause()
        logger.info(f"统一编号查询完成: {cus_ciq_no}")

    def export_detail(self, cus_ciq_no: str, save_path) -> Path:
        """勾选该条报关数据后点击【详情导出】，保存详情文件"""
        frame = self._find_frame()
        self.check_box(
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).get_by_role("checkbox")
        )
        path = self.save_download(
            lambda: frame.get_by_role("button", name=re.compile(DETAIL_EXPORT_BUTTON)).click(),
            save_path,
        )
        logger.info(f"详情导出完成: {cus_ciq_no} -> {path}")
        return path

    def download_documents(self, cus_ciq_no: str, save_dir, declaration_path=None,
                           doc_types=None, resolve_path=None) -> dict:
        """依次下载报关单、放行通知单、电子委托协议

        需先执行 query_by_no 查询到该条数据；详情导出后行勾选一般还在，
        这里重新勾选一次，避免列表刷新导致选择丢失。

        单一窗口打印的报关单不是最终正确版本（以中国电子口岸导出为准），
        传入 declaration_path 时报关单存到该临时路径，其余单据仍存 save_dir。

        Args:
            doc_types: 只下载指定单据（如 (“放行通知单”, “电子委托协议”)），None 表示全部
            resolve_path: 回调 (单据名称, 默认路径) -> 目标路径，用于按 4.3 规则命名
        """
        frame = self._find_frame()
        self.check_box(
            frame.locator(RESULT_TABLE).get_by_role("row", name=cus_ciq_no).get_by_role("checkbox")
        )

        out_dir = Path(save_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        saved = {}
        for doc_name, button_name in DOCUMENT_BUTTONS:
            if doc_types and doc_name not in doc_types:
                continue
            target = out_dir / f"{cus_ciq_no}-{doc_name}.pdf"
            if doc_name == "报关单" and declaration_path:
                target = Path(declaration_path)
            if resolve_path:
                target = Path(resolve_path(doc_name, target))
            frame = self._find_frame()
            frame.get_by_role("button", name=re.compile(button_name)).click()
            path = self.save_download(
                lambda: frame.get_by_text(PREVIEW_TEXT).click(),
                target,
                pdf_preview=True,
            )
            saved[doc_name] = path
            logger.info(f"单据下载完成: {doc_name} -> {path}")
        return saved

    def download_invoice_packing(self, cus_ciq_no: str, save_dir, resolve_path=None) -> dict:
        """点击结果列表里的统一编号进入单据页，从【随附单据】下载发票与装箱单

        录制脚本里点到的是“发票/文件名”单元格，实际要行内的【下载】才会出文件；
        下载完成后关闭单据页并返回报关数据查询，便于继续处理下一条。
        传入 resolve_path 时按 4.3 规则命名目标文件。
        """
        list_frame = self._find_frame()
        list_frame.get_by_role("link", name=cus_ciq_no).click()

        detail = self._find_frame_with_button(ATTACHMENT_BUTTON)
        self._confirm_dialog(detail)
        detail.get_by_role("button", name=re.compile(ATTACHMENT_BUTTON)).click()

        out_dir = Path(save_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        saved = {}
        for doc_name in ATTACHMENT_DOCS:
            row = detail.get_by_role("row", name=re.compile(doc_name)).first
            try:
                row.wait_for(timeout=self.page_timeout)
            except PlaywrightTimeoutError:
                raise RuntimeError(f"随附单据中未找到 {doc_name}: {cus_ciq_no}") from None

            path = self.save_download(
                lambda: self._click_row_download(row, detail),
                (Path(resolve_path(doc_name, out_dir / f"{cus_ciq_no}-{doc_name}.pdf"))
                 if resolve_path else out_dir / f"{cus_ciq_no}-{doc_name}.pdf"),
            )
            saved[doc_name] = path
            logger.info(f"附件下载完成: {doc_name} -> {path}")

        self._close_declaration(detail)
        self.back_to_query()
        return saved

    def back_to_query(self) -> None:
        """返回报关数据查询（综合查询菜单可能是收起状态，需要先展开）"""
        target = self.page.get_by_title("报关数据查询")
        if target.count() == 0 or not target.first.is_visible():
            self.page.locator("a").filter(has_text="综合查询").first.click()
        self.page.get_by_title("报关数据查询").click()
        logger.info("已返回报关数据查询")

    def _close_declaration(self, frame) -> None:
        """关闭单据页；上传/保存/关闭三个按钮相邻，需按名称精确点【关闭】"""
        for locator in (
            frame.get_by_role("button", name="关闭", exact=True),
            frame.get_by_role("link", name="关闭", exact=True),
            frame.get_by_text("关闭", exact=True),
        ):
            if locator.count() > 0:
                locator.first.click()
                logger.info("已关闭单据页")
                return
        raise RuntimeError("单据页未找到【关闭】按钮")

    def _click_row_download(self, row, detail) -> None:
        """点击行内【下载】；行内没有下载入口时先选中行再点工具栏下载"""
        for locator in (
            row.get_by_text("下载", exact=True),
            row.get_by_role("link", name=re.compile("下载")),
            row.get_by_role("button", name=re.compile("下载")),
        ):
            if locator.count() > 0:
                locator.first.click()
                return

        row.click()
        detail.get_by_role("button", name=re.compile("下载")).first.click()

    def _confirm_dialog(self, frame, timeout: int = 2000) -> bool:
        """点击统一编号后可能弹出信息框，点【确定】关闭；未出现则跳过"""
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            for scope in (frame, self.page):
                try:
                    confirm = scope.locator("a.layui-layer-btn0")
                    if confirm.count() > 0 and confirm.first.is_visible():
                        confirm.first.click()
                        logger.info("已点击信息弹窗的【确定】")
                        return True
                except Exception:
                    continue
            self.page.wait_for_timeout(200)
        return False

    def _find_frame_with_button(self, button_name: str, timeout: int = None):
        """在所有 frame 中查找包含指定按钮的页面（单据页 iframe 名称不固定）"""
        pattern = re.compile(button_name)
        deadline = time.time() + (timeout or self.page_timeout) / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    button = frame.get_by_role("button", name=pattern)
                    if button.count() > 0 and button.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(500)
        raise RuntimeError(f"未找到包含【{button_name}】按钮的单据页")

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
