#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：逐单下载单据（开发文档 8.6，执行顺序按业务确认）

1) 货物申报模块：逐单按统一编号查询 → 详情导出 → 放行通知单、电子委托协议、发票、装箱单
   先暂存到 06_Temp/{公司}/下载单据暂存/{统一编号}/；
2) 全部单号下载完成后关闭货物申报页面并返回单一窗口首页；
3) 统一解析各单详情导出（得到 BIP 字段、合同号/出口日期、随附单证与保税核注清单编号），
   统计需要下载保税核注清单的单据；
4) 保税核注清单模块：批量下载，申报日期直接取该单“最近操作时间段”的起止（YYYYMMDD），
   完成后关闭账册窗口并返回首页；
5) 电子口岸模块：逐单下载含出口退税联的报关单；
6) 按 4.3 命名归档到 03_Output/{公司}/下载单据/{YYYYMM}/，并回写运行结果表。

- 每类单据单独重试与校验，失败只影响该类单据；
- 必需单据缺失时该单运行状态记“失败-下载失败”，已下载文件保留；
- 结果表按报关单号 upsert，写入详情解析出的 BIP/报关单字段、文件路径与状态。
"""
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from core.task_base import TaskBase
from data_processing.declaration import ManualError, ParseError, parse_detail
from data_processing.result_table import open_result_table, year_of
from data_processing.rules import (
    DOC_BONDED,
    DOC_DECLARATION,
    DOC_ENTRUST,
    DOC_INVOICE,
    DOC_PACKING,
    DOC_RELEASE,
    DOWNLOADED,
    FAIL_DOWNLOAD,
    FAIL_PARSE,
    MANUAL,
    PATH_COLUMNS,
    need_bonded,
    required_docs,
    result_row_from_declaration,
    result_row_from_list_item,
)
from log.loggingConfig import logger
from utils.credential import get_password
from utils.naming import build_document_path
from web.browser import BrowserManager
from web.chinaport.declaration import ChinaPortDeclarationExport
from web.chinaport.login import DEFAULT_URL as CHINAPORT_DEFAULT_URL, ChinaPortLogin
from web.single_window.bonded_list import DEFAULT_HOME_URL, BondedListQuery
from web.single_window.declaration import DeclarationExport

DOC_EXT = ".pdf"


class DownloadSubtask(TaskBase):
    """逐单下载单一窗口与电子口岸单据"""

    def __init__(self, context, company: dict, mode: str = "normal"):
        company_name = str(company.get("公司名称") or "").strip()
        super().__init__(f"逐单下载单据-{company_name}", context)
        self.company = company
        self.company_name = company_name
        # 文件路径统一使用公司全称
        self.company_path_name = str(company.get("公司全称") or company_name).strip()
        self.mode = mode
        self.config = context.config
        self.output_dir = Path(self.config.get_path("output"))
        self.log_dir = Path(self.config.get_path("log"))
        self.temp_dir = Path(self.config.get_path("temp"))
        self.config_dir = self.config.get_path("config")
        self.download_dir = self.output_dir / self.company_path_name / "下载单据"
        self.detail_dir = self.temp_dir / self.company_path_name / "报关单详情"
        self.staging_dir = self.temp_dir / self.company_path_name / "下载单据暂存"
        self.retries = max(int(self.config.get("业务参数.单据下载重试次数", 2)), 1)
        policy = str(self.config.get("业务参数.同名文件处理", "overwrite") or "overwrite").strip()
        self.file_policy = "skip" if mode == "rerun" else policy
        self.bonded_keyword = str(self.config.get("业务参数.保税核注清单触发词", "进料") or "进料")
        self.tables = context.get("result_tables") or {}
        self._bonded = None

    def execute(self) -> bool:
        todo = self.context.get("todo") or []
        if not todo:
            logger.info(f"{self.company_name} 无待处理报关单，跳过下载")
            return True

        goods_page = self.context.get("single_window_page")
        exporter = self.context.get("single_window_exporter")
        if exporter is None and goods_page is not None:
            exporter = self._build_exporter(goods_page)

        # 阶段一：货物申报（详情导出 + 四类单据，全部先暂存）
        records = [self._download_goods_declaration(item, exporter) for item in todo]
        # 阶段二：关闭货物申报页面并回首页
        self._close_goods_page()

        # 阶段三～五：解析详情、归档、保税核注清单、电子口岸（任一阶段异常不影响后续与结果落表）
        for phase in (self._parse_details, self._archive_docs,
                      self._download_bonded_batch, self._download_chinaport):
            try:
                phase(records)
            except Exception as e:
                logger.exception(f"{self.company_name} {phase.__name__} 阶段异常: {e}")

        # 阶段六：回写运行结果表
        for record in records:
            try:
                self._finalize(record)
            except Exception as e:
                logger.error(f"[{record['item'].no}] 结果表写入失败: {e}")

        ok = sum(1 for record in records if record["status"] == DOWNLOADED)
        logger.info(f"{self.company_name} 下载完成: 成功 {ok}/{len(records)} 条")
        self.context.set("download_records", records)
        return True

    # ------------------------------------------------------------------ 阶段一：货物申报
    def _download_goods_declaration(self, item, exporter) -> dict:
        started = datetime.now()
        record = {
            "item": item,
            "decl": None,
            "parse_error": "",
            "detail_path": None,
            "temp_docs": {},
            "docs": {},
            "errors": [],
            "status": DOWNLOADED,
            "bonded_required": False,
            "started": started,
        }
        if exporter is None:
            record["status"] = FAIL_PARSE
            record["errors"].append("单一窗口导出器未初始化")
            return record

        bill_no = item.no
        logger.info(f"[{bill_no}] 开始下载（统一编号={item.unified_no}）")
        self.detail_dir.mkdir(parents=True, exist_ok=True)
        detail_target = self.detail_dir / (
            f"{item.unified_no or bill_no}-报关单详情-{time.strftime('%Y%m%d %H%M%S')}.xlsx"
        )
        try:
            # 按结果表记录的“最近操作时间段”重设时间条件，避免早前分段的单查不到
            window = getattr(item, "update_window", None)
            if window:
                exporter.query_by_no(item.unified_no, window[0], window[1])
            else:
                exporter.query_by_no(item.unified_no)
            record["detail_path"] = self._call(exporter.export_detail, item.unified_no, detail_target)
        except Exception as e:
            logger.error(f"[{bill_no}] 报关单详情导出失败: {e}")
            record["status"] = FAIL_PARSE
            record["errors"].append(f"报关单详情导出失败: {e}")
            return record

        staging = self.staging_dir / (item.unified_no or bill_no)
        self._download_print_docs(record, exporter, staging)
        self._download_invoice_packing(record, exporter, staging)

        logger.info(
            f"[{bill_no}] 货物申报单据暂存: "
            + (", ".join(sorted(record["temp_docs"])) if record["temp_docs"] else "无")
        )
        return record

    def _download_print_docs(self, record, exporter, save_dir) -> None:
        """放行通知单、电子委托协议（打印预览方式下载，先存暂存目录）"""
        for doc_type in (DOC_RELEASE, DOC_ENTRUST):
            try:
                saved = self._call(
                    exporter.download_documents,
                    record["item"].unified_no,
                    save_dir,
                    None,
                    (doc_type,),
                )
                path = saved.get(doc_type) or Path(save_dir) / f"{record['item'].unified_no}-{doc_type}.pdf"
                if self._valid(path):
                    record["temp_docs"][doc_type] = Path(path)
                else:
                    record["errors"].append(f"{doc_type}下载失败")
            except Exception as e:
                logger.error(f"[{record['item'].no}] {doc_type}下载失败: {e}")
                record["errors"].append(f"{doc_type}下载失败: {e}")

    def _download_invoice_packing(self, record, exporter, save_dir) -> None:
        """发票、装箱单（随附单据页下载，两类同页，按组重试）"""
        doc_types = (DOC_INVOICE, DOC_PACKING)
        try:
            saved = self._call(
                exporter.download_invoice_packing,
                record["item"].unified_no,
                save_dir,
            )
        except Exception as e:
            logger.error(f"[{record['item'].no}] 发票/装箱单下载失败: {e}")
            record["errors"].append(f"发票/装箱单下载失败: {e}")
            saved = {}

        for doc_type in doc_types:
            path = saved.get(doc_type) or Path(save_dir) / f"{record['item'].unified_no}-{doc_type}.pdf"
            if self._valid(path):
                record["temp_docs"][doc_type] = Path(path)

    def _close_goods_page(self) -> None:
        """关闭货物申报页面并回单一窗口首页，便于进入保税核注清单等下一个模块"""
        page = self.context.get("single_window_page")
        if page is not None:
            try:
                if not page.is_closed():
                    page.close()
                    logger.info("已关闭货物申报页面")
            except Exception as e:
                logger.warning(f"关闭货物申报页面失败: {e}")
            self.context.set("single_window_page", None)

        portal = self.context.get_browser("single_window")
        if portal is not None:
            try:
                if not portal.page.is_closed():
                    portal.page.goto(DEFAULT_HOME_URL)
                    portal.page.wait_for_timeout(500)
                    logger.info("已返回单一窗口首页")
            except Exception as e:
                logger.warning(f"返回单一窗口首页失败: {e}")

    # ------------------------------------------------------------------ 阶段三：解析详情
    def _parse_details(self, records) -> None:
        for record in records:
            if record["status"] == FAIL_PARSE or not record["detail_path"]:
                continue
            item = record["item"]
            try:
                declaration = parse_detail(
                    record["detail_path"],
                    supervision_mode=item.supervision_mode,
                    export_date_fallback=item.export_date,
                    unified_no=item.unified_no,
                )
            except (ParseError, ManualError) as e:
                logger.error(f"[{item.no}] 报关单详情解析失败: {e}")
                record["parse_error"] = str(e)
                record["errors"].append(f"报关单详情解析失败: {e}")
                record["status"] = FAIL_PARSE
                continue

            record["decl"] = declaration
            record["bonded_required"] = need_bonded(item.supervision_mode, self.bonded_keyword)
            if len(declaration.bonded_nos) > 1:
                record["status"] = MANUAL
                record["errors"].append(f"保税核注清单编号不唯一: {','.join(declaration.bonded_nos)}")
            elif record["bonded_required"] and not declaration.bonded_no:
                record["errors"].append("监管方式含“进料”但随附单证中无保税核注清单编号")
            logger.info(f"[{item.no}] 详情解析完成: {declaration.summary()}")

    # ------------------------------------------------------------------ 阶段四：归档命名
    def _archive_docs(self, records) -> None:
        for record in records:
            declaration = record["decl"]
            if declaration is None:
                continue
            for doc_type, temp_path in list(record["temp_docs"].items()):
                if not self._valid(temp_path):
                    continue
                target = self._target(declaration, doc_type)
                if target is None:
                    record["docs"][doc_type] = self._existing_path(declaration, doc_type)
                    Path(temp_path).unlink(missing_ok=True)
                    logger.info(f"[{declaration.no}] {doc_type} 已存在，按 skip 策略复用")
                    continue
                try:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    Path(temp_path).replace(target)
                    record["docs"][doc_type] = target
                    logger.info(f"[{declaration.no}] {doc_type} 已归档: {target.name}")
                except Exception as e:
                    logger.error(f"[{declaration.no}] {doc_type} 归档失败: {e}")
                    record["errors"].append(f"{doc_type}归档失败: {e}")

    # ------------------------------------------------------------------ 阶段五：保税核注清单
    def _download_bonded_batch(self, records) -> None:
        """统一下载所有需要保税核注清单的单据（一个模块内完成，结束后关闭窗口）"""
        pending = [
            record for record in records
            if record["decl"] is not None
            and record["bonded_required"]
            and len(record["decl"].bonded_nos) == 1
        ]
        if not pending:
            return
        logger.info(f"{self.company_name} 需下载保税核注清单 {len(pending)} 条")

        try:
            bonded = self._bonded_query()
        except Exception as e:
            logger.error(f"打开保税核注清单查询失败: {e}")
            for record in pending:
                record["errors"].append(f"保税核注清单未下载: {e}")
            return

        try:
            for record in pending:
                self._download_bonded_one(record, bonded)
        finally:
            self._close_bonded()

    def _download_bonded_one(self, record, bonded) -> None:
        declaration = record["decl"]
        target = self._target(declaration, DOC_BONDED)
        if target is None:
            record["docs"][DOC_BONDED] = self._existing_path(declaration, DOC_BONDED)
            logger.info(f"[{declaration.no}] 保税核注清单已存在，按 skip 策略复用")
            return

        # 申报日期直接取该单“最近操作时间”起止，格式 YYYYMMDD（bonded.query 内格式化）
        window = getattr(record["item"], "update_window", None)
        if window:
            start, end = window
        else:
            end = date.today()
            start = end - timedelta(days=int(self.config.get("业务参数.查询最近操作天数", 62)) - 1)

        try:
            logger.info(
                f"[{declaration.no}] 保税核注清单查询: 清单号={declaration.bonded_no} "
                f"申报日期 {start:%Y%m%d}~{end:%Y%m%d}"
            )
            enterprise_code = str(self.company.get("海关编码") or "").strip()
            if enterprise_code:
                found = self._call(bonded.query, declaration.bonded_no, start, end, enterprise_code)
            else:
                found = self._call(bonded.query, declaration.bonded_no, start, end)
            if found:
                path = self._call(bonded.print_list, declaration.bonded_no, target)
                if self._valid(path):
                    record["docs"][DOC_BONDED] = path
                else:
                    record["errors"].append(f"保税核注清单下载失败: {declaration.bonded_no}")
            else:
                record["errors"].append(f"保税核注清单未查询到: {declaration.bonded_no}")
        except Exception as e:
            logger.error(f"[{declaration.no}] 保税核注清单下载失败: {e}")
            record["errors"].append(f"保税核注清单下载失败: {e}")

    def _bonded_query(self):
        """打开保税核注清单查询窗口（整个模块只打开一次，批量下载后关闭）"""
        if self._bonded is not None and not self._bonded.page.is_closed():
            return self._bonded

        portal = self.context.get_browser("single_window")
        page = portal.page if portal is not None else None
        if page is None or page.is_closed():
            raise RuntimeError("未找到可用的单一窗口门户页面，无法打开保税核注清单查询")
        bonded = BondedListQuery(page, page_timeout=self._page_timeout(),
                                 download_timeout=self._download_timeout())
        bonded.open()
        self._bonded = bonded
        return bonded

    def _close_bonded(self) -> None:
        """关闭加工贸易账册窗口并返回首页"""
        if self._bonded is None:
            return
        try:
            self._bonded.close()
        except Exception as e:
            logger.warning(f"关闭保税核注清单窗口失败: {e}")
        finally:
            self._bonded = None

    def _close_single_window_browser(self) -> None:
        """关闭单一窗口浏览器（含其 Playwright 实例），为电子口岸浏览器腾出线程"""
        self.context.set("single_window_page", None)
        portal = self.context.browsers.pop("single_window", None)
        if portal is None:
            return
        try:
            portal.close()
            logger.info("已关闭单一窗口浏览器")
        except Exception as e:
            logger.warning(f"关闭单一窗口浏览器失败: {e}")

    # ------------------------------------------------------------------ 电子口岸
    def _download_chinaport(self, records) -> None:
        """单一窗口全部完成后，登录电子口岸逐单下载报关单（含出口退税联）"""
        pending = [r for r in records if r["decl"] is not None]
        if not pending:
            return

        credential_target = str(self.company.get("凭据目标名") or "").strip()
        if not credential_target:
            for record in pending:
                record["errors"].append("公司未配置凭据目标名，电子口岸报关单未下载")
            return

        # 电子口岸需另起浏览器：Playwright 同步 API 不允许同一线程内嵌套启动第二个实例，
        # 单一窗口单据与保税核注清单已全部处理完，先关闭单一窗口浏览器再登录电子口岸
        self._close_single_window_browser()
        browser = BrowserManager(headless=self._headless(), channel=self.config.get("系统配置.浏览器通道", "msedge"))
        try:
            browser.start()
            password = get_password(credential_target)
            url = str(self.config.get("系统配置.电子口岸地址") or CHINAPORT_DEFAULT_URL).strip()
            login = ChinaPortLogin(browser, url=url, page_timeout=self._page_timeout())
            login.login(password, timeout=int(self.config.get("系统配置.登录超时(秒)", 120)) * 1000)
            app = login.open_export_tax_refund()
            exporter = ChinaPortDeclarationExport(
                app, page_timeout=self._page_timeout(), download_timeout=self._download_timeout()
            )

            for record in pending:
                declaration = record["decl"]
                target = self._target(declaration, DOC_DECLARATION)
                if target is None:
                    record["docs"][DOC_DECLARATION] = self._existing_path(declaration, DOC_DECLARATION)
                    logger.info(f"[{declaration.no}] 报关单已存在，按 skip 策略复用")
                    continue
                try:
                    exporter.query(declaration.no)
                    path = self._call(exporter.download_declaration, declaration.no, target)
                    if self._valid(path):
                        record["docs"][DOC_DECLARATION] = path
                    else:
                        record["errors"].append("报关单下载失败: 文件为空")
                except Exception as e:
                    logger.error(f"[{declaration.no}] 电子口岸报关单下载失败: {e}")
                    record["errors"].append(f"报关单下载失败: {e}")
        except Exception as e:
            logger.error(f"{self.company_name} 电子口岸登录/下载失败: {e}")
            for record in pending:
                record["errors"].append(f"电子口岸报关单未下载: {e}")
        finally:
            browser.close()

    # ------------------------------------------------------------------ 阶段六：结果落表
    def _finalize(self, record) -> None:
        item = record["item"]
        declaration = record["decl"]

        if declaration is not None:
            required = required_docs(record["bonded_required"])
            missing = [doc_type for doc_type in required if not self._valid(record["docs"].get(doc_type))]
            if missing and record["status"] == DOWNLOADED:
                record["status"] = FAIL_DOWNLOAD
                record["errors"].append(f"缺失单据: {','.join(missing)}")
            # 详情里有值的字段以详情为准，详情无值的用列表导出兜底，两边都没有则留空
            values = result_row_from_list_item(item)
            values.update({
                key: value
                for key, value in result_row_from_declaration(declaration).items()
                if value not in ("", None)
            })
        else:
            if record["parse_error"]:
                record["status"] = FAIL_PARSE
            values = {}

        # 已下载但未归档的单据（解析失败时）也记录路径，避免重复下载
        for doc_type, path in record["docs"].items():
            column = PATH_COLUMNS.get(doc_type)
            if column:
                values[column] = self._absolute(path)

        values.update({
            "运行状态": record["status"],
            "异常原因": "；".join(record["errors"]),
            "开始时间": record["started"].strftime("%Y-%m-%d %H:%M:%S"),
            "结束时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

        export_date = declaration.export_date if declaration else item.export_date
        table = self._table(export_date)
        action = table.upsert(item.no, values)
        table.save()
        logger.info(
            f"[{item.no}] 结果表{action}: 状态={record['status']}"
            + (f"，异常={values['异常原因']}" if values["异常原因"] else "")
        )

    def _table(self, export_date):
        year = year_of(export_date)
        if year not in self.tables:
            self.tables[year] = open_result_table(self.log_dir, self.company_path_name, year, self.config_dir)
        return self.tables[year]

    # ------------------------------------------------------------------ 公共
    def _build_exporter(self, goods_page) -> DeclarationExport:
        return DeclarationExport(
            goods_page, page_timeout=self._page_timeout(), download_timeout=self._download_timeout()
        )

    def _call(self, func, *args, **kwargs):
        """按“单据下载重试次数”重试，全部失败后抛出最后一次异常"""
        last_error = None
        for attempt in range(1, self.retries + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                if attempt < self.retries:
                    logger.warning(f"{getattr(func, '__name__', func)} 第 {attempt} 次失败: {e}")
                    time.sleep(1.5)
        raise last_error

    def _target(self, doc_ctx, doc_type):
        """生成目标路径；按文件策略判断是否需要跳过（返回 None 表示复用已有文件）"""
        path = build_document_path(
            self.download_dir, doc_ctx.no, doc_ctx.contract_no, doc_ctx.export_date, doc_type, DOC_EXT
        )
        if self.file_policy == "skip" and self._valid(path):
            return None
        if self.file_policy == "rename" and path.exists():
            index = 1
            while True:
                candidate = path.with_name(f"{path.stem}({index}){path.suffix}")
                if not candidate.exists():
                    return candidate
                index += 1
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def _existing_path(self, doc_ctx, doc_type) -> Path:
        return build_document_path(
            self.download_dir, doc_ctx.no, doc_ctx.contract_no, doc_ctx.export_date, doc_type, DOC_EXT
        )

    @staticmethod
    def _valid(path) -> bool:
        return bool(path) and Path(path).exists() and Path(path).stat().st_size > 0

    @staticmethod
    def _absolute(path) -> str:
        """结果表记录文件绝对路径"""
        return str(Path(path).resolve())

    def _page_timeout(self) -> int:
        return int(self.config.get("系统配置.页面超时(秒)", 30)) * 1000

    def _download_timeout(self) -> int:
        return int(self.config.get("系统配置.下载超时(秒)", 60)) * 1000

    def _headless(self) -> bool:
        value = self.config.get("系统配置.无头模式", False)
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in ("true", "1", "yes", "是")
