#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：导出报关单列表 → 拉取 BIP 附件数据 → 生成待处理清单 → 写入运行结果表

对应开发文档 8.4（分段查询与合并导出）与 8.5（待处理清单）。
比对口径按业务确认：本次列表导出 ∩ BIP 无附件记录（CZMC 为空）－ 已成功/需人工 ＋ 补跑清单。
"""
import time
from datetime import date
from pathlib import Path

from core.task_base import TaskBase
from data_processing.bip_attachments import fetch_attachment_export, load_attachment_export
from data_processing.declaration_list import ListFormatError, merge_exports, read_list_file
from data_processing.result_table import COL_UPDATE_WINDOW, open_result_table, year_of
from data_processing.rules import (
    PENDING,
    build_todo,
    build_windows,
    result_row_from_list_item,
)
from log.loggingConfig import logger
from web.single_window.declaration import DeclarationExport
from web.single_window.query import DeclarationQuery

EMPTY_RESULT_HINT = "结果列表为空"


class DeclarationSubtask(TaskBase):
    """导出并生成待处理清单"""

    def __init__(self, context, company: dict, mode: str = "normal", bip_file=None):
        company_name = str(company.get("公司名称") or "").strip()
        super().__init__(f"报关单列表与待处理清单-{company_name}", context)
        self.company = company
        self.company_name = company_name
        # 文件路径统一使用公司全称
        self.company_path_name = str(company.get("公司全称") or company_name).strip()
        self.mode = mode
        # 指定 BIP 附件数据文件时不再直连 Oracle（供测试/重跑比对使用）
        self.bip_file = Path(bip_file) if bip_file else None
        self.config = context.config
        self.output_dir = self.config.get_path("output")
        self.log_dir = self.config.get_path("log")
        self.temp_dir = self.config.get_path("temp")
        self.config_dir = self.config.get_path("config")
        self.result_tables = {}
        # 报关单号 -> 命中的“最近操作时间”分段，逐单查询时重设时间条件
        self.bill_windows = {}

    def execute(self) -> bool:
        goods_page = self.context.get("single_window_page")
        if goods_page is None:
            raise RuntimeError("未找到单一窗口货物申报页面，请先执行登录子任务")

        list_file = self._export_list(goods_page)
        if list_file is None:
            logger.warning(f"{self.company_name} 最近查询窗口内没有报关单")
            self.context.set("todo", [])
            return False

        bip_file = self._fetch_bip_attachments()
        bip_data = load_attachment_export(bip_file)

        list_items = read_list_file(list_file)
        result_status = self._load_result_status()
        todo, stats = build_todo(list_items, bip_data, result_status)

        for item in todo:
            item.update_window = self.bill_windows.get(item.no)
            table = self._result_table(item.export_date)
            action = table.upsert(item.no, {
                **result_row_from_list_item(item),
                COL_UPDATE_WINDOW: item.window_text(),
                "运行状态": PENDING,
                "异常原因": "",
                "开始时间": None,
                "结束时间": None,
            })
            logger.info(f"待处理写入结果表: {item.no}（{action}，最近操作时间段={item.window_text() or '未知'}）")
        for table in self.result_tables.values():
            table.save()

        self.context.set("todo", todo)
        self.context.set("result_tables", self.result_tables)
        self.context.set("candidate_stats", stats)
        self.context.set("bill_windows", self.bill_windows)
        self.context.set("declaration_list_file", list_file)
        self.context.set("bip_attachment_file", bip_file)
        logger.info(f"{self.company_name} 待处理清单生成完成: {len(todo)} 条")
        return bool(todo)

    # ------------------------------------------------------------------ 列表导出
    def _export_list(self, goods_page):
        """按 7 天一段逐段查询并导出，合并为 03_Output/{公司}/报关单列表 下的文件"""
        total_days = int(self.config.get("业务参数.查询最近操作天数", 62))
        step = int(self.config.get("业务参数.查询分段天数", 7))
        retries = max(int(self.config.get("业务参数.单据下载重试次数", 2)), 1)
        page_timeout = int(self.config.get("系统配置.页面超时(秒)", 30)) * 1000
        download_timeout = int(self.config.get("系统配置.下载超时(秒)", 60)) * 1000
        # “查询→全选→列表导出”各步骤之间、以及分段之间的停顿，避免页面未刷新完连续操作卡住
        step_interval = int(self.config.get("业务参数.分段查询间隔(毫秒)", 3000) or 3000)

        exporter = DeclarationExport(
            goods_page, page_timeout=page_timeout, download_timeout=download_timeout,
            step_interval=step_interval,
        )
        query = DeclarationQuery(goods_page)
        windows = build_windows(date.today(), total_days, step)
        logger.info(f"{self.company_name} 报关单列表查询窗口: {len(windows)} 段，共 {total_days} 天，"
                    f"操作间隔 {step_interval / 1000:.0f} 秒")
        self.context.set("single_window_exporter", exporter)

        segment_dir = Path(self.temp_dir) / self.company_path_name / "报关单列表"
        segment_dir.mkdir(parents=True, exist_ok=True)

        exported, missing = [], []
        for start, end in windows:
            target = segment_dir / (
                f"{self.company_path_name}-报关单列表-{start:%Y%m%d}_{end:%Y%m%d}-"
                f"{time.strftime('%Y%m%d %H%M%S')}.xlsx"
            )
            try:
                path = self._export_segment(query, exporter, start, end, target, retries)
            except Exception as e:
                logger.error(f"分段导出失败 {start}~{end}: {e}")
                missing.append(f"{start:%Y%m%d}_{end:%Y%m%d}")
                continue
            if path:
                exported.append(path)
                self._record_windows(path, start, end)

        if missing:
            logger.warning(f"以下分段导出失败，结果可能缺失: {', '.join(missing)}")
        self.context.set("missing_segments", missing)
        if not exported:
            return None

        stamp = time.strftime("%Y%m%d %H%M%S")
        merged_path = (
            Path(self.output_dir) / self.company_path_name / "报关单列表"
            / f"{self.company_path_name}-报关单列表-合并-{stamp}.xlsx"
        )
        try:
            return merge_exports(exported, merged_path)
        except ListFormatError as e:
            logger.error(f"报关单列表合并失败: {e}")
            return exported[0]

    @staticmethod
    def _export_segment(query, exporter, start, end, target, retries):
        """单段查询+导出，失败按配置次数重试；无数据返回 None"""
        last_error = None
        for attempt in range(1, retries + 1):
            try:
                query.search(start, end)
                # 等待结果列表刷新稳定后再全选，避免把上一段的结果导出
                exporter.wait_result_settled()
                return exporter.export_list(target)
            except RuntimeError as e:
                if EMPTY_RESULT_HINT in str(e):
                    logger.info(f"{start}~{end} 无报关单数据")
                    return None
                last_error = e
                logger.warning(f"分段导出第 {attempt} 次失败: {e}")
            except Exception as e:
                last_error = e
                logger.warning(f"分段导出第 {attempt} 次异常: {e}")
        raise last_error or RuntimeError("分段导出失败")

    def _record_windows(self, segment_file, start, end) -> None:
        """记录该段导出的报关单号 -> 最近操作时间段，供逐单查询时重设时间条件"""
        try:
            for item in read_list_file(segment_file):
                self.bill_windows[item.no] = (start, end)
        except Exception as e:
            logger.warning(f"记录分段报关单时间窗失败（不影响导出）: {segment_file} - {e}")

    # ------------------------------------------------------------------ BIP 附件数据
    def _fetch_bip_attachments(self) -> Path:
        if self.bip_file is not None:
            if not self.bip_file.exists():
                raise FileNotFoundError(f"指定的 BIP 已上传附件文件不存在: {self.bip_file}")
            logger.info(f"按指定文件读取 BIP 已上传附件数据（不查询数据库）: {self.bip_file}")
            return self.bip_file

        stamp = time.strftime("%Y%m%d %H%M%S")
        save_path = (
            Path(self.output_dir) / self.company_path_name / "BIP已上传附件表"
            / f"{self.company_path_name}-BIP已上传附件表-{stamp}.xlsx"
        )
        return fetch_attachment_export(self.config, save_path)

    # ------------------------------------------------------------------ 结果表
    def _load_result_status(self) -> dict:
        """读取今年与去年结果表的运行状态（8.5：覆盖两个年度，按单号匹配）"""
        current = date.today().year
        status = {}
        for year in (current, current - 1):
            table = self._result_table(year)
            for row in table.rows():
                bill_no = str(row.get("报关单号") or "").strip()
                if bill_no:
                    status[bill_no] = str(row.get("运行状态") or "").strip()
        logger.info(f"{self.company_name} 结果表已有单号 {len(status)} 个")
        return status

    def _result_table(self, export_date):
        year = year_of(export_date)
        if year not in self.result_tables:
            self.result_tables[year] = open_result_table(
                self.log_dir, self.company_path_name, year, self.config_dir
            )
        return self.result_tables[year]
