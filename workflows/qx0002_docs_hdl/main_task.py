#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QX0002 报关单和放行通知单等单据处理流程 - 主任务（开发文档 8.1）

按公司串行：登录单一窗口 → 导出报关单列表并与 BIP 附件数据比对生成待处理清单
（写入运行结果表）→ 逐单下载单一窗口单据与电子口岸报关单 → 回写运行结果表。
单个单号失败不影响后续单号；只有登录/配置等全局问题才终止当前公司。
"""
from core.context import TaskContext
from core.task_base import TaskBase
from data_processing.rules import DOWNLOADED, FAIL_DOWNLOAD, FAIL_PARSE, MANUAL
from log.loggingConfig import logger
from workflows.qx0002_docs_hdl.subtask_declaration import DeclarationSubtask
from workflows.qx0002_docs_hdl.subtask_download import DownloadSubtask
from workflows.qx0002_docs_hdl.subtask_login import SingleWindowLoginSubtask


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


class Qx0002MainTask(TaskBase):
    """主任务：按“公司盾配置”中启用的公司逐家执行"""

    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__(config.project_name, context)
        self.mode = str(config.get("基础配置.运行模式", "normal") or "normal").strip().lower()

    def execute(self) -> bool:
        companies = [
            company
            for company in self.context.config.read_table("公司盾配置")
            if _as_bool(company.get("启用"), False)
        ]
        if not companies:
            raise ValueError("公司盾配置中没有启用的公司，请检查 config.xlsx")

        summaries = []
        for company in companies:
            company_name = str(company.get("公司名称") or "").strip()
            logger.info(f"开始处理公司: {company_name}，运行模式: {self.mode}")
            try:
                login = SingleWindowLoginSubtask(self.context, company)
                if not login.run():
                    logger.error(f"公司登录失败，继续下一家公司: {company_name}")
                    summaries.append({"公司": company_name, "结果": "登录失败"})
                    continue

                declaration = DeclarationSubtask(self.context, company, self.mode)
                declaration.run()
                todo = self.context.get("todo") or []
                stats = self.context.get("candidate_stats") or {}
                if not todo:
                    logger.info(f"{company_name} 没有待处理报关单")
                    summaries.append({"公司": company_name, "结果": "无待处理", **stats})
                    continue

                download = DownloadSubtask(self.context, company, self.mode)
                if not download.run():
                    summaries.append({"公司": company_name, "结果": "下载阶段异常", **stats})
                    continue
                summaries.append({"公司": company_name, "结果": "完成", **stats,
                                  **self._download_summary()})
            except Exception as e:
                logger.exception(f"公司处理失败，继续下一家公司: {company_name} - {e}")
                summaries.append({"公司": company_name, "结果": f"失败: {e}"})
            finally:
                self.context.close_all_browsers()

        for summary in summaries:
            logger.info("运行汇总: " + "，".join(f"{key}={value}" for key, value in summary.items()))
        return True

    def _download_summary(self) -> dict:
        records = self.context.get("download_records") or []
        counter = {
            "成功下载": sum(1 for record in records if record.get("status") == DOWNLOADED),
            "下载失败": sum(1 for record in records if record.get("status") == FAIL_DOWNLOAD),
            "解析失败": sum(1 for record in records if record.get("status") == FAIL_PARSE),
            "需人工": sum(1 for record in records if record.get("status") == MANUAL),
        }
        return {"处理数": len(records), **counter}
