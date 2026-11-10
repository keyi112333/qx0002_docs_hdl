#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QX0002 报关单和放行通知单等单据处理流程 - 主任务
"""
from core.context import TaskContext
from core.task_base import TaskBase
from log.loggingConfig import logger
from workflows.qx0002_docs_hdl.subtask_login import SingleWindowLoginSubtask


class Qx0002MainTask(TaskBase):
    """主任务：按“公司盾配置”中启用的公司逐家执行"""

    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__(config.project_name, context)

    def execute(self) -> bool:
        companies = [
            company
            for company in self.context.config.read_table("公司盾配置")
            if str(company.get("启用") or "").strip() in ("是", "true", "True", "1")
        ]
        if not companies:
            raise ValueError("公司盾配置中没有启用的公司，请检查 config.xlsx")

        success_count = 0
        for company in companies:
            company_name = company.get("公司名称")
            try:
                login = SingleWindowLoginSubtask(self.context, company)
                if login.run():
                    success_count += 1
                else:
                    logger.error(f"公司登录失败，继续下一家公司: {company_name}")
            except Exception as e:
                logger.error(f"公司登录失败，继续下一家公司: {company_name} - {e}")
            finally:
                self.context.close_all_browsers()

        logger.info(f"单一窗口登录完成: 成功 {success_count}/{len(companies)} 家公司")
        if success_count == 0:
            raise RuntimeError(f"全部 {len(companies)} 家公司登录失败")
        return True
