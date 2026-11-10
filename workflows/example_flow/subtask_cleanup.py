#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：清理资源
"""
from core.task_base import TaskBase
from log.loggingConfig import logger


class CleanupSubtask(TaskBase):
    """子任务：清理资源

    职责：
    - 关闭所有浏览器
    - 清理临时数据
    """

    def __init__(self, context):
        super().__init__("清理资源", context)

    def execute(self) -> bool:
        try:
            # 关闭所有浏览器
            self.context.close_all_browsers()
            logger.info("所有浏览器已关闭")

            # 清理临时数据
            self.context.clear()
            logger.info("临时数据已清理")

            return True
        except Exception as e:
            logger.error(f"清理资源失败: {e}")
            return False