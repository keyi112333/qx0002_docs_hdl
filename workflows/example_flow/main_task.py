#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例流程 - 主任务（自由组装子任务）
"""
from core.task_base import TaskBase
from core.context import TaskContext
from workflows.example_flow.subtask_login import LoginSubtask
from workflows.example_flow.subtask_read_data import ReadDataSubtask
from workflows.example_flow.subtask_input import InputSubtask
from workflows.example_flow.subtask_download import DownloadSubtask
from workflows.example_flow.subtask_cleanup import CleanupSubtask
from log.loggingConfig import logger


class ExampleMainTask(TaskBase):
    """示例流程主任务

    继承 TaskBase，在 execute 中自由组装子任务：
    - 可以根据条件决定是否执行
    - 可以循环执行
    - 可以自由控制顺序
    """

    def __init__(self, config):
        # 创建上下文
        context = TaskContext()
        context.config = config
        super().__init__("示例流程", context)

    def execute(self) -> bool:
        """自由组装子任务执行逻辑"""

        # # 1. 登录系统A
        # login = LoginSubtask(self.context)
        # if not login.run():
        #     logger.error("登录失败，流程终止")
        #     return False

        # # 2. 读取数据
        # read_data = ReadDataSubtask(self.context)
        # if not read_data.run():
        #     logger.error("读取数据失败，流程终止")
        #     return False

        # # 3. 根据数据量决定是否继续
        # data = self.context.get("input_data", [])
        # if len(data) == 0:
        #     logger.warning("数据为空，跳过录入步骤")
        # else:
        #     # 录入数据
        #     input_task = InputSubtask(self.context)
        #     if not input_task.run():
        #         logger.error("录入数据失败，但继续执行后续步骤")

        # # 4. 下载报表
        # download = DownloadSubtask(self.context)
        # download.run()

        # # 5. 清理资源（始终执行）
        # cleanup = CleanupSubtask(self.context)
        # cleanup.run()

        return True