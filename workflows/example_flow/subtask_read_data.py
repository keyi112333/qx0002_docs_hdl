#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子任务：读取Excel数据
"""
from core.task_base import TaskBase
from data_processing.excel import ExcelHandler


class ReadDataSubtask(TaskBase):
    """子任务：读取Excel数据

    职责：
    - 从输入目录读取Excel文件
    - 解析数据并存入上下文供后续子任务使用
    """

    def __init__(self, context):
        super().__init__("读取Excel数据", context)

    def execute(self) -> bool:
        # 从上下文获取配置
        config = self.context.config
        input_file = config.get_path('input', 'data.xlsx')

        # 读取Excel数据
        excel = ExcelHandler(str(input_file))

        # 读取工作表数据（根据实际需求修改sheet名称）
        data = excel.read_sheet("Sheet1")

        # 存入上下文供后续子任务使用
        self.context.set("input_data", data)

        excel.close()

        return len(data) > 0