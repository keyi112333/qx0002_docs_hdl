#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置读取模块
"""
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


class ConfigError(Exception):
    """配置错误"""
    pass


class ConfigLoader:
    """配置加载器"""

    def __init__(self, config_path: Optional[Union[str, Path]] = None):
        self._config: Dict[str, Any] = {}
        self._path: Optional[Path] = None
        if config_path:
            self.load(config_path)

    def load(self, config_path: Union[str, Path]) -> Dict[str, Any]:
        """加载Excel配置文件"""
        if not HAS_OPENPYXL:
            raise ConfigError("需要安装 openpyxl: pip install openpyxl")

        path = Path(config_path)
        if not path.exists():
            raise ConfigError(f"配置文件不存在: {config_path}")

        self._path = path
        workbook = openpyxl.load_workbook(path, data_only=True)

        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            self._config[sheet_name] = self._read_sheet(sheet)

        workbook.close()
        return self._config

    def _read_sheet(self, sheet) -> Dict[str, Any]:
        """读取工作表为字典"""
        result = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                result[row[0]] = row[1]
        return result

    def read_table(self, sheet_name: str, header_row: int = 1) -> List[Dict[str, Any]]:
        """读取二维表工作表（首行为表头），返回字典列表

        Args:
            sheet_name: 工作表名称
            header_row: 表头所在行号

        Returns:
            List[Dict]: 每行数据为一个字典，空行自动忽略
        """
        if not HAS_OPENPYXL:
            raise ConfigError("需要安装 openpyxl: pip install openpyxl")
        if not self._path or not self._path.exists():
            raise ConfigError("配置文件未加载")

        workbook = openpyxl.load_workbook(self._path, data_only=True)
        try:
            if sheet_name not in workbook.sheetnames:
                raise ConfigError(f"工作表不存在: {sheet_name}")
            rows = list(workbook[sheet_name].iter_rows(values_only=True))
        finally:
            workbook.close()

        if len(rows) < header_row:
            return []

        headers = [
            str(value).strip() if value is not None else ""
            for value in rows[header_row - 1]
        ]
        table: List[Dict[str, Any]] = []
        for row in rows[header_row:]:
            if all(value is None or str(value).strip() == "" for value in row):
                continue
            table.append({
                headers[index]: value
                for index, value in enumerate(row)
                if index < len(headers) and headers[index]
            })
        return table

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置项（支持 sheet.key 格式）"""
        if '.' in key:
            sheet, k = key.split('.', 1)
            return self._config.get(sheet, {}).get(k, default)
        return self._config.get(key, default)

    @property
    def config(self) -> Dict[str, Any]:
        return self._config.copy()