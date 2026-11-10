#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置模块
"""
from pathlib import Path
from typing import Dict, List, Optional
import shutil
from config.config_loader import ConfigLoader, ConfigError

# 项目python代码的log输出目录，根据实际情况修改
PYHON_LOG_FOLDER_PATH = r"./log_data/"
# UserConfig路径，后续需修改为项目工作目录下的指定文件路径
USER_CONFIG_PATH = r"./data/UserConfig.xlsx"


class Config:
    """项目配置"""

    def __init__(self, config_path: str = None):
        self._loader: Optional[ConfigLoader] = None
        self._paths: Dict[str, Path] = {}
        self._config_source_path: Optional[Path] = None  # 用户配置文件源路径

        if config_path:
            self.load(config_path)

    def load(self, config_path: str) -> None:
        """加载配置文件"""
        self._config_source_path = Path(config_path)  # 保存源路径
        self._loader = ConfigLoader(config_path)

        # 读取基础配置
        base_config = self._loader._config.get('基础配置', {})
        self.project_name = base_config.get('项目名称', 'XX流程')
        self.base_folder = base_config.get('基础路径', './data')

        # 读取目录配置
        dir_config = self._loader._config.get('目录配置', {})
        self._dir_names = {
            'config': dir_config.get('配置目录', '01_配置'),
            'input': dir_config.get('输入目录', '02_输入'),
            'output': dir_config.get('输出目录', '03_输出'),
            'log': dir_config.get('日志目录', '04_日志'),
            'screenshot': dir_config.get('截图目录', '05_错误截图'),
            'temp': dir_config.get('临时目录', '06_临时文件'),
        }
        self._init_paths()

    def _init_paths(self) -> None:
        """初始化路径"""
        base = Path(self.base_folder) / self.project_name
        for key, name in self._dir_names.items():
            self._paths[key] = base / name

    def get_path(self, path_type: str, filename: str = "") -> Path:
        """获取路径"""
        path = self._paths.get(path_type, self._paths.get('input'))
        if filename:
            return path / filename
        return path

    def get(self, key: str, default=None):
        """获取配置项（支持 sheet.key 格式）

        Args:
            key: 配置键，格式为 "sheet名.配置项"，如 "基础配置.项目名称"
            default: 默认值

        Returns:
            配置值，不存在则返回默认值
        """
        if self._loader:
            return self._loader.get(key, default)
        return default

    def read_table(self, sheet_name: str, header_row: int = 1) -> List[Dict]:
        """读取二维表配置（首行为表头），如“公司盾配置”"""
        if not self._loader:
            raise ConfigError("配置文件未加载")
        return self._loader.read_table(sheet_name, header_row)

    def load_business_config(self, filename: str = "config.xlsx") -> Path:
        """加载配置目录下的业务配置（如公司盾配置、系统配置）

        业务配置文件存放在 UserConfig 指定的配置目录中，
        与 UserConfig 本身的基础配置、目录配置合并使用。
        """
        if not self._loader:
            raise ConfigError("配置文件未加载")

        path = self.get_path('config', filename)
        if not path.exists():
            raise ConfigError(f"业务配置文件不存在: {path}")

        self._loader.load(path)
        return path

    def ensure_directories(self) -> None:
        """确保目录存在，并复制用户配置文件到配置目录"""
        for path in self._paths.values():
            path.mkdir(parents=True, exist_ok=True)

        # 复制用户配置文件到配置目录
        if self._config_source_path and self._config_source_path.exists():
            config_dir = self._paths.get('config')
            if config_dir:
                dest_path = config_dir / self._config_source_path.name
                if not dest_path.exists():
                    shutil.copy2(self._config_source_path, dest_path)


# 全局配置实例
config = Config()