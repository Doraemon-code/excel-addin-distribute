# -*- coding: utf-8 -*-
"""
配置加载模块
从 config.yaml 加载配置并提供全局访问
支持 PyInstaller 打包后读取内嵌的配置文件
"""

import os
import sys
from pathlib import Path
from typing import Any, Optional

import yaml


# 配置文件路径
_CONFIG_FILE: Optional[Path] = None
_config: dict = {}


def get_base_path() -> Path:
    """
    获取基础路径
    - 开发环境：项目根目录
    - 打包后：exe 所在目录或 PyInstaller 临时目录
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后
        # sys._MEIPASS 是 PyInstaller 解压的临时目录
        # sys.executable 是 exe 文件路径
        return Path(sys._MEIPASS)
    else:
        # 开发环境
        return Path(__file__).parent.parent


def get_config_path() -> Path:
    """获取配置文件路径"""
    global _CONFIG_FILE
    if _CONFIG_FILE is None:
        _CONFIG_FILE = get_base_path() / "config.yaml"
    return _CONFIG_FILE


def load_config() -> dict:
    """加载配置文件"""
    global _config

    config_path = get_config_path()
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            _config = yaml.safe_load(f) or {}
    else:
        _config = {}

    return _config


def get(key: str, default: Any = None) -> Any:
    """获取配置项"""
    if not _config:
        load_config()
    return _config.get(key, default)


# ==================== 常用配置属性 ====================

def _get_app_title() -> str:
    return get("app_title", "Excel 加载项安装工具")

def _get_app_version() -> str:
    return get("app_version", "1.0.0")

def _get_xlam_filename() -> str:
    return get("xlam_filename", "MyAddin.xlam")

def _get_version_filename() -> str:
    return get("version_filename", "version.json")

def _get_window_width() -> int:
    return get("window_width", 680)

def _get_window_height() -> int:
    return get("window_height", 560)

def _get_webdav_url() -> str:
    return get("webdav_default_url", "")

def _get_webdav_user() -> str:
    return get("webdav_default_user", "")

def _get_webdav_pass() -> str:
    return get("webdav_default_pass", "")

def _get_webdav_folder() -> str:
    return get("webdav_default_folder", "")


# 导出配置常量（延迟加载）
class Config:
    """配置类，提供延迟加载的配置属性"""

    @property
    def APP_TITLE(self) -> str:
        return _get_app_title()

    @property
    def APP_VERSION(self) -> str:
        return _get_app_version()

    @property
    def XLAM_FILENAME(self) -> str:
        return _get_xlam_filename()

    @property
    def VERSION_FILENAME(self) -> str:
        return _get_version_filename()

    @property
    def WINDOW_WIDTH(self) -> int:
        return _get_window_width()

    @property
    def WINDOW_HEIGHT(self) -> int:
        return _get_window_height()

    @property
    def WEBDAV_DEFAULT_URL(self) -> str:
        return _get_webdav_url()

    @property
    def WEBDAV_DEFAULT_USER(self) -> str:
        return _get_webdav_user()

    @property
    def WEBDAV_DEFAULT_PASS(self) -> str:
        return _get_webdav_pass()

    @property
    def WEBDAV_DEFAULT_FOLDER(self) -> str:
        return _get_webdav_folder()

    @property
    def ADDIN_DIR(self) -> Path:
        return Path(os.environ["APPDATA"]) / "Microsoft" / "AddIns"

    @property
    def TARGET_PATH(self) -> Path:
        return self.ADDIN_DIR / self.XLAM_FILENAME

    @property
    def USER_CONFIG_DIR(self) -> Path:
        return Path(os.environ["APPDATA"]) / "ExcelAddinInstaller"

    @property
    def USER_CONFIG_FILE(self) -> Path:
        return self.USER_CONFIG_DIR / "config.json"


# 全局配置实例
config = Config()