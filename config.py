# -*- coding: utf-8 -*-
"""
配置加载模块
从 config.yaml 加载配置并提供全局访问
"""

import os
from pathlib import Path
from typing import Any

import yaml


# 配置文件路径
CONFIG_FILE = Path(__file__).parent / "config.yaml"

# 默认配置
DEFAULT_CONFIG = {
    "webdav_default_url": "",
    "webdav_default_user": "",
    "webdav_default_pass": "",
    "xlam_filename": "MyAddin.xlam",
    "version_filename": "version.json",
    "app_title": "Excel 加载项安装工具",
    "app_version": "1.0.0",
    "window_width": 680,
    "window_height": 560,
}

# 加载配置
_config: dict = {}


def load_config() -> dict:
    """加载配置文件"""
    global _config

    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            _config = yaml.safe_load(f) or {}
    else:
        _config = {}

    # 合并默认配置
    for key, value in DEFAULT_CONFIG.items():
        if key not in _config:
            _config[key] = value

    return _config


def get(key: str, default: Any = None) -> Any:
    """获取配置项"""
    return _config.get(key, default)


# 加载配置
load_config()

# 导出常用配置常量
APP_TITLE = get("app_title")
APP_VERSION = get("app_version")
XLAM_FILENAME = get("xlam_filename")
VERSION_FILENAME = get("version_filename")
WINDOW_WIDTH = get("window_width")
WINDOW_HEIGHT = get("window_height")

# WebDAV 默认配置
WEBDAV_DEFAULT_URL = get("webdav_default_url", "")
WEBDAV_DEFAULT_USER = get("webdav_default_user", "")
WEBDAV_DEFAULT_PASS = get("webdav_default_pass", "")

# 目标安装路径
ADDIN_DIR = Path(os.environ["APPDATA"]) / "Microsoft" / "AddIns"
TARGET_PATH = ADDIN_DIR / XLAM_FILENAME

# 用户配置文件路径
USER_CONFIG_DIR = Path(os.environ["APPDATA"]) / "ExcelAddinInstaller"
USER_CONFIG_FILE = USER_CONFIG_DIR / "config.json"