# -*- coding: utf-8 -*-
"""
版本管理模块
负责读取和比较版本信息
"""

import json
import os
from pathlib import Path
from packaging.version import Version

from utils.config import config


def is_newer(remote_ver: str, local_ver: str) -> bool:
    """
    判断 remote_ver 是否比 local_ver 更新

    Args:
        remote_ver: 远程版本号字符串
        local_ver: 本地版本号字符串

    Returns:
        bool: 远程版本是否更新
    """
    try:
        return Version(remote_ver) > Version(local_ver)
    except Exception:
        return False


def read_local_version() -> dict:
    """
    读取本地已安装版本信息

    Returns:
        dict: 版本信息字典，未安装时返回 {'version': '未安装'}
    """
    local_version_path = config.ADDIN_DIR / config.VERSION_FILENAME

    if not local_version_path.exists():
        return {'version': '未安装', 'changelog': ''}

    try:
        with open(local_version_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {'version': '未知', 'changelog': ''}


def read_remote_version(webdav_client) -> dict:
    """
    从 WebDAV 下载并解析 version.json

    Args:
        webdav_client: WebDAV 客户端实例

    Returns:
        dict: 远程版本信息字典，失败时返回 None
    """
    try:
        version_json = webdav_client.get_version_info()
        return version_json
    except Exception as e:
        return None


def save_local_version(version_info: dict) -> bool:
    """
    保存版本信息到本地

    Args:
        version_info: 版本信息字典

    Returns:
        bool: 是否保存成功
    """
    local_version_path = config.ADDIN_DIR / config.VERSION_FILENAME

    try:
        os.makedirs(local_version_path.parent, exist_ok=True)
        with open(local_version_path, 'w', encoding='utf-8') as f:
            json.dump(version_info, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False