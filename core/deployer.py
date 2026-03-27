# -*- coding: utf-8 -*-
"""
部署模块
核心部署逻辑，负责安装 xlam 文件和配置注册表
"""

import os
import shutil
import winreg
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from config import XLAM_FILENAME, ADDIN_DIR, TARGET_PATH, VERSION_FILENAME


def get_office_version() -> str:
    """
    自动检测已安装的 Office 版本

    Returns:
        str: Office 版本号字符串，如 '16.0'，未找到时默认返回 '16.0'
    """
    # 常见 Office 版本：16.0 (2016/2019/365), 15.0 (2013), 14.0 (2010)
    for ver in ["16.0", "15.0", "14.0"]:
        try:
            winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                f"Software\\Microsoft\\Office\\{ver}\\Excel\\Options"
            )
            return ver
        except FileNotFoundError:
            continue
    return "16.0"


def register_addin(xlam_path: str, log_callback: Callable[[str], None]) -> bool:
    """
    注册 Excel 加载项到注册表

    Args:
        xlam_path: xlam 文件完整路径
        log_callback: 日志回调函数

    Returns:
        bool: 是否注册成功
    """
    try:
        version = get_office_version()
        key_path = f"Software\\Microsoft\\Office\\{version}\\Excel\\Options"

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # 查找可用的 OPEN 键名，避免覆盖其他加载项
        i = 0
        target_key_name = None

        while True:
            name = "OPEN" if i == 0 else f"OPEN{i}"
            try:
                val, _ = winreg.QueryValueEx(key, name)
                if XLAM_FILENAME in str(val):
                    target_key_name = name  # 复用已有键
                    break
                i += 1
            except FileNotFoundError:
                target_key_name = name  # 找到空位
                break

        # 写入注册表
        winreg.SetValueEx(
            key,
            target_key_name,
            0,
            winreg.REG_SZ,
            f'/R "{xlam_path}"'
        )
        winreg.CloseKey(key)

        log_callback(f"📝 注册表已写入（Office {version}）：{target_key_name}")
        return True

    except Exception as e:
        log_callback(f"❌ 注册表写入失败：{str(e)}")
        return False


def deploy(
    xlam_bytes: bytes,
    log_callback: Callable[[str], None],
    version_info: Optional[dict] = None
) -> bool:
    """
    部署 xlam 文件的完整流程

    Args:
        xlam_bytes: xlam 文件的二进制内容
        log_callback: 日志回调函数
        version_info: 可选，远程 version.json 的内容，安装完成后写入本地

    Returns:
        bool: 是否部署成功
    """
    try:
        # 步骤 1：确保目标目录存在
        os.makedirs(ADDIN_DIR, exist_ok=True)
        log_callback(f"📁 目标目录：{ADDIN_DIR}")

        # 步骤 2：备份原文件
        if TARGET_PATH.exists():
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = Path(str(TARGET_PATH) + f".bak_{ts}")
            shutil.copy2(TARGET_PATH, backup_path)
            log_callback(f"💾 已备份原文件：{backup_path.name}")

        # 步骤 3：写入新文件
        with open(TARGET_PATH, "wb") as f:
            f.write(xlam_bytes)
        log_callback("✅ xlam 文件已保存至目标路径")

        # 步骤 4：写注册表
        if not register_addin(str(TARGET_PATH), log_callback):
            return False

        # 步骤 5：保存本地 version.json
        if version_info:
            local_version_path = ADDIN_DIR / VERSION_FILENAME
            import json
            with open(local_version_path, "w", encoding="utf-8") as f:
                json.dump(version_info, f, ensure_ascii=False, indent=2)
            log_callback(f"📄 本地版本信息已更新：{version_info.get('version', '')}")

        # 步骤 6：完成
        log_callback("🎉 安装完成，请重启 Excel 使更改生效")
        return True

    except Exception as e:
        log_callback(f"❌ 部署失败：{str(e)}")
        return False


def get_installed_version() -> str:
    """
    获取已安装的版本号

    Returns:
        str: 版本号，未安装返回 '未安装'
    """
    import json
    local_version_path = ADDIN_DIR / VERSION_FILENAME

    if not local_version_path.exists():
        return "未安装"

    try:
        with open(local_version_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("version", "未知")
    except Exception:
        return "未知"