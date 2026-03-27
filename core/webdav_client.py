# -*- coding: utf-8 -*-
"""
WebDAV 客户端模块
负责 WebDAV 连接、文件下载和版本信息获取
"""

import json
from io import BytesIO
from typing import Callable, Optional, Tuple

from webdav3.client import Client

from utils.config import config


class WebDAVClient:
    """WebDAV 客户端封装类"""

    def __init__(self, url: str, username: str, password: str):
        """
        初始化 WebDAV 客户端

        Args:
            url: WebDAV 服务器地址
            username: 用户名
            password: 密码
        """
        self.options = {
            'webdav_hostname': url,
            'webdav_login': username,
            'webdav_password': password
        }
        self.client = Client(self.options)

    def test_connection(self) -> Tuple[bool, str]:
        """
        测试 WebDAV 连接

        Returns:
            tuple: (成功与否, 错误信息)
        """
        try:
            # 尝试列出根目录
            self.client.list('/')
            return True, "连接成功"
        except Exception as e:
            return False, str(e)

    def download_file(
        self,
        remote_filename: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> bytes:
        """
        下载文件，返回二进制内容

        Args:
            remote_filename: 远程文件名
            progress_callback: 进度回调函数，参数为 (已下载字节, 总字节)

        Returns:
            bytes: 文件二进制内容
        """
        buffer = BytesIO()
        self.client.download_from(buffer, remote_filename)
        return buffer.getvalue()

    def get_version_info(self) -> dict:
        """
        下载并解析 version.json

        Returns:
            dict: 版本信息字典，失败时返回 None
        """
        try:
            content = self.download_file(config.VERSION_FILENAME)
            return json.loads(content.decode('utf-8'))
        except Exception:
            return None

    def download_xlam(
        self,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> bytes:
        """
        下载 xlam 文件

        Args:
            progress_callback: 进度回调函数

        Returns:
            bytes: xlam 文件二进制内容
        """
        return self.download_file(config.XLAM_FILENAME, progress_callback)

    def get_file_size(self, remote_filename: str) -> int:
        """
        获取远程文件大小

        Args:
            remote_filename: 远程文件名

        Returns:
            int: 文件大小（字节），失败返回 -1
        """
        try:
            info = self.client.info(remote_filename)
            return int(info.get('size', 0))
        except Exception:
            return -1