# -*- coding: utf-8 -*-
"""
远程安装 Tab
提供从 WebDAV 服务器安装和更新 xlam 的功能
"""

import json
import threading
from pathlib import Path
from typing import Callable, Optional

import customtkinter as ctk

from config import (
    WEBDAV_DEFAULT_URL, WEBDAV_DEFAULT_USER, WEBDAV_DEFAULT_PASS,
    USER_CONFIG_DIR, USER_CONFIG_FILE
)
from core.webdav_client import WebDAVClient
from core.version import is_newer, read_local_version, read_remote_version


class RemoteInstallTab(ctk.CTkFrame):
    """远程安装 Tab 页面"""

    def __init__(
        self,
        master,
        log_callback: Callable[[str], None],
        install_callback: Callable[[bytes, Optional[dict]], None]
    ):
        """
        初始化远程安装 Tab

        Args:
            master: 父组件
            log_callback: 日志回调函数
            install_callback: 安装回调函数，接收 (文件二进制内容, 版本信息)
        """
        super().__init__(master)

        self.log = log_callback
        self.install = install_callback
        self.webdav_client: Optional[WebDAVClient] = None
        self.remote_version_info: Optional[dict] = None

        self._create_widgets()
        self._load_user_config()

    def _create_widgets(self):
        """创建界面组件"""
        # WebDAV 配置区
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.pack(fill="x", padx=20, pady=10)

        # 标题
        self.config_title = ctk.CTkLabel(
            self.config_frame,
            text="WebDAV 配置",
            font=("", 14, "bold")
        )
        self.config_title.pack(anchor="w", padx=10, pady=(10, 5))

        # 地址输入
        self.url_label = ctk.CTkLabel(self.config_frame, text="服务器地址：")
        self.url_label.pack(anchor="w", padx=10)
        self.url_entry = ctk.CTkEntry(
            self.config_frame,
            placeholder_text="https://your-server/dav/"
        )
        self.url_entry.pack(fill="x", padx=10, pady=(0, 5))

        # 用户名输入
        self.user_label = ctk.CTkLabel(self.config_frame, text="用户名：")
        self.user_label.pack(anchor="w", padx=10)
        self.user_entry = ctk.CTkEntry(self.config_frame)
        self.user_entry.pack(fill="x", padx=10, pady=(0, 5))

        # 密码输入
        self.pass_label = ctk.CTkLabel(self.config_frame, text="密码：")
        self.pass_label.pack(anchor="w", padx=10)
        self.pass_entry = ctk.CTkEntry(self.config_frame, show="*")
        self.pass_entry.pack(fill="x", padx=10, pady=(0, 10))

        # 配置按钮区
        self.config_btn_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        self.config_btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.save_btn = ctk.CTkButton(
            self.config_btn_frame,
            text="💾 保存配置",
            width=100,
            command=self._on_save_config
        )
        self.save_btn.pack(side="left", padx=(0, 10))

        self.test_btn = ctk.CTkButton(
            self.config_btn_frame,
            text="🔗 测试连接",
            width=100,
            command=self._on_test_connection
        )
        self.test_btn.pack(side="left")

        # 连接状态标签
        self.status_label = ctk.CTkLabel(
            self.config_frame,
            text="",
            text_color="gray"
        )
        self.status_label.pack(anchor="w", padx=10, pady=(0, 10))

        # 版本信息区
        self.version_frame = ctk.CTkFrame(self)
        self.version_frame.pack(fill="x", padx=20, pady=10)

        self.version_title = ctk.CTkLabel(
            self.version_frame,
            text="版本信息",
            font=("", 14, "bold")
        )
        self.version_title.pack(anchor="w", padx=10, pady=(10, 5))

        # 版本显示
        self.version_info_frame = ctk.CTkFrame(self.version_frame, fg_color="transparent")
        self.version_info_frame.pack(fill="x", padx=10, pady=(0, 10))

        # 远程版本
        self.remote_ver_label = ctk.CTkLabel(
            self.version_info_frame,
            text="远程版本：未检查"
        )
        self.remote_ver_label.pack(anchor="w")

        # 本地版本
        self.local_ver_label = ctk.CTkLabel(
            self.version_info_frame,
            text="本地版本：未安装"
        )
        self.local_ver_label.pack(anchor="w")

        # 更新日志
        self.changelog_label = ctk.CTkLabel(
            self.version_info_frame,
            text="更新日志：",
            wraplength=600
        )
        self.changelog_label.pack(anchor="w", pady=(5, 0))

        # 检查更新按钮
        self.check_btn = ctk.CTkButton(
            self.version_frame,
            text="🔍 检查更新",
            width=120,
            command=self._on_check_update
        )
        self.check_btn.pack(anchor="w", padx=10, pady=(0, 10))

        # 操作区
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.pack(fill="x", padx=20, pady=10)

        # 进度条
        self.progress = ctk.CTkProgressBar(self.action_frame)
        self.progress.pack(fill="x", padx=10, pady=10)
        self.progress.set(0)

        # 安装按钮
        self.install_btn = ctk.CTkButton(
            self.action_frame,
            text="⬇️ 一键安装 / 更新",
            width=200,
            height=40,
            font=("", 16),
            command=self._on_install
        )
        self.install_btn.pack(pady=(0, 10))

    def _load_user_config(self):
        """加载用户配置"""
        if USER_CONFIG_FILE.exists():
            try:
                with open(USER_CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # 填充配置
                if config.get("webdav_url"):
                    self.url_entry.insert(0, config["webdav_url"])
                elif WEBDAV_DEFAULT_URL:
                    self.url_entry.insert(0, WEBDAV_DEFAULT_URL)

                if config.get("webdav_user"):
                    self.user_entry.insert(0, config["webdav_user"])
                elif WEBDAV_DEFAULT_USER:
                    self.user_entry.insert(0, WEBDAV_DEFAULT_USER)

                if config.get("webdav_pass"):
                    self.pass_entry.insert(0, config["webdav_pass"])
                elif WEBDAV_DEFAULT_PASS:
                    self.pass_entry.insert(0, WEBDAV_DEFAULT_PASS)

            except Exception:
                pass
        else:
            # 使用默认配置
            if WEBDAV_DEFAULT_URL:
                self.url_entry.insert(0, WEBDAV_DEFAULT_URL)
            if WEBDAV_DEFAULT_USER:
                self.user_entry.insert(0, WEBDAV_DEFAULT_USER)
            if WEBDAV_DEFAULT_PASS:
                self.pass_entry.insert(0, WEBDAV_DEFAULT_PASS)

        # 更新本地版本显示
        self._update_local_version()

    def _save_user_config(self):
        """保存用户配置"""
        config = {
            "webdav_url": self.url_entry.get(),
            "webdav_user": self.user_entry.get(),
            "webdav_pass": self.pass_entry.get()
        }

        USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(USER_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def _on_save_config(self):
        """保存配置按钮点击"""
        self._save_user_config()
        self.status_label.configure(text="✅ 配置已保存", text_color="green")

    def _on_test_connection(self):
        """测试连接按钮点击"""
        url = self.url_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.pass_entry.get()

        if not url:
            self.status_label.configure(text="❌ 请输入服务器地址", text_color="red")
            return

        self.test_btn.configure(state="disabled")
        self.status_label.configure(text="正在连接...", text_color="gray")

        def test():
            try:
                client = WebDAVClient(url, user, password)
                success, msg = client.test_connection()

                if success:
                    self.webdav_client = client
                    self.after(0, lambda: self.status_label.configure(
                        text="✅ 连接成功", text_color="green"
                    ))
                else:
                    self.after(0, lambda: self.status_label.configure(
                        text=f"❌ 连接失败：{msg}", text_color="red"
                    ))
            except Exception as e:
                self.after(0, lambda: self.status_label.configure(
                    text=f"❌ 连接失败：{str(e)}", text_color="red"
                ))
            finally:
                self.after(0, lambda: self.test_btn.configure(state="normal"))

        threading.Thread(target=test, daemon=True).start()

    def _update_local_version(self):
        """更新本地版本显示"""
        local_info = read_local_version()
        self.local_ver_label.configure(
            text=f"本地版本：{local_info.get('version', '未安装')}"
        )

    def _on_check_update(self):
        """检查更新按钮点击"""
        url = self.url_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.pass_entry.get()

        if not url:
            self.log("❌ 请输入服务器地址")
            return

        self.check_btn.configure(state="disabled")
        self.log("🔍 正在检查更新...")

        def check():
            try:
                client = WebDAVClient(url, user, password)
                self.webdav_client = client

                # 获取远程版本
                remote_info = read_remote_version(client)

                if not remote_info:
                    self.after(0, lambda: self.log("❌ 无法获取远程版本信息"))
                    return

                self.remote_version_info = remote_info

                # 更新显示
                remote_ver = remote_info.get("version", "未知")
                changelog = remote_info.get("changelog", "无")

                self.after(0, lambda: self.remote_ver_label.configure(
                    text=f"远程版本：{remote_ver}"
                ))
                self.after(0, lambda: self.changelog_label.configure(
                    text=f"更新日志：{changelog}"
                ))

                # 更新本地版本
                self.after(0, self._update_local_version)

                # 比较版本
                local_info = read_local_version()
                local_ver = local_info.get("version", "0")

                if local_ver == "未安装":
                    self.after(0, lambda: self.log(f"✅ 发现新版本：{remote_ver}（本地未安装）"))
                elif is_newer(remote_ver, local_ver):
                    self.after(0, lambda: self.log(f"✅ 发现新版本：{remote_ver}"))
                else:
                    self.after(0, lambda: self.log("✅ 已是最新版本"))

            except Exception as e:
                self.after(0, lambda: self.log(f"❌ 检查更新失败：{str(e)}"))
            finally:
                self.after(0, lambda: self.check_btn.configure(state="normal"))

        threading.Thread(target=check, daemon=True).start()

    def _on_install(self):
        """安装按钮点击"""
        if not self.webdav_client:
            self.log("❌ 请先测试连接")
            return

        self.install_btn.configure(state="disabled", text="下载中...")
        self.progress.set(0)
        self.log("⬇️ 正在下载 xlam 文件...")

        def download():
            try:
                # 下载文件
                xlam_bytes = self.webdav_client.download_xlam()
                self.after(0, lambda: self.progress.set(0.5))
                self.after(0, lambda: self.log(f"✅ 下载完成（{len(xlam_bytes)} 字节）"))

                # 获取版本信息
                version_info = self.webdav_client.get_version_info()

                self.after(0, lambda: self.progress.set(0.7))

                # 调用安装回调
                self.after(0, lambda: self.install(xlam_bytes, version_info))

                self.after(0, lambda: self.progress.set(1))

            except Exception as e:
                self.after(0, lambda: self.log(f"❌ 下载失败：{str(e)}"))
            finally:
                self.after(0, lambda: self.install_btn.configure(
                    state="normal", text="⬇️ 一键安装 / 更新"
                ))

        threading.Thread(target=download, daemon=True).start()

    def refresh_version(self):
        """刷新版本信息"""
        self._update_local_version()