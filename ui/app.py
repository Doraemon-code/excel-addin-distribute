# -*- coding: utf-8 -*-
"""
主窗口模块
创建应用主界面，包含 Tab 页和日志输出区
"""

import threading
from typing import Optional

import customtkinter as ctk

from utils.config import config
from core.deployer import deploy, get_installed_version
from ui.tab_remote import RemoteInstallTab
from ui.tab_local import LocalInstallTab


class App(ctk.CTk):
    """主应用窗口"""

    def __init__(self):
        super().__init__()

        # 窗口配置
        self.title(f"{config.APP_TITLE} v{config.APP_VERSION}")
        self.geometry(f"{config.WINDOW_WIDTH}x{config.WINDOW_HEIGHT}")
        self.resizable(False, False)

        # 设置主题
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # 创建界面
        self._create_widgets()

    def _create_widgets(self):
        """创建界面组件"""
        # Header 区域
        self.header_frame = ctk.CTkFrame(self, height=50)
        self.header_frame.pack(fill="x", padx=10, pady=(10, 5))

        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text=config.APP_TITLE,
            font=("", 18, "bold")
        )
        self.title_label.pack(side="left", padx=10, pady=10)

        # 本地版本显示
        installed_ver = get_installed_version()
        self.version_label = ctk.CTkLabel(
            self.header_frame,
            text=f"本地已安装版本：{installed_ver}",
            font=("", 12),
            text_color="gray"
        )
        self.version_label.pack(side="right", padx=10, pady=10)

        # Tab View
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=5)

        # 添加 Tab
        self.remote_tab = self.tab_view.add("🌐 远程安装")
        self.local_tab = self.tab_view.add("📁 本地安装")

        # 创建 Tab 内容
        self.remote_install_tab = RemoteInstallTab(
            self.remote_tab,
            log_callback=self._log,
            install_callback=self._on_remote_install
        )
        self.remote_install_tab.pack(fill="both", expand=True)

        self.local_install_tab = LocalInstallTab(
            self.local_tab,
            log_callback=self._log,
            install_callback=self._on_local_install
        )
        self.local_install_tab.pack(fill="both", expand=True)

        # 日志输出区
        self.log_frame = ctk.CTkFrame(self, height=120)
        self.log_frame.pack(fill="x", padx=10, pady=(5, 10))
        self.log_frame.pack_propagate(False)

        self.log_label = ctk.CTkLabel(
            self.log_frame,
            text="操作日志",
            font=("", 12, "bold")
        )
        self.log_label.pack(anchor="w", padx=10, pady=(5, 0))

        self.log_text = ctk.CTkTextbox(
            self.log_frame,
            height=80,
            state="disabled",
            font=("Consolas", 11)
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

    def _log(self, message: str):
        """
        输出日志

        Args:
            message: 日志消息
        """
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_remote_install(self, xlam_bytes: bytes, version_info: Optional[dict] = None):
        """
        远程安装回调

        Args:
            xlam_bytes: xlam 文件二进制内容
            version_info: 版本信息字典
        """
        def install():
            deploy(xlam_bytes, self._log, version_info)
            # 更新版本显示
            self.after(0, self._update_version_display)

        threading.Thread(target=install, daemon=True).start()

    def _on_local_install(self, xlam_bytes: bytes):
        """
        本地安装回调

        Args:
            xlam_bytes: xlam 文件二进制内容
        """
        def install():
            deploy(xlam_bytes, self._log)
            # 更新版本显示
            self.after(0, self._update_version_display)

        threading.Thread(target=install, daemon=True).start()

    def _update_version_display(self):
        """更新版本显示"""
        installed_ver = get_installed_version()
        self.version_label.configure(text=f"本地已安装版本：{installed_ver}")
        self.remote_install_tab.refresh_version()