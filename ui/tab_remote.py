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

from utils.config import config
from core.webdav_client import WebDAVClient


class RemoteInstallTab(ctk.CTkFrame):
    """远程安装 Tab 页面"""

    def __init__(
        self,
        master,
        log_callback: Callable[[str], None],
        install_callback: Callable[[bytes, Optional[dict]], None],
        version_update_callback: Callable[[], None]
    ):
        """
        初始化远程安装 Tab

        Args:
            master: 父组件
            log_callback: 日志回调函数
            install_callback: 安装回调函数，接收 (文件二进制内容, 版本信息)
            version_update_callback: 版本更新回调
        """
        super().__init__(master)

        self.install = install_callback
        self._version_update = version_update_callback
        self.webdav_client: Optional[WebDAVClient] = None
        self.remote_version_info: Optional[dict] = None

        self._create_widgets()
        self._load_user_config()

    def _create_widgets(self):
        """创建界面组件"""
        # 使用 grid 布局确保按钮常驻
        self.grid_rowconfigure(0, weight=0)  # 配置区
        self.grid_rowconfigure(1, weight=1)  # 日志区（可伸缩）
        self.grid_rowconfigure(2, weight=0)  # 操作区（固定）
        self.grid_columnconfigure(0, weight=1)

        # WebDAV 配置区
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5))

        # 配置输入（紧凑布局）
        config_grid = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        config_grid.pack(fill="x", padx=10, pady=10)

        # 服务器地址
        ctk.CTkLabel(config_grid, text="服务器:", font=(config.FONT_FAMILY, 11)).grid(row=0, column=0, sticky="w", pady=2)
        self.url_entry = ctk.CTkEntry(config_grid, show="*", width=280)
        self.url_entry.grid(row=0, column=1, padx=(5, 10), pady=2, sticky="ew")

        # 用户名
        ctk.CTkLabel(config_grid, text="用户名:", font=(config.FONT_FAMILY, 11)).grid(row=1, column=0, sticky="w", pady=2)
        self.user_entry = ctk.CTkEntry(config_grid, show="*", width=280)
        self.user_entry.grid(row=1, column=1, padx=(5, 10), pady=2, sticky="ew")

        # 密码
        ctk.CTkLabel(config_grid, text="密码:", font=(config.FONT_FAMILY, 11)).grid(row=2, column=0, sticky="w", pady=2)
        self.pass_entry = ctk.CTkEntry(config_grid, show="*", width=280)
        self.pass_entry.grid(row=2, column=1, padx=(5, 10), pady=2, sticky="ew")

        config_grid.columnconfigure(1, weight=1)

        # 配置按钮
        btn_frame = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.save_btn = ctk.CTkButton(
            btn_frame, text="💾 保存", width=80, height=28,
            font=(config.FONT_FAMILY, 11), command=self._on_save_config
        )
        self.save_btn.pack(side="left", padx=(0, 5))

        self.test_btn = ctk.CTkButton(
            btn_frame, text="🔗 测试", width=80, height=28,
            font=(config.FONT_FAMILY, 11), command=self._on_test_connection
        )
        self.test_btn.pack(side="left")

        self.status_label = ctk.CTkLabel(btn_frame, text="", text_color="gray", font=(config.FONT_FAMILY, 11))
        self.status_label.pack(side="left", padx=10)

        # 操作区（固定在底部）
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(5, 10))

        # 进度条
        self.progress = ctk.CTkProgressBar(self.action_frame)
        self.progress.pack(fill="x", padx=10, pady=10)
        self.progress.set(0)

        # 安装按钮
        self.install_btn = ctk.CTkButton(
            self.action_frame,
            text="⬇️ 一键安装 / 更新",
            width=200,
            height=36,
            font=(config.FONT_FAMILY, 14),
            command=self._on_install
        )
        self.install_btn.pack(pady=(0, 10))

        # 日志区（中间可伸缩）
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=5)

        self.log_label = ctk.CTkLabel(
            self.log_frame, text="操作日志",
            font=(config.FONT_FAMILY, 11, "bold")
        )
        self.log_label.pack(anchor="w", padx=10, pady=(5, 0))

        self.log_text = ctk.CTkTextbox(
            self.log_frame,
            state="disabled",
            font=(config.FONT_MONO, 10),
            height=60  # 最小高度
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

    def _log(self, message: str):
        """输出日志"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _load_user_config(self):
        """加载用户配置"""
        if config.USER_CONFIG_FILE.exists():
            try:
                with open(config.USER_CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved_config = json.load(f)

                if saved_config.get("webdav_url"):
                    self.url_entry.insert(0, saved_config["webdav_url"])
                elif config.WEBDAV_DEFAULT_URL:
                    self.url_entry.insert(0, config.WEBDAV_DEFAULT_URL)

                if saved_config.get("webdav_user"):
                    self.user_entry.insert(0, saved_config["webdav_user"])
                elif config.WEBDAV_DEFAULT_USER:
                    self.user_entry.insert(0, config.WEBDAV_DEFAULT_USER)

                if saved_config.get("webdav_pass"):
                    self.pass_entry.insert(0, saved_config["webdav_pass"])
                elif config.WEBDAV_DEFAULT_PASS:
                    self.pass_entry.insert(0, config.WEBDAV_DEFAULT_PASS)

            except Exception:
                pass
        else:
            if config.WEBDAV_DEFAULT_URL:
                self.url_entry.insert(0, config.WEBDAV_DEFAULT_URL)
            if config.WEBDAV_DEFAULT_USER:
                self.user_entry.insert(0, config.WEBDAV_DEFAULT_USER)
            if config.WEBDAV_DEFAULT_PASS:
                self.pass_entry.insert(0, config.WEBDAV_DEFAULT_PASS)

    def _save_user_config(self):
        """保存用户配置"""
        saved_config = {
            "webdav_url": self.url_entry.get(),
            "webdav_user": self.user_entry.get(),
            "webdav_pass": self.pass_entry.get()
        }

        config.USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(config.USER_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(saved_config, f, ensure_ascii=False, indent=2)

    def _on_save_config(self):
        """保存配置按钮点击"""
        self._save_user_config()
        self.status_label.configure(text="✅ 已保存", text_color="green")

    def _on_test_connection(self):
        """测试连接按钮点击"""
        url = self.url_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.pass_entry.get()
        folder = config.WEBDAV_DEFAULT_FOLDER

        if not url:
            self.status_label.configure(text="❌ 请输入地址", text_color="red")
            return

        self.test_btn.configure(state="disabled")
        self.status_label.configure(text="连接中...", text_color="gray")

        def test():
            try:
                client = WebDAVClient(url, user, password, folder)
                success, msg = client.test_connection()

                if success:
                    self.webdav_client = client
                    self.after(0, lambda: self.status_label.configure(text="✅ 连接成功", text_color="green"))
                else:
                    self.after(0, lambda: self.status_label.configure(text=f"❌ 失败", text_color="red"))
            except Exception as e:
                self.after(0, lambda: self.status_label.configure(text="❌ 失败", text_color="red"))
            finally:
                self.after(0, lambda: self.test_btn.configure(state="normal"))

        threading.Thread(target=test, daemon=True).start()

    def _on_install(self):
        """安装按钮点击"""
        if not self.webdav_client:
            self._log("❌ 请先测试连接")
            return

        self.install_btn.configure(state="disabled", text="下载中...")
        self.progress.set(0)
        self._log("⬇️ 正在下载 xlam 文件...")

        def download():
            try:
                xlam_bytes = self.webdav_client.download_xlam()
                self.after(0, lambda: self.progress.set(0.5))
                self.after(0, lambda: self._log(f"✅ 下载完成（{len(xlam_bytes)} 字节）"))

                version_info = self.webdav_client.get_version_info()
                self.after(0, lambda: self.progress.set(0.7))

                self.after(0, lambda: self.install(xlam_bytes, version_info))
                self.after(0, lambda: self.progress.set(1))

            except Exception as e:
                self.after(0, lambda: self._log(f"❌ 下载失败：{str(e)}"))
            finally:
                self.after(0, lambda: self.install_btn.configure(state="normal", text="⬇️ 一键安装 / 更新"))

        threading.Thread(target=download, daemon=True).start()

    def refresh_version(self):
        """刷新版本信息"""
        pass