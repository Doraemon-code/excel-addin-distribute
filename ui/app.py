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
from core.version import read_local_version, read_remote_version
from core.webdav_client import WebDAVClient
from ui.tab_remote import RemoteInstallTab
from ui.tab_local import LocalInstallTab


class VersionDialog(ctk.CTkToplevel):
    """版本信息弹窗"""

    def __init__(self, parent, title: str, version: str, date: str, changelog: str = ""):
        super().__init__(parent)

        self.title(title)
        self.geometry("380x280")
        self.resizable(False, False)
        self.attributes("-topmost", True)

        # 居中显示
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 380) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 280) // 2
        self.geometry(f"+{x}+{y}")

        # 主容器
        self.frame = ctk.CTkFrame(self, fg_color="transparent")
        self.frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 版本号
        ctk.CTkLabel(
            self.frame,
            text="版本号",
            font=(config.FONT_FAMILY, 12, "bold"),
            text_color="gray",
            anchor="w"
        ).pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(
            self.frame,
            text=version,
            font=(config.FONT_FAMILY, 16, "bold"),
            anchor="w"
        ).pack(fill="x", pady=(0, 15))

        # 日期
        date_label = "安装日期" if "本地" in title else "发布日期"
        ctk.CTkLabel(
            self.frame,
            text=date_label,
            font=(config.FONT_FAMILY, 12, "bold"),
            text_color="gray",
            anchor="w"
        ).pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(
            self.frame,
            text=date,
            font=(config.FONT_FAMILY, 14),
            anchor="w"
        ).pack(fill="x", pady=(0, 15))

        # 更新日志（仅远程版本显示）
        if changelog:
            ctk.CTkLabel(
                self.frame,
                text="更新日志",
                font=(config.FONT_FAMILY, 12, "bold"),
                text_color="gray",
                anchor="w"
            ).pack(fill="x", pady=(0, 5))

            # 可滚动的日志区域
            self.changelog_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
            self.changelog_frame.pack(fill="both", expand=True, pady=(0, 15))

            self.changelog_text = ctk.CTkTextbox(
                self.changelog_frame,
                font=(config.FONT_FAMILY, 12),
                wrap="word",
                height=80
            )
            self.changelog_text.pack(fill="both", expand=True)
            self.changelog_text.insert("1.0", changelog)
            self.changelog_text.configure(state="disabled")

        # 关闭按钮
        ctk.CTkButton(
            self.frame,
            text="关闭",
            width=100,
            height=32,
            font=(config.FONT_FAMILY, 12),
            command=self.destroy
        ).pack(pady=(5, 0))

        # 绑定 ESC 键关闭
        self.bind("<Escape>", lambda e: self.destroy())

        # 禁用父窗口
        self.grab_set()


class App(ctk.CTk):
    """主应用窗口"""

    def __init__(self):
        super().__init__()

        # 窗口配置
        self.title(f"{config.APP_TITLE} v{config.APP_VERSION}")
        self.geometry(f"{config.WINDOW_WIDTH}x{config.WINDOW_HEIGHT}")
        self.resizable(True, True)
        self.minsize(config.WINDOW_MIN_WIDTH, config.WINDOW_MIN_HEIGHT)

        # 设置主题
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # 版本信息
        self.local_version_info = read_local_version()
        self.remote_version_info: Optional[dict] = None
        self.webdav_client: Optional[WebDAVClient] = None

        # 创建界面
        self._create_widgets()

    def _create_widgets(self):
        """创建界面组件"""
        # Header 区域
        self.header_frame = ctk.CTkFrame(self, height=40)
        self.header_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.header_frame.pack_propagate(False)

        # 标题
        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text=config.APP_TITLE,
            font=(config.FONT_FAMILY, 16, "bold")
        )
        self.title_label.pack(side="left", padx=10)

        # 右侧版本区
        self.version_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.version_frame.pack(side="right", padx=10)

        # 本地版本按钮
        local_ver = self.local_version_info.get("version", "未安装")
        self.local_ver_btn = ctk.CTkButton(
            self.version_frame,
            text=f"📦 本地: {local_ver}",
            width=120,
            height=28,
            font=(config.FONT_FAMILY, 11),
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray85", "gray25"),
            border_width=1,
            border_color=("gray70", "gray40"),
            command=self._show_local_version
        )
        self.local_ver_btn.pack(side="left", padx=(0, 10))

        # 远程版本按钮
        self.remote_ver_btn = ctk.CTkButton(
            self.version_frame,
            text="🌐 远程: 未检查",
            width=130,
            height=28,
            font=(config.FONT_FAMILY, 11),
            fg_color="transparent",
            text_color=("gray50", "gray60"),
            hover_color=("gray85", "gray25"),
            border_width=1,
            border_color=("gray70", "gray40"),
            command=self._show_remote_version
        )
        self.remote_ver_btn.pack(side="left", padx=(0, 10))

        # 检查更新按钮
        self.check_btn = ctk.CTkButton(
            self.version_frame,
            text="🔍 检查",
            width=70,
            height=28,
            font=(config.FONT_FAMILY, 11),
            command=self._on_check_update
        )
        self.check_btn.pack(side="left")

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
            install_callback=self._on_remote_install,
            version_update_callback=self._update_version_display
        )
        self.remote_install_tab.pack(fill="both", expand=True)

        self.local_install_tab = LocalInstallTab(
            self.local_tab,
            log_callback=self._log,
            install_callback=self._on_local_install
        )
        self.local_install_tab.pack(fill="both", expand=True)

    def _log(self, message: str):
        """输出日志（日志输出到 Tab 内的日志区）"""
        pass

    def _show_local_version(self):
        """显示本地版本弹窗"""
        local_ver = self.local_version_info.get("version", "未安装")
        local_date = self.local_version_info.get("releaseDate", "-")
        VersionDialog(self, "本地版本信息", local_ver, local_date)

    def _show_remote_version(self):
        """显示远程版本弹窗"""
        if self.remote_version_info:
            remote_ver = self.remote_version_info.get("version", "未知")
            remote_date = self.remote_version_info.get("releaseDate", "-")
            changelog = self.remote_version_info.get("changelog", "")
            VersionDialog(self, "远程版本信息", remote_ver, remote_date, changelog)
        else:
            VersionDialog(self, "远程版本信息", "未检查", "点击「检查」按钮获取")

    def _on_check_update(self):
        """检查更新按钮点击"""
        url = self.remote_install_tab.url_entry.get().strip()
        user = self.remote_install_tab.user_entry.get().strip()
        password = self.remote_install_tab.pass_entry.get()
        folder = config.WEBDAV_DEFAULT_FOLDER

        if not url:
            self.remote_install_tab._log("❌ 请输入服务器地址")
            return

        self.check_btn.configure(state="disabled")
        self.remote_install_tab._log("🔍 正在检查更新...")

        def check():
            try:
                client = WebDAVClient(url, user, password, folder)
                self.webdav_client = client
                self.remote_install_tab.webdav_client = client

                remote_info = read_remote_version(client)

                if not remote_info:
                    self.after(0, lambda: self.remote_install_tab._log("❌ 无法获取远程版本信息"))
                    return

                self.remote_version_info = remote_info

                remote_ver = remote_info.get("version", "未知")

                # 更新远程版本按钮
                self.after(0, lambda: self.remote_ver_btn.configure(
                    text=f"🌐 远程: {remote_ver}",
                    text_color=("gray10", "gray90")
                ))

                # 比较版本
                local_ver = self.local_version_info.get("version", "0")

                if local_ver == "未安装":
                    self.after(0, lambda: self.remote_install_tab._log(f"✅ 发现新版本：{remote_ver}（本地未安装）"))
                elif local_ver != "未知" and self._is_newer(remote_ver, local_ver):
                    self.after(0, lambda: self.remote_install_tab._log(f"✅ 发现新版本：{remote_ver}"))
                else:
                    self.after(0, lambda: self.remote_install_tab._log("✅ 已是最新版本"))

            except Exception as e:
                self.after(0, lambda: self.remote_install_tab._log(f"❌ 检查更新失败：{str(e)}"))
            finally:
                self.after(0, lambda: self.check_btn.configure(state="normal"))

        threading.Thread(target=check, daemon=True).start()

    def _is_newer(self, remote_ver: str, local_ver: str) -> bool:
        """比较版本号"""
        try:
            from packaging.version import Version
            return Version(remote_ver) > Version(local_ver)
        except Exception:
            return False

    def _on_remote_install(self, xlam_bytes: bytes, version_info: Optional[dict] = None):
        """远程安装回调"""
        def install():
            deploy(xlam_bytes, self.remote_install_tab._log, version_info)
            self.after(0, self._update_version_display)

        threading.Thread(target=install, daemon=True).start()

    def _on_local_install(self, xlam_bytes: bytes):
        """本地安装回调"""
        def install():
            deploy(xlam_bytes, self.local_install_tab._log)
            self.after(0, self._update_version_display)

        threading.Thread(target=install, daemon=True).start()

    def _update_version_display(self):
        """更新版本显示"""
        self.local_version_info = read_local_version()
        local_ver = self.local_version_info.get("version", "未安装")
        self.local_ver_btn.configure(text=f"📦 本地: {local_ver}")