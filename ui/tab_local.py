# -*- coding: utf-8 -*-
"""
本地安装 Tab
提供从本地文件安装 xlam 的功能
"""

import customtkinter as ctk
from tkinter import filedialog
from typing import Callable

from config import XLAM_FILENAME


class LocalInstallTab(ctk.CTkFrame):
    """本地安装 Tab 页面"""

    def __init__(self, master, log_callback: Callable[[str], None], install_callback: Callable[[bytes], None]):
        """
        初始化本地安装 Tab

        Args:
            master: 父组件
            log_callback: 日志回调函数
            install_callback: 安装回调函数，接收文件二进制内容
        """
        super().__init__(master)

        self.log = log_callback
        self.install = install_callback
        self.selected_file = None

        self._create_widgets()

    def _create_widgets(self):
        """创建界面组件"""
        # 说明文字
        self.info_label = ctk.CTkLabel(
            self,
            text="选择本地的 .xlam 文件，工具将自动完成安装",
            font=("", 14)
        )
        self.info_label.pack(pady=(40, 20))

        # 文件选择区
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(fill="x", padx=40, pady=10)

        self.file_entry = ctk.CTkEntry(
            self.file_frame,
            placeholder_text="未选择文件",
            state="disabled",
            height=36
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_btn = ctk.CTkButton(
            self.file_frame,
            text="📂 浏览",
            width=100,
            command=self._on_browse
        )
        self.browse_btn.pack(side="right")

        # 安装按钮
        self.install_btn = ctk.CTkButton(
            self,
            text="⬇️ 安装",
            width=200,
            height=40,
            font=("", 16),
            command=self._on_install,
            state="disabled"
        )
        self.install_btn.pack(pady=40)

    def _on_browse(self):
        """浏览按钮点击事件"""
        file_path = filedialog.askopenfilename(
            title="选择 Excel 加载项文件",
            filetypes=[("Excel 加载项", "*.xlam"), ("所有文件", "*.*")]
        )

        if file_path:
            self.selected_file = file_path
            # 更新显示
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)
            self.file_entry.configure(state="disabled")
            # 启用安装按钮
            self.install_btn.configure(state="normal")

    def _on_install(self):
        """安装按钮点击事件"""
        if not self.selected_file:
            self.log("❌ 请先选择文件")
            return

        try:
            # 禁用按钮
            self.install_btn.configure(state="disabled", text="安装中...")
            self.browse_btn.configure(state="disabled")

            # 读取文件
            with open(self.selected_file, "rb") as f:
                xlam_bytes = f.read()

            self.log(f"📄 已读取文件：{self.selected_file}")

            # 调用安装回调
            self.install(xlam_bytes)

        except Exception as e:
            self.log(f"❌ 读取文件失败：{str(e)}")
        finally:
            # 恢复按钮
            self.install_btn.configure(state="normal", text="⬇️ 安装")
            self.browse_btn.configure(state="normal")

    def reset(self):
        """重置状态"""
        self.selected_file = None
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.configure(state="disabled")
        self.install_btn.configure(state="disabled")