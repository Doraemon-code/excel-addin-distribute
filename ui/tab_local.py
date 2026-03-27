# -*- coding: utf-8 -*-
"""
本地安装 Tab
提供从本地文件安装 xlam 的功能
"""

import customtkinter as ctk
from tkinter import filedialog
from typing import Callable

from utils.config import config


class LocalInstallTab(ctk.CTkFrame):
    """本地安装 Tab 页面"""

    def __init__(
        self,
        master,
        log_callback: Callable[[str], None],
        install_callback: Callable[[bytes], None]
    ):
        """
        初始化本地安装 Tab

        Args:
            master: 父组件
            log_callback: 日志回调函数
            install_callback: 安装回调函数，接收文件二进制内容
        """
        super().__init__(master)

        self.install = install_callback
        self.selected_file = None

        self._create_widgets()

    def _create_widgets(self):
        """创建界面组件"""
        # 使用 grid 布局确保按钮常驻
        self.grid_rowconfigure(0, weight=0)  # 文件选择区
        self.grid_rowconfigure(1, weight=1)  # 日志区（可伸缩）
        self.grid_rowconfigure(2, weight=0)  # 操作区（固定）
        self.grid_columnconfigure(0, weight=1)

        # 文件选择区
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=15)

        self.info_label = ctk.CTkLabel(
            self.file_frame,
            text="选择本地的 .xlam 文件进行安装",
            font=(config.FONT_FAMILY, 12)
        )
        self.info_label.pack(pady=(10, 10))

        select_frame = ctk.CTkFrame(self.file_frame, fg_color="transparent")
        select_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.file_entry = ctk.CTkEntry(
            select_frame,
            placeholder_text="未选择文件",
            state="disabled",
            height=32
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_btn = ctk.CTkButton(
            select_frame,
            text="📂 浏览",
            width=80,
            height=32,
            font=(config.FONT_FAMILY, 11),
            command=self._on_browse
        )
        self.browse_btn.pack(side="right")

        # 操作区（固定在底部）
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(5, 10))

        self.install_btn = ctk.CTkButton(
            self.action_frame,
            text="⬇️ 安装",
            width=200,
            height=36,
            font=(config.FONT_FAMILY, 14),
            command=self._on_install,
            state="disabled"
        )
        self.install_btn.pack(pady=10)

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
            height=60
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

    def _log(self, message: str):
        """输出日志"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_browse(self):
        """浏览按钮点击事件"""
        addin_dir = str(config.ADDIN_DIR)
        if not config.ADDIN_DIR.exists():
            addin_dir = None

        file_path = filedialog.askopenfilename(
            title="选择 Excel 加载项文件",
            initialdir=addin_dir,
            filetypes=[("Excel 加载项", "*.xlam"), ("所有文件", "*.*")]
        )

        if file_path:
            self.selected_file = file_path
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)
            self.file_entry.configure(state="disabled")
            self.install_btn.configure(state="normal")

    def _on_install(self):
        """安装按钮点击事件"""
        if not self.selected_file:
            self._log("❌ 请先选择文件")
            return

        try:
            self.install_btn.configure(state="disabled", text="安装中...")
            self.browse_btn.configure(state="disabled")

            with open(self.selected_file, "rb") as f:
                xlam_bytes = f.read()

            self._log(f"📄 已读取文件：{self.selected_file}")
            self.install(xlam_bytes)

        except Exception as e:
            self._log(f"❌ 读取文件失败：{str(e)}")
        finally:
            self.install_btn.configure(state="normal", text="⬇️ 安装")
            self.browse_btn.configure(state="normal")

    def reset(self):
        """重置状态"""
        self.selected_file = None
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.configure(state="disabled")
        self.install_btn.configure(state="disabled")