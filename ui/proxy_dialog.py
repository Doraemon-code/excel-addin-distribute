# -*- coding: utf-8 -*-
"""
代理设置对话框
提供代理配置的 UI 界面
"""

import customtkinter as ctk

from utils.config import config


class ProxyDialog(ctk.CTkToplevel):
    """代理设置弹窗"""

    def __init__(self, master, on_save_callback=None):
        """
        初始化代理设置对话框

        Args:
            master: 父窗口
            on_save_callback: 保存回调函数
        """
        super().__init__(master)

        self.on_save = on_save_callback
        self.result = None

        self.title("代理设置")
        self.geometry("320x320")
        self.resizable(False, False)

        # 居中显示
        self.transient(master)
        self.grab_set()

        self._create_widgets()
        self._load_config()

        # 等待窗口关闭
        self.wait_window()

    def _create_widgets(self):
        """创建界面组件"""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # 启用代理开关
        self.enabled_switch = ctk.CTkSwitch(
            main_frame,
            text="启用代理",
            font=(config.FONT_FAMILY, 12)
        )
        self.enabled_switch.pack(anchor="w", pady=(0, 10))

        # 代理类型
        type_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            type_frame,
            text="代理类型:",
            font=(config.FONT_FAMILY, 11)
        ).pack(side="left")

        self.type_menu = ctk.CTkOptionMenu(
            type_frame,
            values=["http", "socks5"],
            width=100,
            font=(config.FONT_FAMILY, 11)
        )
        self.type_menu.pack(side="left", padx=(10, 0))

        # 代理 IP
        host_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        host_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            host_frame,
            text="代理 IP:",
            font=(config.FONT_FAMILY, 11),
            width=70
        ).pack(side="left")

        self.host_entry = ctk.CTkEntry(
            host_frame,
            font=(config.FONT_FAMILY, 11)
        )
        self.host_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # 代理端口
        port_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        port_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            port_frame,
            text="代理端口:",
            font=(config.FONT_FAMILY, 11),
            width=70
        ).pack(side="left")

        self.port_entry = ctk.CTkEntry(
            port_frame,
            font=(config.FONT_FAMILY, 11)
        )
        self.port_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # 代理用户名（可选）
        user_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        user_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            user_frame,
            text="用户名:",
            font=(config.FONT_FAMILY, 11),
            width=70
        ).pack(side="left")

        self.user_entry = ctk.CTkEntry(
            user_frame,
            font=(config.FONT_FAMILY, 11),
            placeholder_text="（可选）"
        )
        self.user_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # 代理密码（可选）
        pass_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        pass_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(
            pass_frame,
            text="密码:",
            font=(config.FONT_FAMILY, 11),
            width=70
        ).pack(side="left")

        self.pass_entry = ctk.CTkEntry(
            pass_frame,
            show="*",
            font=(config.FONT_FAMILY, 11),
            placeholder_text="（可选）"
        )
        self.pass_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # 按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 0))

        self.save_btn = ctk.CTkButton(
            btn_frame,
            text="保存",
            width=80,
            font=(config.FONT_FAMILY, 11),
            command=self._on_save
        )
        self.save_btn.pack(side="left", padx=(0, 10))

        self.cancel_btn = ctk.CTkButton(
            btn_frame,
            text="取消",
            width=80,
            fg_color="gray",
            hover_color="darkgray",
            font=(config.FONT_FAMILY, 11),
            command=self._on_cancel
        )
        self.cancel_btn.pack(side="left")

    def _load_config(self):
        """加载当前配置"""
        # 尝试从用户配置文件加载
        import json

        proxy_config = {
            "enabled": config.PROXY_ENABLED,
            "type": config.PROXY_TYPE,
            "host": config.PROXY_HOST,
            "port": config.PROXY_PORT,
            "user": config.PROXY_USER,
            "pass": config.PROXY_PASS,
        }

        if config.USER_CONFIG_FILE.exists():
            try:
                with open(config.USER_CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    proxy_config["enabled"] = saved.get("proxy_enabled", proxy_config["enabled"])
                    proxy_config["type"] = saved.get("proxy_type", proxy_config["type"])
                    proxy_config["host"] = saved.get("proxy_host", proxy_config["host"])
                    proxy_config["port"] = saved.get("proxy_port", proxy_config["port"])
                    proxy_config["user"] = saved.get("proxy_user", proxy_config["user"])
                    proxy_config["pass"] = saved.get("proxy_pass", proxy_config["pass"])
            except Exception:
                pass

        # 设置 UI 值
        if proxy_config["enabled"]:
            self.enabled_switch.select()
        else:
            self.enabled_switch.deselect()

        self.type_menu.set(proxy_config["type"])
        self.host_entry.insert(0, str(proxy_config["host"]))
        self.port_entry.insert(0, str(proxy_config["port"]))
        self.user_entry.insert(0, proxy_config["user"])
        self.pass_entry.insert(0, proxy_config["pass"])

    def _on_save(self):
        """保存按钮点击"""
        self.result = {
            "proxy_enabled": self.enabled_switch.get(),
            "proxy_type": self.type_menu.get(),
            "proxy_host": self.host_entry.get().strip(),
            "proxy_port": int(self.port_entry.get().strip() or 10810),
            "proxy_user": self.user_entry.get().strip(),
            "proxy_pass": self.pass_entry.get(),
        }

        if self.on_save:
            self.on_save(self.result)

        self.destroy()

    def _on_cancel(self):
        """取消按钮点击"""
        self.result = None
        self.destroy()