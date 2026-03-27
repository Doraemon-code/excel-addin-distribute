# -*- coding: utf-8 -*-
"""
Excel 加载项安装工具
入口文件
"""

import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from ui.app import App


def main():
    """程序入口"""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()