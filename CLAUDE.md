# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Excel 加载项部署工具 —— 用于自动部署和更新 Excel 加载项（.xlam 文件）的 Windows 桌面工具，包含 Git Hooks 实现 VBA 代码版本管理。

## 常用命令

### 运行与开发

```bash
# 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

### 打包

```bash
pyinstaller --onefile --windowed \
  --name ExcelAddinInstaller \
  --icon app.ico \
  --add-data "config.yaml;." \
  main.py
```

### Git Hooks（VBA 版本管理）

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh

# 手动导出 VBA 代码
python scripts/export_vba.py                     # 自动查找 .xlam
python scripts/export_vba.py MyAddin.xlam        # 指定文件
python scripts/export_vba.py MyAddin.xlam src/vba  # 指定输出目录
```

## 架构

```
main.py ──→ ui/app.py (主窗口)
              ├── ui/tab_remote.py (远程安装 Tab)
              └── ui/tab_local.py (本地安装 Tab)

core/
├── deployer.py    # 核心部署：复制文件 + 注册表配置
├── webdav_client.py  # WebDAV 文件下载
└── version.py     # 版本读取与比较

utils/config.py   # 配置加载（支持 PyInstaller 内嵌）
```

### 关键设计

1. **配置分离**：
   - `config.yaml`：内嵌配置（打包后不可见），包含 WebDAV 默认值
   - `%APPDATA%\ExcelAddinInstaller\config.json`：用户配置，优先级更高

2. **部署流程**（`core/deployer.py`）：
   - 目标目录：`%APPDATA%\Microsoft\AddIns`
   - 备份原文件 → 写入新文件 → 注册表配置 → 保存 version.json

3. **UI 线程模型**：
   - 所有网络/IO 操作在后台线程执行
   - 使用 `self.after(0, callback)` 回调主线程更新 UI

4. **PyInstaller 支持**：
   - `utils/config.py` 的 `get_base_path()` 检测打包环境
   - 打包后从 `sys._MEIPASS` 读取内嵌的 `config.yaml`

## 配置

### config.yaml

```yaml
webdav_default_url: ""
webdav_default_user: ""
webdav_default_pass: ""
webdav_default_folder: "Addin"
xlam_filename: "MyAddin.xlam"
version_filename: "version.json"
app_title: "Excel 加载项安装工具"
app_version: "1.0.0"
window_width: 680
window_height: 420
font_family: "Microsoft YaHei UI"
```

### export_vba.py

- `SKIP_MODULES`：跳过 `ThisWorkbook` 和 `Sheet` 开头的内置模块
- `EXT_MAP`：VBA 模块类型 → 文件扩展名映射
- 输出目录：`src/vba/`

### pre-commit hook

- 仅当暂存区包含 `.xlam` 文件时触发
- 导出失败仅警告，不阻断提交
- 导出文件自动 `git add`