# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Excel 加载项部署工具 —— 用于自动部署和更新 Excel 加载项（.xlam 文件）的 Windows 桌面工具，包含 Git Hooks 实现 VBA 代码版本管理。

## 项目结构

```
.
├── main.py                  # 入口文件
├── config.yaml              # 配置文件
├── ui/
│   ├── app.py               # 主窗口
│   ├── tab_remote.py        # 远程安装 Tab
│   └── tab_local.py         # 本地安装 Tab
├── core/
│   ├── deployer.py          # 核心部署逻辑
│   ├── webdav_client.py     # WebDAV 客户端
│   └── version.py           # 版本比较
├── utils/
│   └── config.py            # 配置加载模块
├── hooks/
│   ├── pre-commit           # Git 提交前钩子
│   └── hooks-config.json    # 钩子配置
├── scripts/
│   ├── export_vba.py        # VBA 代码导出工具
│   ├── install_hook.sh      # Unix 版 hook 安装脚本
│   └── install_hook.bat     # Windows 版 hook 安装脚本
├── example/                 # 开发示例
│   ├── customUI14.xml       # Ribbon 定义示例
│   └── RibbonCallbacks.bas  # VBA 回调函数示例
└── src/vba/                 # VBA 代码导出目录（自动生成）
```

## 核心功能

### 1. Excel 加载项安装工具（已实现）

技术栈：
- UI: customtkinter
- WebDAV: webdavclient3
- HTTP: requests
- 打包：PyInstaller

功能：
- 远程安装（WebDAV 下载）
- 本地安装（文件选择，默认定位到 Addin 文件夹）
- 版本检测与更新
- Office 注册表配置

配置文件：`config.yaml`

### 2. VBA 代码版本管理（已实现）

使用 Git Hooks 在 `git commit` 前自动导出 xlam 文件中的 VBA 代码到文本文件，实现代码版本追踪。

**工作流程：**
```
git add MyAddin.xlam → git commit → pre-commit 触发 → 导出 .bas/.cls 文件 → 自动加入暂存区
```

**依赖：** `pip install oletools`

## 常用命令

### Hook 安装

```bash
# Unix/Linux/macOS
sh scripts/install_hook.sh

# Windows
scripts\install_hook.bat
```

### 手动导出 VBA 代码

```bash
# 自动查找根目录下的 .xlam 文件
python scripts/export_vba.py

# 指定 xlam 文件路径
python scripts/export_vba.py MyAddin.xlam

# 指定输出目录
python scripts/export_vba.py MyAddin.xlam src/vba
```

### 检查 VBA 导出状态

```bash
# 查看已导出的模块
ls src/vba/

# 查看导出清单
cat src/vba/manifest.txt
```

## 配置说明

### config.yaml

所有配置统一在 `config.yaml` 中管理：

```yaml
webdav_default_url: ""
webdav_default_user: ""
webdav_default_pass: ""
xlam_filename: "MyAddin.xlam"
version_filename: "version.json"
app_title: "Excel 加载项安装工具"
app_version: "1.0.0"
window_width: 680
window_height: 560
```

通过 `utils/config.py` 加载，提供 `config` 对象访问配置。

### export_vba.py 配置

- `SKIP_MODULES`: 内置模块黑名单（`ThisWorkbook`, `Sheet` 开头）
- `EXT_MAP`: VBA 模块类型 → 文件扩展名映射（`.bas`, `.cls`, `.frm`）
- 默认输出目录：`src/vba/`

### pre-commit hook

- 仅当暂存区包含 `.xlam` 文件变更时触发
- 导出失败不阻断提交，仅输出警告
- 导出文件自动 `git add`

## 文件说明

| 文件 | 用途 |
|------|------|
| `hooks/pre-commit` | Git 钩子，调用 export_vba.py |
| `scripts/export_vba.py` | 使用 oletools 解析 xlam 并导出 VBA 代码 |
| `scripts/install_hook.sh` | Unix 版 hook 安装脚本 |
| `scripts/install_hook.bat` | Windows 版 hook 安装脚本 |
