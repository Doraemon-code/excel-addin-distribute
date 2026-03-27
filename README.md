<div align="center">

# Excel 加载项安装工具

Windows 桌面工具，用于自动部署和更新 Excel 加载项（.xlam 文件）

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

</div>

---

## ✨ 功能特性

- 🌐 **远程安装** - 从 WebDAV 服务器下载并安装
- 📁 **本地安装** - 从本地文件安装，自动定位 Addin 文件夹
- 🔄 **版本检测** - 自动检测远程版本，支持一键更新
- 📝 **注册表配置** - 自动配置 Office Excel 加载项
- 🔧 **VBA 版本管理** - Git Hooks 自动导出 VBA 代码

---

## 🚀 快速开始

### 环境要求

- Python 3.10+
- Windows 操作系统

### 安装运行

```bash
# 克隆仓库
git clone https://github.com/your-repo/ExcelAddinInstaller.git
cd ExcelAddinInstaller

# 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

---

## ⚙️ 配置说明

### config.yaml

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `webdav_default_url` | WebDAV 服务器地址 | 内嵌配置 |
| `webdav_default_user` | WebDAV 用户名 | 内嵌配置 |
| `webdav_default_pass` | WebDAV 密码 | 内嵌配置 |
| `webdav_default_folder` | WebDAV 文件夹 | `Addin` |
| `xlam_filename` | 加载项文件名 | `MyAddin.xlam` |
| `version_filename` | 版本信息文件 | `version.json` |
| `app_title` | 应用标题 | Excel 加载项安装工具 |
| `app_version` | 应用版本 | 1.0.0 |

### 用户配置

用户自定义配置保存在：`%APPDATA%\ExcelAddinInstaller\config.json`

---

## 📦 WebDAV 部署

### 服务器目录结构

```
WebDAV/
└── Addin/
    ├── MyAddin.xlam    # 加载项主文件
    └── version.json    # 版本信息
```

### version.json 格式

```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "初始版本发布"
}
```

| 字段 | 说明 |
|------|------|
| `version` | 版本号（语义化版本） |
| `releaseDate` | 发布日期 |
| `filename` | 文件名（需与实际一致） |
| `changelog` | 更新日志 |

---

## 📖 使用指南

### 管理员流程

```mermaid
graph LR
    A[开发 xlam] --> B[更新版本号]
    B --> C[上传 WebDAV]
    C --> D[分发 exe]
```

1. 开发 xlam 文件和 VBA 代码
2. 更新 `version.json` 中的版本号
3. 上传文件到 WebDAV 服务器
4. 分发 `ExcelAddinInstaller.exe` 给用户

### 用户流程

1. 运行 `ExcelAddinInstaller.exe`
2. 点击版本按钮检查更新
3. 点击「一键安装」
4. 重启 Excel

---

## 🔧 开发指南

### 打包命令

```bash
pyinstaller --onefile --windowed \
  --name ExcelAddinInstaller \
  --icon app.ico \
  --add-data "config.yaml;." \
  main.py
```

> 打包后 config.yaml 内嵌在 exe 中，用户无法查看默认配置。

### Git Hooks（VBA 版本管理）

在 `git commit` 前自动导出 xlam 中的 VBA 代码：

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh
```

#### 手动导出 VBA

```bash
# 自动查找 xlam 文件
python scripts/export_vba.py

# 指定文件
python scripts/export_vba.py MyAddin.xlam

# 指定输出目录
python scripts/export_vba.py MyAddin.xlam src/vba
```

---

## 📁 项目结构

```
ExcelAddinInstaller/
├── main.py                  # 入口文件
├── config.yaml              # 配置文件
├── requirements.txt         # 依赖列表
├── ui/
│   ├── app.py               # 主窗口
│   ├── tab_remote.py        # 远程安装 Tab
│   └── tab_local.py         # 本地安装 Tab
├── core/
│   ├── deployer.py          # 核心部署逻辑
│   ├── webdav_client.py     # WebDAV 客户端
│   └── version.py           # 版本比较
├── utils/
│   └── config.py            # 配置加载
├── hooks/
│   ├── pre-commit           # Git 钩子
│   └── hooks-config.json    # 钩子配置
├── scripts/
│   ├── export_vba.py        # VBA 导出工具
│   ├── install_hook.sh      # Unix 安装脚本
│   └── install_hook.bat     # Windows 安装脚本
├── example/                 # 开发示例
│   ├── README.md
│   ├── customUI14.xml       # Ribbon 示例
│   └── RibbonCallbacks.bas  # VBA 回调示例
└── src/vba/                 # VBA 导出目录
```

---

## ❓ 常见问题

| 现象 | 原因 | 解决方法 |
|------|------|----------|
| 点击按钮报"错误的参数号" | VBA 函数签名不正确 | 函数必须包含 `control As IRibbonControl` 参数 |
| 修改 xlam 后代码被还原 | Excel 退出时覆盖文件 | 用 Excel 菜单"打开"方式编辑，而非双击 |
| 安装后 Excel 没有新标签页 | Excel 未完全重启 | 完全关闭 Excel（含任务栏）后重新打开 |
| Git Hooks 未触发 | hooks 未正确安装 | 手动复制 `hooks/pre-commit` 到 `.git/hooks/` |

---

## 📄 许可证

[MIT License](LICENSE)

---

<div align="center">

Made with ❤️ by [Your Name]

</div>