# Excel Addin Installer

[English](#english) | [中文](#中文)

---

<a name="english"></a>

## English

A Windows desktop tool for installing and updating Excel add-ins (.xlam files).

### Features

- **Remote Installation** — Download and install add-ins via WebDAV
- **Local Installation** — Install from local .xlam files
- **Auto Registry Config** — Automatically configure Excel add-in registry entries
- **VBA Version Control** — Optional Git hooks to export VBA code for versioning

### Quick Start

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```

### Build

```bash
pyinstaller --onefile --windowed \
  --name ExcelAddinInstaller \
  --icon app.ico \
  --add-data "config.yaml;." \
  main.py
```

### Configuration

| File | Location | Priority |
|------|----------|----------|
| `config.yaml` | Embedded in exe | Default values |
| `config.json` | `%APPDATA%\ExcelAddinInstaller\` | User overrides |

### WebDAV Deployment

Upload to your WebDAV server:

```
Addin/
├── MyAddin.xlam
└── version.json
```

**version.json format:**

```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "Release notes"
}
```

### Git Hooks (VBA Version Control)

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh
```

**What it does:**

The pre-commit hook automatically exports VBA code from `.xlam` files before each commit:

1. Detects if any `.xlam` files are staged
2. Exports VBA modules (`.bas`, `.cls`) to `src/vba/`
3. Adds exported files to the commit automatically

This enables version control for VBA code inside binary `.xlam` files.

### License

MIT

---

<a name="中文"></a>

## 中文

Windows 桌面工具，用于安装和更新 Excel 加载项（.xlam 文件）。

### 功能特性

- **远程安装** — 通过 WebDAV 下载并安装加载项
- **本地安装** — 从本地 .xlam 文件安装
- **自动注册配置** — 自动完成 Excel 加载项注册表配置
- **VBA 版本管理** — 可选的 Git Hooks，自动导出 VBA 代码用于版本控制

### 快速开始

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

### 配置说明

| 文件 | 位置 | 优先级 |
|------|------|--------|
| `config.yaml` | 内嵌于 exe | 默认值 |
| `config.json` | `%APPDATA%\ExcelAddinInstaller\` | 用户配置（更高） |

### WebDAV 部署

上传到 WebDAV 服务器：

```
Addin/
├── MyAddin.xlam
└── version.json
```

**version.json 格式：**

```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "更新说明"
}
```

### Git Hooks（VBA 版本管理）

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh
```

**作用说明：**

pre-commit 钩子会在每次提交前自动导出 `.xlam` 文件中的 VBA 代码：

1. 检测暂存区是否有 `.xlam` 文件变更
2. 自动导出 VBA 模块（`.bas`、`.cls`）到 `src/vba/`
3. 将导出文件自动加入本次提交

这样就能对二进制 `.xlam` 文件内的 VBA 代码进行版本控制。

### 开发加载项

创建 Excel 加载项的示例文件位于 `example/` 目录：

| 文件 | 说明 |
|------|------|
| `customUI14.xml` | Ribbon 工具栏定义（Office 2010+） |
| `RibbonCallbacks.bas` | VBA 回调函数示例 |

详细步骤请参考 [`example/README.md`](example/README.md)。

### 许可证

MIT