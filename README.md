<div align="center">

# Excel 加载项安装工具

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)]()

Windows 桌面工具，用于安装和更新 Excel 加载项（.xlam 文件），支持 WebDAV 远程部署

**[English](README.en.md)** | 简体中文

</div>

---

## 功能特性

- **远程安装** — 通过 WebDAV 下载并安装
- **本地安装** — 从本地 .xlam 文件安装
- **自动注册配置** — 自动完成 Excel 加载项注册表设置
- **VBA 版本管理** — Git Hooks 自动导出 VBA 代码用于版本控制

---

## 操作流程

### 管理员（发布端）

#### 1. 构建 xlam 文件（含 Ribbon UI）

使用 [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) 创建加载项：

1. 打开 Excel → 新建空白工作簿
2. `Alt + F11` → 插入模块 → 粘贴 VBA 代码
3. 另存为 `MyAddin.xlam`（Excel 加载项格式）
4. 用 **Office RibbonX Editor** 打开 xlam
5. Insert → Office 2010+ Custom UI Part
6. 粘贴 CustomUI XML → Validate → Save

> **重要：** Ribbon 回调函数必须包含 `IRibbonControl` 参数：
> ```vba
> Public Sub MyCallback(control As IRibbonControl)
>     ' 代码
> End Sub
> ```

> **正确编辑 xlam：**
> - ❌ 不要双击打开（加载项模式下 Ctrl+S 不保存代码）
> - ✅ 完全关闭 Excel → 打开空白工作簿 → 文件 → 打开 → 选择 xlam → 编辑 → 保存

#### 2. 部署到 WebDAV

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

#### 3. 分发安装程序

将 `ExcelAddinInstaller.exe` 发给用户。

---

### 用户（安装端）

1. 运行 `ExcelAddinInstaller.exe`
2. 选择安装方式：
   - **远程安装**：填写 WebDAV 地址/账号（如有默认值可直接使用）
   - **本地安装**：选择本地 .xlam 文件
3. （远程）点击「检查更新」→ 确认新版本 → 点击「一键安装」
4. **完全重启 Excel**，检查是否出现 Ribbon 标签页

---

## WebDAV 配置注意事项

| 项目 | 说明 |
|------|------|
| 版本号必须递增 | 否则客户端检测不到更新 |
| 文件名必须一致 | `version.json` 的 filename 字段必须与实际文件名一致 |
| 目录名保持一致 | 默认使用 `Addin/` 目录（可通过 `webdav_default_folder` 配置） |
| 资源可访问 | 确保 WebDAV 资源可下载（URL、鉴权、权限、证书等） |

---

## 配置说明

| 文件 | 位置 | 优先级 |
|------|------|--------|
| `config.yaml` | 内嵌于 exe | 默认值 |
| `config.json` | `%APPDATA%\ExcelAddinInstaller\` | 用户配置（更高） |

---

## 开发运行

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
  --icon icon/app.ico \
  main.py
```

### Git Hooks（VBA 版本管理）

```bash
# Windows
scripts\install_hook.bat

# Unix/Linux/macOS
sh scripts/install_hook.sh
```

pre-commit 钩子会在提交前自动导出 `.xlam` 中的 VBA 代码到 `src/vba/`，实现版本控制。

---

## 许可证

[MIT License](LICENSE)