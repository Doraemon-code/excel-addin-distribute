# Excel 加载项分发与部署流程

## 角色说明

| 角色 | 职责 |
|------|------|
| **管理员** | 开发维护 xlam、管理 WebDAV 服务器、分发 exe、执行版本管理 hooks |
| **用户** | 运行 exe 完成一键安装，后续检查更新 |

---

## 👨‍💼 管理员部分

### 第一步：开发 xlam 文件

#### 1.1 编写 VBA 代码

1. 新建或打开 `MyAddin.xlam`
2. 按 `Alt + F11` 打开 VBA 编辑器
3. 插入模块，编写所有宏函数，注意每个函数必须是 `Public Sub` 且第一个参数为 `IRibbonControl`：

```vba
Public Sub GenerateIndex(control As IRibbonControl)
    ' 生成目录逻辑
End Sub

Public Sub FontFormat(control As IRibbonControl)
    ' 字体格式化逻辑
End Sub
```

4. 保存并关闭 Excel

#### 1.2 配置工具栏 XML

1. 用 **Office RibbonX Editor** 打开 `MyAddin.xlam`
2. 点击 `Insert` → `Insert Office 2010+ Custom UI Part`
3. 将 CustomUI XML 粘贴到编辑区，注意：
   - 根节点必须是 `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">`
   - `onAction` 只写函数名，不写路径
4. 点击 `Validate` 确认无报错
5. 点击 `Save` 保存

#### 1.3 维护 version.json

每次更新 xlam 后，同步更新 WebDAV 上的 `version.json`：

```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "新增正则提取功能，修复字体格式化问题"
}
```

> ⚠️ version 字段使用语义化版本号（major.minor.patch），每次发布必须递增，否则用户端检测不到更新。

### 第二步：上传到 WebDAV 服务器

WebDAV 目录结构如下：

```
WebDAV 根目录/
├── MyAddin.xlam       ← 加载项主文件
└── version.json       ← 版本信息文件（必须同步更新）
```

### 第三步：分发 exe 工具

将 `ExcelAddinInstaller.exe` 通过以下任意方式分发给用户：

- 局域网共享文件夹
- 企业内网网盘
- 企业邮件附件

### 第四步：后续更新流程

```
修改 VBA 代码 或 工具栏 XML
       ↓
用 RibbonX Editor 保存 xlam
       ↓
更新 version.json 中的版本号和 changelog
       ↓
将 xlam 和 version.json 上传覆盖 WebDAV 旧文件
       ↓
通知用户打开 exe → 点击「检查更新」
```

---

## 👤 用户部分

### 首次安装（只需操作一次）

1. 双击运行管理员提供的 `ExcelAddinInstaller.exe`（无需管理员权限）
2. 切换到「远程安装」Tab：
   - 填写管理员提供的 WebDAV 服务器地址、用户名、密码
   - 点击「测试连接」确认可以连通
3. 点击「一键安装」，等待进度条完成
4. 看到「✅ 安装完成，请重启 Excel」提示后，关闭并重新打开 Excel
5. Excel 顶部出现新的标签页，安装完成 ✅

> 如果没有网络或管理员提供了本地 xlam 文件，
> 切换到「本地安装」Tab，浏览选择文件后点击「安装」即可。

### 后续更新

1. 打开 `ExcelAddinInstaller.exe`
2. 点击「检查更新」
3. 有新版本时点击「立即更新」，重启 Excel 生效
4. 提示「已是最新版本」则无需任何操作

### 常见问题

| 现象 | 原因 | 解决方法 |
|------|------|----------|
| 安装后 Excel 没有新标签页 | Excel 未完全重启 | 完全关闭 Excel（含任务栏）后重新打开 |
| 点击按钮提示"错误的参数号或无效的属性赋值" | VBA 函数签名不正确 | 函数必须包含 `control As IRibbonControl` 参数 |
| 点击按钮提示找不到宏 | 加载项未正确加载 | 重新运行 exe 再安装一次 |
| 修改 xlam 后代码被还原 | Excel 退出时覆盖文件 | 用 Excel 菜单"打开"方式编辑，而非双击文件 |
| 测试连接失败 | 网络或地址问题 | 确认当前网络可访问公司内网，检查地址是否正确 |
| 更新后旧功能不见了 | 新版 xlam 有问题 | 联系管理员 |

> **重要**：所有 Ribbon 按钮回调函数必须包含 `IRibbonControl` 参数：
> ```vba
> ' 正确
> Public Sub GenerateIndex(control As IRibbonControl)
>     ' 代码
> End Sub
>
> ' 错误
> Public Sub GenerateIndex()
>     ' 代码
> End Sub
> ```

### 开发示例

参见 `example/` 目录，包含：
- `customUI14.xml` - Ribbon 工具栏定义示例
- `RibbonCallbacks.bas` - VBA 回调函数模块示例
- `README.md` - 开发指南

---

## 🔧 Git Hooks 设计（VBA 代码版本管理）

### 背景

VBA 代码存储在 xlam 文件中，无法直接用 git 进行版本控制。需要通过 git hooks 在提交前自动导出 VBA 代码到文本文件，实现代码的版本管理。

### 目录结构

```
ExcelAddinInstaller/
├── .git/
│   └── hooks/
│       ├── pre-commit         ← 提交前钩子
│       └── hooks-config.json  ← hooks 配置文件
├── vba-src/                   ← VBA 代码导出目录
│   ├── modules/               ← 标准模块
│   ├── classes/               ← 类模块
│   ├── sheets/                ← 工作表模块
│   └── ThisWorkbook.bas       ← 工作簿模块
└── MyAddin.xlam               ← 加载项文件
```

### hooks-config.json 配置

```json
{
  "xlam_path": "./MyAddin.xlam",
  "export_dir": "./vba-src",
  "export_format": "bas",
  "include_metadata": true,
  "auto_commit": false
}
```

### pre-commit 钩子功能

1. **检测 xlam 文件变更**：检查是否有 xlam 文件被修改
2. **导出 VBA 代码**：使用 VBA-Exporter 工具导出所有模块到文本文件
3. **暂存导出文件**：将导出的代码文件自动加入暂存区
4. **版本比对**：生成代码 diff，便于 review

### pre-commit 脚本实现

```bash
#!/bin/bash

# Excel VBA 代码导出钩子
# 用途：在提交 xlam 文件前自动导出 VBA 代码到文本文件

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(dirname "$(dirname "$SCRIPT_DIR")")"
CONFIG_FILE="$SCRIPT_DIR/hooks-config.json"

# 读取配置
XLAM_PATH=$(jq -r '.xlam_path' "$CONFIG_FILE")
EXPORT_DIR=$(jq -r '.export_dir' "$CONFIG_FILE")

echo "🔍 检测 VBA 代码变更..."

# 检查是否有 .xlam 文件被修改
CHANGED_XLAM=$(git diff --cached --name-only --diff-filter=M | grep '\.xlam$')

if [ -z "$CHANGED_XLAM" ]; then
    echo "✅ 无 xlam 文件变更，跳过 VBA 导出"
    exit 0
fi

# 检查 VBA-Exporter 是否可用
if ! command -v vba-exporter &> /dev/null; then
    echo "⚠️  警告：vba-exporter 未安装，无法导出 VBA 代码"
    echo "   请运行：pip install vba-exporter"
    exit 0
fi

# 导出 VBA 代码
echo "📦 正在导出 VBA 代码..."
vba-exporter export \
    --input "$REPO_ROOT/$XLAM_PATH" \
    --output "$REPO_ROOT/$EXPORT_DIR" \
    --format bas \
    --structure

if [ $? -eq 0 ]; then
    echo "✅ VBA 代码导出成功"
    # 自动添加导出的文件到暂存区
    git add "$REPO_ROOT/$EXPORT_DIR"
    echo "📝 已添加导出文件到暂存区"
else
    echo "❌ VBA 代码导出失败"
    exit 1
fi
```

### VBA-Exporter 工具

使用 Python 实现的 VBA 代码导出工具：

```python
# vba_exporter.py
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

def export_vba_code(xlam_path: str, output_dir: str):
    """
    从 xlam 文件中导出 VBA 代码

    xlam 文件本质是 ZIP 压缩包，包含：
    - xl/vbaProject.bin: VBA 项目二进制数据
    - customUI/customUI.xml: 工具栏 XML
    """
    with zipfile.ZipFile(xlam_path, 'r') as zip_ref:
        # 提取 VBA 项目
        if 'xl/vbaProject.bin' in zip_ref.namelist():
            zip_ref.extract('xl/vbaProject.bin', output_dir)

        # 提取 CustomUI
        for name in zip_ref.namelist():
            if name.startswith('customUI/'):
                zip_ref.extract(name, output_dir)
```

### 使用流程

```
管理员修改 VBA 代码
       ↓
保存 MyAddin.xlam
       ↓
git add MyAddin.xlam
       ↓
git commit
       ↓
🔔 pre-commit 钩子触发
       ↓
自动导出 VBA 代码到 vba-src/
       ↓
导出文件自动加入暂存区
       ↓
提交完成（包含 xlam + VBA 源码）
```

### 导出文件结构示例

```
vba-src/
├── modules/
│   ├── Module1.bas
│   ├── Module2.bas
│   └── RibbonCallbacks.bas
├── classes/
│   └── clsHelper.cls
├── sheets/
│   ├── Sheet1.frm
│   └── Sheet2.frm
├── ThisWorkbook.bas
├── customUI/
│   └── customUI14.xml
└── vba-modules.json    ← 模块清单（含哈希值）
```

### 模块清单示例

```json
{
  "exported_at": "2026-03-27T10:30:00Z",
  "xlam_hash": "sha256:abc123...",
  "modules": [
    {
      "name": "Module1",
      "type": "standard",
      "file": "modules/Module1.bas",
      "hash": "sha256:def456..."
    },
    {
      "name": "clsHelper",
      "type": "class",
      "file": "classes/clsHelper.cls",
      "hash": "sha256:ghi789..."
    }
  ]
}
```

### 安装钩子

首次使用时运行：

```bash
# 复制钩子到 .git/hooks 目录
cp hooks/pre-commit .git/hooks/
cp hooks/hooks-config.json .git/hooks/

# 赋予执行权限
chmod +x .git/hooks/pre-commit
```

### 注意事项

1. **Windows 环境**：git hooks 需要使用 Unix 风格的 shebang (`#!/bin/bash`)，Git for Windows 会自动处理
2. **依赖安装**：需要安装 `vba-exporter` 工具或自行实现导出逻辑
3. **二进制文件**：vbaProject.bin 是二进制文件，git 只能存储其变更，无法 diff
4. **大型项目**：导出过程可能需要几秒钟，复杂项目建议添加进度提示
