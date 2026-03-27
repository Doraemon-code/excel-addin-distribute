# Excel 加载项开发示例

本目录包含创建 Excel 加载项的示例文件。

## 文件说明

| 文件 | 说明 |
|------|------|
| `customUI14.xml` | Ribbon 工具栏定义（Office 2010+） |
| `RibbonCallbacks.bas` | VBA 回调函数模块示例 |

## 创建加载项步骤

### 1. 创建 xlam 文件

1. 打开 Excel，新建空白工作簿
2. 按 `Alt + F11` 打开 VBA 编辑器
3. 右键 VBAProject → 插入 → 模块
4. 将 `RibbonCallbacks.bas` 内容粘贴进去
5. 关闭 VBA 编辑器
6. 文件 → 另存为 → 选择"Excel 加载项 (*.xlam)"
7. 保存为 `MyAddin.xlam`

### 2. 添加 CustomUI

1. 用 **Office RibbonX Editor** 打开 `MyAddin.xlam`
2. 右键 → Insert → Insert Office 2010+ Custom UI Part
3. 将 `customUI14.xml` 内容粘贴到编辑区
4. 点击 **Validate** 确认无错误
5. 点击 **Save** 保存

### 3. 测试

1. 双击 `MyAddin.xlam` 或用安装工具安装
2. 重启 Excel
3. 检查是否出现"小叮当当"标签页
4. 点击按钮测试功能

## 重要提醒

**所有 Ribbon 按钮回调函数必须包含 `IRibbonControl` 参数：**

```vba
' 正确
Public Sub GenerateIndex(control As IRibbonControl)
    ' 代码
End Sub

' 错误（会报"错误的参数号或无效的属性赋值"）
Public Sub GenerateIndex()
    ' 代码
End Sub
```

## 导入 .bas 文件

在 VBA 编辑器中也可以直接导入 `.bas` 文件：
1. 文件 → 导入文件
2. 选择 `RibbonCallbacks.bas`

## 编辑 xlam 文件的正确方法

**重要**：双击打开 xlam 文件时，Excel 会以"加载项模式"加载，此时 `Ctrl + S` 不会保存代码修改。

正确步骤：
1. **完全关闭 Excel**（包括任务栏中所有实例）
2. 打开一个新的 Excel 空白工作簿
3. **文件 → 打开 → 浏览** → 选择 xlam 文件
   - 注意：使用"打开"菜单，不是双击文件
4. 按 `Alt + F11` 打开 VBA 编辑器
5. 修改代码
6. `Ctrl + S` 保存
7. 关闭 Excel

| 操作方式 | Excel 加载模式 | Ctrl+S 是否保存代码 |
|----------|---------------|-------------------|
| 双击 xlam 文件 | 加载项模式 | ❌ 不保存 |
| Excel 菜单"打开" | 工作簿模式 | ✅ 保存 |