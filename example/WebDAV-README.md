# WebDAV 文件夹部署说明

## 文件清单

WebDAV 服务器的 `Addin` 文件夹中需要放置以下文件：

```
Addin/
├── MyAddin.xlam    ← 加载项主文件（必需）
└── version.json    ← 版本信息文件（必需）
```

## 文件说明

### MyAddin.xlam

Excel 加载项主文件，包含：
- VBA 代码（宏函数）
- Ribbon 工具栏定义（CustomUI XML）

**创建方法：** 参见 `example/README.md`

### version.json

版本信息文件，用于版本检测和更新提示。

**格式：**
```json
{
  "version": "1.0.0",
  "releaseDate": "2026-03-27",
  "filename": "MyAddin.xlam",
  "changelog": "更新说明"
}
```

| 字段 | 说明 |
|------|------|
| version | 版本号，使用语义化版本（major.minor.patch） |
| releaseDate | 发布日期 |
| filename | 文件名（与实际文件名一致） |
| changelog | 更新日志，显示在安装工具中 |

## 部署步骤

1. 将 `MyAddin.xlam` 上传到 WebDAV 的 `Addin` 文件夹
2. 根据实际情况修改 `version.json` 并上传
3. 确保文件可被下载

## 更新流程

```
修改 VBA 代码
    ↓
保存 MyAddin.xlam
    ↓
更新 version.json 中的版本号和 changelog
    ↓
上传覆盖 WebDAV 中的文件
    ↓
通知用户点击「检查更新」
```

## 注意事项

- version 字段必须递增，否则用户端检测不到更新
- 文件名必须与 version.json 中的 filename 一致
- changelog 建议简洁明了，描述主要变更