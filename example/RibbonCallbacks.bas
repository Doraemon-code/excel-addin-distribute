Attribute VB_Name = "RibbonCallbacks"
'============================================================
' Ribbon 回调函数模块
' 所有按钮回调函数必须包含 IRibbonControl 参数
'============================================================

Option Explicit

'------------------------------------------------------------
' 目录功能
'------------------------------------------------------------

Public Sub GenerateIndex(control As IRibbonControl)
    ' TODO: 实现生成目录功能
    MsgBox "生成目录功能待实现", vbInformation, "提示"
End Sub

Public Sub JumpSheetIndex(control As IRibbonControl)
    ' TODO: 实现回到目录功能
    MsgBox "回到目录功能待实现", vbInformation, "提示"
End Sub

'------------------------------------------------------------
' 格式化功能
'------------------------------------------------------------

Public Sub FirstObsFormat(control As IRibbonControl)
    ' TODO: 实现首行格式化功能
    MsgBox "首行格式化功能待实现", vbInformation, "提示"
End Sub

Public Sub FontFormat(control As IRibbonControl)
    ' TODO: 实现字体格式化功能
    MsgBox "字体格式化功能待实现", vbInformation, "提示"
End Sub

Public Sub RemoveAllAutoFilters(control As IRibbonControl)
    ' TODO: 实现移除筛选功能
    MsgBox "移除筛选功能待实现", vbInformation, "提示"
End Sub

Public Sub RemoveLineBreaks(control As IRibbonControl)
    ' TODO: 实现移除换行符功能
    MsgBox "移除换行符功能待实现", vbInformation, "提示"
End Sub

'------------------------------------------------------------
' 汇总功能
'------------------------------------------------------------

Public Sub Classification(control As IRibbonControl)
    ' TODO: 实现分类计数功能
    MsgBox "分类计数功能待实现", vbInformation, "提示"
End Sub

Public Sub DupSum(control As IRibbonControl)
    ' TODO: 实现去重计数功能
    MsgBox "去重计数功能待实现", vbInformation, "提示"
End Sub

'------------------------------------------------------------
' 正则功能
'------------------------------------------------------------

Public Sub ExtractTextByRegex(control As IRibbonControl)
    ' TODO: 实现正则提取功能
    MsgBox "正则提取功能待实现", vbInformation, "提示"
End Sub

'------------------------------------------------------------
' MedDRA 功能
'------------------------------------------------------------

Public Sub SortAndGroupCount(control As IRibbonControl)
    ' TODO: 实现同义词检测功能
    MsgBox "同义词检测功能待实现", vbInformation, "提示"
End Sub

'------------------------------------------------------------
' 帮助功能
'------------------------------------------------------------

Public Sub ShowHelp(control As IRibbonControl)
    ' TODO: 实现帮助功能
    MsgBox "CallMe 功能待实现", vbInformation, "提示"
End Sub