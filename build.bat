@echo off
chcp 65001 >nul
echo ========================================
echo   Excel 加载项安装工具 - 打包脚本
echo ========================================
echo.

:: 检查虚拟环境
if not exist ".venv\Scripts\activate.bat" (
    echo [错误] 未找到虚拟环境，请先运行: python -m venv .venv
    pause
    exit /b 1
)

:: 激活虚拟环境
call .venv\Scripts\activate.bat

:: 检查 PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [安装] 正在安装 PyInstaller...
    pip install pyinstaller
)

:: 清理旧文件
if exist "dist" rd /s /q "dist"
if exist "build" rd /s /q "build"
if exist "*.spec" del /q "*.spec"

echo.
echo [打包] 正在构建可执行文件...
echo.

:: 打包命令（配置已内嵌，无需外部文件）
pyinstaller --onefile --windowed ^
    --name "ExcelAddinInstaller" ^
    --distpath "dist" ^
    --workpath "build" ^
    main.py

if errorlevel 1 (
    echo.
    echo [失败] 打包过程中出现错误
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包完成！
echo   输出文件: dist\ExcelAddinInstaller.exe
echo ========================================
echo.

:: 打开输出目录
explorer "dist"

pause