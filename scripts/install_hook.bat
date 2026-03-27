@echo off
REM scripts\install_hook.bat
REM Windows 版一键安装脚本
REM 用法：双击运行，或在仓库根目录执行 scripts\install_hook.bat

setlocal enabledelayedexpansion

REM 检查是否在 git 仓库中
git rev-parse --show-toplevel >nul 2>&1
if errorlevel 1 (
    echo [错误] 请在 git 仓库目录中运行此脚本
    pause
    exit /b 1
)

for /f "delims=" %%i in ('git rev-parse --show-toplevel') do set REPO_ROOT=%%i

set HOOK_SRC=%REPO_ROOT%\hooks\pre-commit
set HOOK_DST=%REPO_ROOT%\.git\hooks\pre-commit

REM 检查源文件
if not exist "%HOOK_SRC%" (
    echo [错误] 找不到 hooks\pre-commit
    pause
    exit /b 1
)

REM 备份已有 hook
if exist "%HOOK_DST%" (
    for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set D=%%c%%b%%a
    for /f "tokens=1-3 delims=:." %%a in ("%time%") do set T=%%a%%b%%c
    set BACKUP=%HOOK_DST%.bak_!D!_!T: =0!
    copy "%HOOK_DST%" "!BACKUP!" >nul
    echo [备份] 已备份原有 hook
)

REM 复制 hook
copy "%HOOK_SRC%" "%HOOK_DST%" >nul
echo.
echo [成功] pre-commit hook 已安装！
echo.
echo 接下来的步骤：
echo   1. 安装依赖：pip install oletools
echo   2. 将 xlam 加入追踪：git add MyAddin.xlam
echo   3. 正常提交，VBA 代码会自动导出到 src\vba\
echo.
pause
