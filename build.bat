@echo off
chcp 65001 >nul
echo ========================================
echo   Excel Addin Installer - Build Script
echo ========================================
echo.

:: Check virtual environment
if not exist ".venv\Scripts\pyinstaller.exe" (
    echo [Error] PyInstaller not found in .venv
    echo [Info] Please run: pip install -r requirements.txt
    pause
    exit /b 1
)

:: Clean old files
if exist "dist" rd /s /q "dist"
if exist "build" rd /s /q "build"
if exist "*.spec" del /q "*.spec"

echo.
echo [Build] Creating executable...
echo.

:: Build (config embedded, no external files needed)
.venv\Scripts\pyinstaller.exe --onefile --windowed ^
    --name "ExcelAddinInstaller" ^
    --icon "icon\app.ico" ^
    --distpath "dist" ^
    --workpath "build" ^
    main.py

if errorlevel 1 (
    echo.
    echo [Failed] Build error occurred
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Build completed!
echo   Output: dist\ExcelAddinInstaller.exe
echo ========================================
echo.

:: Open output directory
explorer "dist"

pause