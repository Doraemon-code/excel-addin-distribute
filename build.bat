@echo off
chcp 65001 >/dev/null
echo ========================================
echo   Excel Addin Installer - Build Script
echo ========================================
echo.

:: Check virtual environment
if not exist ".venv\Scripts\activate.bat" (
    echo [Error] Virtual environment not found. Run: python -m venv .venv
    pause
    exit /b 1
)

:: Activate virtual environment
call .venv\Scripts\activate.bat

:: Check PyInstaller
python -c "import PyInstaller" 2>/dev/null
if errorlevel 1 (
    echo [Install] Installing PyInstaller...
    pip install pyinstaller
)

:: Clean old files
if exist "dist" rd /s /q "dist"
if exist "build" rd /s /q "build"
if exist "*.spec" del /q "*.spec"

echo.
echo [Build] Creating executable...
echo.

:: Build (config embedded, no external files needed)
pyinstaller --onefile --windowed ^
    --name "ExcelAddinInstaller" ^
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
