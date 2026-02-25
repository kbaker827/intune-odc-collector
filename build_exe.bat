@echo off
echo ===========================================
echo Building Intune ODC Log Collector
echo ===========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Install PyInstaller if not already installed
pip install pyinstaller --quiet

REM Build the executable with spec file
echo.
echo Building executable...
pyinstaller --clean IntuneODCCollector.spec

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ===========================================
echo Build successful!
echo ===========================================
echo.
echo Executable: dist\IntuneODCCollector.exe
echo.
pause
