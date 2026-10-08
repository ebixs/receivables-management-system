@echo off
setlocal

echo ============================================================
echo   Building the Receivables Management desktop app
echo ============================================================
echo.

echo [1/4] Installing requirements...
pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install requirements.
    pause
    exit /b 1
)

echo.
echo [2/4] Building the executable...
pyinstaller --onefile --noconsole --name "receivables_management" app_desktop.py
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo [3/4] Assembling the release folder...
if exist release rmdir /s /q release
mkdir release
copy dist\receivables_management.exe release\ >nul
xcopy /E /I /Y ..\dashboard release\dashboard >nul

echo.
echo [4/4] Cleaning up build files...
rmdir /s /q build >nul 2>&1
del receivables_management.spec >nul 2>&1

echo.
echo Done.
pause
