@echo off
setlocal
cd /d "%~dp0"

python -m PyInstaller --noconfirm --clean --onefile --windowed --name TrinaryImageApp --collect-all tkinterdnd2 main.py

if errorlevel 1 (
  echo.
  echo Build failed.
  pause
  exit /b 1
)

echo.
echo Build complete: dist\TrinaryImageApp.exe
pause
