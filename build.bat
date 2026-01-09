@echo off
setlocal
cd /d "%~dp0"

REM Use the venv one directory up (..\ .venv). Adjust if your venv lives elsewhere.
set "PYTHON=.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
  echo ERROR: Could not find venv python at: %PYTHON%
  echo Fix build.bat to point to your working venv python.exe
  pause
  exit /b 1
)

"%PYTHON%" -c "import sys; print('Building with:', sys.executable, sys.version)"
"%PYTHON%" -c "import pygame; print('pygame:', pygame.__version__)" || (
  echo ERROR: pygame not available in this venv.
  pause
  exit /b 1
)

taskkill /F /IM selection_r3.exe >nul 2>&1
rmdir /S /Q build >nul 2>&1
rmdir /S /Q dist  >nul 2>&1
del /Q selection_r3.spec >nul 2>&1

"%PYTHON%" -m pip install --upgrade pip
"%PYTHON%" -m pip install pyinstaller pygame ttkbootstrap

"%PYTHON%" -m PyInstaller --noconfirm --onefile --windowed ^
  --collect-all pygame ^
  --add-data "assets;assets" ^
  selection_r3.py

echo.
echo Done. Check dist\
pause
