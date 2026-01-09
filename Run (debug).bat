@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python selection_r5.py
pause