@echo off
rem ==== exe を作成します (dist\GrepTool\GrepTool.exe) ====
cd /d %~dp0
if not exist .venv (
    py -3 -m venv .venv || python -m venv .venv
)
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller --noconfirm --clean --windowed --name GrepTool ^
    --collect-all pypdfium2 --collect-all pypdfium2_raw ^
    --hidden-import win32timezone ^
    --exclude-module tkinter ^
    GrepTool.py
if errorlevel 1 (
    echo ビルドに失敗しました。
    pause
    exit /b 1
)
echo.
echo 完成: dist\GrepTool\GrepTool.exe
echo dist\GrepTool フォルダごと配布してください。
pause
