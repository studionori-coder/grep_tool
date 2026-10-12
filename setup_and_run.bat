@echo off
rem ==== 初回: 仮想環境を作って必要なライブラリを入れ、ツールを起動します ====
cd /d %~dp0
if not exist .venv (
    echo 仮想環境を作成しています...
    py -3 -m venv .venv || python -m venv .venv
    if errorlevel 1 (
        echo Python 3.10 以上をインストールしてください: https://www.python.org/downloads/
        pause
        exit /b 1
    )
    .venv\Scripts\python -m pip install --upgrade pip
    .venv\Scripts\python -m pip install -r requirements.txt
)
start "" .venv\Scripts\pythonw GrepTool.py %*
