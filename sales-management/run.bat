@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [setup] venv を作成します...
    python -m venv .venv
    if errorlevel 1 (
        echo Python が見つかりません。Python 3.10 以降をインストールしてください。
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

if not exist ".venv\.installed" (
    echo [setup] パッケージをインストールします...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    echo done > .venv\.installed
)

python app.py
if errorlevel 1 (
    echo.
    echo アプリが異常終了しました。エラーを確認してください。
    pause
)
endlocal
