@echo off
setlocal
cd /d "%~dp0"

rem ------------------------------------------------------------------
rem  Python 3.12 を優先する。
rem  3.14 は Tkinter の Toplevel ポップアップが固まる既知の不具合あり。
rem ------------------------------------------------------------------
set "PY_LAUNCHER="
where py >nul 2>nul
if not errorlevel 1 (
    py -3.12 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PY_LAUNCHER=py -3.12"
    )
)

if not defined PY_LAUNCHER (
    rem 3.12 が見つからない場合は警告して system python にフォールバック
    echo [warn] Python 3.12 が見つかりません。
    echo        Python 3.14 では受注タブのポップアップが固まる不具合があります。
    echo        https://www.python.org/downloads/ から 3.12 のインストールを推奨します。
    set "PY_LAUNCHER=python"
)

if not exist ".venv\Scripts\python.exe" (
    echo [setup] venv を作成します（%PY_LAUNCHER%）...
    %PY_LAUNCHER% -m venv .venv
    if errorlevel 1 (
        echo Python が見つかりません。Python 3.12 をインストールしてください。
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

rem venv 内の Python バージョンを確認
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set "VENV_PY_VER=%%v"
echo [info] venv Python: %VENV_PY_VER%
echo %VENV_PY_VER% | findstr /b "3.14" >nul
if not errorlevel 1 (
    echo.
    echo ============================================================
    echo  WARNING: venv が Python 3.14 で作成されています。
    echo  受注タブの新規受注／出荷確定ダイアログが固まる場合があります。
    echo  対処: このフォルダの .venv フォルダを削除して、Python 3.12 を
    echo        インストール後に run.bat を再実行してください。
    echo ============================================================
    echo.
    pause
)

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
