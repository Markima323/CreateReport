@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "BOOTSTRAP_PY=python"
%BOOTSTRAP_PY% --version >nul 2>nul
if errorlevel 1 (
    set "BOOTSTRAP_PY=py -3"
    %BOOTSTRAP_PY% --version >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python was not found. Please install Python 3 first.
        pause
        exit /b 1
    )
)

set "VENV_DIR=.venv_create_report"
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    %BOOTSTRAP_PY% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

set "PY_CMD=%VENV_DIR%\Scripts\python.exe"

echo [2/3] Installing dependencies...
%PY_CMD% -m pip install --disable-pip-version-check --quiet -r "bin\requirements.txt"
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo [3/3] Running Excel template -> Word report pipeline...
%PY_CMD% "bin\process_excel_to_word.py"
if errorlevel 1 (
    echo [ERROR] Pipeline failed. Check the error messages above.
    pause
    exit /b 1
)

echo Completed. Word reports are in the "word" folder.
pause
exit /b 0
