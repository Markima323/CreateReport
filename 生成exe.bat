@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "VENV_DIR=.venv_create_report"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        py -3 -m venv "%VENV_DIR%"
        if errorlevel 1 (
            echo ERROR: Python 3 is required.
            pause
            exit /b 1
        )
    )
)

"%PYTHON_EXE%" -m pip install --disable-pip-version-check --quiet -r "bin\requirements.txt"
if errorlevel 1 (
    echo ERROR: Failed to install build dependencies.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -X utf8 "bin\build_exe.py"
if errorlevel 1 (
    echo ERROR: Failed to build EXE.
    pause
    exit /b 1
)

echo Build completed: dist\CreateReport.exe
pause
exit /b 0
