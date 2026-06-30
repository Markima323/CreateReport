@echo off
setlocal EnableExtensions
if exist "%SystemRoot%\System32\chcp.com" "%SystemRoot%\System32\chcp.com" 65001 >nul 2>nul
cd /d "%~dp0"
if errorlevel 1 (
    echo [ERROR] Failed to enter project folder.
    pause
    exit /b 1
)

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"
set "VENV_DIR=.venv_create_report"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "REQUIREMENTS=bin\requirements.txt"

echo CreateReport dependency installer
echo Project root: %CD%
echo.

if not exist "%REQUIREMENTS%" (
    echo [ERROR] Requirements file not found: %REQUIREMENTS%
    pause
    exit /b 1
)

if exist "%PYTHON_EXE%" (
    echo [1/4] Reusing virtual environment: %VENV_DIR%
    goto venv_ready
)

call :find_python
if errorlevel 1 (
    echo [ERROR] Python 3 was not found.
    echo Please install Python 3, then run this file again.
    pause
    exit /b 1
)

echo [1/4] Creating virtual environment: %VENV_DIR%
"%BOOTSTRAP_PY%" %BOOTSTRAP_ARGS% -m venv "%VENV_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment.
    pause
    exit /b 1
)

:venv_ready
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment Python not found: %PYTHON_EXE%
    pause
    exit /b 1
)

echo [2/4] Preparing pip...
"%PYTHON_EXE%" -m ensurepip --upgrade >nul 2>nul
"%PYTHON_EXE%" -m pip --version
if errorlevel 1 (
    echo [ERROR] pip is not available in the virtual environment.
    pause
    exit /b 1
)

echo [3/4] Installing dependencies from %REQUIREMENTS%...
"%PYTHON_EXE%" -m pip install --upgrade -r "%REQUIREMENTS%"
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo [4/4] Verifying runtime imports...
"%PYTHON_EXE%" -X utf8 -c "import docx, openpyxl, pymupdf, tkinterdnd2, win32com.client; from PIL import Image; print('Runtime dependency check passed.')"
if errorlevel 1 (
    echo [ERROR] Dependency verification failed.
    pause
    exit /b 1
)

echo.
echo Dependencies are ready.
echo You can now run the main launcher bat file.
pause
exit /b 0

:find_python
set "BOOTSTRAP_PY="
set "BOOTSTRAP_ARGS="

python --version >nul 2>nul
if not errorlevel 1 (
    set "BOOTSTRAP_PY=python"
    exit /b 0
)

py -3 --version >nul 2>nul
if not errorlevel 1 (
    set "BOOTSTRAP_PY=py"
    set "BOOTSTRAP_ARGS=-3"
    exit /b 0
)

if exist "%SystemRoot%\py.exe" (
    "%SystemRoot%\py.exe" -3 --version >nul 2>nul
    if not errorlevel 1 (
        set "BOOTSTRAP_PY=%SystemRoot%\py.exe"
        set "BOOTSTRAP_ARGS=-3"
        exit /b 0
    )
)

for %%P in (
    "%LocalAppData%\Programs\Python\Python313\python.exe"
    "%LocalAppData%\Programs\Python\Python312\python.exe"
    "%LocalAppData%\Programs\Python\Python311\python.exe"
    "%LocalAppData%\Programs\Python\Python310\python.exe"
    "%LocalAppData%\Programs\Python\Python39\python.exe"
    "%ProgramFiles%\Python313\python.exe"
    "%ProgramFiles%\Python312\python.exe"
    "%ProgramFiles%\Python311\python.exe"
    "%ProgramFiles%\Python310\python.exe"
    "%ProgramFiles%\Python39\python.exe"
    "%ProgramFiles(x86)%\Python313\python.exe"
    "%ProgramFiles(x86)%\Python312\python.exe"
    "%ProgramFiles(x86)%\Python311\python.exe"
    "%ProgramFiles(x86)%\Python310\python.exe"
    "%ProgramFiles(x86)%\Python39\python.exe"
) do (
    if exist %%P (
        %%P --version >nul 2>nul
        if not errorlevel 1 (
            set "BOOTSTRAP_PY=%%~P"
            exit /b 0
        )
    )
)

exit /b 1
