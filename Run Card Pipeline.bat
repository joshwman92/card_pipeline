@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "LUCAS_ASSIGNMENT_CONFIG_PATH=%~dp0assignment_companies.json"

set "VENV_OK=0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>nul
    if !ERRORLEVEL! EQU 0 set "VENV_OK=1"
    if "!VENV_OK!" EQU "1" (
        ".venv\Scripts\python.exe" app.py
        exit /b %ERRORLEVEL%
    ) else (
        echo Existing .venv Python is not usable. Falling back to system Python.
        echo Run install_dependencies.bat to rebuild .venv.
        echo.
    )
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3 app.py
    exit /b %ERRORLEVEL%
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        python app.py
        exit /b %ERRORLEVEL%
    ) else (
        echo L.U.C.A.S could not find a working Python.
        echo.
        echo Run install_dependencies.bat after installing Python 3.11 or newer from python.org.
        echo Make sure "Add python.exe to PATH" is checked during install.
        echo.
        pause
        exit /b 1
    )
)
